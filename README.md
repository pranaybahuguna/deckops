# deckops: LLM-driven PowerPoint generation and editing (POC)

The model never writes a .pptx. It writes a **plan** (what the deck should say). A pure
**compiler** turns the plan into typed **operations**, and a sandboxed **executor** applies them
to the deck. Untouched slides come out byte-identical, every change is an auditable op, and
the model never has to produce coordinates or XML.

```
python3 run_demo.py                          # end-to-end run with an offline stub planner (~6 s)
python3 run_demo.py --llm-config llm.json    # use a real OpenAI-compatible model
python3 -m pytest -q tests                   # pure-function tests: no model, no sandbox
```

`llm.json` (git-ignored): `{"base_url": "https://…", "model": "…", "api_key": "…"}`.
No environment variables are read anywhere.

**Requirements:** Linux, Python 3.11+, `python-pptx`, `lxml`, `pydantic`, `Pillow`, `pytest`,
LibreOffice (`soffice`), `pdftoppm` (poppler-utils), `unshare` + `prlimit` (util-linux).

## What the demo proves

| Check | Result |
|---|---|
| Network reachable inside sandbox | blocked |
| Translate slides 2–5 to French (mixed-format runs, line break, table, group shape, notes) | valid deck |
| Untouched slides, master, images | byte-identical to input |
| Translation overflows a text box | detected; repaired by model shortening, else font-size fallback |
| Op whose target changed since planning (stale fingerprint) | whole batch refused, nothing written |
| Zip bomb (600 MB unpacked, 600 KB on disk) | rejected before any parser runs |
| Create a 4-slide deck from a template by role (title/body), no coordinates | valid deck |
| Edit: 8 process cards → 6 cards in a 3×2 grid, longer text | only that slide's XML changes |

The model step is a deterministic stub unless `--llm-config` is given.

## Architecture

```
            request + deck                        plan JSON
  User ───────────────────▶ CONTROL PLANE ◀──────────────────▶ LLM
                            orchestrator / planner / compiler
                            (never opens a .pptx; JSON only)
                                    │  sandbox.run(cmd, file names)  ▲ JSON report
                                    ▼                                │
                            DATA PLANE (sandbox: no network, rlimits)
                            dataplane_main → zipguard → probe / executor / validate(+LibreOffice)
                                    │
                            WORKSPACE (local folder per run; S3 prefix in production)
                            source.pptx  index.json  plan.json  ops.json  out.pptx  reports  previews
```

Pipeline: **intake → probe → plan → compile → execute → validate → repair (≤2) → deliver**.
The model appears only in *plan* and *repair*. Compile is a pure function, and execute always replays from
the original source.

## Code map

| Module | Runs in | Responsibility |
|---|---|---|
| `deckops/refs.py` | both | NodeRef addressing (`slide:256/shape:4/para:0/run:1`), fingerprints, tagged-run text |
| `deckops/probe.py` | sandbox | .pptx → DeckIndex JSON (slides, shapes, text, tables, notes, layouts, opaque objects, external links) |
| `deckops/ops.py` | both | The op language: pydantic models for each operation |
| `deckops/compiler.py` | control plane | Plan → ops, pure: `compile_translate`, `compile_create`, `compile_card_grid` |
| `deckops/executor.py` | sandbox | Bind all targets + check fingerprints, apply in phases, save keeping original bytes |
| `deckops/validate.py` | sandbox | Structure checks, font-metric overflow prediction, LibreOffice previews |
| `deckops/zipguard.py` | sandbox | Archive safety checks before parsing |
| `dataplane_main.py` | sandbox | The only entry point that reads deck bytes |
| `controlplane/sandbox.py` | control plane | Launches the data plane isolated. **Swap point for a microVM runtime** |
| `controlplane/planner.py` | control plane | Model step: offline stub + OpenAI-compatible client |
| `controlplane/orchestrator.py` | control plane | `translate()`, `create()`, `edit()` pipelines |
| `examples/edit_card_grid.py` | — | Edit-mode example (8 cards → 6) |

### Op language (current)

`set_paragraph_text` · `set_run_text` · `set_font_size` · `set_geometry` · `delete_shape` · `add_slide` · `delete_slide`

Adding a capability = an op model in `ops.py` + bind/apply in `executor.py` + a compile function.

## Design decisions found by building it

1. **Byte preservation needs help.** python-pptx re-serializes every part it parses. `save_preserving`
   compares parts by meaning (canonical XML; relationships with resolved targets; effective content
   types) and writes the original bytes for anything unchanged.
2. **Translate by paragraph, not by run.** Word order changes across languages. Paragraphs go to the
   model as tagged runs (`<r0>…</r0><r1>…</r1>`), and runs may be reordered while each keeps its formatting.
3. **Bind, then apply.** All targets are resolved against the source before any mutation, so ops in one
   batch can't shift each other's positional indices.
4. **Validate differentially.** Overflow prediction is approximate, and real decks have pre-existing problems.
   Only issues the edit made worse are reported.
5. **Autofit differs by renderer.** PowerPoint shrinks text on edit; LibreOffice doesn't. Findings carry
   `masked_by_autofit`.
6. **The model never emits geometry.** Layout (e.g. re-flowing cards into a grid) is computed by the compiler
   from the slide's own margins and gutters.

## Sandbox: POC vs production

The POC isolates the data plane with `unshare -rn` (user + network namespace: no network) and
`prlimit` (memory/CPU/file-size caps), a scrubbed environment and a timeout. It does **not** isolate the
filesystem or the kernel. For production, replace `controlplane/sandbox.py:run()` with a microVM runtime
(e.g. Kata Containers / Firecracker). Replace the local workspace with object storage that the VM downloads
from and uploads to. The command-line + JSON contract stays the same.

## Orchestration

The orchestrator is plain Python functions on purpose. The pipeline is a fixed chain with one bounded
loop and the model at two points. In production, wrap the stages as nodes of whatever workflow
framework the host application already uses (e.g. pydantic-graph or LangGraph), with persisted state so a
run survives a restarted worker. LangChain isn't needed: the model calls are structured-JSON requests
validated by pydantic.
