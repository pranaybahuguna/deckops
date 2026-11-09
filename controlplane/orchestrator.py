"""Control plane: the pipeline. Never opens a .pptx itself - every byte-touching
step goes through sandbox.run(). Holds only JSON (index, plan, ops, reports).

Stages: intake -> probe -> plan -> compile -> execute -> validate -> repair -> deliver
"""
from __future__ import annotations

import json
import os
import shutil
import time
import uuid

from deckops.compiler import compile_create, compile_translate
from deckops.probe import iter_paragraphs

from . import sandbox

MAX_REPAIR = 2


class Run:
    def __init__(self, root, planner, log=print):
        self.id = time.strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
        self.ws = os.path.abspath(os.path.join(root, self.id))
        os.makedirs(self.ws)
        self.planner, self.log, self.trace = planner, log, []

    def p(self, name):
        return os.path.join(self.ws, name)

    def save(self, name, obj):
        json.dump(obj, open(self.p(name), "w"), indent=1, ensure_ascii=False)

    def box(self, *args):
        """Run a data-plane command in the sandbox; report path is always the last arg."""
        t0 = time.time()
        rc, rep = sandbox.run(list(args), self.ws, self.p(args[-1]))
        self.trace.append({"stage": args[0], "rc": rc, "seconds": round(time.time() - t0, 2)})
        self.log(f"  [sandbox] {args[0]:<8} rc={rc} {round(time.time() - t0, 1)}s")
        if rc not in (0, 3):
            raise sandbox.SandboxError(json.dumps(rep)[:1500])
        return rc, rep


def _changelog(index, ops):
    paras = {p["ref"]: (s["index"], p["text"]) for s in index["slides"] for p in iter_paragraphs({"slides": [s]})}
    out = []
    for op in ops:
        if op["op"] == "set_paragraph_text":
            from deckops.refs import parse_tagged
            toks = parse_tagged(op["value"])
            after = "".join(t[2] if t[0] == "r" else (t[1] if t[0] == "plain" else "\n" if t[0] == "br" else "") for t in toks)
            si, before = paras[op["target"]]
            out.append({"slide": si + 1, "ref": op["target"], "before": before, "after": after})
        else:
            out.append({k: v for k, v in op.items() if k != "expect"})
    return out


def translate(root, src_path, lang, slide_numbers, planner, log=print):
    run = Run(root, planner, log)
    log(f"run {run.id}  translate -> {lang}, slides {slide_numbers}")
    shutil.copy(src_path, run.p("source.pptx"))                                  # 0 intake
    _, index = run.box("probe", "source.pptx", "index.json")                      # 1 probe
    scope = {n - 1 for n in slide_numbers}
    paras = list(iter_paragraphs(index, scope))
    log(f"  probe: {len(index['slides'])} slides, {len(paras)} paragraphs in scope, "
        f"{len(index['unsupported'])} opaque objects, {len(index['external_rels'])} external links")
    plan = planner.translate(paras, lang)                                         # 2-4 plan
    run.save("plan.json", plan)
    ops = compile_translate(plan, index)                                          # 5 compile (pure)
    log(f"  compile: {len(ops)} ops")
    by_ref = {p["ref"]: p for p in paras}
    # Differential validation: the source is the baseline. Only flag shapes the edit
    # made worse - the heuristic is approximate and many real decks already "overflow".
    scope_arg = ",".join(map(str, sorted(scope)))
    _, base = run.box("validate", "source.pptx", scope_arg, "-", "baseline.json")
    base_ratio = {f["ref"]: f["ratio"] for f in base["findings"]}
    log(f"  baseline: {len(base_ratio)} shapes already tight in the source")
    for attempt in range(MAX_REPAIR + 1):
        run.save("ops.json", ops)
        rc, ex = run.box("execute", "source.pptx", "ops.json", "out.pptx", "exec.json")   # 6 execute
        if rc == 3:
            raise RuntimeError(f"execution refused: {ex['errors']}")
        _, val = run.box("validate", "out.pptx", ",".join(map(str, sorted(scope))),
                         f"render{attempt}", "validate.json")                     # 7 validate
        over = [f for f in val["findings"] if f["kind"] == "overflow"
                and f["ratio"] > max(1.0, base_ratio.get(f["ref"], 0)) * 1.03]
        val["findings"] = over
        log(f"  validate (attempt {attempt}): {len(over)} new overflow"
            f"{' (template has autofit: PowerPoint may shrink it on edit, LibreOffice will not)' if any(f['masked_by_autofit'] for f in over) else ''}"
            f", structural={val['structural'] or 'ok'}")
        if not over or attempt == MAX_REPAIR:
            break
        # 8 repair: first ask the model for shorter phrasing, then deterministic font fallback
        current = {o["target"]: o["value"] for o in ops if o["op"] == "set_paragraph_text"}
        if attempt == 0:
            victims = [dict(by_ref[r], tagged=current.get(r, by_ref[r]["tagged"]))
                       for f in over for r in by_ref if r.startswith(f["ref"] + "/")]
            short = planner.shorten(victims, lang, 25)
            plan_map = {p["ref"]: p for p in plan["paragraphs"]}
            for p in short["paragraphs"]:
                plan_map[p["ref"]] = p
            plan["paragraphs"] = list(plan_map.values())
            ops = compile_translate(plan, index)
            log(f"  repair: model shortened {len(short['paragraphs'])} paragraphs")
        else:
            ops = [o for o in ops if not (o["op"] == "set_font_size" and o["target"] in {f["ref"] for f in over})]
            ops += [{"op": "set_font_size", "target": f["ref"], "size_pt": f["suggest_size_pt"]} for f in over]
            log(f"  repair: font size fallback on {len(over)} shapes")
    manifest = {
        "run": run.id, "mode": "translate", "planner": planner.name, "target_lang": lang,
        "ops": len(ops), "exec": ex, "remaining_findings": val["findings"],
        "structural": val["structural"], "previews": val.get("previews"),
        "changelog": _changelog(index, ops), "trace": run.trace,
    }
    run.save("manifest.json", manifest)                                           # 9 deliver
    return run, manifest


def create(root, template_path, brief, planner, log=print, family_map=None):
    run = Run(root, planner, log)
    log(f"run {run.id}  create from template")
    shutil.copy(template_path, run.p("source.pptx"))
    _, index = run.box("probe", "source.pptx", "index.json")
    plan = planner.plan_create(brief)
    run.save("plan.json", plan)
    ops = compile_create(plan, index["layouts"], family_map)
    ops += [{"op": "delete_slide", "target": s["ref"]} for s in index["slides"]]  # template sample slides
    log(f"  compile: {len(ops)} ops -> layouts {[o['layout'] for o in ops if o['op'] == 'add_slide']}")
    run.save("ops.json", ops)
    rc, ex = run.box("execute", "source.pptx", "ops.json", "out.pptx", "exec.json")
    if rc == 3:
        raise RuntimeError(f"execution refused: {ex['errors']}")
    _, val = run.box("validate", "out.pptx", "all", "render0", "validate.json")
    log(f"  validate: {len(val['findings'])} findings, structural={val['structural'] or 'ok'}")
    manifest = {"run": run.id, "mode": "create", "planner": planner.name, "ops": len(ops), "exec": ex,
                "findings": val["findings"], "structural": val["structural"],
                "previews": val.get("previews"), "changelog": _changelog(index, ops), "trace": run.trace}
    run.save("manifest.json", manifest)
    return run, manifest


def edit(root, src_path, plan, compile_fn, slide_numbers, planner_name, log=print):
    """Edit mode: plan (from the model) -> compile_fn -> sandboxed execute/validate."""
    run = Run(root, None, log)
    log(f"run {run.id}  edit slides {slide_numbers}")
    shutil.copy(src_path, run.p("source.pptx"))
    _, index = run.box("probe", "source.pptx", "index.json")
    run.save("plan.json", plan)
    ops = compile_fn(plan, index)
    log(f"  compile: {len(ops)} ops " + str({k: sum(o['op'] == k for o in ops) for k in {o['op'] for o in ops}}))
    run.save("ops.json", ops)
    scope = ",".join(str(n - 1) for n in slide_numbers)
    _, base = run.box("validate", "source.pptx", scope, "-", "baseline.json")
    base_ratio = {f["ref"]: f["ratio"] for f in base["findings"]}
    rc, ex = run.box("execute", "source.pptx", "ops.json", "out.pptx", "exec.json")
    if rc == 3:
        raise RuntimeError(f"execution refused: {ex['errors']}")
    _, val = run.box("validate", "out.pptx", scope, "render0", "validate.json")
    over = [f for f in val["findings"] if f["ratio"] > max(1.0, base_ratio.get(f["ref"], 0)) * 1.03]
    val["structural"] = [p for p in val["structural"] if p not in base["structural"]]   # only what WE broke
    log(f"  validate: {len(over)} new overflow, new structural problems={val['structural'] or 'none'}"
        f" (pre-existing in source: {base['structural'] or 'none'})")
    manifest = {"run": run.id, "mode": "edit", "planner": planner_name, "ops": len(ops), "exec": ex,
                "findings": over, "structural": val["structural"], "previews": val.get("previews"),
                "changelog": _changelog(index, ops), "trace": run.trace}
    run.save("manifest.json", manifest)
    return run, manifest
