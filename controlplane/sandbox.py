"""Run a data-plane command in a sandbox.

POC sandbox: Linux user + network namespace (`unshare -rn`, so no network at all),
rlimits on memory/CPU/file size, a scrubbed environment, and a hard timeout.

Production: replace `run()` with "claim a Kata microVM, hand it file refs + JSON,
wait for JSON back". The contract (argv in, files + JSON report out) is the same,
which is the point - the control plane never learns which sandbox it got.

Known gap of the POC sandbox: no mount namespace, so the process can still READ
the host filesystem. Kata closes that; so would bubblewrap.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENTRY = os.path.join(ROOT, "dataplane_main.py")


class SandboxError(Exception):
    pass


def run(args: list[str], workspace: str, report_path: str, timeout=240, mem_mb=4096, cpu_s=180):
    workspace = os.path.abspath(workspace)
    home = os.path.join(workspace, ".home")
    os.makedirs(home, exist_ok=True)
    cmd = [
        "unshare", "-rn", "--",
        "prlimit", f"--as={mem_mb * 1024 * 1024}", f"--cpu={cpu_s}",
        f"--fsize={512 * 1024 * 1024}", "--nofile=512", "--",
        sys.executable, "-E", "-s", ENTRY, *args,
    ]
    env = {"HOME": home, "PATH": "/usr/local/bin:/usr/bin:/bin", "LANG": "C.UTF-8", "TMPDIR": home}
    t0 = time.time()
    try:
        proc = subprocess.run(cmd, cwd=workspace, env=env, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        raise SandboxError(f"{args[0]} timed out after {timeout}s")
    elapsed = round(time.time() - t0, 2)
    try:
        report = json.load(open(report_path))
    except Exception:
        raise SandboxError(f"{args[0]} produced no report (rc={proc.returncode}): {proc.stderr[-800:]}")
    if isinstance(report, dict):
        report.setdefault("_sandbox", {"rc": proc.returncode, "seconds": elapsed})
    return proc.returncode, report
