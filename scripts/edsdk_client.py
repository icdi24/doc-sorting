#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Minimal Python client for the Tencent local Office editor SDK (edsdk).

The editor SDK ships with the built-in skill `tencent-local-office-edit` as an
`edsdk.py` CLI in front of a local JSON-RPC service. This module locates that
`edsdk.py` automatically, loads it, and exposes:

    json_result = call(tool_name, arguments_dict)

`call()` performs one JSON-RPC `tools/call` and returns the decoded JSON payload
(the tool's first text content block), or ``{'__raw__': str}`` when the payload
is not JSON.

CLI
---
    python edsdk_client.py --pool                       # editor pool status
    python edsdk_client.py --list [prefix]              # tool inventory
    python edsdk_client.py --schema <tool>              # tool parameter schema
    python edsdk_client.py <tool> --fid <file.docx> --json '<args>'
    python edsdk_client.py <tool> --fid <file.docx> --json-file <args.json>
    python edsdk_client.py <tool> --json '<json containing file_id>'

Notes
-----
* ``edsdk._rpc('tools/call', ...)`` returns the *result object* directly
  (it already unwraps the JSON-RPC envelope). Do not look for ``obj['result']``.
  This client accepts either shape defensively.
* All ``file_id`` values are the absolute document path (forward slashes work).
* Chinese paths: prefer ``--json-file`` (UTF-8) over ``--json`` when invoking
  from a shell that may mangle non-ASCII quoting.
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import sys

SKILL_NAME = "tencent-local-office-edit"

_SEARCH_ROOTS = [
    r"D:/workBuddy",
    r"C:/Program Files/WorkBuddy",
    r"C:/Program Files (x86)/WorkBuddy",
    os.path.expanduser("~/AppData/Local/Programs/WorkBuddy"),
    os.path.expanduser("~/AppData/Local/workbuddy"),
    os.path.expanduser("~/.workbuddy/plugins"),
    "/Applications/WorkBuddy.app",
    "/opt/WorkBuddy",
]

_EXACT_PATHS = [
    os.environ.get("EDSDK_DIR", ""),
    os.path.expanduser(
        "~/AppData/Local/Programs/WorkBuddy/resources/app.asar.unpacked/resources/"
        "plugins/workbuddy-builtin/skills/" + SKILL_NAME
    ),
    "/Applications/WorkBuddy.app/Contents/Resources/app.asar.unpacked/resources/"
    "plugins/workbuddy-builtin/skills/" + SKILL_NAME,
]

_edsdk = None
_dir = None


def find_edsdk_dir() -> str:
    """Return the directory containing edsdk.py, or raise FileNotFoundError."""
    for p in _EXACT_PATHS:
        if p and os.path.isfile(os.path.join(p, "edsdk.py")):
            return p
    for root in _SEARCH_ROOTS:
        if not os.path.isdir(root):
            continue
        hits = glob.glob(
            os.path.join(root, "**", SKILL_NAME, "edsdk.py"), recursive=True
        )
        if hits:
            return os.path.dirname(sorted(hits)[0])
    raise FileNotFoundError(
        "Could not locate edsdk.py. Set EDSDK_DIR to the "
        f"'{SKILL_NAME}' skill directory."
    )


def load():
    """Import and cache the edsdk module."""
    global _edsdk, _dir
    if _edsdk is None:
        _dir = find_edsdk_dir()
        if _dir not in sys.path:
            sys.path.insert(0, _dir)
        import edsdk  # type: ignore

        _edsdk = edsdk
    return _edsdk


def call(tool: str, args: dict, *, timeout: int | None = None):
    """Invoke one editor-SDK tool and return its decoded JSON payload.

    Raises RuntimeError when the tool reports ``isError``.
    """
    mod = load()
    payload = {"name": tool, "arguments": args or {}}
    obj = mod._rpc("tools/call", payload)
    # Defensive: some builds wrap in {'result': ...}
    if isinstance(obj, dict) and set(obj.keys()) == {"result"}:
        obj = obj["result"]
    if not isinstance(obj, dict):
        return {"__raw__": repr(obj)}
    if obj.get("isError"):
        text = "\n".join(c.get("text", "") for c in (obj.get("content") or []))
        raise RuntimeError(f"{tool} failed: {text[:400]}")
    text = "\n".join(c.get("text", "") for c in (obj.get("content") or []))
    if text == "":
        return {}
    try:
        return json.loads(text)
    except (ValueError, TypeError):
        return {"__raw__": text}


def call_ok(tool: str, args: dict) -> bool:
    """Boolean variant of call() that never raises."""
    try:
        call(tool, args)
        return True
    except Exception:  # noqa: BLE001 - best effort batch helper
        return False


# --------------------------------------------------------------------------- CLI


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Tencent local Office editor-SDK client")
    ap.add_argument("tool", nargs="?", help="tool name to invoke")
    ap.add_argument("--fid", help="document path, injected as file_id")
    ap.add_argument("--json", dest="args_json", help="tool arguments as JSON")
    ap.add_argument("--json-file", dest="args_file", help="tool arguments JSON file")
    ap.add_argument("--out", help="write the JSON payload to this file")
    ap.add_argument("--pool", action="store_true", help="show editor pool status")
    ap.add_argument("--list", nargs="?", const="", help="list tools (optional prefix)")
    ap.add_argument("--schema", help="print a tool's parameter schema")
    ap.add_argument("--where", action="store_true", help="print the edsdk directory")
    a = ap.parse_args(argv)

    if a.where:
        print(find_edsdk_dir())
        return 0

    if a.pool:
        print(json.dumps(call("get_pool_status", {}), ensure_ascii=False, indent=2))
        return 0

    if a.list is not None:
        mod = load()
        tools = mod._rpc("tools/list", {})
        if isinstance(tools, dict) and "tools" in tools:
            names = [t["name"] for t in tools["tools"]]
        elif isinstance(tools, list):
            names = [t.get("name", str(t)) for t in tools]
        else:
            names = [str(tools)]
        sel = sorted(n for n in names if n.startswith(a.list))
        print("\n".join(sel))
        return 0

    if a.schema:
        mod = load()
        info = mod._rpc("tools/list", {})
        items = info.get("tools", info) if isinstance(info, dict) else info
        for t in items:
            if t.get("name") == a.schema:
                print(json.dumps(t, ensure_ascii=False, indent=2))
                return 0
        print(f"tool not found: {a.schema}", file=sys.stderr)
        return 2

    if not a.tool:
        ap.print_help()
        return 2

    if a.args_file:
        with open(a.args_file, encoding="utf-8") as fh:
            args = json.load(fh)
    elif a.args_json:
        args = json.loads(a.args_json)
    else:
        args = {}
    if a.fid:
        args.setdefault("file_id", a.fid)

    result = call(a.tool, args)
    dumped = json.dumps(result, ensure_ascii=False, indent=2)
    if a.out:
        with open(a.out, "w", encoding="utf-8") as fh:
            fh.write(dumped)
        print(f"wrote {a.out}")
    else:
        print(dumped)
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
