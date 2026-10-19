#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Insert paragraphs/captions/blank-fillers through the editor SDK, safely.

Insertion is the one job the editor SDK does better than raw XML surgery: it
handles paragraph splitting, numbering and inline images for you. The catch is
that every insertion shifts every later offset, so a batch prepared against a
stale snapshot merges or lands text in the wrong place.

This tool therefore enforces three rules:

1. **Descending order.** Ops are always applied from the highest ``idx`` down,
   so an insertion never invalidates an ``idx`` that has not run yet.
2. **One op, one call.** Each op is a single atomic
   ``doc_insert_paragraph_with_text`` - no two-step "write text then restyle",
   which is exactly where the drift creeps in.
3. **Anchor-resolved, never hand-computed.** Offsets come from ``doc_find``
   (``anchors``), ``doc_get_images`` (``images`` / ``captions``) or a
   ``next_index`` returned by a previous op - never from arithmetic on a
   character count.

Subcommands
-----------
    anchors <file.docx> "程序截图：" "运行结果："      # editor offsets per anchor
    images  <file.docx>                                # image index/offset list
    captions <file.docx> --captions-file caps.json     # caption under each image
    run     plan.json                                  # execute an op plan
    plan    plan.json                                  # resolve + print, no writes

Plan format (run/plan)
----------------------
{
  "file_id": "D:/path/report.docx",
  "save": true,
  "ops": [
    {"after_image": 5429, "text": "图 1-1 CSRankings 类的代码",
     "para": {"jc": "center"}, "font": {"size_pt": 9}},
    {"idx": 14702, "text": "运行结果：本次运行共完成 12 组测试。",
     "font": {"size_pt": 12, "bold": true}},
    {"idx": 20011, "runs": [{"text": "私有成员", "format": {"bold": true}},
                            {"text": "只能被本类访问。"}]}
  ]
}

``after_image`` uses the image's editor offset from ``doc_get_images``; the
caption is inserted at ``index + 1`` so it lands immediately *below* the image.

captions file (captions subcommand)
-----------------------------------
{"file_id": "...", "save": true, "image_offset": 1, "caption_style":
  {"para": {"jc": "center"}, "font": {"size_pt": 9}},
 "captions": {"5429": "图 1-1 xxx", "5459": "图 1-2 yyy"}}
"""

from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from edsdk_client import call, call_ok  # noqa: E402


def build_args(op: dict, file_id: str) -> dict:
    args = {"file_id": op.get("file_id", file_id)}
    if "after_image" in op:
        args["idx"] = int(op["after_image"]) + int(op.get("image_offset", 1))
    else:
        args["idx"] = int(op["idx"])
    if op.get("text") is not None:
        args["text"] = op["text"]
    if op.get("runs"):
        args["runs"] = op["runs"]
    if op.get("level") is not None:
        args["level"] = op["level"]
    if op.get("para"):
        args["paragraph_property"] = op["para"]
    if op.get("font"):
        args["text_format"] = _font_to_text_format(op["font"])
    return args


def _font_to_text_format(font: dict) -> dict:
    """Accept both {size_pt} and the SDK's native {font_size}."""
    out = {}
    for k, v in font.items():
        if k == "size_pt":
            out["font_size"] = v
        else:
            out[k] = v
    return out


def run_plan(plan: dict, *, dry: bool) -> int:
    file_id = plan["file_id"]
    ops = list(plan.get("ops") or [])
    if plan.get("captions"):
        style = plan.get("caption_style", {})
        for img_idx, text in plan["captions"].items():
            ops.append({
                "after_image": int(img_idx),
                "image_offset": plan.get("image_offset", 1),
                "text": text,
                **style,
            })

    prepared = [(build_args(op, file_id), op.get("text", "")[:48]) for op in ops]
    # Rule 1: descending by insertion point
    prepared.sort(key=lambda t: t[0]["idx"], reverse=True)

    if dry:
        for args, label in prepared:
            print("idx=%-7d %s" % (args["idx"], label))
        print("(dry run - %d ops, nothing written)" % len(prepared))
        return 0

    ok = err = 0
    errors = []
    for args, label in prepared:
        try:
            res = call("doc_insert_paragraph_with_text", args)
            ok += 1
            ni = res.get("next_index") if isinstance(res, dict) else None
            if ni is not None and not isinstance(ni, int):
                errors.append((args["idx"], "non-int next_index"))
        except Exception as exc:  # noqa: BLE001
            err += 1
            errors.append((args["idx"], str(exc)[:160]))
    print("inserted ok=%d err=%d" % (ok, err))
    for idx, msg in errors[:15]:
        print("  ERR idx=%s %s" % (idx, msg))

    if plan.get("save"):
        print("save:", "ok" if call_ok("save_file", {"file_id": file_id}) else "FAILED")
    return 1 if err else 0


def cmd_anchors(a) -> int:
    for text in a.texts:
        res = call("doc_find", {"file_id": a.docx, "text": text})
        locs = res.get("locations", []) if isinstance(res, dict) else []
        print("%s -> %d match(es)" % (text, len(locs)))
        for loc in locs:
            print("   begin=%-7d end=%-7d para=%s  %s"
                  % (loc.get("begin"), loc.get("end"),
                     loc.get("paragraph_id", "?"),
                     (loc.get("related_text") or "")[:60]))
    return 0


def cmd_images(a) -> int:
    res = call("doc_get_images", {"file_id": a.docx})
    imgs = res.get("images", []) if isinstance(res, dict) else []
    print("image_offset\tindex\tcaption_idx")
    for i, im in enumerate(imgs, 1):
        print("%d\t%s\t%s" % (i, im.get("index"),
                              int(im["index"]) + a.image_offset))
    print("# %d images" % len(imgs))
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Safe docx insertion via editor SDK")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("anchors")
    p.add_argument("docx")
    p.add_argument("texts", nargs="+")
    p.set_defaults(fn=cmd_anchors)

    p = sub.add_parser("images")
    p.add_argument("docx")
    p.add_argument("--image-offset", type=int, default=1)
    p.set_defaults(fn=cmd_images)

    p = sub.add_parser("run")
    p.add_argument("plan")
    p.add_argument("--dry-run", action="store_true")
    p.set_defaults(fn=None)

    p = sub.add_parser("plan")
    p.add_argument("plan")
    p.set_defaults(fn=None)

    p = sub.add_parser("captions")
    p.add_argument("docx")
    p.add_argument("--captions-file", required=True)
    p.add_argument("--dry-run", action="store_true")
    p.set_defaults(fn=None)

    a = ap.parse_args(argv)

    if a.cmd in ("run", "plan"):
        with open(a.plan, encoding="utf-8") as fh:
            plan = json.load(fh)
        return run_plan(plan, dry=(a.cmd == "plan" or a.dry_run))

    if a.cmd == "captions":
        with open(a.captions_file, encoding="utf-8") as fh:
            plan = json.load(fh)
        plan.setdefault("file_id", a.docx)
        return run_plan(plan, dry=a.dry_run)

    return a.fn(a)


if __name__ == "__main__":
    raise SystemExit(main())
