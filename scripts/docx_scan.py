#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Report the paragraph-level structure of a .docx (text + effective format).

Works directly on the OOXML package, so it never touches the live editor and
never mutates the file. Use it to *understand* a document before editing it.

What it prints (TSV by default):
    idx  begin  end  size_pt  bold  align  indent  img  style  text

* ``begin`` / ``end`` are approximate **editor character offsets** measured from
  the start of ``w:body`` content using the editor's own counting model:
      - every ``w:t`` text node contributes ``len(text)``
      - ``w:drawing`` / ``w:pict`` / ``w:tab`` / ``w:br`` / ``w:sym`` /
        ``w:noBreakHyphen`` / ``w:object`` each contribute 1
      - every paragraph contributes +1, every table cell and every table row +1
  Add ``--base N`` to shift them into the editor's absolute coordinate space
  (see ``references/edsdk.md`` → "Offset calibration"). These offsets are only
  needed when you must call the editor SDK by index; prefer ``doc_find`` anchors.

Options
-------
    --scope all|body|table        restrict to body paragraphs, tables, or both
    --only TEXT                   only print paragraphs whose text contains TEXT
    --find TEXT                   print matching editor-style offsets only
    --images                      list image paragraphs (index, size)
    --classify                    tag each paragraph with the matching rule name
    --spec rules.json             rules used by --classify
    --json OUT                    also dump the full model as JSON
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
INLINE_ONE = ("drawing", "pict", "tab", "br", "noBreakHyphen", "sym", "object")


def w(tag: str) -> str:
    return "{%s}%s" % (W, tag)


# --------------------------------------------------------------- XML utilities


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(w("t")))


def inline_units(el) -> int:
    n = 0
    for node in el.iter():
        if node.tag == w("t"):
            n += len(node.text or "")
        elif node.tag in {w(x) for x in INLINE_ONE}:
            n += 1
    return n


def _val(el):
    return el.get(w("val")) if el is not None else None


def first_run(p):
    for r in p.iter(w("r")):
        return r
    return None


def run_size_halfpt(r):
    if r is None:
        return None
    rpr = r.find(w("rPr"))
    if rpr is None:
        return None
    sz = rpr.find(w("sz"))
    return int(_val(sz)) if _val(sz) is not None else None


def para_mark_size_halfpt(p):
    ppr = p.find(w("pPr"))
    if ppr is None:
        return None
    rpr = ppr.find(w("rPr"))
    if rpr is None:
        return None
    sz = rpr.find(w("sz"))
    return int(_val(sz)) if _val(sz) is not None else None


def describe(p) -> dict:
    ppr = p.find(w("pPr"))
    style = align = indent = None
    if ppr is not None:
        style = _val(ppr.find(w("pStyle")))
        align = _val(ppr.find(w("jc")))
        ind = ppr.find(w("ind"))
        if ind is not None:
            indent = {
                k.split("}")[-1]: v for k, v in ind.attrib.items()
            }
    r = first_run(p)
    rpr = r.find(w("rPr")) if r is not None else None
    bold = rpr is not None and rpr.find(w("b")) is not None
    size = run_size_halfpt(r) or para_mark_size_halfpt(p)
    has_img = any(
        node.tag in (w("drawing"), w("pict")) for node in p.iter()
    )
    return {
        "style": style,
        "align": align,
        "indent": indent,
        "size_halfpt": size,
        "size_pt": (size / 2) if size else None,
        "bold": bold,
        "has_image": has_img,
    }


# ------------------------------------------------------------- model building


def walk(el, base: int, out: list, scope: str):
    off = base
    for child in el:
        if child.tag == w("tcPr") or child.tag == w("tblPr"):
            continue
        if child.tag == w("p"):
            u = inline_units(child)
            if scope in ("all", "body"):
                rec = {"slot": "body", "begin": off, "end": off + u}
                rec.update(describe(child))
                rec["text"] = para_text(child).replace("\r", "")
                out.append(rec)
            off += u + 1
        elif child.tag == w("tbl"):
            for tr in child.findall(w("tr")):
                for tc in tr.findall(w("tc")):
                    off = walk(tc, off, out, scope)
                    off += 1
                off += 1
    return off


def build_model(path: str, scope: str = "all", base: int = 0):
    with zipfile.ZipFile(path) as z:
        root = _lxml_fromstring(z.read("word/document.xml"))
    body = root.find(w("body"))
    recs: list = []
    walk(body, base, recs, scope)
    return recs


def _lxml_fromstring(data):
    try:
        from lxml import etree  # type: ignore

        return etree.fromstring(data)
    except ImportError:  # pragma: no cover
        import xml.etree.ElementTree as ET

        return ET.fromstring(data)


# ------------------------------------------------------------------ classify


def load_spec(path):
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def match_rule(rec: dict, match: dict) -> bool:
    text = rec["text"]
    stripped = text.strip()
    if "text_regex" in match and not re.search(match["text_regex"], stripped):
        return False
    if "text_in" in match and stripped not in match["text_in"]:
        return False
    if "text_equals" in match and stripped != match["text_equals"]:
        return False
    if "text_contains" in match and match["text_contains"] not in text:
        return False
    if "style_id" in match and rec["style"] != match["style_id"]:
        return False
    if "size_pt" in match and rec["size_pt"] != match["size_pt"]:
        return False
    if "has_image" in match and rec["has_image"] != match["has_image"]:
        return False
    if "is_blank" in match and (not stripped) != match["is_blank"]:
        return False
    return True


def classify(recs, spec):
    for r in recs:
        r["rule"] = ""
        for rule in spec.get("rules", []):
            if match_rule(r, rule.get("match", {})):
                r["rule"] = rule.get("name", "?")
                break
    return recs


# --------------------------------------------------------------------- CLI


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Scan a .docx paragraph model")
    ap.add_argument("docx")
    ap.add_argument("--scope", default="all", choices=["all", "body", "table"])
    ap.add_argument("--base", type=int, default=0)
    ap.add_argument("--only")
    ap.add_argument("--find")
    ap.add_argument("--images", action="store_true")
    ap.add_argument("--classify", action="store_true")
    ap.add_argument("--spec")
    ap.add_argument("--json")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args(argv)

    recs = build_model(a.docx, a.scope, a.base)

    if a.classify:
        if not a.spec:
            print("--classify needs --spec", file=sys.stderr)
            return 2
        classify(recs, load_spec(a.spec))

    if a.images:
        recs = [r for r in recs if r["has_image"]]
    if a.only:
        recs = [r for r in recs if a.only in r["text"]]
    if a.find:
        for r in recs:
            i = r["text"].find(a.find)
            if i >= 0:
                print(r["begin"] + i)
        return 0

    if a.json:
        with open(a.json, "w", encoding="utf-8") as fh:
            json.dump(recs, fh, ensure_ascii=False, indent=1)

    if not a.quiet:
        hdr = "idx\tbegin\tend\tsize_pt\tbold\talign\timg\trule\tindent\ttext"
        print(hdr)
        for i, r in enumerate(recs, 1):
            ind = ""
            if r["indent"]:
                ind = ",".join("%s=%s" % (k, v) for k, v in r["indent"].items())
            print(
                "%d\t%d\t%d\t%s\t%s\t%s\t%s\t%s\t%s\t%s"
                % (
                    i,
                    r["begin"],
                    r["end"],
                    r["size_pt"] if r["size_pt"] is not None else "",
                    "Y" if r["bold"] else "",
                    r["align"] or "",
                    "IMG" if r["has_image"] else "",
                    r["rule"],
                    ind,
                    r["text"][:110].replace("\t", " ").replace("\n", " "),
                )
            )

    # summary
    from collections import Counter

    c = Counter(r.get("rule") or "-" for r in recs if r["text"].strip())
    print("---", file=sys.stderr)
    print("paragraphs: %d" % len(recs), file=sys.stderr)
    print(
        "non-empty by rule: "
        + ", ".join("%s=%d" % kv for kv in sorted(c.items())),
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
