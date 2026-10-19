#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Rule-driven, coordinate-free formatting of a .docx via direct OOXML surgery.

Why not the editor SDK for bulk formatting?
-------------------------------------------
The editor SDK addresses paragraphs by *character offsets*. Every write shifts
those offsets, so a batch prepared against one snapshot silently corrupts
neighbouring content when applied (indents landing on captions, table cells and
blank paragraphs; sizes/bold landing on question text). Offsets are also only
reliable while the editor's own model agrees with the on-disk XML.

This tool never uses offsets. It selects paragraphs by *content pattern*
(regex / exact text / style id / current font size / has-image / blank) and
rewrites their formatting in the XML. That makes it:

* **idempotent** - running it twice yields the same document
* **safe**     - a rule can only touch paragraphs it explicitly matches
* **auditable**- it prints a per-rule count, and ``--check`` shows a diff plan

Serialization uses **lxml**, which preserves the original namespace prefixes.
``xml.etree.ElementTree`` drops prefixes that are only referenced via
``mc:Ignorable``, after which Word reports "the file may be corrupted".

Rule spec (JSON)
----------------
{
  "defaults": {"base_pt": 10.5},
  "rules": [
    {"name": "figure-caption",
     "match":  {"text_regex": "^图\\\\s*\\\\d+-\\\\d+"},
     "format": {"size_pt": 9, "bold": false, "align": "center",
                "clear_indent": true}},
    {"name": "table-caption",
     "match":  {"text_regex": "^表\\\\s*\\\\d+-\\\\d+"},
     "format": {"size_pt": 9, "bold": true, "align": "center",
                "clear_indent": true}},
    {"name": "label",
     "match":  {"text_in": ["程序截图：", "运行结果：", "文字说明："]},
     "format": {"size_pt": 12, "bold": true, "clear_indent": true}},
    {"name": "body",
     "match":  {"style_id": "00000e", "size_pt": 10.5},
     "format": {"first_line_indent_chars": 2}}
  ],
  "demote": {"first_line_indent_chars": 2}   // optional: applied to paragraphs
                                             // not matched by any rule
}

``match`` keys (all optional, ANDed): text_regex, text_in, text_equals,
text_contains, style_id, size_pt, has_image, is_blank.

``format`` keys:
  size_pt                 set w:sz / w:szCs (half-points = pt*2) on the
                          paragraph mark AND every run
  bold                    true / false - set or remove w:b + w:bCs
  align                   "center" | "left" | "right" | "both" | null (drop w:jc)
  clear_indent            true - remove the whole w:ind element
  first_line_indent_chars N - set w:ind firstLine=20*N*base_pt and
                          firstLineChars=N*100  (Word's "首行缩进 N 字符")
  left_indent_chars       N - analogous, w:ind leftChars=N*100

Usage
-----
    python docx_format.py in.docx --spec rules.json --dry-run
    python docx_format.py in.docx --spec rules.json --out clean.docx
    python docx_format.py in.docx --spec rules.json --inplace
    python docx_format.py in.docx --spec rules.json --inplace --clear-stray-indent
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
INLINE_ONE = ("drawing", "pict", "tab", "br", "noBreakHyphen", "sym", "object")


def w(tag: str):
    return "{%s}%s" % (W, tag)


# ------------------------------------------------------------------ read side


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(w("t")))


def run_size_halfpt(r):
    rpr = r.find(w("rPr"))
    if rpr is None:
        return None
    sz = rpr.find(w("sz"))
    return int(sz.get(w("val"))) if sz is not None and sz.get(w("val")) else None


def para_mark_size_halfpt(p):
    ppr = p.find(w("pPr"))
    if ppr is None:
        return None
    rpr = ppr.find(w("rPr"))
    if rpr is None:
        return None
    sz = rpr.find(w("sz"))
    return int(sz.get(w("val"))) if sz is not None and sz.get(w("val")) else None


def rec_of(p) -> dict:
    ppr = p.find(w("pPr"))
    style = None
    if ppr is not None:
        style = ppr.find(w("pStyle"))
        style = style.get(w("val")) if style is not None else None
    size = None
    for r in p.iter(w("r")):
        size = run_size_halfpt(r)
        break
    if size is None:
        size = para_mark_size_halfpt(p)
    text = para_text(p).replace("\r", "")
    return {
        "text": text,
        "style": style,
        "size_pt": (size / 2) if size else None,
        "size_halfpt": size,
        "has_image": any(n.tag in (w("drawing"), w("pict")) for n in p.iter()),
        "is_blank": not text.strip(),
    }


def match_rule(rec: dict, match: dict) -> bool:
    s = rec["text"].strip()
    if "text_regex" in match and not re.search(match["text_regex"], s):
        return False
    if "text_in" in match and s not in match["text_in"]:
        return False
    if "text_equals" in match and s != match["text_equals"]:
        return False
    if "text_contains" in match and match["text_contains"] not in rec["text"]:
        return False
    if "style_id" in match and rec["style"] != match["style_id"]:
        return False
    if "size_pt" in match and rec["size_pt"] != match["size_pt"]:
        return False
    if "has_image" in match and rec["has_image"] != match["has_image"]:
        return False
    if "is_blank" in match and rec["is_blank"] != match["is_blank"]:
        return False
    return True


# ----------------------------------------------------------------- write side


def get_ppr(p):
    ppr = p.find(w("pPr"))
    if ppr is None:
        ppr = etree.Element(w("pPr"))
        p.insert(0, ppr)
    return ppr


def para_rpr(ppr):
    """rPr belonging to the paragraph mark (controls the paragraph's default)."""
    rpr = ppr.find(w("rPr"))
    if rpr is None:
        rpr = etree.SubElement(ppr, w("rPr"))
    return rpr


def run_rpr(r):
    rpr = r.find(w("rPr"))
    if rpr is None:
        rpr = etree.Element(w("rPr"))
        r.insert(0, rpr)
    return rpr


def set_size(rpr, half_pt: int):
    for tag in ("sz", "szCs"):
        e = rpr.find(w(tag))
        if e is None:
            e = etree.SubElement(rpr, w(tag))
        e.set(w("val"), str(int(half_pt)))


def set_bold(rpr, on: bool):
    for tag in ("b", "bCs"):
        e = rpr.find(w(tag))
        if on:
            if e is None:
                e = etree.Element(w(tag))
                rpr.insert(0, e)
            if w("val") in e.attrib:  # <w:b w:val="0"/> means OFF
                del e.attrib[w("val")]
        else:
            if e is not None:
                rpr.remove(e)


def set_align(ppr, val):
    jc = ppr.find(w("jc"))
    if val is None:
        if jc is not None:
            ppr.remove(jc)
        return
    if jc is None:
        # w:jc must precede w:rPr / w:sectPr
        jc = etree.Element(w("jc"))
        anchor = None
        for tag in ("rPr", "sectPr"):
            e = ppr.find(w(tag))
            if e is not None:
                anchor = e
                break
        if anchor is not None:
            ppr.insert(ppr.index(anchor), jc)
        else:
            ppr.append(jc)
    jc.set(w("val"), val)


def get_ind(ppr, create=True):
    ind = ppr.find(w("ind"))
    if ind is None and create:
        ind = etree.Element(w("ind"))
        anchor = None
        for tag in ("jc", "rPr", "sectPr"):
            e = ppr.find(w(tag))
            if e is not None:
                anchor = e
                break
        if anchor is not None:
            ppr.insert(ppr.index(anchor), ind)
        else:
            ppr.append(ind)
    return ind


def clear_indent(ppr) -> int:
    ind = ppr.find(w("ind"))
    if ind is not None:
        ppr.remove(ind)
        return 1
    return 0


def set_indent_chars(ppr, first_line=None, left=None, base_pt=10.5):
    ind = get_ind(ppr)
    if first_line is not None:
        ind.set(w("firstLine"), str(int(round(20 * first_line * base_pt))))
        ind.set(w("firstLineChars"), str(int(first_line * 100)))
    if left is not None:
        ind.set(w("left"), str(int(round(20 * left * base_pt))))
        ind.set(w("leftChars"), str(int(left * 100)))


def apply_format(p, fmt: dict, base_pt: float) -> None:
    ppr = get_ppr(p)
    if fmt.get("clear_indent"):
        clear_indent(ppr)
    if "size_pt" in fmt:
        hp = int(round(float(fmt["size_pt"]) * 2))
        set_size(para_rpr(ppr), hp)
        for r in p.iter(w("r")):
            set_size(run_rpr(r), hp)
    if "bold" in fmt:
        set_bold(para_rpr(ppr), bool(fmt["bold"]))
        for r in p.iter(w("r")):
            set_bold(run_rpr(r), bool(fmt["bold"]))
    if "align" in fmt:
        set_align(ppr, fmt["align"])
    if "first_line_indent_chars" in fmt:
        set_indent_chars(ppr, first_line=float(fmt["first_line_indent_chars"]),
                         base_pt=base_pt)
    if "left_indent_chars" in fmt:
        set_indent_chars(ppr, left=float(fmt["left_indent_chars"]), base_pt=base_pt)


def is_stray_indent(ppr) -> bool:
    """True for w:ind written by the SDK's first_line_indent_chars path.

    The SDK emits ``firstLineChars`` WITHOUT ``firstLine``; the canonical
    Word form has both. Such an element is the fingerprint of a mis-applied
    batch and is safe to remove.
    """
    ind = ppr.find(w("ind"))
    if ind is None:
        return False
    return bool(ind.get(w("firstLineChars"))) and not ind.get(w("firstLine"))


# ------------------------------------------------------------------------ run


def process(root, spec: dict, clear_stray: bool, dry: bool):
    rules = spec.get("rules", [])
    demote = spec.get("demote")
    base_pt = float(spec.get("defaults", {}).get("base_pt", 10.5))
    stats: dict = {"rules": {}, "stray_removed": 0, "demoted": 0, "cr_removed": 0}
    plan = []
    for p in root.iter(w("p")):
        rec = rec_of(p)
        hit = None
        for rule in rules:
            if match_rule(rec, rule.get("match", {})):
                hit = rule
                break
        if hit:
            name = hit.get("name", "rule")
            stats["rules"][name] = stats["rules"].get(name, 0) + 1
            if dry:
                plan.append((name, rec["text"][:70]))
            else:
                apply_format(p, hit.get("format", {}), base_pt)
        else:
            ppr = p.find(w("pPr"))
            if demote and rec["text"].strip():
                stats["demoted"] += 1
                if not dry:
                    apply_format(p, demote, base_pt)
            elif clear_stray and ppr is not None and is_stray_indent(ppr):
                stats["stray_removed"] += 1
                if not dry:
                    clear_indent(ppr)
    # strip the literal CR the SDK can inject when a range covers the para mark
    for t in root.iter(w("t")):
        if t.text and "\r" in t.text:
            stats["cr_removed"] += t.text.count("\r")
            if not dry:
                t.text = t.text.replace("\r", "")
    return stats, plan


def rewrite(src: str, dst: str, transform) -> None:
    """Copy the package, replacing word/document.xml with transform(root).

    When ``dst == src`` (in-place mode) the output is staged in a sibling temp
    file and moved into place afterwards - Windows refuses to open the same
    path for writing while it is still open for reading.
    """
    same = os.path.abspath(src) == os.path.abspath(dst)
    stage = dst + ".stage~" if same else dst

    with zipfile.ZipFile(src) as zin:
        root = etree.fromstring(zin.read("word/document.xml"))
        root = transform(root)
        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                              standalone=True)
        with zipfile.ZipFile(stage, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, data)
                else:
                    zout.writestr(item, zin.read(item.filename))

    if same:
        os.replace(stage, dst)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Rule-driven docx formatting")
    ap.add_argument("docx")
    ap.add_argument("--spec", required=True)
    ap.add_argument("--out")
    ap.add_argument("--inplace", action="store_true",
                    help="overwrite the input (creates <name>.bak first)")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--clear-stray-indent", action="store_true",
                    help="remove SDK-style firstLineChars-only indents from "
                         "paragraphs matched by no rule")
    a = ap.parse_args(argv)

    with open(a.spec, encoding="utf-8") as fh:
        spec = json.load(fh)

    if a.inplace:
        dst = a.docx
    else:
        dst = a.out or os.path.splitext(a.docx)[0] + ".formatted.docx"

    holder = {}

    def transform(root):
        stats, plan = process(root, spec, a.clear_stray_indent, a.dry_run)
        holder["stats"] = stats
        holder["plan"] = plan
        return root

    if a.dry_run:
        with zipfile.ZipFile(a.docx) as z:
            root = etree.fromstring(z.read("word/document.xml"))
        stats, plan = process(root, spec, a.clear_stray_indent, True)
        print(json.dumps(stats, ensure_ascii=False, indent=2))
        for name, txt in plan[:80]:
            print("  %-16s %s" % (name, txt))
        if len(plan) > 80:
            print("  ... %d more" % (len(plan) - 80))
        return 0

    if a.inplace:
        bak = a.docx + ".bak"
        if not os.path.exists(bak):
            shutil.copy2(a.docx, bak)
            print("backup -> %s" % bak)

    rewrite(a.docx, dst, transform)
    print(json.dumps(holder["stats"], ensure_ascii=False, indent=2))
    print("written -> %s" % dst)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
