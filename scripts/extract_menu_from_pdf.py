#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import io
import json
import os
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

from pypdf import PdfReader
from pypdf.generic import ContentStream, NameObject
from PIL import Image


def slugify(name: str) -> str:
    s = re.sub(r"\s+", "_", name.strip())
    s = re.sub(r"[^0-9A-Za-z_\u4e00-\u9fff-]+", "", s)
    return s[:80] or "item"


def parse_min_order(min_order: Optional[str]) -> Optional[float]:
    if not min_order:
        return None
    m = re.search(r"(\d+(?:\.\d+)?)", str(min_order))
    return float(m.group(1)) if m else None


@dataclass
class Item:
    category: str
    name: str
    unit_price: Optional[float]
    min_order: Optional[str]
    image: Optional[str] = None


def _mul_affine(m2: List[float], m1: List[float]) -> List[float]:
    # PDF affine matrix: [a b c d e f]
    a2, b2, c2, d2, e2, f2 = m2
    a1, b1, c1, d1, e1, f1 = m1
    return [
        a2 * a1 + c2 * b1,
        b2 * a1 + d2 * b1,
        a2 * c1 + c2 * d1,
        b2 * c1 + d2 * d1,
        a2 * e1 + c2 * f1 + e2,
        b2 * e1 + d2 * f1 + f2,
    ]


def extract_images_from_page_sorted_by_rows(
    page, reader: PdfReader, expected_row_sizes: List[int]
) -> List[Tuple[str, bytes]]:
    """
    Return list of (ext, bytes) sorted by position (top-to-bottom, left-to-right).
    This is much more stable than iterating XObjects, and matches the menu grid.
    """
    resources = page.get("/Resources") or {}
    xobj = resources.get("/XObject") or {}
    try:
        xobj = xobj.get_object()
    except Exception:
        pass

    cs = ContentStream(page.get_contents(), reader)

    def ident() -> List[float]:
        return [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]

    stack = [ident()]
    placed: List[Tuple[float, float, float, float, Any]] = []  # (y, x, w, h, image_obj)

    for operands, op in cs.operations:
        if op == b"q":
            stack.append(stack[-1].copy())
        elif op == b"Q":
            if len(stack) > 1:
                stack.pop()
        elif op == b"cm":
            m = [float(x) for x in operands]
            stack[-1] = _mul_affine(m, stack[-1])
        elif op == b"Do":
            name = operands[0]
            obj = xobj.get(name)
            if not obj:
                continue
            o = obj.get_object()
            if o.get("/Subtype") != NameObject("/Image"):
                continue
            width = int(o.get("/Width") or 0)
            height = int(o.get("/Height") or 0)
            if width < 120 or height < 120:
                continue
            a, b, c, d, e, f = stack[-1]
            w = (a * a + b * b) ** 0.5
            h = (c * c + d * d) ** 0.5
            # Use translation part as approximate positioning.
            placed.append((float(f), float(e), float(w), float(h), o))

    if not placed:
        return []

    # Cluster images into visual rows by Y (tolerance relative to average height).
    avg_h = sum(p[3] for p in placed) / max(1, len(placed))
    tol = max(12.0, avg_h * 0.35)
    by_y = sorted(placed, key=lambda t: t[0])

    clusters: List[List[Tuple[float, float, float, float, Any]]] = []
    for p in by_y:
        if not clusters:
            clusters.append([p])
            continue
        if abs(p[0] - clusters[-1][-1][0]) <= tol:
            clusters[-1].append(p)
        else:
            clusters.append([p])

    # Determine which Y direction is "top".
    # We pick the direction that best matches the expected row sizes (in order).
    def score(order: List[List[Tuple[float, float, float, float, Any]]]) -> int:
        sizes = [len(r) for r in order]
        s = 0
        for i, exp in enumerate(expected_row_sizes):
            if i >= len(sizes):
                break
            # reward exact match; otherwise reward closeness
            s += 1000 - min(1000, abs(sizes[i] - exp) * 120)
        return s

    rows_desc = sorted(clusters, key=lambda r: -sum(p[0] for p in r) / len(r))
    rows_asc = sorted(clusters, key=lambda r: sum(p[0] for p in r) / len(r))
    rows = rows_desc if score(rows_desc) >= score(rows_asc) else rows_asc

    # Flatten rows: left-to-right within a row, then next row.
    placed_sorted: List[Tuple[float, float, float, float, Any]] = []
    for row in rows:
        row_sorted = sorted(row, key=lambda t: t[1])
        placed_sorted.extend(row_sorted)

    out: List[Tuple[str, bytes]] = []
    for _, __, ___, ____, o in placed_sorted:
        data = o.get_data()
        filt = o.get("/Filter")
        if isinstance(filt, list):
            filt = filt[0]

        try:
            if filt in ("/DCTDecode", "/JPXDecode", "/FlateDecode", None):
                img = Image.open(io.BytesIO(data))
                buf = io.BytesIO()
                img.convert("RGB").save(buf, format="PNG", optimize=True)
                out.append(("png", buf.getvalue()))
            else:
                img = Image.open(io.BytesIO(data))
                buf = io.BytesIO()
                img.convert("RGB").save(buf, format="PNG", optimize=True)
                out.append(("png", buf.getvalue()))
        except Exception:
            out.append(("bin", data))

    return out


def parse_menu_text(page_text: str) -> Tuple[str, List[Dict[str, Any]]]:
    """
    Parse one page text into (category, sections[]).
    Each section has names[], prices[], mins[].
    """
    text = (page_text or "").strip()
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    category = lines[0] if lines else "未分类"

    sections: List[Dict[str, Any]] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.startswith("产品"):
            names = line.replace("产品", "", 1).strip().split()
            prices: List[str] = []
            mins: List[str] = []

            j = i + 1
            while j < len(lines) and not lines[j].startswith("单价"):
                j += 1
            if j < len(lines):
                prices = [p for p in lines[j].split() if re.fullmatch(r"\d+(?:\.\d+)?", p)]

            k = j + 1
            while k < len(lines) and not lines[k].startswith("起订量"):
                k += 1
            if k < len(lines):
                mins = lines[k].replace("起订量", "", 1).strip().split()

            sections.append({"names": names, "prices": prices, "mins": mins})
            i = k + 1
        else:
            i += 1

    return category, sections


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out-json", required=True)
    ap.add_argument("--assets-dir", required=True)
    args = ap.parse_args()

    os.makedirs(args.assets_dir, exist_ok=True)
    reader = PdfReader(args.pdf)

    # We store image paths relative to the web root (dessert-quote-web/),
    # so that opening index.html via file:// works correctly.
    web_root = os.path.abspath(os.path.join(os.path.dirname(args.out_json), ".."))
    assets_dir_abs = os.path.abspath(args.assets_dir)
    assets_rel_from_root = os.path.relpath(assets_dir_abs, web_root).replace("\\", "/")

    items: List[Item] = []

    for page_idx, page in enumerate(reader.pages, start=1):
        category, sections = parse_menu_text(page.extract_text() or "")
        expected = [len(s.get("names", [])) for s in sections]
        # Extract images (sorted row-by-row) and assign sequentially to names on this page
        page_images = extract_images_from_page_sorted_by_rows(page, reader, expected)
        img_cursor = 0

        for sec in sections:
            names: List[str] = sec["names"]
            prices: List[str] = sec["prices"]
            mins: List[str] = sec["mins"]
            n = max(len(names), len(prices), len(mins))
            for idx in range(n):
                if idx >= len(names):
                    continue
                name = names[idx]
                if not name:
                    continue
                unit_price = float(prices[idx]) if idx < len(prices) else None
                min_order = mins[idx] if idx < len(mins) else None

                image_path = None
                if img_cursor < len(page_images):
                    ext, data = page_images[img_cursor]
                    img_cursor += 1
                    filename = f"p{page_idx:02d}_{slugify(name)}.{ext if ext != 'bin' else 'png'}"
                    full = os.path.join(args.assets_dir, filename)
                    # write decoded png or raw bytes
                    with open(full, "wb") as f:
                        f.write(data)
                    image_path = f"{assets_rel_from_root}/{filename}"

                items.append(
                    Item(
                        category=category,
                        name=name,
                        unit_price=unit_price,
                        min_order=min_order,
                        image=image_path,
                    )
                )

    # de-dupe by (category,name)
    seen = set()
    dedup: List[Dict[str, Any]] = []
    for it in items:
        key = (it.category, it.name)
        if key in seen:
            continue
        seen.add(key)
        dedup.append(
            {
                "category": it.category,
                "name": it.name,
                "unitPrice": it.unit_price,
                "minOrder": it.min_order,
                "minOrderNum": parse_min_order(it.min_order),
                "image": it.image,
            }
        )

    with open(args.out_json, "w", encoding="utf-8") as f:
        json.dump(dedup, f, ensure_ascii=False, indent=2)

    print(f"items: {len(dedup)}")


if __name__ == "__main__":
    main()
