#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import os
from pathlib import Path
from typing import Dict, List, Optional

from PIL import Image


def make_thumb(src: Path, dst: Path, max_size: int = 240, quality: int = 70) -> None:
    img = Image.open(src)
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGB")
    if img.mode == "RGBA":
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        img = bg

    img.thumbnail((max_size, max_size))
    dst.parent.mkdir(parents=True, exist_ok=True)
    # Use WebP for much smaller file sizes
    img.save(dst, "WEBP", quality=quality, method=6)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--menu-json", required=True)
    ap.add_argument("--web-root", required=True)
    ap.add_argument("--thumb-dir", default="assets/menu-thumbs")
    ap.add_argument("--max-size", type=int, default=240)
    ap.add_argument("--quality", type=int, default=70)
    args = ap.parse_args()

    web_root = Path(args.web_root).resolve()
    menu_json = Path(args.menu_json).resolve()
    thumb_dir = (web_root / args.thumb_dir).resolve()

    data: List[Dict] = json.loads(menu_json.read_text(encoding="utf-8"))
    updated = 0
    generated = 0

    for item in data:
        img_rel: Optional[str] = item.get("image")
        if not img_rel:
            continue
        img_path = (web_root / img_rel).resolve()
        if not img_path.exists():
            continue

        # Create deterministic thumb filename based on original relative path
        stem = Path(img_rel).name
        thumb_name = f"{Path(stem).stem}.webp"
        thumb_path = thumb_dir / thumb_name

        if not thumb_path.exists() or thumb_path.stat().st_mtime < img_path.stat().st_mtime:
            make_thumb(img_path, thumb_path, max_size=args.max_size, quality=args.quality)
            generated += 1

        thumb_rel = os.path.relpath(thumb_path, web_root).replace("\\", "/")
        if item.get("imageThumb") != thumb_rel:
            item["imageThumb"] = thumb_rel
            updated += 1

    menu_json.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"generated_thumbs={generated} updated_items={updated} thumb_dir={thumb_dir}")


if __name__ == "__main__":
    main()

