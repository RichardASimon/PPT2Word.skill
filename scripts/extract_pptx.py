#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extract all media files and build slide-to-image mapping from a PPTX file.

Usage:
    python extract_pptx.py <pptx_file> <output_dir>

Outputs:
    <output_dir>/slide_images.json  — {slide_num: [image_filenames]}
    <output_dir>/ppt/media/*       — all extracted images (EMF -> PNG converted)
"""

import json
import os
import sys
import zipfile
import re
from xml.etree import ElementTree as ET

NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def extract_media(pptx_path, output_dir):
    """Extract all media from PPTX, convert EMF to PNG."""
    media_dir = os.path.join(output_dir, "ppt", "media")
    os.makedirs(media_dir, exist_ok=True)

    with zipfile.ZipFile(pptx_path) as z:
        media_files = [f for f in z.namelist() if f.startswith("ppt/media/")]
        for f in media_files:
            basename = os.path.basename(f)
            data = z.read(f)
            out_path = os.path.join(media_dir, basename)
            with open(out_path, "wb") as out:
                out.write(data)
            print(f"  Extracted: {basename} ({len(data):,} bytes)")

    # Convert EMF -> PNG and fix JPEG compatibility
    try:
        from PIL import Image
        # EMF -> PNG
        emf_files = [f for f in os.listdir(media_dir) if f.endswith(".emf")]
        for emf in emf_files:
            png_name = emf.replace(".emf", ".png")
            png_path = os.path.join(media_dir, png_name)
            try:
                img = Image.open(os.path.join(media_dir, emf))
                img.save(png_path, "PNG")
                print(f"  Converted: {emf} -> {png_name}")
            except Exception as e:
                print(f"  Warning: could not convert {emf}: {e}")
        # Fix JPEG headers for python-docx compatibility
        jpeg_files = [f for f in os.listdir(media_dir) if f.endswith((".jpeg", ".jpg"))]
        for jpg in jpeg_files:
            path = os.path.join(media_dir, jpg)
            try:
                img = Image.open(path)
                if img.mode == "RGBA":
                    img = img.convert("RGB")
                img.save(path, "JPEG", quality=95)
            except Exception as e:
                print(f"  Warning: could not fix {jpg}: {e}")
        if jpeg_files:
            print(f"  Fixed {len(jpeg_files)} JPEG files for Word compatibility")
    except ImportError:
        print("  Note: PIL not installed, skipping image conversion")


def build_image_mapping(pptx_path, output_dir):
    """Build {slide_num: [image_filenames]} mapping from PPTX XML."""
    slide_images = {}

    with zipfile.ZipFile(pptx_path) as z:
        slide_files = sorted(
            [f for f in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", f)]
        )

        for sf in slide_files:
            slide_num = int(re.search(r"(\d+)", os.path.basename(sf)).group(1))
            rel_file = f"ppt/slides/_rels/slide{slide_num}.xml.rels"

            if rel_file not in z.namelist():
                continue

            # Build rId -> media path map
            rel_root = ET.fromstring(z.read(rel_file).decode("utf-8"))
            rid_map = {}
            for rel in rel_root:
                rid = rel.get("Id")
                target = rel.get("Target")
                if rid and target:
                    if target.startswith("../"):
                        target = target.replace("../", "ppt/")
                    rid_map[rid] = target

            # Find image references
            root = ET.fromstring(z.read(sf).decode("utf-8"))
            images = []
            for blip in root.findall(".//a:blip", NAMESPACES):
                embed = blip.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                )
                if embed and embed in rid_map:
                    fname = rid_map[embed].split("/")[-1]
                    if fname.endswith(".emf"):
                        fname = fname.replace(".emf", ".png")
                    images.append(fname)

            if images:
                slide_images[slide_num] = images

    # Save mapping
    mapping_path = os.path.join(output_dir, "slide_images.json")
    with open(mapping_path, "w") as f:
        json.dump(slide_images, f, indent=2)

    print(f"\n  Image mapping saved: {len(slide_images)} slides with images")
    return slide_images


def main():
    if len(sys.argv) < 3:
        print("Usage: python extract_pptx.py <pptx_file> <output_dir>")
        sys.exit(1)

    pptx_path = sys.argv[1]
    output_dir = sys.argv[2]

    if not os.path.exists(pptx_path):
        print(f"Error: file not found: {pptx_path}")
        sys.exit(1)

    print(f"Extracting media from: {pptx_path}")
    extract_media(pptx_path, output_dir)

    print(f"\nBuilding slide-image mapping...")
    build_image_mapping(pptx_path, output_dir)

    print(f"\nDone. Output in: {output_dir}")


if __name__ == "__main__":
    main()
