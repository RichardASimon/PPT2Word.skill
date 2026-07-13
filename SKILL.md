---
name: ppt2word
description: "Convert PowerPoint (.pptx) files to professional Word (.docx) documents. Use whenever the user wants to convert a PPT/PPTX to Word, transform slides into a document, preserve tables/images/SmartArt from slides in a Word file, or mentions 'PPT to Word', 'PPTX to DOCX', 'slides to document', 'convert presentation to report'. Also triggers when the user has a .pptx file and asks to make it a 'report', 'document', 'doc', or 'Word file'. Works with any .pptx file regardless of language (Chinese, English, Portuguese, etc.)."
---

# PPT to Word Converter

Converts `.pptx` presentations to professionally formatted `.docx` documents, preserving text, tables, images, and SmartArt diagrams in their original logical order.

## When to Use

Trigger this skill when:
- The user has a `.pptx` file and wants a `.docx` version
- The user asks to "convert PPT to Word" (or similar phrasing)
- The user wants to transform a presentation into a report/document
- The user needs to preserve tables, charts, or images from slides in a Word file

## Dependencies

Install these before running (only needed once):

```bash
pip install "markitdown[pptx]" python-docx Pillow
```

The script needs no other external tools — PPTX media extraction is done via Python's built-in `zipfile`.

## Workflow

### Step 1: Extract Text Content

Run markitdown to extract all slide text and tables. Pipe through Python to handle encoding:

```bash
python -m markitdown "input.pptx" 2>&1 | python -c "
import sys
text = sys.stdin.buffer.read().decode('gb18030', errors='replace')
with open('content.md', 'w', encoding='utf-8') as f:
    f.write(text)
print(f'Extracted: {len(text)} chars')
"
```

The `gb18030` decode handles both UTF-8 and GBK outputs (it's a superset). Using `buffer.read()` avoids Windows console encoding issues that can truncate the output with plain `>` redirection.

### Step 2: Extract Media Files

PPTX is a ZIP archive. Extract all images/charts/SmartArt (EMF files are SmartArt/vector charts):

```python
import zipfile, os
with zipfile.ZipFile("input.pptx") as z:
    os.makedirs("media", exist_ok=True)
    for f in z.namelist():
        if f.startswith("ppt/media/"):
            z.extract(f, "media")
```

EMF files must to be converted to PNG for Word embedding:

```python
from PIL import Image
import os
for f in os.listdir("media/ppt/media/"):
    if f.endswith(".emf"):
        img = Image.open(f"media/ppt/media/{f}")
        img.save(f"media/ppt/media/{f.replace('.emf','.png')}", "PNG")
```

### Step 3: Build Slide-to-Image Mapping

Parse the PPTX XML to know which images appear on which slide. This preserves the original visual logic:

```python
import zipfile, re
from xml.etree import ElementTree as ET

ns = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def get_slide_images(pptx, slide_num):
    rel_file = f'ppt/slides/_rels/slide{slide_num}.xml.rels'
    slide_file = f'ppt/slides/slide{slide_num}.xml'
    with zipfile.ZipFile(pptx) as z:
        # Build rId -> media path map from relationship file
        rel_root = ET.fromstring(z.read(rel_file))
        rid_map = {}
        for rel in rel_root:
            rid = rel.get('Id')
            target = rel.get('Target')
            if target and target.startswith('../'):
                target = target.replace('../', 'ppt/')
            rid_map[rid] = target
        # Find all image references in slide
        root = ET.fromstring(z.read(slide_file))
        images = []
        for blip in root.findall('.//a:blip', ns):
            embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if embed in rid_map:
                fname = rid_map[embed].split('/')[-1]
                if fname.endswith('.emf'):
                    fname = fname.replace('.emf', '.png')
                images.append(fname)
    return images

# Build full mapping
slide_images = {}
for sn in range(1, 93):
    imgs = get_slide_images("input.pptx", sn)
    if imgs:
        slide_images[sn] = imgs
```

### Step 4: Generate Word Document

Use the bundled `scripts/generate_docx.py` as the base generation script. It handles:

1. **Cover page** — title, subtitle, version, date, company
2. **Table of contents** — auto-generated from part structure
3. **Structured content** — H1/H2/H3 headings from slide titles
4. **Tables** — properly formatted with header row and alternating row colors
5. **Images** — inserted at their original slide positions
6. **Notes/Sources** — collected at the end of the document (not inline)
7. **Headers/Footers** — with page numbers

The script reads from two inputs:
- The UTF-8 markdown file (text + tables)
- The `slide_images.json` dict (slide number → image filenames)

Run it:

```bash
python scripts/generate_docx.py content_utf8.md slide_images.json media/ppt/media/ output.docx
```

## Output Format

The generated Word document follows this structure:

```
[Cover Page]
  Title (28pt, navy)
  Subtitle (16pt, teal)
  Version + Date
  Company name
  [Cover images/logos]

[Table of Contents]
  Part 01 — <Topic>
  Part 02 — <Topic>
  ...

[Part 01] ← page break
  Section heading (22pt, navy)
  ───────────────────────── (teal divider)
    Subsection (16pt, teal)
    Body text (10.5pt, black)
    [Image]
    [Table with styled header]
    ...

[Part 02]
  ...

[参考资料 / References]
  【Slide 5】
    https://source-link...
  【Slide 13】
    https://source-link...
```

## Handling Common Issues

### Encoding
Always check encoding first. markitdown output for Chinese PPTs is often GBK. If you see `UnicodeDecodeError: 'utf-8'`, convert with `gb18030`.

### EMF (SmartArt/Charts)
EMF files are vector graphics (SmartArt, flowcharts, diagrams). PIL can read most EMF files and convert to PNG. If PIL fails, the image is skipped with a warning — the document is still generated.

### JPEG Compatibility
Some JPEGs in PPTX use non-standard headers that python-docx can't embed. If you see `Warning: could not insert`, re-save the JPEG with PIL:

```python
from PIL import Image
img = Image.open(path)
img = img.convert('RGB')
img.save(path, 'JPEG', quality=95)
```

### Large Files
For presentations with 50+ slides and many images, the output `.docx` can be large (20-50MB). This is normal. Users can compress images in Word via Format → Compress Pictures after generation.

### Tables Split by Blank Lines
Markitdown may insert blank lines between table rows. The parser must treat consecutive `|`-starting lines as one table, ignoring blank lines within them. See `scripts/generate_docx.py` for the correct implementation.

### Corrupted Images
Some JPEGs in PPTX may use non-standard headers. Re-save them with PIL before embedding:

```python
from PIL import Image
img = Image.open(path)
if img.mode == 'RGBA':
    img = img.convert('RGB')
img.save(path, 'JPEG', quality=95)
```

## Customization

The `scripts/generate_docx.py` script is designed to be edited. Key customization points at the top of the file:

- `REPORT_TITLE` / `REPORT_SUBTITLE` — cover page text
- `COMPANY` — company name
- `PART_NAMES` — mapping of part numbers to section titles
- Color scheme (`header_color`, `alt_color` in table function)
- Font sizes and families

## Performance

- Typical 50-slide PPTX: ~2 minutes
- 100+ slides with images: ~3-5 minutes
- Output quality improves with the richness of extraction (all media preserved)
