# PPT2Word Skill

> Convert PowerPoint (.pptx) presentations to professional Word (.docx) documents — with tables, images, and SmartArt preserved.

[![Source](https://img.shields.io/badge/Source-aggre--cloud.com-blue)](https://acdatech.com)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8%2B-brightgreen)](https://python.org)

---

## Overview

PPT2Word is a Claude Code skill that transforms `.pptx` files into beautifully formatted `.docx` documents. Unlike simple text extraction, this skill preserves the **complete visual logic** of your presentation:

- All text content with hierarchical heading structure (H1/H2/H3)
- All tables with professional formatting (styled headers, alternating row colors)
- All images, charts, and SmartArt diagrams (including EMF vector graphics)
- Notes and source citations (collected at end of document as References)

**Origin**: This skill was created and open-sourced by [aggre-cloud (聚云科技)](https://acdatech.com), a technology company focused on cloud computing, big data, IoT, and AI technology transfer and digital transformation solutions.

**🇨🇳 [中文文档 / Chinese README](README_CN.md)**

---

## Table of Contents

- [How It Works](#how-it-works)
- [Installation](#installation)
- [Usage](#usage)
  - [Standalone Python Scripts](#standalone-python-scripts)
  - [As a Claude Code Skill](#as-a-claude-code-skill)
- [Output Format](#output-format)
- [Limitations](#limitations)
- [Performance](#performance)
- [Contributing](#contributing)
- [License](#license)

---

## How It Works

The conversion pipeline has 3 stages:

```
PPTX File
    │
    ├─► [Stage 1] markitdown ──► UTF-8 Markdown (text + tables)
    │
    ├─► [Stage 2] extract_pptx.py ──► Media files + slide-to-image mapping
    │       • Unzip PPTX (it's a ZIP archive)
    │       • Convert EMF (SmartArt) → PNG
    │       • Fix JPEG headers for Word compatibility
    │       • Parse XML to build {slide_num: [images]} mapping
    │
    └─► [Stage 3] generate_docx.py ──► Professional Word document
            • Cover page + Table of contents
            • Hierarchical headings from slide structure
            • Styled tables with blue headers
            • Images inserted at original positions
            • Notes collected at end as References
```

---

## Installation

### Prerequisites

```bash
pip install "markitdown[pptx]" python-docx Pillow
```

### Install the Skill

#### Option A: Via Claude Code skills install

```bash
npx skills add https://github.com/RichardASimon/PPT2Word.skill --skill ppt2word
```

#### Option B: Manual clone

```bash
git clone https://github.com/RichardASimon/PPT2Word.skill.git
cd PPT2Word.skill
# The skill is ready to use — Claude Code auto-detects skills in .agents/skills/
```

#### Option C: Copy to your project

```bash
cp -r PPT2Word.skill your-project/.agents/skills/ppt2word
```

---

## Usage

### Standalone Python Scripts

You can run the pipeline manually without Claude Code:

#### Step 1: Extract Text Content

```bash
python -m markitdown "your_presentation.pptx" 2>&1 | python -c "
import sys
text = sys.stdin.buffer.read().decode('gb18030', errors='replace')
with open('content.md', 'w', encoding='utf-8') as f:
    f.write(text)
print(f'Extracted: {len(text)} chars')
"
```

#### Step 2: Extract Media & Build Image Mapping

```bash
python scripts/extract_pptx.py "your_presentation.pptx" output_dir/
```

This creates:
- `output_dir/ppt/media/` — all images (EMF converted to PNG, JPEGs fixed)
- `output_dir/slide_images.json` — mapping of which images belong to which slide

#### Step 3: Generate Word Document

```bash
python scripts/generate_docx.py content.md output_dir/slide_images.json output_dir/ppt/media/ output.docx
```

**Arguments:**
| Arg | Description |
|-----|-------------|
| `content.md` | UTF-8 markdown from markitdown |
| `slide_images.json` | Slide-to-image mapping from extract_pptx.py |
| `media_dir` | Directory containing extracted images |
| `output.docx` | Output Word file path |

#### Quick One-Liner (All 3 Steps)

```bash
python -m markitdown input.pptx 2>&1 | python -c "
import sys; t=sys.stdin.buffer.read().decode('gb18030',errors='replace')
open('c.md','w',encoding='utf-8').write(t)" && \
python scripts/extract_pptx.py input.pptx out/ && \
python scripts/generate_docx.py c.md out/slide_images.json out/ppt/media/ output.docx
```

#### To Skip Images (Text + Tables Only)

```bash
python scripts/generate_docx.py content.md output.docx
```

When called with only 2 arguments, the script generates a text-only document without images.

---

### As a Claude Code Skill

Once installed, simply tell Claude what you want in natural language:

**Examples:**

```
"Convert my presentation.pptx to a Word document"
"把这个PPT转换成格式美观的Word文档"
"I need this PPTX as a report.docx, keep all the images and tables"
"Turn slides into a Word file with professional formatting"
"Make this presentation into a document with the SmartArt preserved"
```

Claude will automatically:
1. Run the full 3-stage pipeline
2. Handle encoding issues (Chinese GBK, etc.)
3. Report what was converted (tables, images, notes count)
4. Tell you where the output file is

**No commands to remember** — the skill triggers on intent.

---

## Output Format

The generated Word document follows a professional report structure:

```
[Cover Page]
    Title (28pt, navy #1E2761)
    Subtitle (16pt, teal #028090)
    Version + Date
    Company name
    [Logo/cover images]

[Table of Contents]
    Part 01 — Brazil Renewable Energy Market
    Part 02 — Energy Storage Regulations
    ...

═══════════════════════════════════════
[Part 01]                                    ← page break
═══════════════════════════════════════
    Section Heading (22pt, navy)
    ───────────────────────── (teal divider)
        Subsection (16pt, teal)
        Body text (10.5pt, black)
        [Image — centered, 14cm wide]
        [Table — blue header + alternating rows]
        ...

    Subsection (16pt, teal)
        Body text...
        [Image]
        ...

═══════════════════════════════════════
[Part 02]                                    ← page break
═══════════════════════════════════════
    ...

[参考资料 / References]
    【Slide 5】
        https://source-link...
    【Slide 13】
        https://source-link...
```

### Color Scheme

| Element | Color | Hex |
|---------|-------|-----|
| H1 Headers | Navy | `#1E2761` |
| H2 Headers | Teal | `#028090` |
| H3 Headers | Charcoal | `#36454F` |
| Table Header BG | Navy | `#1E2761` |
| Table Alt Row | Light Blue-Gray | `#F0F4F8` |
| Body Text | Black | `#000000` |

---

## Limitations

### What the Skill Handles Well

✅ Text extraction in any language (Chinese, English, Portuguese, etc.)
✅ Tables of any size (tested with 28-row × 6-column tables)
✅ Images: PNG, JPEG, and EMF (SmartArt/vector graphics → PNG)
✅ Complex slide structures with multiple images per slide
✅ Source notes and citations
✅ Multi-part presentations with hierarchical sections

### Known Limitations

❌ **Animated/Interactive Content**: Slide transitions, animations, triggered sequences, and hyperlinks to external URLs are not preserved (hyperlinked text is kept; tooltip/hover effects are lost).

❌ **Embedded Video/Audio**: Audio clips and video files embedded in slides are not extracted. The skill processes still images only.

❌ **3D Models**: Rotatable 3D models in PPTX are not supported.

❌ **Exact Visual Replication**: The output is a *Word document*, not a pixel-perfect copy of slides. Slide layouts (text boxes positioned side-by-side, custom shapes as design elements) are converted to linear document flow. Text appears in reading order.

❌ **Complex SmartArt Fidelity**: While EMF SmartArt is converted to PNG and preserved as an image, deeply nested or unusual SmartArt layouts may not render perfectly during EMF→PNG conversion (depends on PIL's EMF support).

❌ **Font Fidelity**: Original PPTX fonts are replaced with Arial (Latin) + SimSun/SimHei (Chinese). Custom fonts used in the original presentation are not embedded.

❌ **Slide Backgrounds/Design Themes**: Slide master backgrounds, gradients, and decorative design elements are stripped. The output uses a clean white Word document style.

❌ **Equations/Math**: OLE equations and complex mathematical notation may not be fully captured (depends on markitdown's extraction quality).

❌ **Very Large Files**: Presentations with 200+ slides and hundreds of high-resolution images will produce 50MB+ Word files. Processing may take 5-10 minutes.

❌ **Password-Protected PPTX**: Cannot process encrypted or password-protected presentations.

❌ **Linked (Not Embedded) Images**: Images linked from external files (not embedded in the PPTX) cannot be extracted.

### Workarounds

- For **exact slide fidelity** (design-heavy decks), consider exporting slides as images and inserting them into Word manually.
- For **equations**, review the output and use Word's equation editor to re-insert any missing formulas.
- For **large files**, compress images after generation: Word → Format Picture → Compress → 150ppi.

---

## Performance

| Presentation Size | Processing Time | Output Size |
|-------------------|----------------|-------------|
| 20 slides, no images | ~30 seconds | ~500 KB |
| 50 slides, 30 images | ~2 minutes | ~5 MB |
| 92 slides, 112 images | ~3 minutes | ~31 MB |
| 150+ slides, 200+ images | ~5-10 minutes | ~50+ MB |

*Tested on Windows 10, Python 3.14, SSD, 16GB RAM*

---

## Project Structure

```
ppt2word/
├── SKILL.md                          # Skill definition (what Claude reads)
├── README.md                         # This file
├── LICENSE                           # MIT License
└── scripts/
    ├── extract_pptx.py               # Stage 2: Media extraction + mapping
    └── generate_docx.py              # Stage 3: Word document generation
```

- `SKILL.md` — Contains the instructions Claude follows when the skill triggers
- `scripts/extract_pptx.py` — Standalone script: extracts images from PPTX, converts EMF→PNG, builds slide-image mapping
- `scripts/generate_docx.py` — Standalone script: generates Word doc from markdown + images

---

## Contributing

Found a bug or want to improve the skill? Contributions are welcome!

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Test with your own PPTX files
4. Submit a Pull Request

**Known areas for improvement:**
- [ ] Add support for PowerPoint charts (not just SmartArt)
- [ ] Add option to preserve slide layout as images
- [ ] Improve equation extraction
- [ ] Add batch processing for multiple PPTX files

---

## Credits

- **Created by**: [aggre-cloud (聚云科技)](https://acdatech.com)
- **Skill architecture**: Anthropic Claude Code skills framework
- **Text extraction**: [markitdown](https://github.com/microsoft/markitdown) by Microsoft
- **Image processing**: [Pillow (PIL)](https://python-pillow.org/)
- **Word generation**: [python-docx](https://python-docx.readthedocs.io/)

---

## License

MIT License — free to use, modify, and distribute. See [LICENSE](LICENSE) for details.

---

<p align="center">
  <sub>Built with ❤️ by <a href="https://acdatech.com">acdatech.com</a></sub>
</p>
