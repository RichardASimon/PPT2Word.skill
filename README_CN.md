# PPT2Word 技能

> 将 PowerPoint (.pptx) 演示文稿转换为专业 Word (.docx) 文档 — 完整保留表格、图片和 SmartArt 图表。

[![来源](https://img.shields.io/badge/来源-aggre--cloud-blue)](https://acdatech.com)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8%2B-brightgreen)](https://python.org)

---

## 简介

PPT2Word 是一款 Claude Code 技能，可将 `.pptx` 文件转换为格式精美的 `.docx` 文档。不同于简单的文本提取，本技能完整保留演示文稿的**视觉逻辑**：

- 所有文本内容，带层级标题结构（H1/H2/H3）
- 所有表格，专业格式（蓝色表头、交替行底色）
- 所有图片、图表和 SmartArt 图形（含 EMF 矢量图）
- 备注和引用来源（收集在文档末尾的参考资料章节）

**来源**: 本技能由 [aggre-cloud 聚云科技](https://acdatech.com) 创建并开源。聚云科技是一家专注于**云计算、大数据、物联网、人工智能**等领域技术成果转化与数字化转型解决方案的科技公司。

**🇺🇸 [English README](README.md)**

---

## 目录

- [工作原理](#工作原理)
- [安装](#安装)
- [使用方法](#使用方法)
  - [独立 Python 脚本运行](#独立-python-脚本运行)
  - [作为 Claude Code 技能使用](#作为-claude-code-技能使用)
- [输出格式](#输出格式)
- [极限与限制](#极限与限制)
- [性能](#性能)
- [贡献](#贡献)
- [许可证](#许可证)

---

## 工作原理

转换流程分为 3 个阶段：

```
PPTX 文件
    │
    ├─► [阶段 1] markitdown ──► UTF-8 Markdown（文本 + 表格）
    │
    ├─► [阶段 2] extract_pptx.py ──► 媒体文件 + 幻灯片-图片映射
    │       • 解压 PPTX（本质是 ZIP 压缩包）
    │       • 转换 EMF（SmartArt）→ PNG
    │       • 修复 JPEG 头文件以兼容 Word
    │       • 解析 XML 构建 {幻灯片编号: [图片列表]} 映射
    │
    └─► [阶段 3] generate_docx.py ──► 专业 Word 文档
            • 封面页 + 目录
            • 基于幻灯片结构的层级标题
            • 带样式表格（蓝色表头）
            • 在原始位置插入图片
            • 备注收集在末尾作为参考资料
```

---

## 安装

### 前置依赖

```bash
pip install "markitdown[pptx]" python-docx Pillow
```

### 安装技能

#### 方式一：通过 Claude Code skills install

```bash
npx skills add https://github.com/RichardASimon/PPT2Word.skill --skill ppt2word
```

#### 方式二：手动克隆

```bash
git clone https://github.com/RichardASimon/PPT2Word.skill.git
cd PPT2Word.skill
# 技能已就绪 — Claude Code 自动识别 .agents/skills/ 目录
```

#### 方式三：复制到项目

```bash
cp -r PPT2Word.skill your-project/.agents/skills/ppt2word
```

---

## 使用方法

### 独立 Python 脚本运行

无需 Claude Code，可手动运行完整流程：

#### 步骤 1：提取文本内容

```bash
python -m markitdown "your_presentation.pptx" 2>&1 | python -c "
import sys
text = sys.stdin.buffer.read().decode('gb18030', errors='replace')
with open('content.md', 'w', encoding='utf-8') as f:
    f.write(text)
print(f'已提取: {len(text)} 字符')
"
```

#### 步骤 2：提取媒体文件 + 构建图片映射

```bash
python scripts/extract_pptx.py "your_presentation.pptx" output_dir/
```

输出：
- `output_dir/ppt/media/` — 所有图片（EMF 已转 PNG，JPEG 已修复）
- `output_dir/slide_images.json` — 每张幻灯片对应的图片映射

#### 步骤 3：生成 Word 文档

```bash
python scripts/generate_docx.py content.md output_dir/slide_images.json output_dir/ppt/media/ output.docx
```

**参数说明：**
| 参数 | 说明 |
|------|------|
| `content.md` | markitdown 输出的 UTF-8 markdown |
| `slide_images.json` | extract_pptx.py 生成的图片映射 |
| `media_dir` | 提取的图片目录 |
| `output.docx` | 输出 Word 文件路径 |

#### 一键运行（全部 3 步）

```bash
python -m markitdown input.pptx 2>&1 | python -c "
import sys; t=sys.stdin.buffer.read().decode('gb18030',errors='replace')
open('c.md','w',encoding='utf-8').write(t)" && \
python scripts/extract_pptx.py input.pptx out/ && \
python scripts/generate_docx.py c.md out/slide_images.json out/ppt/media/ output.docx
```

#### 跳过图片（仅文本 + 表格）

```bash
python scripts/generate_docx.py content.md output.docx
```

仅传 2 个参数时，脚本生成纯文本文档，不插入图片。

---

### 作为 Claude Code 技能使用

安装后，直接用自然语言告诉 Claude 你想做什么：

**触发示例：**

```
"把 presentation.pptx 转成 Word 文档"
"把这个PPT转换成格式美观的Word文档"
"I need this PPTX as a report.docx, keep all the images and tables"
"Turn slides into a Word file with professional formatting"
"Make this presentation into a document with the SmartArt preserved"
"帮我转成docx，图片和表格都要保留"
```

Claude 会自动：
1. 运行完整的 3 阶段流程
2. 处理编码问题（中文 GBK 等）
3. 报告转换结果（表格、图片、备注数量）
4. 告诉你输出文件位置

**无需记忆命令** — 技能基于意图触发。

---

## 输出格式

生成的 Word 文档遵循专业报告结构：

```
[封面页]
    标题（28pt，深蓝 #1E2761）
    副标题（16pt，青色 #028090）
    版本 + 日期
    公司名称
     Logo/封面图片 ]

[目录]
    第一部分 — 巴西可再生能源市场
    第二部分 — 储能市场监管
    ...

═══════════════════════════════════════
[第一部分]                              ← 分页符
═══════════════════════════════════════
    章节标题（22pt，深蓝）
    ─────────────────────────（青色分隔线）
        子章节（16pt，青色）
        正文（10.5pt，黑色）
        [图片 — 居中，14cm 宽]
        [表格 — 蓝色表头 + 交替行]
        ...

    子章节（16pt，青色）
        正文...
        [图片]
        ...

═══════════════════════════════════════
[第二部分]                              ← 分页符
═══════════════════════════════════════
    ...

[参考资料 / References]
    【幻灯片 5】
        https://source-link...
    【幻灯片 13】
        https://source-link...
```

### 配色方案

| 元素 | 颜色 | 色值 |
|------|------|------|
| H1 标题 | 深蓝 | `#1E2761` |
| H2 标题 | 青色 | `#028090` |
| H3 标题 | 炭灰 | `#36454F` |
| 表格表头背景 | 深蓝 | `#1E2761` |
| 表格交替行 | 浅蓝灰 | `#F0F4F8` |
| 正文 | 黑色 | `#000000` |

---

## 极限与限制

### 技能擅长处理

✅ 任意语言的文本提取（中文、英文、葡萄牙文等）
✅ 任意大小的表格（实测 28 行 × 6 列）
✅ 图片：PNG、JPEG、EMF（SmartArt/矢量图 → PNG）
✅ 单张幻灯片含多张图片的复杂结构
✅ 来源备注和引用
✅ 多部分、多层次结构的演示文稿

### 已知限制

❌ **动画/交互内容**：幻灯片切换动画、触发器动画、外部链接等不保留（超链接文本保留，悬停效果丢失）

❌ **嵌入视频/音频**：音频和视频文件不提取，仅处理静态图片

❌ **3D 模型**：不支持 PPTX 中的可旋转 3D 模型

❌ **非像素级复制**：输出是 *Word 文档*，不是幻灯片的像素级复制品。幻灯片布局（文本框并排、自定义形状等设计元素）会转换为线性文档流，文本按阅读顺序排列

❌ **复杂 SmartArt 保真度**：虽然 EMF 格式的 SmartArt 会转为 PNG 图片保留，但嵌套层级较深或特殊布局的 SmartArt 在 EMF→PNG 转换时可能渲染不完美（取决于 PIL 的 EMF 支持能力）

❌ **字体保真度**：原始 PPTX 中的字体被替换为 Arial（拉丁）+ SimSun/SimHei（中文），不嵌入自定义字体

❌ **幻灯片背景/设计主题**：幻灯片母版背景、渐变、装饰性设计元素会被去除，输出为干净的白色 Word 文档风格

❌ **公式/数学符号**：OLE 公式和复杂数学符号可能无法完全提取（取决于 markitdown 的提取质量）

❌ **超大文件**：200+ 幻灯片、数百张高清图片的演示文稿会生成 50MB+ 的 Word 文件，处理时间可能达 5-10 分钟

❌ **加密 PPTX**：无法处理加密或密码保护的演示文稿

❌ **链接（非嵌入）图片**：引用外部文件的链接图片无法提取

### 替代方案

- 对于**设计密集型幻灯片**（对视觉保真度要求高），建议将幻灯片导出为图片后手动插入 Word
- 对于**公式**，生成后请检查并使用 Word 的公式编辑器重新插入缺失的公式
- 对于**大文件**，生成后可压缩图片：Word → 设置图片格式 → 压缩 → 150ppi

---

## 性能

| 演示文稿规模 | 处理时间 | 输出大小 |
|-------------|---------|---------|
| 20 张幻灯片，无图片 | ~30 秒 | ~500 KB |
| 50 张幻灯片，30 张图片 | ~2 分钟 | ~5 MB |
| 92 张幻灯片，112 张图片 | ~3 分钟 | ~31 MB |
| 150+ 张幻灯片，200+ 张图片 | ~5-10 分钟 | ~50+ MB |

*测试环境：Windows 10, Python 3.14, SSD, 16GB RAM*

---

## 项目结构

```
ppt2word/
├── SKILL.md                          # 技能定义（Claude 读取的指令）
├── README.md                         # 英文文档
├── README_CN.md                      # 中文文档（本文件）
├── LICENSE                           # MIT 许可证
└── scripts/
    ├── extract_pptx.py               # 阶段 2：媒体提取 + 映射构建
    └── generate_docx.py              # 阶段 3：Word 文档生成
```

- `SKILL.md` — 技能触发时 Claude 执行的指令
- `scripts/extract_pptx.py` — 独立脚本：从 PPTX 提取图片、EMF→PNG、构建幻灯片-图片映射
- `scripts/generate_docx.py` — 独立脚本：基于 markdown + 图片生成 Word 文档

---

## 贡献

发现问题或想改进技能？欢迎贡献！

1. Fork 本仓库
2. 创建功能分支 (`git checkout -b feature/your-feature`)
3. 用自己的 PPTX 文件测试
4. 提交 Pull Request

**已知可改进方向：**
- [ ] 增加对 PowerPoint 图表的支持（不仅是 SmartArt）
- [ ] 增加将幻灯片布局保留为图片的选项
- [ ] 改进公式提取
- [ ] 增加批量处理多个 PPTX 文件的功能

---

## 致谢

- **创建者**: [aggre-cloud 聚云科技](https://acdatech.com)
- **技能架构**: Anthropic Claude Code skills 框架
- **文本提取**: [markitdown](https://github.com/microsoft/markitdown) by Microsoft
- **图片处理**: [Pillow (PIL)](https://python-pillow.org/)
- **Word 生成**: [python-docx](https://python-docx.readthedocs.io/)

---

## 许可证

MIT License — 可自由使用、修改和分发。详见 [LICENSE](LICENSE)。

---

<p align="center">
  <sub>由 <a href="https://acdatech.com">aggre-cloud 聚云科技</a> 用心构建</a></sub>
</p>
