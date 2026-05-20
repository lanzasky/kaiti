---
name: kaiti-skill
description: Use when writing, rewriting, formatting, or preparing defense PPT blueprints for Chinese graduate proposal reports, especially food engineering/digital twin proposals requiring balanced literature review, academic DOCX formatting, and 16-slide oral defense planning.
---

# Kaiti Skill

This skill packages the proposal workflow developed for `预制菜热处理数字孪生开题报告`.

## Core Workflow

Use `kaiti_workflow.py` when the user asks for any of:

1. Rewriting a proposal draft into a balanced, template-aligned Markdown report.
2. Formatting a generated Word document against a Chinese academic template.
3. Generating a 16-slide defense PPT blueprint with image placement and oral scripts.

## Python API

```python
from kaiti_skill import (
    rewrite_and_balance_proposal,
    format_academic_docx,
    generate_ppt_blueprint,
)
```

### Rewrite

```python
rewrite_and_balance_proposal(input_file, output_file)
```

Reads `.docx`, `.md`, or `.txt`; writes a Markdown proposal. If `OPENAI_API_KEY` is available, the function calls the OpenAI Responses API with the embedded V8-style rewrite prompt. Otherwise it writes a structured prompt package for Codex/human review.

### Format DOCX

```python
format_academic_docx(input_docx, reference_template, output_docx)
```

Applies academic formatting:

- A4 page and template margins.
- Songti Chinese, Times New Roman Latin/digits, small-four body size.
- 1.5 line spacing and two-character first-line indent.
- Black outside page border via WordprocessingML.
- Three-line academic tables: top/bottom 1.5 pt, header line 0.75 pt, no vertical inner borders.
- Hanging indent for references.

### PPT Blueprint

```python
generate_ppt_blueprint(input_docx, image_dir, output_blueprint)
```

Builds a strict 16-slide defense guide containing:

- 【幻灯片标题】
- 【指定插入图片】
- 【PPT 核心文字】
- 【汇报讲稿】

If `output_blueprint` ends with `.docx`, the function creates a Word guide and embeds the assigned local images when available. Otherwise it writes Markdown.

## CLI

```powershell
python -m kaiti_skill.kaiti_workflow format input.docx template.docx output.docx
python -m kaiti_skill.kaiti_workflow blueprint input.docx figures output.docx
python -m kaiti_skill.kaiti_workflow rewrite input.docx output.md
```
