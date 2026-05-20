---
name: kaiti-skill
description: Use when the user needs a generic academic proposal engine driven by a Gemini Deep Research draft, a local references folder, and a school template; supports proposal synthesis, academic DOCX formatting with three-line tables, and 15-16 slide defense PPT blueprint generation.
---

# Kaiti Skill

This skill is a generic academic proposal assistant. It is not tied to any discipline, topic, model, equipment, or project name.

## When To Use

Use this skill when the user provides:

- `gemini_draft_file`: a Gemini Deep Research draft, outline, or initial proposal.
- `references_dir`: a folder or file list containing local reference materials.
- `template_file`: the target university or institute proposal template.

## Core API

```python
from kaiti_skill import (
    synthesize_and_rewrite_proposal,
    format_academic_docx,
    generate_generic_ppt_blueprint,
)
```

### 1. Proposal Synthesis

```python
synthesize_and_rewrite_proposal(
    gemini_draft_file,
    references_dir,
    template_file,
    output_file,
)
```

The function:

- reads the Gemini draft to infer the new topic, research object, objectives, and route;
- extracts readable snippets from the local references folder;
- extracts headings and hierarchy from the template;
- calls an LLM when `OPENAI_API_KEY` is configured;
- writes a Markdown proposal aligned to the template.

If no API key is configured, it writes a complete prompt package to `output_file` so Codex can continue the generation transparently.

### 2. Academic DOCX Formatting

```python
format_academic_docx(input_docx, template_file, output_docx)
```

The formatting engine applies cross-disciplinary Chinese academic formatting:

- template page size and margins;
- Songti Chinese body text, Times New Roman Latin/digits, small-four body size;
- 1.5 line spacing and first-line indent;
- black outside page border;
- academic three-line tables: top/bottom 1.5 pt, header line 0.75 pt, no vertical borders, centered cell text;
- centered figure/table captions and hanging reference indents.

### 3. Generic PPT Blueprint

```python
generate_generic_ppt_blueprint(input_docx, image_list, output_blueprint)
```

The function reads the proposal text and creates a 15-16 page Word or Markdown defense guide with:

- 【幻灯片标题】
- 【指定插入图片】
- 【PPT 核心文字】
- 【汇报讲稿】

It assigns images by filename semantics. When no matching image is found, it leaves a red placeholder.

## CLI

```powershell
python -m kaiti_skill.kaiti_workflow synthesize draft.docx references template.docx proposal.md
python -m kaiti_skill.kaiti_workflow format proposal.docx template.docx formatted.docx
python -m kaiti_skill.kaiti_workflow blueprint proposal.docx figures blueprint.docx
```
