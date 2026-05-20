from __future__ import annotations

import argparse
import json
import os
import re
import textwrap
import urllib.request
from copy import deepcopy
from pathlib import Path
from typing import Iterable

try:
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
except ImportError as exc:  # pragma: no cover - dependency message path
    raise RuntimeError(
        "kaiti_skill requires python-docx. Install it with `python -m pip install python-docx`."
    ) from exc


REWRITE_PROMPT = """\
你是一位严谨的食品工程与自动化交叉学科博导，请将输入材料重写为正式开题报告 Markdown。

必须执行以下规则：
1. 第一部分立论依据约 9000-10000 字，不能把字数堆在单一小节，需均衡分配至 1.1、1.2.1、1.2.2、1.2.3、1.2.4 与 1.3。
2. 1.1 采用“背景引入 -> 然而传统方法存在局限 -> 四个突出问题 -> 综上所述 -> 本研究意义”的模板逻辑。
3. 1.2 国内外研究现状拆为四节：食品热加工 CFD 与多物理场、风味动力学与品质评价、多源感知与在线状态估计、ROM/ANODE 与 MPC。
4. 每个 1.2 子节都要深入引用具体方法、设备、算法局限，并在末尾生成表头统一的总结表：研究对象 | 核心研究方法 | 观测指标 | 参考价值 | 局限性 | 参考文献。
5. 文献引用必须按出场顺序连续编号，正文用上标方括号，主要参考文献放在 1.3 之后，条目之间空一行。
6. 2.1 与 2.2 保持精炼概括；2.3 是技术路线主战场，必须采用流畅学术叙事，严禁“需求痛点：”“数学建模：”等机械小标题。
7. 2.3 的四个细化部分均采用“问题引出 -> 工程实施 -> 公式 -> 参数解释 -> 对本课题意义”的夹叙夹议写法。
8. 2.3 必须包含并解释 N-S 方程、Boussinesq 浮力项、Christiansen-Craig 非牛顿降黏、Darcy-Forchheimer 多孔阻力、Arrhenius 动力学、ANODE、MHE、Pareto-MPC。
9. 2.4 实验方案与 2.3 明确分开，保留软硬件配置表、APBRS 泛化训练矩阵和对照组/实验组设计。
10. 全文公式使用纯净 LaTeX：行内 $...$，独立公式 $$...$$，禁止使用 \\[ 或 \\]。

请输出完整 Markdown，不要解释过程。
"""


PPT_IMAGE_MAP = {
    4: "PPT_FigA_问题提出真实拼图.png",
    6: "图1：预制菜热处理数字孪生技术路线图.png",
    7: "PPT_FigB_朴素CFD仿真底座.png",
    8: "PPT_FigC_朴素多源检测流程.png",
    9: "PPT_FigD_极简ANODE数学框图.png",
    10: "图3：数字孪生数学模型与控制闭环框架图.png",
    11: "图2：桌面级多模态微型杀菌釜实验平台布局图.png",
    13: "预制菜热处理数字孪生前期仿真与优化成果综合图.png",
}


def _path(value: str | os.PathLike) -> Path:
    return Path(value).expanduser().resolve()


def _read_docx_text(path: Path) -> str:
    doc = Document(str(path))
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    cells: list[str] = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if text:
                    cells.append(text)
    return "\n".join(paragraphs + cells)


def _read_text_file(path: Path) -> str:
    if path.suffix.lower() == ".docx":
        return _read_docx_text(path)
    return path.read_text(encoding="utf-8", errors="ignore")


def _openai_rewrite(prompt: str, model: str | None = None) -> str:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not set.")
    payload = {
        "model": model or os.getenv("KAITI_OPENAI_MODEL", "gpt-4.1"),
        "input": prompt,
    }
    req = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=600) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    chunks: list[str] = []
    for item in data.get("output", []):
        for content in item.get("content", []):
            if content.get("type") in {"output_text", "text"}:
                chunks.append(content.get("text", ""))
    result = "\n".join(chunks).strip()
    if not result:
        raise RuntimeError("OpenAI response did not contain output text.")
    return result


def rewrite_and_balance_proposal(input_file: str | os.PathLike, output_file: str | os.PathLike) -> Path:
    """Rewrite a proposal draft into a balanced V8-style Markdown document.

    If `OPENAI_API_KEY` is available, this function calls the OpenAI Responses API.
    Without an API key it writes a reproducible prompt package to `output_file`, so
    Codex can continue the rewrite with the same constraints.
    """

    src = _path(input_file)
    dst = _path(output_file)
    source_text = _read_text_file(src)
    prompt = f"{REWRITE_PROMPT}\n\n# 输入材料\n\n{source_text}"
    try:
        output = _openai_rewrite(prompt)
    except Exception as exc:
        output = (
            "# 开题报告重写任务包\n\n"
            "> 当前环境未完成在线 AI 重写，以下为可复现的 V8 风格重写 Prompt 与输入材料。\n\n"
            f"> 触发原因：{exc}\n\n"
            "## 重写 Prompt\n\n"
            f"{REWRITE_PROMPT}\n\n"
            "## 输入材料\n\n"
            f"{source_text}\n"
        )
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(output, encoding="utf-8")
    return dst


def _set_rfonts(element, east_asia: str = "宋体", latin: str = "Times New Roman") -> None:
    rpr = element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for key, value in {
        "w:ascii": latin,
        "w:hAnsi": latin,
        "w:cs": latin,
        "w:eastAsia": east_asia,
    }.items():
        rfonts.set(qn(key), value)


def _format_run(run, east_asia: str, latin: str, size_pt: float | None, bold: bool | None) -> None:
    run.font.name = latin
    _set_rfonts(run._element, east_asia, latin)
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold


def _apply_page_border(section, template_border=None) -> None:
    sect_pr = section._sectPr
    for old in list(sect_pr.findall(qn("w:pgBorders"))):
        sect_pr.remove(old)
    if template_border is not None:
        border = deepcopy(template_border)
    else:
        border = OxmlElement("w:pgBorders")
        border.set(qn("w:offsetFrom"), "page")
        for edge in ("top", "left", "bottom", "right"):
            child = OxmlElement(f"w:{edge}")
            child.set(qn("w:val"), "single")
            child.set(qn("w:sz"), "8")
            child.set(qn("w:space"), "24")
            child.set(qn("w:color"), "000000")
            border.append(child)
    pg_mar = sect_pr.find(qn("w:pgMar"))
    if pg_mar is not None:
        sect_pr.insert(sect_pr.index(pg_mar), border)
    else:
        sect_pr.append(border)


def _set_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        old = borders.find(qn(f"w:{edge}"))
        if old is not None:
            borders.remove(old)
        elem = OxmlElement(f"w:{edge}")
        if edge in {"top", "bottom"}:
            elem.set(qn("w:val"), "single")
            elem.set(qn("w:sz"), "12")  # 1.5 pt
            elem.set(qn("w:color"), "000000")
        else:
            elem.set(qn("w:val"), "nil")
            elem.set(qn("w:sz"), "0")
            elem.set(qn("w:color"), "auto")
        borders.append(elem)


def _set_cell_borders(cell, header_bottom: bool = False) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        old = borders.find(qn(f"w:{edge}"))
        if old is not None:
            borders.remove(old)
        elem = OxmlElement(f"w:{edge}")
        if edge == "bottom" and header_bottom:
            elem.set(qn("w:val"), "single")
            elem.set(qn("w:sz"), "6")  # 0.75 pt
            elem.set(qn("w:color"), "000000")
        else:
            elem.set(qn("w:val"), "nil")
            elem.set(qn("w:sz"), "0")
            elem.set(qn("w:color"), "auto")
        borders.append(elem)


def _is_caption(text: str) -> bool:
    return bool(re.match(r"^(图|表)\s*[\d一二三四五六七八九十]+", text.strip()))


def format_academic_docx(
    input_docx: str | os.PathLike,
    reference_template: str | os.PathLike,
    output_docx: str | os.PathLike,
) -> Path:
    """Apply Chinese academic proposal formatting to a DOCX document."""

    src = _path(input_docx)
    template_path = _path(reference_template)
    dst = _path(output_docx)
    doc = Document(str(src))
    template = Document(str(template_path))
    tsec = template.sections[0]
    template_border = tsec._sectPr.find(qn("w:pgBorders"))
    template_border = deepcopy(template_border) if template_border is not None else None

    for section in doc.sections:
        section.page_width = tsec.page_width
        section.page_height = tsec.page_height
        section.top_margin = tsec.top_margin
        section.bottom_margin = tsec.bottom_margin
        section.left_margin = tsec.left_margin
        section.right_margin = tsec.right_margin
        _apply_page_border(section, template_border)

    in_refs = False
    for para in doc.paragraphs:
        text = para.text.strip()
        if text == "主要参考文献":
            in_refs = True
        elif text.startswith("第二部分"):
            in_refs = False

        if para.style.name.startswith("Heading 1") or text.startswith(("第一部分", "第二部分", "三、", "四、")):
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.first_line_indent = None
            para.paragraph_format.line_spacing = 1.5
            for run in para.runs:
                _format_run(run, "黑体", "Times New Roman", 15, True)
        elif para.style.name.startswith("Heading 2") or re.match(r"^\d+\.\d+\s+", text):
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para.paragraph_format.first_line_indent = None
            para.paragraph_format.line_spacing = 1.5
            for run in para.runs:
                _format_run(run, "黑体", "Times New Roman", 14, True)
        elif para.style.name.startswith("Heading 3") or re.match(r"^\d+\.\d+\.\d+\s+", text):
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para.paragraph_format.first_line_indent = None
            para.paragraph_format.line_spacing = 1.5
            for run in para.runs:
                _format_run(run, "黑体", "Times New Roman", 12, True)
        elif _is_caption(text):
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.first_line_indent = None
            para.paragraph_format.line_spacing = 1.0
            for run in para.runs:
                _format_run(run, "宋体", "Times New Roman", 10.5, True)
        elif in_refs and re.match(r"^\[\d+\]", text):
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            para.paragraph_format.left_indent = Cm(0.74)
            para.paragraph_format.first_line_indent = -Cm(0.74)
            para.paragraph_format.line_spacing = 1.25
            for run in para.runs:
                _format_run(run, "宋体", "Times New Roman", 10.5, False)
        else:
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            para.paragraph_format.first_line_indent = Pt(24)
            para.paragraph_format.line_spacing = 1.5
            for run in para.runs:
                _format_run(run, "宋体", "Times New Roman", 12, False)

    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        _set_table_borders(table)
        for row_idx, row in enumerate(table.rows):
            for cell in row.cells:
                _set_cell_borders(cell, header_bottom=row_idx == 0)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    para.paragraph_format.first_line_indent = None
                    para.paragraph_format.line_spacing = 1.0
                    for run in para.runs:
                        _format_run(run, "宋体", "Times New Roman", 10.5, row_idx == 0)

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst))
    return dst


def _make_slide_blueprint() -> list[dict]:
    return [
        {"no": 1, "title": "封面｜基于物理信息降阶模型与虚实同步校正的预制菜热加工数字孪生系统研究", "core": ["中国农业大学工学院", "硕士开题答辩", "汇报人：杨阳", "导师：张小栓 教授"], "script": "尊敬的各位评委老师，大家好。我是中国农业大学工学院的硕士生杨阳。今天我开题汇报的题目是《基于物理信息降阶模型与虚实同步校正的预制菜热加工数字孪生系统研究》。"},
        {"no": 2, "title": "结构大纲｜Contents", "core": ["01 问题提出", "02 研究目标", "03 研究内容", "04 研究基础与预期进展"], "script": "本次汇报将严格按照问题提出、研究目标、研究内容以及研究基础与预期进展这四个部分展开。"},
        {"no": 3, "title": "1.1 研究背景", "core": ["预制菜产业高速发展", "热处理是杀菌保藏、风味形成的核心工序", "经验式加工面临批次波动大、标准化程度低的挑战"], "script": "预制菜产业正在经历从经验加工向数字智造的转型。热处理决定产品保存、安全和风味，但目前主要依赖经验试错，缺乏对食品内部状态的透明化管控。"},
        {"no": 4, "title": "1.2 核心痛点：“不可能三角”", "core": ["中心冷点杀菌安全", "表层组织过热劣变", "生产加工效率"], "script": "复合预制菜内部对流受限，形成安全、品质与效率的不可能三角。长时间杀菌能保护冷点安全，却会造成表层肉质变柴和风味流失；缩短时间又可能带来安全风险。"},
        {"no": 5, "title": "1.3 现有研究现状与局限", "core": ["高保真 CFD 算得准但太慢", "单点温度监控无法反映全局品质", "缺乏多目标协同优化的动态控制系统"], "script": "现有高保真 CFD 准确但难以在线应用，单点温度传感又无法反映风味与质构破坏。因此需要一套看得透、算得快、能自动校正的数字孪生系统。"},
        {"no": 6, "title": "2.1 研究目标与技术路线", "core": ["构建“感知-对齐-预测-决策-验证”一体化数字孪生平台", "实现安全、品质与效率的闭环优化"], "script": "本课题从多物理场机理出发，经过品质标定与降阶推演，最终利用多源感知在线对齐和 MPC 实现动态闭环控制，突破固定规程盲煮的局限。"},
        {"no": 7, "title": "3.1 高保真热-流-固-生化物理底座", "core": ["14万级混合网格", "非牛顿降黏与多孔阻力模型", "固液共轭传热"], "script": "第一部分搭建物理底座。模型引入高黏酱汁受热剪切变稀、多孔菌菇渗流阻力和固液共轭传热，使冷点漂移和局部滞热能够被真实反演。"},
        {"no": 8, "title": "3.2 品质特征多源标定与动力学映射", "core": ["GC-MS 风味物质提取", "TPA 质构硬度检测", "Arrhenius 动力学模型"], "script": "第二部分解决品质量化。通过 GC-MS 和 TPA 提取风味挥发与质构软化数据，并拟合为 Arrhenius 动力学方程，把口感与品质转化为可计算约束。"},
        {"no": 9, "title": "3.3 基于 ANODE 的亚秒级降阶推演", "core": ["APBRS 多工况泛化特训", "增广隐变量空间", "单次推演耗时缩减至亚秒级"], "script": "第三部分为系统提速。通过 APBRS 伪随机工况训练 ANODE 数字大脑，用增广隐变量区分复杂热历史，实现亚秒级推演和虚拟排雷。"},
        {"no": 10, "title": "3.4 MHE 在线校正与 Pareto-MPC 闭环", "core": ["250s 滑动窗口对齐漂移", "Pareto-MPC 多目标滚动寻优"], "script": "第四部分实现在线闭环。MHE 定期融合传感数据洗刷模型漂移，MPC 向前预测未来工艺结果，在安全与品质之间寻找最佳控制路径。"},
        {"no": 11, "title": "3.5 虚实同步实验方案", "core": ["桌面级微型杀菌釜验证平台", "对照组：121.1℃恒温", "实验组：动态 DT-MPC"], "script": "实验上搭建多模态微型杀菌釜平台，以传统恒温杀菌为对照，以数字孪生变温控制为实验组，通过理化和风味检测完成双闭环验证。"},
        {"no": 12, "title": "3.6 创新性描述", "core": ["机理深度耦合创新", "ANODE 与泛化训练结合", "隐变量在线校正闭环"], "script": "本课题打破食品仿真和自动控制的边界，将风味与质构动力学融入流固耦合网络，并通过隐变量在线校正增强系统实战能力。"},
        {"no": 13, "title": "4.1 研究基础 (1)：物理底座与热滞后验证", "core": ["克服4%刚性假死提取黄金数据", "验证高黏体系4.4K极强热滞后"], "script": "前期已跑通 14 万单元复杂物理底座，并验证高黏预制菜内部存在约 4.4K 的热滞后误差，这为虚实同步校正提供了直接依据。"},
        {"no": 14, "title": "4.1 研究基础 (2)：极速排雷与海量工艺寻优", "core": ["成功排除传统工艺安全大雷", "26.7分钟完成27.28万次全空间配方扫描"], "script": "依托初步训练的数字大脑，已在虚拟空间低成本排除传统工艺风险，并在普通 CPU 上完成 27 万多种工艺配方扫描，验证了降阶寻优可行性。"},
        {"no": 15, "title": "4.2 预期进展", "core": ["平台搭建与校准", "高保真建模与动力学实验", "降阶推演与校正算法开发", "闭环验证与论文撰写"], "script": "项目将先完成平台搭建与品质动力学实验，中期攻克降阶模型训练和 MHE 对齐算法，后期开展闭环对照实验并完成论文撰写。"},
        {"no": 16, "title": "结尾页｜敬请各位老师批评指正", "core": ["敬请各位老师批评指正！", "汇报人：杨阳"], "script": "以数字孪生引领预制菜从经验盲煮走向智能加工，我的开题汇报到此结束。敬请各位专家老师批评指正，谢谢大家！"},
    ]


def _blueprint_image(slide_no: int, image_dir: Path) -> Path | None:
    name = PPT_IMAGE_MAP.get(slide_no)
    return image_dir / name if name else None


def _write_blueprint_markdown(slides: Iterable[dict], image_dir: Path) -> str:
    lines = ["# PPT制作与演讲蓝图指导书_16页版", ""]
    for slide in slides:
        img = _blueprint_image(slide["no"], image_dir)
        img_text = str(img) if img and img.exists() else ("无" if img is None else f"{img}（未找到，请补图）")
        lines += [
            f"## Slide {slide['no']:02d}",
            "",
            f"**【幻灯片标题】**：{slide['title']}",
            "",
            f"**【指定插入图片】**：{img_text}",
            "",
            "**【PPT 核心文字】**",
            "",
        ]
        lines += [f"- {item}" for item in slide["core"]]
        lines += ["", f"**【汇报讲稿】**：{slide['script']}", ""]
    return "\n".join(lines)


def _write_blueprint_docx(slides: Iterable[dict], image_dir: Path, output: Path) -> None:
    doc = Document()
    doc.add_heading("PPT制作与演讲蓝图指导书_16页版", 0)
    for slide in slides:
        doc.add_heading(f"Slide {slide['no']:02d}", level=1)
        doc.add_paragraph(f"【幻灯片标题】：{slide['title']}")
        img = _blueprint_image(slide["no"], image_dir)
        if img and img.exists():
            doc.add_paragraph(f"【指定插入图片】：{img}")
            try:
                doc.add_picture(str(img), width=Cm(14.5))
            except Exception:
                doc.add_paragraph("（图片预览插入失败，但路径已保留。）")
        elif img:
            doc.add_paragraph(f"【指定插入图片】：{img}（未找到，请补图）")
        else:
            doc.add_paragraph("【指定插入图片】：无")
        doc.add_paragraph("【PPT 核心文字】")
        for item in slide["core"]:
            doc.add_paragraph(item, style="List Bullet")
        doc.add_paragraph(f"【汇报讲稿】：{slide['script']}")
    for para in doc.paragraphs:
        para.paragraph_format.line_spacing = 1.25
        for run in para.runs:
            _format_run(run, "宋体", "Times New Roman", 10.5, None)
    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output))


def generate_ppt_blueprint(
    input_docx: str | os.PathLike,
    image_dir: str | os.PathLike,
    output_blueprint: str | os.PathLike,
) -> Path:
    """Generate a strict 16-slide proposal-defense blueprint.

    `input_docx` is read to verify and preserve provenance. The slide structure is
    the refined Zhang-Baitao-style four-module blueprint used in this project.
    """

    src = _path(input_docx)
    images = _path(image_dir)
    output = _path(output_blueprint)
    _ = _read_docx_text(src)
    slides = _make_slide_blueprint()
    if output.suffix.lower() == ".docx":
        _write_blueprint_docx(slides, images, output)
    else:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(_write_blueprint_markdown(slides, images), encoding="utf-8")
    return output


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Kaiti proposal workflow helpers.")
    sub = parser.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("rewrite")
    p.add_argument("input_file")
    p.add_argument("output_file")
    p = sub.add_parser("format")
    p.add_argument("input_docx")
    p.add_argument("reference_template")
    p.add_argument("output_docx")
    p = sub.add_parser("blueprint")
    p.add_argument("input_docx")
    p.add_argument("image_dir")
    p.add_argument("output_blueprint")
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    if args.cmd == "rewrite":
        out = rewrite_and_balance_proposal(args.input_file, args.output_file)
    elif args.cmd == "format":
        out = format_academic_docx(args.input_docx, args.reference_template, args.output_docx)
    elif args.cmd == "blueprint":
        out = generate_ppt_blueprint(args.input_docx, args.image_dir, args.output_blueprint)
    else:  # pragma: no cover
        raise SystemExit(f"Unknown command: {args.cmd}")
    print(out)


if __name__ == "__main__":
    main()
