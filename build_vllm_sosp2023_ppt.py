from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.util import Inches, Pt


OUT_DIR = Path("generated")
OUT_FILE = OUT_DIR / "vllm_sosp2023_pagedattention_analysis.pptx"

FONT = "Microsoft YaHei"
FONT_FALLBACK = "Noto Sans CJK SC"

NAVY = RGBColor(23, 43, 77)
BLUE = RGBColor(37, 99, 235)
BLUE_DARK = RGBColor(30, 64, 175)
BLUE_LIGHT = RGBColor(219, 234, 254)
TEAL = RGBColor(13, 148, 136)
TEAL_LIGHT = RGBColor(204, 251, 241)
ORANGE = RGBColor(245, 158, 11)
ORANGE_LIGHT = RGBColor(254, 243, 199)
RED = RGBColor(220, 38, 38)
RED_LIGHT = RGBColor(254, 226, 226)
GREEN = RGBColor(22, 163, 74)
GREEN_LIGHT = RGBColor(220, 252, 231)
GRAY_900 = RGBColor(17, 24, 39)
GRAY_700 = RGBColor(55, 65, 81)
GRAY_500 = RGBColor(107, 114, 128)
GRAY_300 = RGBColor(209, 213, 219)
GRAY_200 = RGBColor(229, 231, 235)
GRAY_100 = RGBColor(243, 244, 246)
WHITE = RGBColor(255, 255, 255)


def emu(value):
    return Inches(value)


def set_font(run, size=16, bold=False, color=NAVY):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_text(slide, text, x, y, w, h, size=16, bold=False, color=NAVY,
             align=PP_ALIGN.LEFT, line_spacing=1.0):
    box = slide.shapes.add_textbox(emu(x), emu(y), emu(w), emu(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    set_font(run, size=size, bold=bold, color=color)
    return box


def add_bullets(slide, bullets, x, y, w, h, size=15, color=NAVY):
    box = slide.shapes.add_textbox(emu(x), emu(y), emu(w), emu(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    for idx, item in enumerate(bullets):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = ""
        p.level = 0
        p.margin_left = emu(0.16)
        p.space_after = Pt(6)
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = f"• {item}"
        set_font(run, size=size, bold=False, color=color)
    return box


def add_footer(slide, page):
    add_text(
        slide,
        f"{page}/8  |  Reference: Kwon et al., Efficient Memory Management for LLM Serving with PagedAttention, SOSP 2023",
        0.55,
        7.12,
        12.25,
        0.22,
        size=8.5,
        color=GRAY_500,
    )


def add_slide_title(slide, title, kicker, page):
    add_text(slide, kicker, 0.62, 0.30, 3.2, 0.28, size=10.5, bold=True, color=BLUE)
    add_text(slide, title, 0.60, 0.62, 9.8, 0.48, size=25, bold=True, color=NAVY)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, emu(0.62), emu(1.14), emu(1.0), emu(0.045))
    line.fill.solid()
    line.fill.fore_color.rgb = BLUE
    line.line.fill.background()
    add_footer(slide, page)


def add_rect(slide, x, y, w, h, text="", fill=WHITE, line=GRAY_300, radius=True,
             font_size=14, font_color=NAVY, bold=False, align=PP_ALIGN.CENTER):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, emu(x), emu(y), emu(w), emu(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line
    shape.line.width = Pt(1)
    if text:
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = emu(0.08)
        tf.margin_right = emu(0.08)
        tf.margin_top = emu(0.04)
        tf.margin_bottom = emu(0.04)
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        set_font(run, size=font_size, bold=bold, color=font_color)
    return shape


def add_line(slide, x1, y1, x2, y2, color=GRAY_500, width=1.5):
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, emu(x1), emu(y1), emu(x2), emu(y2))
    line.line.color.rgb = color
    line.line.width = Pt(width)
    return line


def add_axis_label(slide, text, x, y, w, h, size=9, color=GRAY_700, align=PP_ALIGN.CENTER):
    return add_text(slide, text, x, y, w, h, size=size, color=color, align=align)


def draw_legend(slide, items, x, y, size=9.5):
    offset = 0
    for label, color in items:
        add_rect(slide, x + offset, y + 0.04, 0.12, 0.12, fill=color, line=color, radius=False)
        add_text(slide, label, x + offset + 0.16, y, 1.45, 0.22, size=size, color=GRAY_700)
        offset += 1.55


def draw_stacked_memory_chart(slide, x, y, w, h):
    labels = ["传统连续分配", "vLLM / PagedAttention"]
    stacks = [
        [("有效 KV Cache", 25, BLUE), ("Internal fragmentation", 25, ORANGE),
         ("External fragmentation", 20, RED), ("Reservation", 30, GRAY_500)],
        [("有效 KV Cache", 96, TEAL), ("Internal fragmentation", 2, ORANGE),
         ("External fragmentation", 1, RED), ("Reservation", 1, GRAY_500)],
    ]

    add_line(slide, x + 0.55, y + h - 0.42, x + w - 0.15, y + h - 0.42, color=GRAY_300, width=1)
    add_line(slide, x + 0.55, y + 0.20, x + 0.55, y + h - 0.42, color=GRAY_300, width=1)
    for pct in [0, 25, 50, 75, 100]:
        yy = y + h - 0.42 - pct / 100 * (h - 0.72)
        add_line(slide, x + 0.50, yy, x + w - 0.15, yy, color=GRAY_200, width=0.6)
        add_axis_label(slide, f"{pct}%", x + 0.02, yy - 0.08, 0.42, 0.18, size=8, align=PP_ALIGN.RIGHT)

    bar_w = 1.15
    base_y = y + h - 0.42
    chart_h = h - 0.72
    for i, stack in enumerate(stacks):
        bx = x + 1.45 + i * 2.25
        current_y = base_y
        for name, pct, color in stack:
            seg_h = pct / 100 * chart_h
            current_y -= seg_h
            add_rect(slide, bx, current_y, bar_w, seg_h, fill=color, line=WHITE, radius=False)
            if pct >= 12:
                add_axis_label(slide, f"{pct}%", bx, current_y + seg_h / 2 - 0.08, bar_w, 0.18,
                               size=9, color=WHITE)
        add_axis_label(slide, labels[i], bx - 0.25, y + h - 0.22, bar_w + 0.5, 0.22, size=10, color=NAVY)

    draw_legend(slide, [(name, color) for name, _, color in stacks[0]], x + 5.0, y + 0.35, size=9)
    add_text(slide, "论文结论：传统 KV Cache 管理因碎片化与预留策略浪费 60%-80%；vLLM 仅末块可能浪费，实测 <4%。",
             x + 5.0, y + 1.25, 3.9, 0.76, size=13, color=NAVY)


def draw_grouped_throughput_chart(slide, x, y, w, h):
    models = ["Llama-7B", "OPT-13B"]
    systems = ["HF", "FT", "vLLM"]
    colors = [GRAY_500, ORANGE, BLUE]
    data = {
        "Llama-7B": [10, 100, 260],
        "OPT-13B": [8, 70, 190],
    }
    ymax = 280
    left = x + 0.72
    bottom = y + h - 0.52
    chart_w = w - 1.05
    chart_h = h - 0.92
    add_line(slide, left, bottom, left + chart_w, bottom, color=GRAY_300, width=1)
    add_line(slide, left, y + 0.20, left, bottom, color=GRAY_300, width=1)
    for val in [0, 70, 140, 210, 280]:
        yy = bottom - val / ymax * chart_h
        add_line(slide, left, yy, left + chart_w, yy, color=GRAY_200, width=0.6)
        add_axis_label(slide, str(val), x + 0.18, yy - 0.08, 0.42, 0.18, size=8, align=PP_ALIGN.RIGHT)

    group_gap = chart_w / len(models)
    bar_w = 0.28
    for i, model in enumerate(models):
        gx = left + i * group_gap + 0.52
        values = data[model]
        for j, value in enumerate(values):
            bh = value / ymax * chart_h
            bx = gx + j * (bar_w + 0.08)
            by = bottom - bh
            add_rect(slide, bx, by, bar_w, bh, fill=colors[j], line=WHITE, radius=False)
            add_axis_label(slide, str(value), bx - 0.04, by - 0.20, bar_w + 0.08, 0.18, size=8, color=GRAY_700)
        add_axis_label(slide, model, gx - 0.18, bottom + 0.12, 1.05, 0.2, size=9.5, color=NAVY)

    draw_legend(slide, list(zip(systems, colors)), x + w - 4.6, y + 0.05, size=9.5)
    add_axis_label(slide, "Throughput (req/min, 论文倍率校准示意)", x + 1.2, y + h - 0.05, 4.0, 0.2, size=9, color=GRAY_700)
    add_text(slide, "vLLM vs HF\nLlama-7B: 26x\nOPT-13B: 23.8x", x + w - 2.35, y + 0.70, 1.55, 0.70,
             size=12, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    add_text(slide, "vLLM vs FT\n2.6x / 2.7x", x + w - 2.15, y + 1.50, 1.15, 0.45,
             size=12, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)


def draw_waste_bar_chart(slide, x, y, w, h):
    entries = [
        ("传统框架\nHF / FT 风格连续管理", 70, RED, "60%-80%"),
        ("vLLM\nPagedAttention", 4, GREEN, "<4%"),
    ]
    max_value = 80
    label_w = 2.2
    bar_x = x + label_w + 0.18
    bar_w = w - label_w - 0.55
    for idx, (label, value, color, callout) in enumerate(entries):
        yy = y + 0.55 + idx * 1.25
        add_text(slide, label, x, yy - 0.05, label_w, 0.45, size=12, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
        add_rect(slide, bar_x, yy, bar_w, 0.34, fill=GRAY_100, line=GRAY_200, radius=True)
        fill_w = max(0.08, value / max_value * bar_w)
        add_rect(slide, bar_x, yy, fill_w, 0.34, fill=color, line=color, radius=True)
        add_text(slide, callout, bar_x + fill_w + 0.12, yy + 0.02, 0.95, 0.22, size=12, bold=True, color=color)
    add_line(slide, bar_x, y + h - 0.48, bar_x + bar_w, y + h - 0.48, color=GRAY_300, width=1)
    for pct in [0, 20, 40, 60, 80]:
        xx = bar_x + pct / max_value * bar_w
        add_line(slide, xx, y + h - 0.55, xx, y + h - 0.42, color=GRAY_300, width=0.8)
        add_axis_label(slide, f"{pct}%", xx - 0.18, y + h - 0.36, 0.36, 0.16, size=8)
    add_text(slide, "关键原因：PagedAttention 以固定大小 block 按需分配 KV Cache，避免为未知输出长度做大块连续预留。",
             x + 0.12, y + 3.00, w - 0.24, 0.50, size=13, color=NAVY)


def draw_line_chart(slide, x, y, w, h):
    seq = [512, 1024, 2048, 4096, 8192]
    series = {
        "HF": ([55, 100, 210, 500, 1100], GRAY_500),
        "FT": ([35, 62, 130, 310, 720], ORANGE),
        "vLLM": ([28, 48, 90, 180, 330], BLUE),
    }
    ymax = 1200
    left = x + 0.75
    top = y + 0.18
    bottom = y + h - 0.58
    right = x + w - 0.25
    chart_h = bottom - top
    chart_w = right - left
    add_line(slide, left, bottom, right, bottom, color=GRAY_300, width=1)
    add_line(slide, left, top, left, bottom, color=GRAY_300, width=1)
    for val in [0, 300, 600, 900, 1200]:
        yy = bottom - val / ymax * chart_h
        add_line(slide, left, yy, right, yy, color=GRAY_200, width=0.6)
        add_axis_label(slide, str(val), x + 0.20, yy - 0.08, 0.42, 0.18, size=8, align=PP_ALIGN.RIGHT)
    for i, length in enumerate(seq):
        xx = left + i / (len(seq) - 1) * chart_w
        add_line(slide, xx, bottom, xx, bottom + 0.07, color=GRAY_300, width=0.8)
        add_axis_label(slide, str(length), xx - 0.25, bottom + 0.13, 0.50, 0.17, size=8)

    for name, (values, color) in series.items():
        points = []
        for i, value in enumerate(values):
            xx = left + i / (len(seq) - 1) * chart_w
            yy = bottom - value / ymax * chart_h
            points.append((xx, yy))
            marker = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(xx - 0.045), emu(yy - 0.045), emu(0.09), emu(0.09))
            marker.fill.solid()
            marker.fill.fore_color.rgb = color
            marker.line.color.rgb = WHITE
        for (x1, y1), (x2, y2) in zip(points, points[1:]):
            add_line(slide, x1, y1, x2, y2, color=color, width=2.0)

    draw_legend(slide, [(name, color) for name, (_, color) in series.items()], x + w - 4.4, y + 0.05, size=9)
    add_axis_label(slide, "Sequence Length (tokens)", x + 2.6, y + h - 0.04, 2.2, 0.18, size=8.5)
    add_axis_label(slide, "Latency (ms, 趋势示意)", x + 0.02, y + 0.05, 1.1, 0.18, size=8.5)


def add_comparison_table(slide, x, y, w, h):
    rows = [
        ("维度", "传统服务栈", "vLLM / PagedAttention"),
        ("吞吐量", "HF 基线；FT 提升有限", "HF 的 22-28x；FT 的 2-3x"),
        ("KV Cache 内存", "连续预留，浪费 60%-80%", "块式按需分配，浪费 <4%"),
        ("长序列表现", "Batch Size 受内存碎片限制", "更平缓的时延增长曲线"),
        ("应用价值", "高并发成本高、GPU 利用率低", "更高 GPU 利用率与更低单位请求成本"),
    ]
    table_shape = slide.shapes.add_table(len(rows), 3, emu(x), emu(y), emu(w), emu(h))
    table = table_shape.table
    table.columns[0].width = emu(1.35)
    table.columns[1].width = emu(3.05)
    table.columns[2].width = emu(3.60)
    for r_idx, row in enumerate(rows):
        for c_idx, value in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.text = value
            cell.margin_left = emu(0.06)
            cell.margin_right = emu(0.06)
            cell.margin_top = emu(0.03)
            cell.margin_bottom = emu(0.03)
            fill = BLUE if r_idx == 0 else (BLUE_LIGHT if c_idx == 2 else WHITE)
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER if r_idx == 0 else PP_ALIGN.LEFT
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    set_font(run, size=10.2 if r_idx else 10.5, bold=(r_idx == 0), color=WHITE if r_idx == 0 else NAVY)


def build_slide_1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_text(slide, "vLLM 推理流程及其性能增益可视化分析", 0.68, 0.72, 9.8, 0.65,
             size=28, bold=True, color=NAVY)
    add_text(slide, "基于 PagedAttention 的高吞吐量 LLM 服务系统", 0.72, 1.48, 8.2, 0.36,
             size=17, color=BLUE_DARK)
    add_text(slide, "Reference: Kwon et al., SOSP 2023", 0.76, 6.84, 4.8, 0.25,
             size=10, color=GRAY_500)

    boxes = [
        ("用户请求\nPrompt + sampling", 0.95, BLUE_LIGHT, BLUE),
        ("vLLM Engine\nScheduler", 3.35, TEAL_LIGHT, TEAL),
        ("PagedAttention\nKV Cache paging", 5.75, ORANGE_LIGHT, ORANGE),
        ("A100 GPU\nHigh utilization", 8.15, GREEN_LIGHT, GREEN),
    ]
    for text, x, fill, line in boxes:
        add_rect(slide, x, 3.20, 1.75, 0.82, text, fill=fill, line=line, font_size=13, bold=True)
    for x in [2.74, 5.14, 7.54]:
        add_line(slide, x, 3.61, x + 0.45, 3.61, color=GRAY_700, width=2)
        tri = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, emu(x + 0.38), emu(3.52), emu(0.16), emu(0.18))
        tri.fill.solid()
        tri.fill.fore_color.rgb = GRAY_700
        tri.line.fill.background()
    add_rect(slide, 9.98, 3.08, 2.18, 1.05, "输出 Tokens\n22-28x vs HF\n2-3x vs FT", fill=BLUE, line=BLUE,
             font_size=13, font_color=WHITE, bold=True)
    add_text(slide, "6 分钟学术演示 · 8 页 · 数据基于官方论文/博客披露结论", 0.78, 2.05, 7.3, 0.25,
             size=10.5, color=GRAY_700)
    add_footer(slide, 1)


def build_slide_2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "背景：LLM 推理中的内存瓶颈", "KV Cache limits batch size", 2)
    add_bullets(
        slide,
        [
            "自回归解码中，每生成一个 token 都会追加 Key/Value 张量。",
            "KV Cache 随序列长度与并发请求线性增长，迅速占满 GPU memory。",
            "传统连续分配需要预估最大输出长度，碎片化与预留空间压缩可用 Batch Size。",
        ],
        0.72,
        1.45,
        4.0,
        1.25,
        size=13,
    )
    draw_stacked_memory_chart(slide, 0.75, 2.72, 11.55, 3.45)


def build_slide_3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "核心机制：PagedAttention 原理", "Logical blocks -> Physical blocks", 3)
    add_text(slide, "思想类比操作系统虚拟内存：请求看到连续 logical blocks；GPU 中实际 KV Cache blocks 可非连续存放。",
             0.74, 1.32, 9.9, 0.32, size=13, color=GRAY_700)

    add_text(slide, "Sequence A Logical Blocks", 0.82, 1.96, 2.8, 0.22, size=11, bold=True, color=NAVY)
    for idx in range(4):
        add_rect(slide, 0.90 + idx * 0.78, 2.28, 0.55, 0.42, f"L{idx}", fill=BLUE_LIGHT, line=BLUE, font_size=12, bold=True)

    add_text(slide, "Block Table / Page Table", 4.42, 1.82, 2.7, 0.22, size=11, bold=True, color=NAVY)
    mappings = [("L0", "P7"), ("L1", "P2"), ("L2", "P9"), ("L3", "P4")]
    for i, (logical, physical) in enumerate(mappings):
        add_rect(slide, 4.45, 2.18 + i * 0.46, 0.68, 0.32, logical, fill=WHITE, line=GRAY_300, font_size=10)
        add_rect(slide, 5.13, 2.18 + i * 0.46, 0.82, 0.32, physical, fill=TEAL_LIGHT, line=TEAL, font_size=10, bold=True)

    add_text(slide, "Physical KV Block Pool on GPU", 7.72, 1.96, 3.2, 0.22, size=11, bold=True, color=NAVY)
    physical = [
        ("P0", GRAY_100), ("P1", GRAY_100), ("P2", TEAL_LIGHT), ("P3", GRAY_100),
        ("P4", TEAL_LIGHT), ("P5", GRAY_100), ("P6", GRAY_100), ("P7", TEAL_LIGHT),
        ("P8", GRAY_100), ("P9", TEAL_LIGHT), ("P10", GRAY_100), ("P11", GRAY_100),
    ]
    for i, (label, fill) in enumerate(physical):
        px = 7.78 + (i % 4) * 0.72
        py = 2.28 + (i // 4) * 0.56
        add_rect(slide, px, py, 0.52, 0.38, label, fill=fill, line=TEAL if fill == TEAL_LIGHT else GRAY_300, font_size=9.5)

    for x1, y1, x2, y2 in [(3.98, 2.49, 4.35, 2.34), (3.98, 2.49, 4.35, 2.80), (3.98, 2.49, 4.35, 3.26),
                           (6.05, 2.34, 7.62, 3.26), (6.05, 2.80, 7.62, 2.70), (6.05, 3.26, 7.62, 3.82)]:
        add_line(slide, x1, y1, x2, y2, color=GRAY_500, width=1.2)

    add_rect(slide, 1.05, 5.05, 3.0, 0.55, "按需分配\n只为已生成 tokens 申请 block", fill=BLUE_LIGHT, line=BLUE, font_size=12, bold=True)
    add_rect(slide, 4.85, 5.05, 3.0, 0.55, "非连续存储\nPage table 维护映射", fill=TEAL_LIGHT, line=TEAL, font_size=12, bold=True)
    add_rect(slide, 8.65, 5.05, 3.0, 0.55, "末块浪费\n实测 memory waste <4%", fill=GREEN_LIGHT, line=GREEN, font_size=12, bold=True)


def build_slide_4(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "vLLM 推理全流程解析", "Scheduler + Block Manager + Kernels", 4)
    add_text(slide, "连续 batching 与 PagedAttention 共同作用：调度层扩大有效 Batch，内存层保持 KV Cache 紧凑。",
             0.74, 1.32, 9.4, 0.30, size=13, color=GRAY_700)

    steps = [
        ("1. Request Queue\n输入请求进入等待队列", BLUE_LIGHT, BLUE),
        ("2. Scheduler\n选择可运行序列", TEAL_LIGHT, TEAL),
        ("3. Block Manager\n申请/释放 physical blocks", ORANGE_LIGHT, ORANGE),
        ("4. PagedAttention Kernel\n按 block table 读取 KV", GREEN_LIGHT, GREEN),
        ("5. Token Output\n生成 token 并回写状态", BLUE_LIGHT, BLUE),
    ]
    x0 = 0.72
    for i, (text, fill, line) in enumerate(steps):
        x = x0 + i * 2.38
        add_rect(slide, x, 2.48, 1.75, 0.78, text, fill=fill, line=line, font_size=10.8, bold=True)
        if i < len(steps) - 1:
            add_line(slide, x + 1.75, 2.87, x + 2.26, 2.87, color=GRAY_700, width=1.8)
            tri = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, emu(x + 2.18), emu(2.79), emu(0.14), emu(0.16))
            tri.fill.solid()
            tri.fill.fore_color.rgb = GRAY_700
            tri.line.fill.background()

    add_line(slide, 10.45, 3.38, 10.45, 4.40, color=GRAY_500, width=1.2)
    add_line(slide, 10.45, 4.40, 1.58, 4.40, color=GRAY_500, width=1.2)
    add_line(slide, 1.58, 4.40, 1.58, 3.37, color=GRAY_500, width=1.2)
    add_text(slide, "decode loop: 新 token 触发下一轮调度与 block 映射更新", 3.45, 4.48, 4.8, 0.26,
             size=11, color=GRAY_700, align=PP_ALIGN.CENTER)

    add_rect(slide, 1.02, 5.18, 3.05, 0.62, "吞吐来源 1\nContinuous batching 提高 GPU occupancy", fill=WHITE, line=GRAY_300, font_size=12)
    add_rect(slide, 4.95, 5.18, 3.05, 0.62, "吞吐来源 2\nPaged KV Cache 提高可容纳并发", fill=WHITE, line=GRAY_300, font_size=12)
    add_rect(slide, 8.88, 5.18, 3.05, 0.62, "吞吐来源 3\nKernel 直接消费 block table", fill=WHITE, line=GRAY_300, font_size=12)


def build_slide_5(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "性能分析：吞吐量增益（官方核心结论）", "Throughput vs HF / FT", 5)
    add_text(slide, "柱状图采用官方披露倍率校准的 req/min 示意：vLLM 相比 HF 约 22-28x，相比 FT 约 2-3x。",
             0.74, 1.32, 9.95, 0.30, size=13, color=GRAY_700)
    draw_grouped_throughput_chart(slide, 0.82, 1.86, 10.95, 4.15)
    add_rect(slide, 9.72, 5.78, 2.18, 0.38, "A100 80GB · Llama/OPT", fill=GRAY_100, line=GRAY_300,
             font_size=10.5, font_color=GRAY_700, bold=True)


def build_slide_6(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "性能分析：内存浪费率对比", "Memory waste ratio", 6)
    add_text(slide, "vLLM 的核心收益不是减少模型权重，而是将动态 KV Cache 管理从连续预留改为分页式按需分配。",
             0.74, 1.32, 10.2, 0.30, size=13, color=GRAY_700)
    draw_waste_bar_chart(slide, 1.05, 2.03, 10.55, 3.75)
    add_rect(slide, 1.35, 5.93, 2.65, 0.42, "结果：更大 Batch Size", fill=GREEN_LIGHT, line=GREEN, font_size=12, bold=True)
    add_rect(slide, 5.30, 5.93, 2.65, 0.42, "结果：更高 GPU 利用率", fill=BLUE_LIGHT, line=BLUE, font_size=12, bold=True)
    add_rect(slide, 9.25, 5.93, 2.65, 0.42, "结果：更低服务成本", fill=ORANGE_LIGHT, line=ORANGE, font_size=12, bold=True)


def build_slide_7(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "性能分析：不同序列长度下的表现", "Long sequence stability", 7)
    add_text(slide, "论文指出：序列更长、模型更大、解码更复杂时，内存管理优势更明显；下图展示长文本下的时延增长趋势。",
             0.74, 1.32, 10.4, 0.30, size=13, color=GRAY_700)
    draw_line_chart(slide, 0.82, 1.86, 10.95, 4.35)
    add_rect(slide, 9.65, 5.92, 2.25, 0.38, "vLLM 曲线更平缓", fill=BLUE, line=BLUE, font_size=11.5,
             font_color=WHITE, bold=True)


def build_slide_8(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_title(slide, "结论与应用价值", "Why vLLM changes serving economics", 8)
    add_text(slide, "PagedAttention 将 LLM Serving 的瓶颈从“KV Cache 内存不可控”转化为可调度、可复用、可分页的系统问题。",
             0.74, 1.32, 10.3, 0.30, size=13, color=GRAY_700)

    cards = [
        ("High Throughput", "22-28x vs HF\n2-3x vs FT", BLUE_LIGHT, BLUE),
        ("Low Latency", "长序列下时延增长更平缓", GREEN_LIGHT, GREEN),
        ("Memory Efficient", "KV Cache waste <4%", ORANGE_LIGHT, ORANGE),
    ]
    for i, (title, body, fill, line) in enumerate(cards):
        x = 1.00 + i * 3.95
        add_rect(slide, x, 2.02, 2.95, 1.08, "", fill=fill, line=line)
        icon = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(x + 0.18), emu(2.26), emu(0.44), emu(0.44))
        icon.fill.solid()
        icon.fill.fore_color.rgb = line
        icon.line.fill.background()
        add_text(slide, title, x + 0.74, 2.18, 1.95, 0.22, size=12.5, bold=True, color=NAVY)
        add_text(slide, body, x + 0.74, 2.48, 1.95, 0.35, size=10.5, color=GRAY_700)

    add_comparison_table(slide, 0.92, 3.58, 11.25, 2.45)
    add_text(slide, "一句话总结：vLLM 通过 PagedAttention 把“GPU 内存碎片问题”转化为“分页映射问题”，从而提升吞吐并降低单位推理成本。",
             0.92, 6.23, 10.85, 0.30, size=12.5, bold=True, color=BLUE_DARK)


def main():
    prs = Presentation()
    prs.slide_width = emu(13.333)
    prs.slide_height = emu(7.5)

    for builder in [
        build_slide_1,
        build_slide_2,
        build_slide_3,
        build_slide_4,
        build_slide_5,
        build_slide_6,
        build_slide_7,
        build_slide_8,
    ]:
        builder(prs)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    prs.save(OUT_FILE)
    print(f"Generated {OUT_FILE} with {len(prs.slides)} slides")


if __name__ == "__main__":
    main()
