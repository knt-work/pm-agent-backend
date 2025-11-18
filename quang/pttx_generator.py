from __future__ import annotations
import os
from datetime import datetime
import re
import json

try:
    from dateutil import parser as dateparser
except Exception:
    dateparser = None

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.dml.color import RGBColor

from pptx.chart.data import ChartData, CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

WIDESCREEN_16x9 = True
DEFAULT_FONT_NAME = "Calibri"
DEFAULT_TEXT_COLOR = "#1B1B1B"
DEFAULT_BG = "#FFFFFF"


def parse_time(x):
    if isinstance(x, (int, float)):
        return datetime.fromtimestamp(x)
    if isinstance(x, datetime):
        return x
    if isinstance(x, str):
        m = re.match(r"^\s*(\d{4})\s*[Qq]\s*([1-4])\s*$", x)
        if m:
            year = int(m.group(1))
            q = int(m.group(2))
            month = {1: 1, 2: 4, 3: 7, 4: 10}[q]
            return datetime(year, month, 1)
    if dateparser and isinstance(x, str):
        return dateparser.parse(x)
    return datetime.fromisoformat(x)


def to_rgb(hex_color: str) -> RGBColor:
    s = (hex_color or "").lstrip("#")
    if len(s) == 3:
        s = "".join([c * 2 for c in s])
    if len(s) != 6:
        s = "1B1B1B"
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def script_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def inches_rect(position):
    x = Inches(position.get("x", 1.0))
    y = Inches(position.get("y", 1.0))
    w = Inches(position.get("w", 10.0))
    h = Inches(position.get("h", 1.0))
    return x, y, w, h


def apply_text_style(paragraph, style: dict | None):
    if not style:
        style = {}
    font = getattr(paragraph, "font", None)
    if font is None:
        font = paragraph
    f = style.get("font", {})
    font.name = f.get("name", DEFAULT_FONT_NAME)
    if "size" in f and f["size"]:
        font.size = Pt(f["size"])
    if "bold" in f:
        font.bold = bool(f["bold"])
    if "italic" in f:
        font.italic = bool(f["italic"])
    if "underline" in f:
        font.underline = bool(f["underline"])
    color = style.get("text", DEFAULT_TEXT_COLOR)
    font.color.rgb = to_rgb(color)
    align = style.get("align")
    if align and hasattr(paragraph, "alignment"):
        paragraph.alignment = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
        }.get(align, PP_ALIGN.LEFT)


def _style_chart_legend(chart, position="right"):
    chart.has_legend = True
    pos = {
        "right": XL_LEGEND_POSITION.RIGHT,
        "left": XL_LEGEND_POSITION.LEFT,
        "top": XL_LEGEND_POSITION.TOP,
        "bottom": XL_LEGEND_POSITION.BOTTOM,
        "corner": XL_LEGEND_POSITION.CORNER,
    }.get(position, XL_LEGEND_POSITION.RIGHT)
    chart.legend.position = pos
    chart.legend.include_in_layout = True


def render_textbox(slide, element):
    variant = element.get("variant", "paragraph")
    position = element.get("position", {"x": 1.0, "y": 1.0, "w": 10.0, "h": 1.0})
    style = element.get("style", {})

    x, y, w, h = inches_rect(position)
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.clear()
    tf.word_wrap = True
    try:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass

    margins = style.get("margins", {})
    try:
        if "left" in margins:
            tf.margin_left = Inches(margins["left"])
        if "right" in margins:
            tf.margin_right = Inches(margins["right"])
        if "top" in margins:
            tf.margin_top = Inches(margins["top"])
        if "bottom" in margins:
            tf.margin_bottom = Inches(margins["bottom"])
    except Exception:
        pass

    if variant in ("heading", "paragraph"):
        p = tf.paragraphs[0]
        p.text = element.get("text", "")
        if variant == "heading":
            style = {
                "font": {
                    "size": style.get("font", {}).get("size", 28),
                    "bold": True,
                    "name": style.get("font", {}).get("name", DEFAULT_FONT_NAME),
                },
                "align": style.get("align", "left"),
                "text": style.get("text", DEFAULT_TEXT_COLOR),
                "margins": style.get("margins", {}),
            }
        apply_text_style(p, style)

    elif variant == "bullets":
        items = element.get("items", [])
        for i, item in enumerate(items):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = item.get("text", "")
            p.level = int(item.get("level", 0))
            apply_text_style(p, style)
            p.font.size = Pt(style.get("font", {}).get("size", 16))

    elif variant == "rich":
        runs = element.get("runs", [])
        p = tf.paragraphs[0]
        p.text = ""
        for r in runs:
            run = p.add_run()
            run.text = r.get("text", "")
            rfont = r.get("font", {})
            run.font.name = rfont.get("name", DEFAULT_FONT_NAME)
            if "size" in rfont:
                run.font.size = Pt(rfont["size"])
            if "bold" in r:
                run.font.bold = bool(r["bold"])
            if "italic" in r:
                run.font.italic = bool(r["italic"])
            if "underline" in r:
                run.font.underline = bool(r["underline"])
            run.font.color.rgb = to_rgb(
                r.get("color", style.get("text", DEFAULT_TEXT_COLOR))
            )
        apply_text_style(p, style)


def render_table(slide, element):
    headers = element.get("headers", [])
    rows = element.get("rows", [])
    ncols = max(len(headers), max((len(r) for r in rows), default=0))
    nrows = 1 + len(rows)

    x, y, w, h = inches_rect(
        element.get("position", {"x": 1, "y": 1, "w": 10, "h": 2})
    )
    shape = slide.shapes.add_table(nrows, ncols, x, y, w, h)
    table = shape.table

    col_widths = element.get("column_widths")
    if col_widths:
        for i, width in enumerate(col_widths[:ncols]):
            table.columns[i].width = Inches(width)

    for c in range(ncols):
        cell = table.cell(0, c)
        hv = headers[c] if c < len(headers) else ""
        cell.text = "" if hv is None else str(hv)
        p = cell.text_frame.paragraphs[0]
        apply_text_style(
            p,
            {
                "font": {"size": 12, "bold": True},
                "align": element.get("style", {}).get("align", "left"),
            },
        )

    for r, row in enumerate(rows, start=1):
        for c in range(ncols):
            cell = table.cell(r, c)
            tv = row[c] if c < len(row) else ""
            cell.text = "" if tv is None else str(tv)
            p = cell.text_frame.paragraphs[0]
            apply_text_style(
                p,
                {
                    "font": {"size": 12},
                    "align": element.get("style", {}).get("align", "left"),
                },
            )


def render_chart_pie(slide, element):
    legend_pos = element.get("legend", "right")
    legend_pad = float(element.get("legendPadInches", 0.7))
    x, y, w, h = _padded_chart_frame(
        element.get("position", {"x": 6, "y": 2, "w": 5, "h": 4}),
        legend_pos,
        legend_pad,
    )

    data = element.get("data", [])
    chart_data = ChartData()
    chart_data.categories = [str(d.get("label", "")) for d in data]
    chart_data.add_series(
        element.get("title", ""),
        [(0 if d.get("value") is None else d.get("value")) for d in data],
    )

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, x, y, w, h, chart_data
    ).chart

    if element.get("title"):
        chart.has_title = True
        chart.chart_title.text_frame.text = element["title"]

    _style_chart_legend(chart, legend_pos)

    series = chart.series[0]
    for i, d in enumerate(data):
        col = d.get("color")
        if col:
            pt = series.points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = to_rgb(col)

    try:
        plot = chart.plots[0]
        if element.get("showLabels"):
            plot.has_data_labels = True
            if element.get("labels") == "percent":
                plot.data_labels.number_format = "0%"
                plot.data_labels.show_percentage = True
    except Exception:
        pass


def render_chart_bar(slide, element):
    cats = element.get("x", {}).get("categories", [])
    series = element.get("series", [])
    opts = element.get("options", {})
    stacked = bool(opts.get("stacked", False))
    orientation = opts.get("orientation", "vertical")

    if orientation == "horizontal":
        chart_type = (
            XL_CHART_TYPE.BAR_STACKED if stacked else XL_CHART_TYPE.BAR_CLUSTERED
        )
    else:
        chart_type = (
            XL_CHART_TYPE.COLUMN_STACKED
            if stacked
            else XL_CHART_TYPE.COLUMN_CLUSTERED
        )

    legend_pos = element.get("legend", "right")
    legend_pad = float(element.get("legendPadInches", 0.7))
    x, y, w, h = _padded_chart_frame(
        element.get("position", {"x": 1, "y": 2, "w": 11, "h": 4}),
        legend_pos,
        legend_pad,
    )

    chart_data = CategoryChartData()
    chart_data.categories = ["" if v is None else str(v) for v in cats]
    for s in series:
        vals = [0 if v is None else v for v in s.get("data", [])]
        chart_data.add_series(s.get("name", ""), vals)

    chart = slide.shapes.add_chart(
        chart_type, x, y, w, h, chart_data
    ).chart

    if element.get("title"):
        chart.has_title = True
        chart.chart_title.text_frame.text = element["title"]

    _style_chart_legend(chart, legend_pos)

    for i, s in enumerate(series):
        col = s.get("color")
        if col:
            chart.series[i].format.fill.solid()
            chart.series[i].format.fill.fore_color.rgb = to_rgb(col)


def render_chart_line(slide, element):
    cats = element.get("x", {}).get("categories", [])
    series = element.get("series", [])
    opts = element.get("options", {})
    markers = bool(opts.get("markers", True))
    smooth = bool(opts.get("smooth", False))

    chart_type = XL_CHART_TYPE.LINE_MARKERS if markers else XL_CHART_TYPE.LINE

    legend_pos = element.get("legend", "right")
    legend_pad = float(element.get("legendPadInches", 0.7))
    x, y, w, h = _padded_chart_frame(
        element.get("position", {"x": 1, "y": 2, "w": 11, "h": 4}),
        legend_pos,
        legend_pad,
    )

    chart_data = CategoryChartData()
    chart_data.categories = ["" if v is None else str(v) for v in cats]
    for s in series:
        vals = [0 if v is None else v for v in s.get("data", [])]
        chart_data.add_series(s.get("name", ""), vals)

    chart = slide.shapes.add_chart(
        chart_type, x, y, w, h, chart_data
    ).chart

    if element.get("title"):
        chart.has_title = True
        chart.chart_title.text_frame.text = element["title"]

    _style_chart_legend(chart, legend_pos)

    for i, s in enumerate(series):
        ser = chart.series[i]
        if smooth:
            ser.smooth = True

        col = s.get("color")
        if col:
            ln = ser.format.line
            ln.color.rgb = to_rgb(col)
            try:
                from pptx.util import Pt as _Pt

                ln.width = _Pt(2)
            except Exception:
                pass

            if markers:
                mk = ser.marker
                try:
                    mk.format.fill.solid()
                    mk.format.fill.fore_color.rgb = to_rgb(col)
                except Exception:
                    pass
                try:
                    mk.format.line.color.rgb = to_rgb(col)
                except Exception:
                    pass


def render_chart_gantt(slide, element):
    pos = element.get(
        "position", {"x": 1.0, "y": 1.2, "w": 11.3, "h": 5.0}
    )
    x, y, w, h = inches_rect(pos)

    gutter = float(element.get("gutterInches", 1.0))
    content_x = x + Inches(gutter)
    content_w = w - Inches(gutter)
    if content_w < Inches(1.0):
        content_w = Inches(1.0)

    lanes_raw = element.get("units", {}).get("yRange", {}).get("lanes", [])
    lanes = (
        [_shorten_lane(v) for v in lanes_raw]
        if element.get("shortenLanes", True)
        else lanes_raw
    )

    lane_count = max(len(lanes), 1)
    lane_height = h / lane_count

    t0 = parse_time(element.get("units", {}).get("xRange", {}).get("t0"))
    t1 = parse_time(element.get("units", {}).get("xRange", {}).get("t1"))
    total_days = max((t1 - t0).days, 1)

    def x_pos(dt_str: str):
        dt = parse_time(dt_str)
        d = (dt - t0).days
        return content_x + content_w * (d / total_days)

    def y_pos(lane_index: int):
        return y + lane_height * lane_index

    for i in range(lane_count):
        yy = y_pos(i)
        if i % 2 == 0:
            bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, content_x, yy, content_w, lane_height
            )
            bg.fill.solid()
            bg.fill.fore_color.rgb = to_rgb("#F3F4F6")
            bg.line.fill.background()

        tb = slide.shapes.add_textbox(
            x, yy, Inches(gutter) - Inches(0.1), lane_height
        )
        tf = tb.text_frame
        tf.text = lanes[i] if i < len(lanes) else f"Lane {i+1}"
        p = tf.paragraphs[0]
        apply_text_style(
            p, {"font": {"size": 12, "bold": True}, "align": "right"}
        )

    grid = element.get("grid", {})
    quarter_boundaries = grid.get("quarters", [])
    show_labels = grid.get("showLabels", True)
    label_offset = float(grid.get("labelOffsetInches", 0.30))
    label_font_size = int(grid.get("labelFontSize", 12))
    draw_axis_line = bool(grid.get("topAxisLine", True))

    for d in quarter_boundaries:
        xx = x_pos(d)
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, xx, y, Inches(0.018), h
        )
        line.fill.solid()
        line.fill.fore_color.rgb = to_rgb("#E0E0E0")
        line.line.fill.background()

    if draw_axis_line:
        top_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            content_x,
            y - Inches(0.06),
            content_w,
            Inches(0.012),
        )
        top_line.fill.solid()
        top_line.fill.fore_color.rgb = to_rgb("#D1D5DB")
        top_line.line.fill.background()

    if show_labels and len(quarter_boundaries) >= 2:
        for i in range(len(quarter_boundaries) - 1):
            d_left = quarter_boundaries[i]
            d_right = quarter_boundaries[i + 1]
            xc = (x_pos(d_left) + x_pos(d_right)) / 2.0
            dt = parse_time(d_left)
            label = _quarter_label(dt)

            tb = slide.shapes.add_textbox(
                xc - Inches(0.5),
                y - Inches(label_offset),
                Inches(1.0),
                Inches(0.28),
            )
            tf = tb.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = label
            apply_text_style(
                p,
                {
                    "font": {"size": label_font_size, "bold": False},
                    "align": "center",
                },
            )

    chevron_head = float(element.get("chevronHead", 0.28))

    def draw_chevron(left, top, width, height, fill_hex, text, text_color=DEFAULT_TEXT_COLOR):
        chev = slide.shapes.add_shape(
            MSO_SHAPE.CHEVRON, left, top, width, height
        )
        try:
            chev.adjustments[0] = chevron_head
        except Exception:
            pass
        chev.fill.solid()
        chev.fill.fore_color.rgb = to_rgb(fill_hex)
        chev.line.fill.background()
        tf = chev.text_frame
        tf.text = text
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.font.name = DEFAULT_FONT_NAME
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = to_rgb(text_color)
        tf.margin_left = Inches(0.08)
        tf.margin_top = Inches(0.02)
        return chev

    for it in element.get("items", []):
        left = x_pos(it["start"]["x"])
        right = x_pos(it["end"]["x"])
        if right < left:
            left, right = right, left
        width = right - left

        lane_idx = int(it["start"].get("y", 0))
        height = lane_height * float(it.get("size", {}).get("height", 0.6))
        top = y_pos(lane_idx) + (lane_height - height) / 2.0

        fill = it.get("style", {}).get("fill", "#90CAF9")
        text_color = it.get("style", {}).get("text", DEFAULT_TEXT_COLOR)
        label = it.get("label", it.get("id", ""))
        draw_chevron(left, top, width, height, fill, label, text_color)


def render_element(slide, element):
    etype = element.get("type")
    if etype == "text":
        render_textbox(slide, element)
    elif etype == "table":
        render_table(slide, element)
    elif etype == "chart":
        subtype = element.get("subtype")
        if subtype == "gantt":
            render_chart_gantt(slide, element)
        elif subtype == "pie":
            render_chart_pie(slide, element)
        elif subtype == "bar":
            render_chart_bar(slide, element)
        elif subtype == "line":
            render_chart_line(slide, element)


def add_title_if_any(slide, title_text: str | None):
    if not title_text:
        return
    tb = slide.shapes.add_textbox(
        Inches(1.0),
        Inches(0.35),
        Inches(11.3),
        Inches(0.6),
    )
    tf = tb.text_frame
    tf.text = title_text
    p = tf.paragraphs[0]
    apply_text_style(
        p,
        {
            "font": {"size": 28, "bold": True},
            "align": "left",
            "text": "#FFFFFF",
        },
    )


def render_cover_slide(prs, s):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    band_h = Inches(1.2)
    band = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, band_h
    )
    band.fill.solid()
    band.fill.fore_color.rgb = to_rgb(s.get("accentColor", "#111827"))
    band.line.fill.background()

    title = next(
        (
            el
            for el in s.get("elements", [])
            if el.get("type") == "text"
            and el.get("variant") == "heading"
        ),
        None,
    )
    subtitle = next(
        (
            el
            for el in s.get("elements", [])
            if el.get("type") == "text"
            and el.get("variant") in ("paragraph", "rich")
        ),
        None,
    )

    tb = slide.shapes.add_textbox(
        Inches(1.0),
        Inches(1.6),
        prs.slide_width - Inches(2.0),
        Inches(1.6),
    )
    tf = tb.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (title or {}).get("text", "")
    p.font.name = DEFAULT_FONT_NAME
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = to_rgb("#111827")
    p.alignment = PP_ALIGN.LEFT

    if subtitle:
        sb = slide.shapes.add_textbox(
            Inches(1.0),
            Inches(3.0),
            prs.slide_width - Inches(2.0),
            Inches(0.9),
        )
        stf = sb.text_frame
        stf.clear()
        stf.word_wrap = True
        sp = stf.paragraphs[0]
        sp.text = subtitle.get("text", "")
        sp.font.name = DEFAULT_FONT_NAME
        sp.font.size = Pt(18)
        sp.font.color.rgb = to_rgb("#374151")
        sp.alignment = PP_ALIGN.LEFT

    meta = s.get("meta", {})
    if meta:
        fb = slide.shapes.add_textbox(
            Inches(1.0),
            prs.slide_height - Inches(0.9),
            prs.slide_width - Inches(2.0),
            Inches(0.5),
        )
        ft = fb.text_frame
        ft.clear()
        fp = ft.paragraphs[0]
        fp.text = f'{meta.get("owner","")}  •  {meta.get("date","")}'
        fp.font.size = Pt(12)
        fp.font.color.rgb = to_rgb("#6B7280")
        fp.alignment = PP_ALIGN.LEFT

    return slide


def group_elements_by_slide_flat(elements):
    slides = {}
    for el in elements or []:
        idx = int(el.get("slide", 1))
        slides.setdefault(idx, []).append(el)
    return [{"title": None, "elements": slides[i]} for i in sorted(slides)]


def prepare_content_slide(slide):
    shapes = list(slide.shapes)
    for shape in shapes:
        if getattr(shape, "is_placeholder", False) and shape.is_placeholder:
            sp = shape._element
            sp.getparent().remove(sp)
        elif getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            shape.text = ""

def extract_slide_title_from_heading(slide_spec):
    if not slide_spec.get("elements"):
        return slide_spec.get("title"), None

    for el in slide_spec["elements"]:
        if el.get("type") == "text" and el.get("variant") == "heading":
            heading_text = el.get("text")
            heading_style = el.get("style", {})
            return heading_text, heading_style
    return slide_spec.get("title"), None

def build_ppt_from_spec(spec: dict, output_filename: str = "spec_demo_output.pptx") -> str:
    template_path = os.path.join(script_dir(), "FSOFT_SLIDE_TEMPLATE.pptx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    prs = Presentation(template_path)
    if len(prs.slides) < 3:
        raise ValueError("Template must have at least 3 slides")

    content_template_slide = prs.slides[2]
    xml_slides = prs.slides._sldIdLst
    xml_list = list(xml_slides)
    thankyou_id = xml_list[len(xml_list) - 1]

    slides_spec = spec.get("slides")
    if not slides_spec:
        slides_spec = group_elements_by_slide_flat(spec.get("elements", []))
        if slides_spec and (spec.get("metadata") or {}).get("title"):
            slides_spec[0]["title"] = (spec.get("metadata") or {}).get("title")

    first_dynamic = True
    for s in slides_spec:
        if s.get("layout") == "cover":
            continue
        if first_dynamic:
            slide = content_template_slide
            prepare_content_slide(slide)
            first_dynamic = False
        else:
            slide = prs.slides.add_slide(content_template_slide.slide_layout)
            prepare_content_slide(slide)
        title_text, heading_style = extract_slide_title_from_heading(s)

        if heading_style:
            # force white text if not provided
            if "text" not in heading_style:
                heading_style["text"] = "#FFFFFF"

            tb = slide.shapes.add_textbox(
                Inches(1.0), Inches(0.35), Inches(11.3), Inches(0.6)
            )
            tf = tb.text_frame
            tf.text = title_text
            p = tf.paragraphs[0]
            apply_text_style(p, {
                "font": {
                    "size": heading_style.get("font", {}).get("size", 20),
                    "bold": heading_style.get("font", {}).get("bold", True),
                    "name": heading_style.get("font", {}).get("name", DEFAULT_FONT_NAME)
                },
                "align": heading_style.get("align", "left"),
                "text": heading_style.get("text", "#FFFFFF")
            })
        else:
            add_title_if_any(slide, title_text)

        # render all elements except the heading (which we used as the title)
        for el in s.get("elements", []):
            if el.get("variant") == "heading":
                continue
            render_element(slide, el)

    xml_slides = prs.slides._sldIdLst
    ids_now = list(xml_slides)
    if thankyou_id in ids_now:
        xml_slides.remove(thankyou_id)
        xml_slides.append(thankyou_id)

    out_path = os.path.join(script_dir(), output_filename)
    try:
        prs.save(out_path)
    except PermissionError:
        base, ext = os.path.splitext(out_path)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = f"{base}_{ts}{ext}"
        prs.save(alt)
        out_path = alt

    try:
        print(f"Saved PowerPoint: {os.path.abspath(out_path)}")
    except Exception:
        print(f"Saved PowerPoint: {os.path.abspath(out_path)}")
    print(f"Slides: {len(prs.slides)}")
    return out_path


def _padded_chart_frame(position, legend_pos="right", pad_in=0.0):
    from pptx.util import Inches as _In

    x, y, w, h = inches_rect(position)
    pad = _In(max(pad_in, 0.0))
    if legend_pos == "right":
        w = max(_In(1.0), w - pad)
    elif legend_pos == "left":
        x = x + pad
        w = max(_In(1.0), w - pad)
    elif legend_pos == "top":
        y = y + pad
        h = max(_In(1.0), h - pad)
    elif legend_pos == "bottom":
        h = max(_In(1.0), h - pad)
    return x, y, w, h


def _shorten_lane(s: str) -> str:
    s = s.strip()
    if s.lower().startswith("phase "):
        parts = s.split(":", 1)
        head = parts[0]
        tail = parts[1].strip() if len(parts) > 1 else ""
        num = head.split()[-1]
        short = f"P{num} – {tail}" if tail else f"P{num}"
    else:
        short = s
    return short[:32] + ("…" if len(short) > 32 else "")


def _quarter_label(dt):
    q = ((dt.month - 1) // 3) + 1
    return f"Q{q} {dt.year}"

def load_json_spec(filename):
    path = os.path.join(script_dir(), filename)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

SAMPLE_TEN_SLIDES = load_json_spec("mock_pptx.json") 

SAMPLE_FLAT_WITH_SLIDE_TAGS = {
    "version": "1.0",
    "metadata": {"title": "Flat model (slide-tagged)"},
    "elements": [
        {
            "type": "text",
            "variant": "heading",
            "text": "Slide 1 title",
            "position": {"x": 1, "y": 0.6, "w": 10, "h": 0.8},
            "slide": 1,
        },
        {
            "type": "text",
            "variant": "paragraph",
            "text": "Hello slide 1",
            "position": {"x": 1, "y": 1.4, "w": 10, "h": 0.6},
            "slide": 1,
        },
        {
            "type": "text",
            "variant": "heading",
            "text": "Slide 2 title",
            "position": {"x": 1, "y": 0.6, "w": 10, "h": 0.8},
            "slide": 2,
        },
        {
            "type": "text",
            "variant": "paragraph",
            "text": "Hello slide 2",
            "position": {"x": 1, "y": 1.4, "w": 10, "h": 0.6},
            "slide": 2,
        },
    ],
}


if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser(
        description="Render PPTX from JSON spec or built-in samples"
    )
    parser.add_argument(
        "--in",
        dest="in_path",
        help="Path to JSON spec (defaults to quang/mock_pptx.json)",
    )
    parser.add_argument(
        "--out",
        dest="out_path",
        default="spec_output.pptx",
        help="Output PPTX filename",
    )
    parser.add_argument(
        "--sample",
        dest="sample",
        choices=["ten", "flat"],
        help="Build a built-in sample deck if no JSON present",
    )
    args = parser.parse_args()

    in_path = args.in_path
    if not in_path:
        default_json = os.path.join(script_dir(), "mock_pptx.json")
        if os.path.exists(default_json):
            in_path = default_json

    if in_path:
        with open(in_path, "r", encoding="utf-8") as f:
            spec = json.load(f)
        build_ppt_from_spec(spec, output_filename=args.out_path)
    else:
        if args.sample == "flat":
            build_ppt_from_spec(
                SAMPLE_FLAT_WITH_SLIDE_TAGS, output_filename=args.out_path
            )
        else:
            build_ppt_from_spec(SAMPLE_TEN_SLIDES, output_filename=args.out_path)
