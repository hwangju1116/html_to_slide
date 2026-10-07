import base64
import inspect
import io
import logging
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

SKILL_ROOT = str(Path(__file__).resolve().parent.parent)
if SKILL_ROOT not in sys.path:
  sys.path.insert(0, SKILL_ROOT)

from scripts.color_utils import (
    hex_to_pptx_color as _hex_to_pptx_color,
    hex_to_rgb_tuple as _hex_to_rgb_tuple,
    parse_color_value,
)

logger = logging.getLogger(__name__)

def build_styled_native_table(
    slide,
    table_data: Dict[str, Any],
    left: Inches,
    top: Inches,
    width: Inches,
    height: Inches,
    card_border_rgb: Optional[RGBColor] = None,
    font_name: str = "Pretendard",
    col_ratios: Optional[List[float]] = None,
):
  is_dark = table_data.get("is_dark", False)
  slide_bg = table_data.get("slide_bg_rgb", (15, 23, 42) if is_dark else (255, 255, 255))
  blend_rgb = (slide_bg[0], slide_bg[1], slide_bg[2])

  if card_border_rgb is None:
    card_border_rgb = RGBColor(51, 65, 85) if is_dark else RGBColor(233, 219, 237)

  if left < Inches(0.4):
    left = Inches(0.4)
  max_avail_w = Inches(13.333) - left - Inches(0.4)
  if width > max_avail_w:
    width = max_avail_w

  skip_card_wrapper = bool(table_data.get("skip_card_wrapper", False))

  if not skip_card_wrapper and (table_data.get("in_card") or table_data.get("inCard")):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card_bg = table_data.get("card_bg_rgb") or (RGBColor(30, 41, 59) if is_dark else RGBColor(255, 255, 255))
    card.fill.fore_color.rgb = card_bg
    card.line.color.rgb = card_border_rgb
    card.line.width = Pt(1.0)

  card_title = (table_data.get("card_title") or table_data.get("cardTitle")) if not skip_card_wrapper else None
  if card_title:
    tx = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.12), width - Inches(0.4), Inches(0.38))
    p = tx.text_frame.paragraphs[0]
    p.text = card_title
    p.font.name = font_name
    p.font.size = Pt(13.0)
    p.font.bold = True
    ct_col = table_data.get("cardTitleColor")
    p.font.color.rgb = parse_color_value(ct_col, RGBColor(147, 197, 253) if is_dark else RGBColor(37, 99, 235))

  if skip_card_wrapper:
    t_top = top
    t_height = height
    t_left = left
    t_width = width
  else:
    t_top = top + (Inches(0.52) if card_title else Inches(0.12))
    t_height = height - (Inches(0.64) if card_title else Inches(0.24))
    t_left = left + Inches(0.12)
    t_width = width - Inches(0.24)

  num_rows = table_data["num_rows"]
  num_cols = table_data["num_cols"]
  table_shape = slide.shapes.add_table(num_rows, num_cols, t_left, t_top, t_width, t_height)
  tbl = table_shape.table
  try:
    for st_id in tbl._tbl.xpath("./a:tblPr/a:tableStyleId"):
      st_id.getparent().remove(st_id)
  except Exception:
    pass

  if col_ratios and len(col_ratios) == num_cols:
    total_r = sum(col_ratios)
    for c_idx, r in enumerate(col_ratios):
      tbl.columns[c_idx].width = int(t_width * (r / total_r))
  else:
    eq_w = int(t_width / num_cols)
    for c_idx in range(num_cols):
      tbl.columns[c_idx].width = eq_w

  if table_data.get("headers"):
    for c_idx, h in enumerate(table_data["headers"]):
      h_text = h["text"] if isinstance(h, dict) else str(h)
      h_bg = h.get("bgColor") if isinstance(h, dict) else None
      h_col = h.get("color") if isinstance(h, dict) else None

      if h_bg:
        header_bg_rgb = parse_color_value(h_bg, default=RGBColor(15, 23, 42) if is_dark else RGBColor(241, 245, 249), bg_blend_rgb=blend_rgb)
      elif is_dark:
        header_bg_rgb = RGBColor(15, 23, 42)
      else:
        header_bg_rgb = RGBColor(241, 245, 249)

      if h_col:
        header_fg_rgb = parse_color_value(h_col, default=RGBColor(148, 163, 184) if is_dark else RGBColor(30, 41, 59))
      elif is_dark:
        header_fg_rgb = RGBColor(148, 163, 184)
      else:
        header_fg_rgb = RGBColor(30, 41, 59)

      cell = tbl.cell(0, c_idx)
      cell.text = h_text
      cell.fill.solid()
      cell.fill.fore_color.rgb = header_bg_rgb
      cell.vertical_anchor = MSO_ANCHOR.MIDDLE
      cell.margin_left = Inches(0.12)
      cell.margin_right = Inches(0.12)
      cell.margin_top = Inches(0.08)
      cell.margin_bottom = Inches(0.08)
      p = cell.text_frame.paragraphs[0]
      p.font.name = font_name
      p.font.size = Pt(11.0)
      p.font.bold = True
      p.font.color.rgb = header_fg_rgb

  for r_idx, row in enumerate(table_data["rows"]):
    row_num = r_idx + (1 if table_data.get("headers") else 0)
    for c_idx, cell_data in enumerate(row):
      if c_idx >= num_cols:
        break
      cell = tbl.cell(row_num, c_idx)
      cell.vertical_anchor = MSO_ANCHOR.MIDDLE
      cell.margin_left = Inches(0.12)
      cell.margin_right = Inches(0.12)
      cell.margin_top = Inches(0.08)
      cell.margin_bottom = Inches(0.08)
      cell.fill.solid()

      cell_bg = cell_data.get("bgColor")
      if cell_bg:
        row_bg_rgb = parse_color_value(cell_bg, default=RGBColor(30, 41, 59) if is_dark else RGBColor(255, 255, 255), bg_blend_rgb=blend_rgb)
      elif is_dark:
        row_bg_rgb = RGBColor(30, 41, 59) if r_idx % 2 == 0 else RGBColor(22, 30, 46)
      else:
        row_bg_rgb = RGBColor(255, 255, 255)
      cell.fill.fore_color.rgb = row_bg_rgb

      tf = cell.text_frame
      tf.word_wrap = True
      p = tf.paragraphs[0]
      p.font.name = font_name
      p.font.size = Pt(11.5)

      default_cell_fg = RGBColor(241, 245, 249) if is_dark else RGBColor(17, 17, 21)
      if "color" in cell_data:
        col_rgb = parse_color_value(cell_data["color"], default=default_cell_fg)
      elif "color_rgb" in cell_data:
        col_rgb = RGBColor(*cell_data["color_rgb"])
      else:
        col_rgb = default_cell_fg

      bg_lum = (row_bg_rgb[0] * 0.299 + row_bg_rgb[1] * 0.587 + row_bg_rgb[2] * 0.114)
      fg_lum = (col_rgb[0] * 0.299 + col_rgb[1] * 0.587 + col_rgb[2] * 0.114)
      if bg_lum < 120 and fg_lum < 90:
        col_rgb = RGBColor(241, 245, 249)
      elif bg_lum >= 150 and fg_lum > 180:
        col_rgb = RGBColor(17, 17, 21)

      if cell_data.get("lines") and len(cell_data["lines"]) > 0:
        for l_idx, line_info in enumerate(cell_data["lines"]):
          lp = tf.paragraphs[0] if l_idx == 0 else tf.add_paragraph()
          lp.text = line_info.get("text", "")
          lp.font.name = font_name
          lp.font.size = Pt(float(line_info.get("fontSizePt") or (11.0 if l_idx == 0 else 9.5)))
          lp.font.bold = bool(line_info.get("isBold", l_idx == 0 and cell_data.get("is_bold")))
          l_col = parse_color_value(line_info.get("color"), default=col_rgb)
          l_lum = (l_col[0] * 0.299 + l_col[1] * 0.587 + l_col[2] * 0.114)
          if bg_lum < 120 and l_lum < 75:
            l_col = RGBColor(203, 213, 225)
          elif bg_lum >= 150 and l_lum > 190:
            l_col = RGBColor(51, 65, 85)
          lp.font.color.rgb = l_col
      elif cell_data.get("has_badge"):
        p.text = cell_data.get("badge_text") or cell_data["text"]
        p.font.bold = True
        p.font.color.rgb = RGBColor(96, 165, 250) if is_dark else RGBColor(37, 99, 235)
        p.alignment = PP_ALIGN.CENTER
      elif cell_data.get("runs") and len(cell_data["runs"]) > 1:
        p.text = ""
        for r in cell_data["runs"]:
          run = p.add_run()
          run.text = r["text"]
          run.font.name = font_name
          run.font.size = Pt(11.5)
          run.font.bold = r["bold"]
          run.font.color.rgb = col_rgb
      else:
        p.text = cell_data["text"]
        p.font.bold = cell_data["is_bold"]
        p.font.color.rgb = col_rgb

  return table_shape


def build_styled_native_chart(
    slide,
    chart_data: Dict[str, Any],
    font_name: str = "Pretendard",
    is_dark: bool = False,
    slide_bg_rgb: tuple[int, int, int] = (15, 23, 42),
):
  import io
  from pptx.chart.data import CategoryChartData
  from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
  from pptx.enum.dml import MSO_LINE_DASH_STYLE
  from pptx.oxml import parse_xml
  from pptx.oxml.ns import nsdecls

  r = chart_data.get("rect", {})
  left = Inches(max(0.2, float(r.get("left", 0.8))))
  top = Inches(max(0.2, float(r.get("top", 1.8))))
  width = Inches(max(1.2, float(r.get("width", 5.0))))
  height = Inches(max(1.0, float(r.get("height", 3.2))))

  cfg = chart_data.get("chartConfig")
  blend_rgb = (slide_bg_rgb[0], slide_bg_rgb[1], slide_bg_rgb[2])

  if cfg and isinstance(cfg, dict):
    try:
      c_type = str(cfg.get("type", "bar")).lower()
      options = cfg.get("options") if isinstance(cfg.get("options"), dict) else {}
      index_axis = str(cfg.get("indexAxis") or options.get("indexAxis", "x")).lower()

      xl_type = None
      if c_type == "line":
        xl_type = XL_CHART_TYPE.LINE_MARKERS
      elif c_type == "bar" and index_axis == "y":
        xl_type = XL_CHART_TYPE.BAR_CLUSTERED
      elif c_type == "bar":
        xl_type = XL_CHART_TYPE.COLUMN_CLUSTERED
      elif c_type == "doughnut":
        xl_type = XL_CHART_TYPE.DOUGHNUT
      elif c_type == "pie":
        xl_type = XL_CHART_TYPE.PIE
      elif c_type == "radar":
        xl_type = XL_CHART_TYPE.RADAR

      data_dict = cfg.get("data") if isinstance(cfg.get("data"), dict) else {}
      raw_labels = cfg.get("labels") if isinstance(cfg.get("labels"), list) else (data_dict.get("labels") or [])
      raw_datasets = cfg.get("datasets") if isinstance(cfg.get("datasets"), list) else (data_dict.get("datasets") or [])

      if xl_type is not None and len(raw_labels) > 0 and len(raw_datasets) > 0:
        is_horiz_bar = (xl_type == XL_CHART_TYPE.BAR_CLUSTERED)
        labels = [str(lbl) for lbl in (reversed(raw_labels) if is_horiz_bar else raw_labels)]

        chart_data_obj = CategoryChartData()
        chart_data_obj.categories = labels

        for ds in raw_datasets:
          raw_vals = ds.get("data") or []
          vals_iter = reversed(raw_vals) if is_horiz_bar else raw_vals
          vals = []
          for v in vals_iter:
            try:
              vals.append(float(v))
            except Exception:
              vals.append(0.0)
          chart_data_obj.add_series(str(ds.get("label", "")), tuple(vals))

        chart_shape = slide.shapes.add_chart(xl_type, left, top, width, height, chart_data_obj)
        chart = chart_shape.chart

        try:
          cs = chart._element
          for existing_spPr in cs.xpath("./c:spPr"):
            cs.remove(existing_spPr)
          spPr_cs = parse_xml(f"<c:spPr {nsdecls('c', 'a')}><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>")
          chart_el_list = cs.xpath("./c:chart")
          if chart_el_list:
            cs.insert(cs.index(chart_el_list[0]) + 1, spPr_cs)
          else:
            cs.append(spPr_cs)

          pa_list = cs.xpath("./c:chart/c:plotArea")
          if pa_list:
            pa = pa_list[0]
            for existing_pa_spPr in pa.xpath("./c:spPr"):
              pa.remove(existing_pa_spPr)
            spPr_pa = parse_xml(f"<c:spPr {nsdecls('c', 'a')}><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>")
            pa.append(spPr_pa)
        except Exception:
          pass

        if xl_type == XL_CHART_TYPE.DOUGHNUT:
          try:
            hole_pct = 65
            cutout_val = cfg.get("cutout") or options.get("cutout")
            if cutout_val:
              m_cut = re.search(r"\d+", str(cutout_val))
              if m_cut:
                hole_pct = max(10, min(90, int(m_cut.group(0))))
            dc_list = chart._element.xpath("./c:chart/c:plotArea/c:doughnutChart")
            if dc_list:
              dc = dc_list[0]
              for hs in dc.xpath("./c:holeSize"):
                dc.remove(hs)
              dc.append(parse_xml(f'<c:holeSize {nsdecls("c")} val="{hole_pct}"/>'))
          except Exception:
            pass

        plugins_cfg = options.get("plugins") if isinstance(options.get("plugins"), dict) else {}
        legend_cfg = plugins_cfg.get("legend") if isinstance(plugins_cfg.get("legend"), dict) else {}
        default_show_legend = len(raw_datasets) > 1 or xl_type in (XL_CHART_TYPE.DOUGHNUT, XL_CHART_TYPE.PIE)
        if "legendDisplay" in cfg and cfg["legendDisplay"] is not None:
          show_legend = bool(cfg["legendDisplay"])
        else:
          show_legend = bool(legend_cfg.get("display", default_show_legend))
        if show_legend:
          chart.has_legend = True
          pos_str = str(cfg.get("legendPosition") or legend_cfg.get("position", "top")).lower()
          if pos_str == "bottom":
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
          elif pos_str == "right":
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
          elif pos_str == "left":
            chart.legend.position = XL_LEGEND_POSITION.LEFT
          else:
            chart.legend.position = XL_LEGEND_POSITION.TOP
          chart.legend.include_in_layout = False
          chart.legend.font.name = font_name
          chart.legend.font.size = Pt(9.5)
          chart.legend.font.color.rgb = RGBColor(203, 213, 225) if is_dark else RGBColor(51, 65, 85)
        else:
          chart.has_legend = False

        if xl_type not in (XL_CHART_TYPE.DOUGHNUT, XL_CHART_TYPE.PIE, XL_CHART_TYPE.RADAR):
          axis_fg = RGBColor(148, 163, 184) if is_dark else RGBColor(100, 116, 139)
          grid_hex = "1E293B" if is_dark else "E2E8F0"
          scales_cfg = options.get("scales") if isinstance(options.get("scales"), dict) else {}

          try:
            cat_ax = chart.category_axis
            cat_ax.tick_labels.font.name = font_name
            cat_ax.tick_labels.font.size = Pt(9.0)
            cat_ax.tick_labels.font.color.rgb = axis_fg
            cat_ax.has_major_gridlines = False
            for sp in cat_ax._element.xpath("./c:spPr"):
              cat_ax._element.remove(sp)
            cat_ax._element.append(
                parse_xml(f'<c:spPr {nsdecls("c", "a")}><a:ln w="9525"><a:solidFill><a:srgbClr val="{grid_hex}"/></a:solidFill></a:ln></c:spPr>')
            )
          except Exception:
            pass

          try:
            val_ax = chart.value_axis
            val_scale_key = "x" if is_horiz_bar else "y"
            val_scale_cfg = scales_cfg.get(val_scale_key) if isinstance(scales_cfg.get(val_scale_key), dict) else {}
            explicit_display = cfg.get(f"{val_scale_key}Display")
            if explicit_display is False or val_scale_cfg.get("display") is False:
              val_ax.visible = False
              val_ax.has_major_gridlines = False
            else:
              val_ax.tick_labels.font.name = font_name
              val_ax.tick_labels.font.size = Pt(9.0)
              val_ax.tick_labels.font.color.rgb = axis_fg
              val_ax.has_major_gridlines = True
              mg = val_ax.major_gridlines
              for sp in mg._element.xpath("./c:spPr"):
                mg._element.remove(sp)
              mg._element.append(
                  parse_xml(f'<c:spPr {nsdecls("c", "a")}><a:ln w="9525"><a:solidFill><a:srgbClr val="{grid_hex}"/></a:solidFill></a:ln></c:spPr>')
              )
              for sp in val_ax._element.xpath("./c:spPr"):
                val_ax._element.remove(sp)
              val_ax._element.append(
                  parse_xml(f'<c:spPr {nsdecls("c", "a")}><a:ln><a:noFill/></a:ln></c:spPr>')
              )
          except Exception:
            pass

        default_palette = [
            RGBColor(225, 29, 72),
            RGBColor(59, 130, 246),
            RGBColor(16, 185, 129),
            RGBColor(245, 158, 11),
            RGBColor(139, 92, 246),
        ]
        try:
          has_multi_pt_colors = len(raw_datasets) == 1 and isinstance(raw_datasets[0].get("backgroundColor"), list) and len(raw_datasets[0].get("backgroundColor")) > 1
          chart.plots[0].vary_by_categories = bool(has_multi_pt_colors or xl_type in (XL_CHART_TYPE.DOUGHNUT, XL_CHART_TYPE.PIE))
        except Exception:
          pass
        for s_idx, series in enumerate(chart.series):
          if s_idx >= len(raw_datasets):
            break
          ds = raw_datasets[s_idx]
          fallback_col = default_palette[s_idx % len(default_palette)]
          bg_col = ds.get("backgroundColor")
          border_col = ds.get("borderColor")

          if xl_type == XL_CHART_TYPE.LINE_MARKERS:
            line_col_str = border_col if isinstance(border_col, str) else (bg_col if isinstance(bg_col, str) else None)
            line_rgb = parse_color_value(line_col_str, default=fallback_col, bg_blend_rgb=blend_rgb)
            series.format.line.color.rgb = line_rgb
            series.format.line.width = Pt(float(ds.get("borderWidth", 2.5)))
            if ds.get("borderDash"):
              series.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            if float(ds.get("tension", 0) or 0) > 0:
              try:
                for sm in series._element.xpath("./c:smooth"):
                  series._element.remove(sm)
                series._element.append(parse_xml(f'<c:smooth {nsdecls("c")} val="1"/>'))
              except Exception:
                pass
            try:
              hex_col = f"{line_rgb[0]:02X}{line_rgb[1]:02X}{line_rgb[2]:02X}"
              for mk in series._element.xpath("./c:marker"):
                series._element.remove(mk)
              mk_xml = parse_xml(
                  f'<c:marker {nsdecls("c", "a")}>'
                  f'<c:symbol val="circle"/><c:size val="6"/>'
                  f'<c:spPr><a:solidFill><a:srgbClr val="{hex_col}"/></a:solidFill>'
                  f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{hex_col}"/></a:solidFill></a:ln></c:spPr>'
                  f'</c:marker>'
              )
              series._element.insert(2, mk_xml)
            except Exception:
              pass
          else:
            if isinstance(bg_col, list) and len(bg_col) > 0:
              col_list = list(reversed(bg_col)) if is_horiz_bar else list(bg_col)
              for pt_idx, pt in enumerate(series.points):
                c_str = col_list[pt_idx % len(col_list)]
                pt_rgb = parse_color_value(c_str, default=fallback_col, bg_blend_rgb=blend_rgb)
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = pt_rgb
                if xl_type in (XL_CHART_TYPE.DOUGHNUT, XL_CHART_TYPE.PIE):
                  pt.format.line.color.rgb = RGBColor(*blend_rgb)
                  pt.format.line.width = Pt(1.5)
                else:
                  pt.format.line.fill.background()
            else:
              c_str = bg_col if isinstance(bg_col, str) else (border_col if isinstance(border_col, str) else None)
              s_rgb = parse_color_value(c_str, default=fallback_col, bg_blend_rgb=blend_rgb)
              series.format.fill.solid()
              series.format.fill.fore_color.rgb = s_rgb
              series.format.line.fill.background()
              for pt in series.points:
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = s_rgb
                pt.format.line.fill.background()

        return chart_shape
    except Exception as chart_err:
      logger.warning("[Native Chart Builder Warning]: %s", chart_err)

  data_url = chart_data.get("dataUrl")
  if data_url and isinstance(data_url, str) and "base64," in data_url:
    try:
      b64_part = data_url.split("base64,", 1)[1]
      img_bytes = base64.b64decode(b64_part)
      return slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=width, height=height)
    except Exception as pic_err:
      logger.warning("[Chart Picture Fallback Warning]: %s", pic_err)
  return None


def apply_semantic_styles_to_table(
    table_shape,
    table_data: Dict[str, Any],
    font_name: str = "Pretendard",
    is_dark: bool = False,
    slide_bg_rgb: tuple[int, int, int] = (15, 23, 42),
) -> None:
  tbl = table_shape.table
  try:
    for st_id in tbl._tbl.xpath("./a:tblPr/a:tableStyleId"):
      st_id.getparent().remove(st_id)
  except Exception:
    pass
  blend_rgb = (slide_bg_rgb[0], slide_bg_rgb[1], slide_bg_rgb[2])

  if table_data.get("headers") and len(tbl.rows) > 0:
    for c_idx, h in enumerate(table_data["headers"]):
      if c_idx >= len(tbl.rows[0].cells):
        break
      c = tbl.rows[0].cells[c_idx]
      c.fill.solid()
      h_bg = h.get("bgColor") if isinstance(h, dict) else None
      h_col = h.get("color") if isinstance(h, dict) else None
      if h_bg:
        h_bg_rgb = parse_color_value(h_bg, default=RGBColor(15, 23, 42) if is_dark else RGBColor(241, 245, 249), bg_blend_rgb=blend_rgb)
      else:
        h_bg_rgb = RGBColor(15, 23, 42) if is_dark else RGBColor(241, 245, 249)
      c.fill.fore_color.rgb = h_bg_rgb

      h_fg_rgb = parse_color_value(h_col, default=RGBColor(148, 163, 184) if is_dark else RGBColor(30, 41, 59))
      for p in c.text_frame.paragraphs:
        p.font.name = font_name
        p.font.size = Pt(11.0)
        p.font.bold = True
        p.font.color.rgb = h_fg_rgb
        for r in p.runs:
          r.font.name = font_name
          r.font.size = Pt(11.0)
          r.font.bold = True
          r.font.color.rgb = h_fg_rgb

  for r_idx, row_data in enumerate(table_data.get("rows", [])):
    row_num = r_idx + (1 if table_data.get("headers") else 0)
    if row_num >= len(tbl.rows):
      break
    for c_idx, cell_data in enumerate(row_data):
      if c_idx >= len(tbl.columns):
        break
      cell = tbl.cell(row_num, c_idx)
      cell.fill.solid()

      cell_bg = cell_data.get("bgColor")
      if cell_bg:
        row_bg_rgb = parse_color_value(cell_bg, default=RGBColor(30, 41, 59) if is_dark else RGBColor(255, 255, 255), bg_blend_rgb=blend_rgb)
      elif is_dark:
        row_bg_rgb = RGBColor(30, 41, 59) if r_idx % 2 == 0 else RGBColor(22, 30, 46)
      else:
        row_bg_rgb = RGBColor(255, 255, 255) if r_idx % 2 == 0 else RGBColor(248, 250, 252)
      cell.fill.fore_color.rgb = row_bg_rgb

      default_cell_fg = RGBColor(241, 245, 249) if is_dark else RGBColor(17, 17, 21)
      if "color" in cell_data:
        col_rgb = parse_color_value(cell_data["color"], default=default_cell_fg)
      elif "color_rgb" in cell_data:
        col_rgb = RGBColor(*cell_data["color_rgb"])
      else:
        col_rgb = default_cell_fg

      bg_lum = (row_bg_rgb[0] * 0.299 + row_bg_rgb[1] * 0.587 + row_bg_rgb[2] * 0.114)
      fg_lum = (col_rgb[0] * 0.299 + col_rgb[1] * 0.587 + col_rgb[2] * 0.114)
      if bg_lum < 120 and fg_lum < 90:
        col_rgb = RGBColor(241, 245, 249)
      elif bg_lum >= 150 and fg_lum > 180:
        col_rgb = RGBColor(17, 17, 21)

      if not cell.text.strip() and cell_data.get("text"):
        cell.text = cell_data["text"]

      for p in cell.text_frame.paragraphs:
        p.font.name = font_name
        p.font.size = Pt(11.5)
        if cell_data.get("has_badge"):
          p.font.bold = True
          p.font.color.rgb = RGBColor(96, 165, 250) if is_dark else RGBColor(37, 99, 235)
          p.alignment = PP_ALIGN.CENTER
        for r in p.runs:
          r.font.name = font_name
          r.font.size = Pt(11.5)
          if cell_data.get("has_badge"):
            r.font.bold = True
            r.font.color.rgb = RGBColor(96, 165, 250) if is_dark else RGBColor(37, 99, 235)
          else:
            r.font.color.rgb = col_rgb
            if cell_data.get("is_bold"):
              r.font.bold = True


def ensure_slide_canvas_background(
    slide,
    prs: Optional[Presentation] = None,
    bg_color: Optional[RGBColor] = None,
    is_dark: bool = False,
) -> None:
  if bg_color is None:
    bg_color = RGBColor(15, 23, 42) if is_dark else RGBColor(255, 255, 255)

  bg_shapes = [
      s for s in slide.shapes
      if s.shape_type == 1 and s.left.inches <= 0.1 and s.top.inches <= 0.1 and s.width.inches >= 13.0 and s.height.inches >= 7.0
  ]
  if len(bg_shapes) > 1:
    for dup in bg_shapes[1:]:
      try:
        slide.shapes._spTree.remove(dup._element)
      except Exception:
        pass

  if bg_shapes:
    full_bleed_bg = bg_shapes[0]
    full_bleed_bg.fill.solid()
    full_bleed_bg.fill.fore_color.rgb = bg_color
    full_bleed_bg.line.fill.background()
  else:
    w = prs.slide_width if prs else Inches(13.333333)
    h = prs.slide_height if prs else Inches(7.5)
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), w, h)
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = bg_color
    bg_shape.line.fill.background()
    try:
      spTree = slide.shapes._spTree
      spTree.insert(2, bg_shape._element)
    except Exception:
      pass


def embed_extracted_slide_images(
    slide,
    geom: Optional[Dict[str, Any]],
    screenshot_source: Optional[Union[bytes, str]],
) -> None:
  if not geom or not screenshot_source or not geom.get("images"):
    return

  try:
    from PIL import Image
    import os

    if isinstance(screenshot_source, bytes):
      img_stream = io.BytesIO(screenshot_source)
    elif isinstance(screenshot_source, str) and os.path.exists(screenshot_source):
      img_stream = screenshot_source
    else:
      return

    with Image.open(img_stream) as full_img:
      img_w, img_h = full_img.size
      for img_info in geom.get("images", []):
        r = img_info.get("rect", {})
        left_in = float(r.get("left", 0.0))
        top_in = float(r.get("top", 0.0))
        width_in = float(r.get("width", 0.0))
        height_in = float(r.get("height", 0.0))
        if width_in < 0.25 or height_in < 0.25:
          continue

        already_embedded = False
        for s in slide.shapes:
          if getattr(s, "shape_type", None) == 13:
            s_l = s.left.inches if hasattr(s.left, "inches") else s.left / 914400.0
            s_t = s.top.inches if hasattr(s.top, "inches") else s.top / 914400.0
            if abs(s_l - left_in) < 0.35 and abs(s_t - top_in) < 0.35:
              already_embedded = True
              break
        if already_embedded:
          continue

        px_l = max(0, int(round((left_in / 13.333333) * img_w)))
        px_t = max(0, int(round((top_in / 7.5) * img_h)))
        px_r = min(img_w, int(round(((left_in + width_in) / 13.333333) * img_w)))
        px_b = min(img_h, int(round(((top_in + height_in) / 7.5) * img_h)))
        if px_r - px_l < 10 or px_b - px_t < 10:
          continue

        cropped = full_img.crop((px_l, px_t, px_r, px_b))
        bio = io.BytesIO()
        cropped.save(bio, format="PNG")
        bio.seek(0)
        slide.shapes.add_picture(
            bio,
            Inches(left_in),
            Inches(top_in),
            Inches(width_in),
            Inches(height_in),
        )
  except Exception as e:
    logger.warning("Failed to embed extracted slide images: %s", e)


def build_slide_from_geometry(
    slide,
    geom: Dict[str, Any],
    font_name: str = "Pretendard",
    brand_color_rgb: Optional[RGBColor] = None,
    screenshot_source: Optional[Union[bytes, str]] = None,
) -> None:
  if not geom:
    return

  if brand_color_rgb is None:
    brand_color_rgb = RGBColor(37, 99, 235)

  bg_color = parse_color_value(geom.get("slideBgColor"), RGBColor(255, 255, 255))
  is_dark = (bg_color[0] * 0.299 + bg_color[1] * 0.587 + bg_color[2] * 0.114) < 128
  ensure_slide_canvas_background(slide, prs=None, bg_color=bg_color, is_dark=is_dark)
  blend_rgb = (bg_color[0], bg_color[1], bg_color[2])

  for b_data in geom.get("headerBadges", []):
    r = b_data["rect"]
    b_txt = b_data["text"]
    raw_bg = (b_data.get("styles", {}).get("backgroundColor") or "").strip()
    has_bg = bool(raw_bg and raw_bg not in ("transparent", "rgba(0, 0, 0, 0)", "none"))
    is_pill = r["left"] > 7.0 or "pill" in str(b_data.get("styles", {})) or "Slide #" in b_txt or has_bg
    bg_rgb = parse_color_value(raw_bg if has_bg else None, None, bg_blend_rgb=blend_rgb)
    fg_rgb = parse_color_value(b_data["styles"]["color"], RGBColor(255, 255, 255) if has_bg else brand_color_rgb)
    b_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(r["left"]),
        Inches(r["top"]),
        Inches(max(0.8, r["width"])),
        Inches(max(0.22, r["height"])),
    )
    if bg_rgb is not None:
      b_shape.fill.solid()
      b_shape.fill.fore_color.rgb = bg_rgb
    else:
      b_shape.fill.background()
    b_border = parse_color_value(b_data["styles"].get("borderColor"), None, bg_blend_rgb=blend_rgb)
    if b_border:
      b_shape.line.color.rgb = b_border
      b_shape.line.width = Pt(1.0)
    else:
      b_shape.line.fill.background()
    tf = b_shape.text_frame
    tf.word_wrap = False
    tf.margin_left = Inches(0.08 if has_bg else 0.0)
    tf.margin_right = Inches(0.08 if has_bg else 0.0)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.text = b_txt
    p.font.name = font_name
    p.font.size = Pt(b_data["styles"].get("fontSizePt", 10.0))
    p.font.bold = True
    p.font.color.rgb = fg_rgb
    p.alignment = PP_ALIGN.CENTER if has_bg else PP_ALIGN.LEFT

  if geom.get("num") and geom["num"]["text"]:
    n_data = geom["num"]
    r = n_data["rect"]
    num_box = slide.shapes.add_textbox(
        Inches(min(12.2, r["left"])),
        Inches(r["top"]),
        Inches(max(0.8, r["width"])),
        Inches(max(0.25, r["height"])),
    )
    p = num_box.text_frame.paragraphs[0]
    p.text = n_data["text"]
    p.font.name = font_name
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = parse_color_value(n_data["styles"].get("color") if n_data.get("styles") else None, RGBColor(148, 163, 184) if is_dark else RGBColor(100, 116, 139))

  header_bottom = 0.0
  if geom.get("title") and geom["title"]["text"]:
    t_data = geom["title"]
    r = t_data["rect"]
    st = t_data["styles"]
    is_center = st.get("textAlign") == "center"
    t_left = Inches(r["left"]) if is_center else Inches(max(0.6, r["left"]))
    t_top = Inches(r["top"])
    t_w = Inches(r["width"]) if is_center else Inches(min(12.0, r["width"]))
    t_h = Inches(max(0.45, r["height"]))
    header_bottom = max(header_bottom, float(r.get("top", 0.0)) + float(r.get("height", 0.45)))
    t_box = slide.shapes.add_textbox(t_left, t_top, t_w, t_h)
    tf = t_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    t_runs = t_data.get("runs")
    default_t_col = parse_color_value(st.get("color"), RGBColor(248, 250, 252) if is_dark else brand_color_rgb)
    if t_runs and len(t_runs) > 1:
      p.text = ""
      for tr_item in t_runs:
        run = p.add_run()
        run.text = tr_item.get("text", "")
        run.font.name = font_name
        run.font.size = Pt(st.get("fontSizePt", 24.0))
        run.font.bold = True
        run.font.color.rgb = parse_color_value(tr_item.get("color"), default_t_col)
    else:
      p.text = t_data["text"]
      p.font.name = font_name
      p.font.size = Pt(st.get("fontSizePt", 24.0))
      p.font.bold = True
      p.font.color.rgb = default_t_col
    if is_center:
      p.alignment = PP_ALIGN.CENTER

  if geom.get("sub") and geom["sub"]["text"]:
    s_data = geom["sub"]
    r = s_data["rect"]
    st = s_data["styles"]
    s_left = Inches(max(0.6, r["left"]))
    s_top = Inches(r["top"])
    s_w = Inches(min(12.0, r["width"]))
    s_h = Inches(max(0.30, r["height"]))
    header_bottom = max(header_bottom, float(r.get("top", 0.0)) + float(r.get("height", 0.30)))
    s_box = slide.shapes.add_textbox(s_left, s_top, s_w, s_h)
    tf = s_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = s_data["text"]
    p.font.name = font_name
    p.font.size = Pt(st.get("fontSizePt", 13.0))
    p.font.color.rgb = parse_color_value(st.get("color"), RGBColor(148, 163, 184) if is_dark else RGBColor(100, 116, 139))

  min_body_top = (header_bottom + 0.14) if (0.0 < header_bottom < 2.5) else 0.0

  for d in geom.get("desc", []):
    r = d["rect"]
    st = d["styles"]
    d_box = slide.shapes.add_textbox(Inches(r["left"]), Inches(r["top"]), Inches(r["width"]), Inches(r["height"]))
    tf = d_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = d["text"]
    p.font.name = font_name
    p.font.size = Pt(st.get("fontSizePt", 12.0))
    p.font.color.rgb = parse_color_value(st.get("color"), RGBColor(203, 213, 225) if is_dark else RGBColor(74, 77, 82))

  for card in geom.get("cards", []):
    r = card["rect"]
    raw_c_top = float(r["top"])
    raw_c_h = float(r["height"])
    if min_body_top > 0.0 and raw_c_top < min_body_top and raw_c_top > 0.8:
      top_shift = min_body_top - raw_c_top
      raw_c_top = min_body_top
      if raw_c_h > 1.2:
        raw_c_h = max(0.8, raw_c_h - top_shift)
    c_left = Inches(r["left"])
    c_top = Inches(raw_c_top)
    c_w = Inches(r["width"])
    c_h = Inches(raw_c_h)

    if card.get("isBridge") and card.get("bridge"):
      b_info = card["bridge"]
      br = b_info["rect"]
      b_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(br["left"]), Inches(br["top"]), Inches(br["width"]), Inches(br["height"]))
      b_shape.fill.solid()
      b_shape.fill.fore_color.rgb = parse_color_value(b_info["styles"]["backgroundColor"], brand_color_rgb)
      b_shape.line.fill.background()
      tf = b_shape.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.05)
      tf.margin_right = Inches(0.05)
      p0 = tf.paragraphs[0]
      p0.text = b_info.get("arrow", "➔")
      p0.font.name = font_name
      p0.font.size = Pt(15.0)
      p0.font.bold = True
      p0.font.color.rgb = RGBColor(255, 255, 255)
      p0.alignment = PP_ALIGN.CENTER
      p1 = tf.add_paragraph()
      p1.text = b_info.get("text", "")
      p1.font.name = font_name
      p1.font.size = Pt(9.0)
      p1.font.bold = True
      p1.font.color.rgb = RGBColor(255, 255, 255)
      p1.alignment = PP_ALIGN.CENTER
      continue

    raw_c_bg = (card.get("styles", {}).get("backgroundColor") or "").strip()
    raw_c_border = (card.get("styles", {}).get("borderColor") or "").strip()
    has_c_bg = bool(raw_c_bg and raw_c_bg not in ("transparent", "rgba(0, 0, 0, 0)", "none"))
    has_c_border = bool(raw_c_border and raw_c_border not in ("transparent", "rgba(0, 0, 0, 0)", "none"))

    c_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, c_left, c_top, c_w, c_h)
    if has_c_bg:
      c_shape.fill.solid()
      c_bg = parse_color_value(raw_c_bg, RGBColor(30, 41, 59) if is_dark else RGBColor(255, 255, 255), bg_blend_rgb=blend_rgb)
      c_shape.fill.fore_color.rgb = c_bg
    else:
      c_shape.fill.background()

    if has_c_border:
      c_border = parse_color_value(raw_c_border, RGBColor(51, 65, 85) if is_dark else RGBColor(226, 232, 240), bg_blend_rgb=blend_rgb)
      c_shape.line.color.rgb = c_border
      c_shape.line.width = Pt(1.0)
    else:
      c_shape.line.fill.background()

    if card.get("isAttachCard") and card.get("attachData"):
      ad = card["attachData"]
      tf = c_shape.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.18)
      tf.margin_right = Inches(0.18)
      tf.margin_top = Inches(0.14)
      tf.margin_bottom = Inches(0.14)
      p0 = tf.paragraphs[0]
      p0.text = ad["target"] + (f"  [{ad['mode']}]" if ad.get("mode") else "")
      p0.font.name = font_name
      p0.font.size = Pt(12.5)
      p0.font.bold = True
      p0.font.color.rgb = brand_color_rgb
      if ad.get("mech"):
        p1 = tf.add_paragraph()
        p1.text = ad["mech"]
        p1.font.name = "Consolas"
        p1.font.size = Pt(9.0)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(24, 107, 242)
      if ad.get("desc"):
        p2 = tf.add_paragraph()
        p2.text = ad["desc"]
        p2.font.name = font_name
        p2.font.size = Pt(10.0)
        p2.font.color.rgb = RGBColor(74, 85, 104)
      continue

    curr_top = c_top
    if card.get("hasTopAccent"):
      bar_color = parse_color_value(card.get("topAccentColor"), brand_color_rgb)
      accent_bar = slide.shapes.add_shape(
          MSO_SHAPE.ROUNDED_RECTANGLE,
          c_left + Inches(0.20),
          c_top + Inches(0.08),
          c_w - Inches(0.40),
          Inches(0.06),
      )
      accent_bar.fill.solid()
      accent_bar.fill.fore_color.rgb = bar_color
      accent_bar.line.fill.background()
      try:
        accent_bar.adjustments[0] = 0.5
      except Exception:
        pass
      curr_top = c_top + Inches(0.18)
    else:
      curr_top = c_top + Inches(0.12)

    if card.get("hasLeftAccent"):
      l_bar_color = parse_color_value(card.get("leftAccentColor"), brand_color_rgb)
      l_bar = slide.shapes.add_shape(
          MSO_SHAPE.ROUNDED_RECTANGLE,
          c_left + Inches(0.06),
          c_top + Inches(0.12),
          Inches(0.06),
          max(Inches(0.24), c_h - Inches(0.24)),
      )
      l_bar.fill.solid()
      l_bar.fill.fore_color.rgb = l_bar_color
      l_bar.line.fill.background()

    has_pipeline = len(card.get("pipelineSteps", [])) > 0

    if card.get("tag") and card["tag"]["text"]:
      tg = card["tag"]
      tg_r = tg["rect"]
      tg_raw_bg = (tg.get("styles", {}).get("backgroundColor") or "").strip()
      tg_has_bg = bool(tg_raw_bg and tg_raw_bg not in ("transparent", "rgba(0, 0, 0, 0)", "none"))
      tg_bg = parse_color_value(tg_raw_bg if tg_has_bg else None, None, bg_blend_rgb=blend_rgb)
      tg_fg = parse_color_value(tg["styles"]["color"], RGBColor(255, 255, 255) if tg_has_bg else brand_color_rgb)
      tg_shape = slide.shapes.add_shape(
          MSO_SHAPE.ROUNDED_RECTANGLE,
          c_left + Inches(0.18),
          curr_top,
          Inches(max(0.8, tg_r["width"])),
          Inches(max(0.22, tg_r["height"])),
      )
      if tg_bg is not None:
        tg_shape.fill.solid()
        tg_shape.fill.fore_color.rgb = tg_bg
      else:
        tg_shape.fill.background()
      tg_shape.line.fill.background()
      tf = tg_shape.text_frame
      tf.word_wrap = False
      tf.margin_left = Inches(0.06 if tg_has_bg else 0.0)
      tf.margin_right = Inches(0.06 if tg_has_bg else 0.0)
      p = tf.paragraphs[0]
      p.text = tg["text"]
      p.font.name = font_name
      p.font.size = Pt(tg["styles"].get("fontSizePt", 9.5))
      p.font.bold = True
      p.font.color.rgb = tg_fg
      curr_top += Inches(max(0.22, tg_r["height"])) + Inches(0.04)

    for b in card.get("badges", []):
      br = b.get("rect", {})
      b_txt = b.get("text", "")
      if b_txt and br:
        b_w_in = max(0.60, br.get("width", 1.0))
        b_w = Inches(b_w_in)
        b_h = Inches(max(0.22, br.get("height", 0.24)))
        b_left = Inches(br.get("left", r.get("left", 0.0) + r.get("width", 1.0) - b_w_in - 0.15))
        b_top = Inches(br.get("top", r.get("top", 0.0) + 0.18))
        b_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            b_left,
            b_top,
            b_w,
            b_h,
        )
        b_shape.fill.solid()
        b_bg = parse_color_value(b.get("styles", {}).get("backgroundColor"), RGBColor(241, 245, 249), bg_blend_rgb=blend_rgb)
        b_shape.fill.fore_color.rgb = b_bg
        b_border = parse_color_value(b.get("styles", {}).get("borderColor"), RGBColor(226, 232, 240))
        b_shape.line.color.rgb = b_border
        b_shape.line.width = Pt(1.0)
        tf = b_shape.text_frame
        tf.word_wrap = False
        tf.margin_left = Inches(0.06)
        tf.margin_right = Inches(0.06)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        p = tf.paragraphs[0]
        p.text = b_txt
        p.font.name = font_name
        p.font.size = Pt(b.get("styles", {}).get("fontSizePt", 9.5))
        p.font.bold = True
        p.font.color.rgb = parse_color_value(b.get("styles", {}).get("color"), RGBColor(78, 86, 95))
        p.alignment = PP_ALIGN.CENTER

    if card.get("title") and card["title"] != "➔" and not has_pipeline:
      b_first_left = min((b["rect"]["left"] for b in card.get("badges", []) if b.get("rect")), default=None)
      if b_first_left is not None:
        title_w = max(Inches(1.0), Inches(b_first_left - (r.get("left", 0.0) + 0.18) - 0.08))
      else:
        title_w = c_w - Inches(0.36)
      t_box = slide.shapes.add_textbox(
          c_left + Inches(0.18),
          curr_top,
          title_w,
          Inches(0.36),
      )
      tf = t_box.text_frame
      tf.word_wrap = True
      p = tf.paragraphs[0]
      p.text = card["title"]
      p.font.name = font_name
      p.font.size = Pt(card.get("titleStyles", {}).get("fontSizePt", 13.5))
      p.font.bold = True
      p.font.color.rgb = parse_color_value(card.get("titleStyles", {}).get("color") if card.get("titleStyles") else None, RGBColor(96, 165, 250) if is_dark else brand_color_rgb)
      curr_top += Inches(0.38)

    if card.get("desc") and card["desc"]["text"] and not has_pipeline:
      d_box = slide.shapes.add_textbox(c_left + Inches(0.18), curr_top, c_w - Inches(0.36), Inches(0.36))
      tf = d_box.text_frame
      tf.word_wrap = True
      p = tf.paragraphs[0]
      p.text = card["desc"]["text"]
      p.font.name = font_name
      p.font.size = Pt(card["desc"]["styles"].get("fontSizePt", 10.5))
      p.font.color.rgb = parse_color_value(card["desc"]["styles"].get("color"), RGBColor(148, 163, 184) if is_dark else RGBColor(100, 116, 139))
      curr_top += Inches(0.40)

    for ps in card.get("pipelineSteps", []):
      psr = ps["rect"]
      ps_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(psr["left"]), Inches(psr["top"]), Inches(psr["width"]), Inches(psr["height"]))
      ps_shape.fill.solid()
      ps_bg = parse_color_value(ps["styles"]["backgroundColor"], RGBColor(255, 255, 255), bg_blend_rgb=blend_rgb)
      ps_shape.fill.fore_color.rgb = ps_bg
      ps_border = parse_color_value(ps["styles"].get("borderColor"), RGBColor(217, 226, 236))
      ps_shape.line.color.rgb = ps_border
      ps_shape.line.width = Pt(1.0)
      tf = ps_shape.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.10)
      tf.margin_right = Inches(0.10)
      is_ps_center = ps["styles"].get("textAlign") == "center"
      p0 = tf.paragraphs[0]
      step_prefix = f"{ps['stepNum']}. " if (ps.get("stepNum") and not str(ps.get("title", "")).startswith(str(ps["stepNum"]))) else ""
      p0.text = f"{step_prefix}{ps['title']}"
      p0.font.name = font_name
      p0.font.size = Pt((ps.get("titleStyles") or {}).get("fontSizePt", 11.0))
      p0.font.bold = True
      t_col_str = (ps.get("titleStyles") or {}).get("color") or ps["styles"].get("color")
      p0.font.color.rgb = parse_color_value(t_col_str, brand_color_rgb)
      if is_ps_center:
        p0.alignment = PP_ALIGN.CENTER
      if ps.get("desc"):
        p1 = tf.add_paragraph()
        p1.text = ps["desc"]
        p1.font.name = font_name
        p1.font.size = Pt((ps.get("descStyles") or {}).get("fontSizePt", 9.5))
        d_col_str = (ps.get("descStyles") or {}).get("color") or ps["styles"].get("color")
        p1.font.color.rgb = parse_color_value(d_col_str, RGBColor(100, 116, 139))
        if is_ps_center:
          p1.alignment = PP_ALIGN.CENTER

    for fl in card.get("flows", []):
      fl_r = fl.get("rect", {})
      fl_st = fl.get("styles", {})
      fl_raw_bg = (fl_st.get("backgroundColor") or "").strip()
      if fl_r and fl_raw_bg and fl_raw_bg not in ("transparent", "rgba(0, 0, 0, 0)", "none"):
        fl_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(fl_r["left"]),
            Inches(fl_r["top"]),
            Inches(fl_r["width"]),
            Inches(fl_r["height"]),
        )
        fl_box.fill.solid()
        fl_box.fill.fore_color.rgb = parse_color_value(fl_raw_bg, RGBColor(248, 250, 252), bg_blend_rgb=blend_rgb)
        fl_border = parse_color_value(fl_st.get("borderColor"), None, bg_blend_rgb=blend_rgb)
        if fl_border:
          fl_box.line.color.rgb = fl_border
          fl_box.line.width = Pt(1.0)
        else:
          fl_box.line.fill.background()
      for node in fl.get("nodes", []):
        nr = node["rect"]
        if node.get("isArrow"):
          arr_box = slide.shapes.add_textbox(Inches(nr["left"]), Inches(nr["top"]), Inches(max(0.24, nr["width"])), Inches(max(0.22, nr["height"])))
          tf = arr_box.text_frame
          tf.word_wrap = False
          tf.margin_left = Inches(0.0)
          tf.margin_right = Inches(0.0)
          p = tf.paragraphs[0]
          p.text = node["text"]
          p.font.name = font_name
          p.font.size = Pt(11.0)
          p.font.bold = True
          p.font.color.rgb = parse_color_value(node.get("styles", {}).get("color"), brand_color_rgb)
          p.alignment = PP_ALIGN.CENTER
        else:
          n_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(nr["left"]), Inches(nr["top"]), Inches(nr["width"]), Inches(nr["height"]))
          n_shape.fill.solid()
          n_bg = parse_color_value(node["styles"]["backgroundColor"], RGBColor(241, 245, 249) if not node.get("isHighlight") else RGBColor(238, 242, 255), bg_blend_rgb=blend_rgb)
          n_shape.fill.fore_color.rgb = n_bg
          n_border = parse_color_value(node["styles"].get("borderColor"), brand_color_rgb if node.get("isHighlight") else RGBColor(203, 213, 225))
          n_shape.line.color.rgb = n_border
          n_shape.line.width = Pt(1.0)
          tf = n_shape.text_frame
          tf.word_wrap = True
          tf.margin_left = Inches(0.06)
          tf.margin_right = Inches(0.06)
          p = tf.paragraphs[0]
          p.text = node["text"]
          p.font.name = font_name
          p.font.size = Pt(node["styles"].get("fontSizePt", 9.5))
          p.font.bold = True
          p.font.color.rgb = parse_color_value(node["styles"].get("color"), brand_color_rgb if node.get("isHighlight") else RGBColor(30, 41, 59))
          p.alignment = PP_ALIGN.CENTER

    for fi in card.get("fItems", []):
      fir = fi["rect"]
      fi_st = fi.get("styles", {})
      fi_raw_bg = (fi_st.get("backgroundColor") or "").strip()
      if fir and fi_raw_bg and fi_raw_bg not in ("transparent", "rgba(0, 0, 0, 0)", "none"):
        fi_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(fir["left"]),
            Inches(fir["top"]),
            Inches(fir["width"]),
            Inches(fir["height"]),
        )
        fi_box.fill.solid()
        fi_box.fill.fore_color.rgb = parse_color_value(fi_raw_bg, RGBColor(248, 250, 252), bg_blend_rgb=blend_rgb)
        fi_border = parse_color_value(fi_st.get("borderColor"), None, bg_blend_rgb=blend_rgb)
        if fi_border:
          fi_box.line.color.rgb = fi_border
          fi_box.line.width = Pt(1.0)
        else:
          fi_box.line.fill.background()
      ic_r = fi.get("iconRect")
      tx_r = fi.get("textRect")
      ic_w = 0.0
      ic_left = 0.0
      if fi.get("icon") and ic_r:
        ic_w = max(0.38, ic_r.get("width", 0.40))
        ic_h = max(0.22, ic_r.get("height", 0.24))
        ic_left = ic_r.get("left", fir.get("left", r.get("left", 0.0) + 0.18))
        ic_top = ic_r.get("top", fir.get("top", r.get("top", 0.0) + 0.18))
        ic_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(ic_left),
            Inches(ic_top),
            Inches(ic_w),
            Inches(ic_h),
        )
        ic_shape.fill.solid()
        ic_bg = parse_color_value(
            fi.get("iconStyles", {}).get("backgroundColor") if fi.get("iconStyles") else None,
            brand_color_rgb,
            bg_blend_rgb=blend_rgb,
        )
        ic_shape.fill.fore_color.rgb = ic_bg
        ic_shape.line.fill.background()
        tf = ic_shape.text_frame
        tf.word_wrap = False
        tf.margin_left = Inches(0.04)
        tf.margin_right = Inches(0.04)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        p = tf.paragraphs[0]
        p.text = fi["icon"]
        p.font.name = font_name
        p.font.size = Pt(fi.get("iconStyles", {}).get("fontSizePt", 9.0) if fi.get("iconStyles") else 9.0)
        p.font.bold = True
        p.font.color.rgb = parse_color_value(
            fi.get("iconStyles", {}).get("color") if fi.get("iconStyles") else None,
            RGBColor(255, 255, 255),
        )
        p.alignment = PP_ALIGN.CENTER

      if fi.get("text"):
        card_r = r.get("left", 0.0) + r.get("width", 4.0) - 0.15
        if ic_w > 0:
          min_tx_left = ic_left + ic_w + 0.08
          tx_left = max(tx_r.get("left", min_tx_left) if tx_r else min_tx_left, min_tx_left)
        else:
          tx_left = tx_r.get("left", fir.get("left", r.get("left", 0.0) + 0.18)) if tx_r else fir.get("left", r.get("left", 0.0) + 0.18)

        tx_top = tx_r.get("top", fir.get("top", r.get("top", 0.0) + 0.18)) if tx_r else fir.get("top", r.get("top", 0.0) + 0.18)
        tx_w = max(0.8, card_r - tx_left)
        tx_h = max(0.24, tx_r.get("height", 0.24) if tx_r else 0.24)

        tx_box = slide.shapes.add_textbox(Inches(tx_left), Inches(tx_top), Inches(tx_w), Inches(tx_h))
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.02)
        tf.margin_right = Inches(0.02)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        p = tf.paragraphs[0]
        p.text = fi["text"]
        p.font.name = font_name
        p.font.size = Pt(fi.get("textStyles", {}).get("fontSizePt", 10.0) if fi.get("textStyles") else 10.0)
        p.font.color.rgb = parse_color_value(
            fi.get("textStyles", {}).get("color") if fi.get("textStyles") else None,
            RGBColor(17, 17, 21),
        )

    for sc in card.get("subCards", []):
      scr = sc["rect"]
      sc_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(scr["left"]), Inches(scr["top"]), Inches(scr["width"]), Inches(scr["height"]))
      sc_shape.fill.solid()
      sc_bg = parse_color_value(sc["styles"]["backgroundColor"], RGBColor(15, 23, 42) if is_dark else RGBColor(248, 250, 252), bg_blend_rgb=blend_rgb)
      sc_shape.fill.fore_color.rgb = sc_bg
      sc_border = parse_color_value(sc["styles"].get("borderColor"), RGBColor(51, 65, 85) if is_dark else RGBColor(226, 232, 240), bg_blend_rgb=blend_rgb)
      sc_shape.line.color.rgb = sc_border
      sc_shape.line.width = Pt(1.0)
      has_sc_left_accent = float(sc["styles"].get("borderLeftWidth") or 0) >= 3
      if has_sc_left_accent:
        sc_l_col = parse_color_value(sc["styles"].get("borderLeftColor"), brand_color_rgb)
        sc_l_bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(scr["left"] + 0.04),
            Inches(scr["top"] + 0.06),
            Inches(0.05),
            Inches(max(0.16, scr["height"] - 0.12)),
        )
        sc_l_bar.fill.solid()
        sc_l_bar.fill.fore_color.rgb = sc_l_col
        sc_l_bar.line.fill.background()
      tf = sc_shape.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.16 if has_sc_left_accent else 0.12)
      tf.margin_right = Inches(0.12)
      tf.margin_top = Inches(0.08)
      tf.margin_bottom = Inches(0.08)
      p_idx = 0
      if sc.get("head"):
        p0 = tf.paragraphs[0]
        p0.text = sc["head"] + (f"  [{sc['badge']}]" if sc.get("badge") else "")
        p0.font.name = font_name
        p0.font.size = Pt(11.0)
        p0.font.bold = True
        head_col_str = (sc.get("headStyles") or {}).get("color") or sc["styles"].get("color")
        p0.font.color.rgb = parse_color_value(head_col_str, RGBColor(248, 250, 252) if is_dark else brand_color_rgb)
        p_idx += 1
      sc_items = sc.get("items", [])
      for it in sc_items:
        p_it = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
        use_bullet = len(sc_items) > 1 and not it.startswith(("•", "-", "1", "2", "3", "KPI"))
        p_it.text = f"• {it}" if use_bullet else it
        p_it.font.name = font_name
        p_it.font.size = Pt(9.5)
        p_it.font.color.rgb = parse_color_value(sc["styles"].get("color"), RGBColor(203, 213, 225) if is_dark else RGBColor(30, 41, 59))
        p_idx += 1

    for cb in card.get("codeBlocks", []):
      cbr = cb["rect"]
      cb_h = Inches(cbr["height"]) if cbr.get("height", 0) > 0.3 else Inches(0.40)
      cb_box = slide.shapes.add_shape(
          MSO_SHAPE.ROUNDED_RECTANGLE,
          Inches(cbr["left"]),
          Inches(cbr["top"]),
          Inches(cbr["width"]),
          cb_h,
      )
      cb_box.fill.solid()
      cb_bg = parse_color_value(cb["styles"]["backgroundColor"], RGBColor(15, 23, 42) if is_dark else RGBColor(241, 245, 249), bg_blend_rgb=blend_rgb)
      cb_box.fill.fore_color.rgb = cb_bg
      cb_border = parse_color_value(cb["styles"].get("borderColor"), RGBColor(51, 65, 85) if is_dark else RGBColor(226, 232, 240))
      cb_box.line.color.rgb = cb_border
      cb_box.line.width = Pt(1.0)
      tf = cb_box.text_frame
      tf.word_wrap = True
      tf.margin_left = Inches(0.08)
      tf.margin_right = Inches(0.08)
      tf.margin_top = Inches(0.04)
      tf.margin_bottom = Inches(0.04)
      for p_idx, line in enumerate(cb["text"].split("\n")):
        p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = "Consolas"
        p.font.size = Pt(cb["styles"].get("fontSizePt", 9.0))
        p.font.color.rgb = parse_color_value(cb["styles"].get("color"), RGBColor(203, 213, 225) if is_dark else RGBColor(30, 41, 59))

    paras = card.get("paragraphs", [])
    if paras and not card.get("pipelineSteps"):
      if card.get("subCards") or card.get("hasCanvas") or card.get("hasSvgOrImg") or card.get("codeBlocks"):
        for para in paras:
          pr = para.get("rect", {})
          if not pr:
            continue
          p_box = slide.shapes.add_textbox(
              Inches(pr.get("left", r["left"] + 0.18)),
              Inches(pr.get("top", r["top"] + 0.40)),
              Inches(max(0.6, pr.get("width", r["width"] - 0.36))),
              Inches(max(0.24, pr.get("height", 0.30))),
          )
          tf = p_box.text_frame
          tf.word_wrap = True
          tf.margin_top = Inches(0.02)
          tf.margin_bottom = Inches(0.02)
          p = tf.paragraphs[0]
          raw_ptxt = para["text"]
          use_bullet = bool(para.get("isList")) and not raw_ptxt.startswith(("•", "-", "1", "2"))
          default_p_col = parse_color_value(para.get("styles", {}).get("color"), RGBColor(203, 213, 225) if is_dark else RGBColor(71, 85, 105))
          p_size = Pt(para.get("styles", {}).get("fontSizePt", 10.0))
          p_runs = para.get("runs")
          if p_runs and len(p_runs) > 1:
            p.text = ""
            if use_bullet:
              b_run = p.add_run()
              b_run.text = "• "
              b_run.font.name = font_name
              b_run.font.size = p_size
              b_run.font.color.rgb = default_p_col
            for pr_item in p_runs:
              run = p.add_run()
              run.text = pr_item.get("text", "")
              run.font.name = font_name
              run.font.size = p_size
              run.font.bold = bool(pr_item.get("isBold"))
              run.font.color.rgb = parse_color_value(pr_item.get("color"), default_p_col)
          else:
            p.text = f"• {raw_ptxt}" if use_bullet else raw_ptxt
            p.font.name = font_name
            p.font.size = p_size
            p.font.bold = bool(para.get("styles", {}).get("isBold"))
            p.font.color.rgb = default_p_col
      else:
        p_box = slide.shapes.add_textbox(
            c_left + Inches(0.18),
            curr_top,
            c_w - Inches(0.36),
            max(Inches(0.4), c_h - (curr_top - c_top) - Inches(0.10)),
        )
        tf = p_box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        for p_idx, para in enumerate(paras):
          p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
          raw_ptxt = para["text"]
          use_bullet = bool(para.get("isList")) and not raw_ptxt.startswith(("•", "-", "1", "2"))
          p_col = parse_color_value(para["styles"]["color"], RGBColor(226, 232, 240) if is_dark else RGBColor(17, 17, 21))
          p_size = Pt(para.get("styles", {}).get("fontSizePt", 10.5))
          p_runs = para.get("runs")
          if p_runs and len(p_runs) > 1:
            p.text = ""
            if use_bullet:
              b_run = p.add_run()
              b_run.text = "• "
              b_run.font.name = font_name
              b_run.font.size = p_size
              b_run.font.color.rgb = p_col
            for pr_item in p_runs:
              run = p.add_run()
              run.text = pr_item.get("text", "")
              run.font.name = font_name
              run.font.size = p_size
              run.font.bold = bool(pr_item.get("isBold"))
              run.font.color.rgb = parse_color_value(pr_item.get("color"), p_col)
          else:
            p.text = f"• {raw_ptxt}" if use_bullet else raw_ptxt
            p.font.name = font_name
            p.font.size = p_size
            p.font.bold = bool(para.get("styles", {}).get("isBold"))
            p.font.color.rgb = p_col

  for tbl in geom.get("tables", []):
    if tbl.get("num_rows", 0) > 0 and tbl.get("num_cols", 0) > 0:
      tr = tbl.get("rect", {})
      t_left_in = float(tr.get("left", 0.6))
      t_top_in = max(min_body_top, float(tr.get("top", 1.6))) if min_body_top > 0.0 else float(tr.get("top", 1.6))
      t_w_in = float(tr.get("width", 12.1))
      t_h_in = float(tr.get("height", 4.5))
      covered_by_card = any(
          float(c["rect"]["left"]) <= t_left_in + 0.25
          and float(c["rect"]["top"]) <= t_top_in + 0.25
          and (float(c["rect"]["left"]) + float(c["rect"]["width"])) >= (t_left_in + t_w_in - 0.25)
          for c in geom.get("cards", [])
      )
      tbl_copy = dict(tbl)
      tbl_copy["is_dark"] = is_dark
      tbl_copy["slide_bg_rgb"] = blend_rgb
      tbl_copy["skip_card_wrapper"] = bool(covered_by_card or not tbl.get("inCard"))
      build_styled_native_table(
          slide,
          tbl_copy,
          Inches(t_left_in),
          Inches(t_top_in),
          Inches(t_w_in),
          Inches(t_h_in),
          font_name=font_name,
      )

  for ch in geom.get("charts", []):
    build_styled_native_chart(
        slide,
        ch,
        font_name=font_name,
        is_dark=is_dark,
        slide_bg_rgb=blend_rgb,
    )

  for hl in geom.get("highlightBoxes", []):
    hl_r = hl.get("rect", {})
    hl_st = hl.get("styles", {})
    hl_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(hl_r.get("left", 0.8)),
        Inches(hl_r.get("top", 6.3)),
        Inches(hl_r.get("width", 11.733)),
        Inches(max(0.36, hl_r.get("height", 0.42))),
    )
    hl_shape.fill.solid()
    hl_bg = parse_color_value(
        hl_st.get("backgroundColor"),
        RGBColor(30, 41, 59) if is_dark else RGBColor(241, 245, 249),
        bg_blend_rgb=blend_rgb,
    )
    hl_border = parse_color_value(
        hl_st.get("borderColor"),
        RGBColor(59, 130, 246) if is_dark else brand_color_rgb,
        bg_blend_rgb=blend_rgb,
    )
    hl_shape.fill.fore_color.rgb = hl_bg
    hl_shape.line.color.rgb = hl_border
    tf = hl_shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.12)
    tf.margin_right = Inches(0.12)
    p = tf.paragraphs[0]
    p.text = hl.get("text", "")
    p.font.name = font_name
    p.font.size = Pt(hl_st.get("fontSizePt", 10.5))
    p.font.bold = True
    p.font.color.rgb = parse_color_value(hl_st.get("color"), RGBColor(147, 197, 253) if is_dark else brand_color_rgb)

  if geom.get("footerDividerTop"):
    f_div = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.55),
        Inches(float(geom["footerDividerTop"])),
        Inches(12.23),
        Inches(0.01),
    )
    f_div.fill.solid()
    f_div.fill.fore_color.rgb = parse_color_value(
        geom.get("footerDividerColor"),
        RGBColor(30, 41, 59) if is_dark else RGBColor(226, 232, 240),
        bg_blend_rgb=blend_rgb,
    )
    f_div.line.fill.background()

  for fn in geom.get("footnotes", []):
    fn_r = fn.get("rect", {})
    fn_st = fn.get("styles", {})
    is_right_fn = fn_r.get("left", 0.0) > 8.0 or fn_st.get("textAlign") == "right"
    fn_w_val = max(1.2, fn_r.get("width", 4.0) + 0.25)
    fn_l_val = max(0.4, (fn_r.get("left", 0.6) + fn_r.get("width", 4.0) - fn_w_val) if is_right_fn else fn_r.get("left", 0.6))
    fn_box = slide.shapes.add_textbox(
        Inches(fn_l_val),
        Inches(fn_r.get("top", 7.0)),
        Inches(fn_w_val),
        Inches(max(0.22, fn_r.get("height", 0.22))),
    )
    tf = fn_box.text_frame
    tf.word_wrap = False
    tf.margin_left = Inches(0.0)
    tf.margin_right = Inches(0.0)
    tf.margin_top = Inches(0.0)
    tf.margin_bottom = Inches(0.0)
    p = tf.paragraphs[0]
    p.text = fn.get("text", "")
    p.font.name = font_name
    p.font.size = Pt(fn_st.get("fontSizePt", 9.0))
    p.font.bold = bool(fn_st.get("isBold"))
    p.font.color.rgb = parse_color_value(
        fn_st.get("color"),
        RGBColor(100, 116, 139) if is_dark else RGBColor(100, 116, 139),
    )
    if is_right_fn:
      p.alignment = PP_ALIGN.RIGHT

  embed_extracted_slide_images(slide, geom, screenshot_source)


def ensure_slide_typography_consistency(
    slide,
    default_font: str = "Pretendard",
    default_color: Optional[RGBColor] = None,
    is_dark: bool = False,
) -> None:
  if default_color is None:
    default_color = RGBColor(241, 245, 249) if is_dark else RGBColor(17, 17, 21)

  for shape in slide.shapes:
    if shape.has_text_frame:
      tf = shape.text_frame
      tf.word_wrap = True
      for p in tf.paragraphs:
        if not p.font.name:
          p.font.name = default_font
        if not p.font.size:
          p.font.size = Pt(11.0)
        has_run_rgb = False
        for run in p.runs:
          try:
            if run.font.color and run.font.color.type == 1:
              has_run_rgb = True
              break
          except Exception:
            pass
        p_has_rgb = False
        try:
          if p.font.color and p.font.color.type == 1:
            p_has_rgb = True
          elif not has_run_rgb:
            p.font.color.rgb = default_color
            p_has_rgb = True
        except Exception:
          pass
        for run in p.runs:
          if not run.font.name:
            run.font.name = default_font
          if not run.font.size:
            run.font.size = p.font.size
          try:
            if (not run.font.color or run.font.color.type != 1) and p_has_rgb:
              run.font.color.rgb = p.font.color.rgb
          except Exception:
            pass
    elif shape.has_table:
      for row in shape.table.rows:
        for cell in row.cells:
          cell.text_frame.word_wrap = True
          for p in cell.text_frame.paragraphs:
            if not p.font.name:
              p.font.name = default_font
            if not p.font.size:
              p.font.size = Pt(10.0)
            has_run_rgb = False
            for run in p.runs:
              try:
                if run.font.color and run.font.color.type == 1:
                  has_run_rgb = True
                  break
              except Exception:
                pass
            p_has_rgb = False
            try:
              if p.font.color and p.font.color.type == 1:
                p_has_rgb = True
              elif not has_run_rgb:
                p.font.color.rgb = default_color
                p_has_rgb = True
            except Exception:
              pass
            for run in p.runs:
              if not run.font.name:
                run.font.name = default_font
              if not run.font.size:
                run.font.size = p.font.size
              try:
                if (not run.font.color or run.font.color.type != 1) and p_has_rgb:
                  run.font.color.rgb = p.font.color.rgb
              except Exception:
                pass


def refine_card_accent_bars(slide) -> None:
  cards = [sh for sh in slide.shapes if sh.height.inches > 0.8 and sh.width.inches > 1.5 and sh.shape_type == 1]
  thin_bars = [
      sh for sh in slide.shapes
      if 0.03 <= sh.height.inches <= 0.14
      and sh.width.inches > 1.2
      and sh.top.inches < 6.5
      and not (sh.has_text_frame and sh.text_frame.text.strip())
  ]
  for b in thin_bars:
    for c in cards:
      if abs(c.left.inches - b.left.inches) < 0.45 and abs(c.top.inches - b.top.inches) < 0.14 and abs(c.width.inches - b.width.inches) < 0.8:
        inset_x = Inches(0.24)
        inset_y = Inches(0.06)
        b.left = c.left + inset_x
        b.top = c.top + inset_y
        b.width = max(Inches(0.5), c.width - (inset_x * 2))
        b.height = Inches(0.06)
        try:
          b.adjustments[0] = 0.5
        except Exception:
          pass
        try:
          b.line.fill.background()
        except Exception:
          pass

        for tb in slide.shapes:
          if tb.has_text_frame and not tb.has_table and tb != c and tb != b:
            if c.left.inches - 0.1 <= tb.left.inches <= c.left.inches + c.width.inches + 0.1:
              if c.top.inches - 0.05 <= tb.top.inches < c.top.inches + 0.16:
                tb.top = c.top + Inches(0.16)
        break


def audit_and_resolve_slide_collisions(slide) -> None:
  bg_shapes = [
      s for s in slide.shapes
      if s.shape_type == 1 and s.left.inches <= 0.1 and s.top.inches <= 0.1 and s.width.inches >= 13.0 and s.height.inches >= 7.0
  ]
  if len(bg_shapes) > 1:
    for dup in bg_shapes[1:]:
      try:
        slide.shapes._spTree.remove(dup._element)
      except Exception:
        pass

  seen_signatures = set()
  for s in list(slide.shapes):
    if s.has_text_frame and not s.has_table:
      txt = s.text_frame.text.strip()
      if txt and len(txt) > 5:
        sig = (round(s.left.inches, 2), round(s.top.inches, 2), txt)
        if sig in seen_signatures:
          try:
            slide.shapes._spTree.remove(s._element)
          except Exception:
            pass
          continue
        seen_signatures.add(sig)

  shapes = [
      s for s in slide.shapes
      if not (s.left.inches <= 0.1 and s.top.inches <= 0.1 and s.width.inches >= 13.0)
  ]

  for i in range(len(shapes)):
    s1 = shapes[i]
    l1, t1, w1, h1 = s1.left.inches, s1.top.inches, s1.width.inches, s1.height.inches
    r1, b1 = l1 + w1, t1 + h1

    for j in range(len(shapes)):
      if i == j:
        continue
      s2 = shapes[j]
      l2, t2, w2, h2 = s2.left.inches, s2.top.inches, s2.width.inches, s2.height.inches
      r2, b2 = l2 + w2, t2 + h2

      if l1 <= l2 + 0.08 and t1 <= t2 + 0.08 and r1 >= r2 - 0.30 and b1 >= b2 - 0.15:
        continue
      if l2 <= l1 + 0.08 and t2 <= t1 + 0.08 and r2 >= r1 - 0.30 and b2 >= b1 - 0.15:
        continue

      if abs(t1 - t2) > 0.12:
        continue

      x_overlap = min(r1, r2) - max(l1, l2)
      y_overlap = min(b1, b2) - max(t1, t2)

      if x_overlap > 0.02 and y_overlap > 0.05:
        if s1.shape_type == 1 and s2.shape_type == 17 and w1 < 1.8 and h1 < 0.55 and l1 < l2 and l2 > l1 + w1 * 0.4:
          new_left = r1 + 0.06
          shift = new_left - l2
          s2.left = Inches(new_left)
          s2.width = Inches(max(0.6, w2 - shift))
          l2, r2 = new_left, new_left + max(0.6, w2 - shift)

        elif s1.shape_type == 17 and s2.shape_type == 1 and w2 < 2.2 and h2 < 0.50 and l1 + 0.4 < l2:
          new_w = max(0.6, l2 - l1 - 0.06)
          s1.width = Inches(new_w)
          w1, r1 = new_w, l1 + new_w

        elif s1.shape_type == 17 and s2.shape_type == 17 and l1 + 0.4 < l2:
          new_w = max(0.6, l2 - l1 - 0.06)
          s1.width = Inches(new_w)
          w1, r1 = new_w, l1 + new_w


def harmonize_deck_presentation_fidelity(
    prs: Presentation,
    default_font: str = "Pretendard",
    default_bg_hex: Optional[str] = None,
) -> None:
  is_dark = False
  deck_bg_rgb = RGBColor(255, 255, 255)
  if default_bg_hex:
    bg_tuple = _hex_to_rgb_tuple(default_bg_hex, (255, 255, 255))
    deck_bg_rgb = RGBColor(*bg_tuple)
    is_dark = (bg_tuple[0] * 0.299 + bg_tuple[1] * 0.587 + bg_tuple[2] * 0.114) < 128
  elif len(prs.slides) > 0 and len(prs.slides[0].shapes) > 0:
    for sh in prs.slides[0].shapes:
      if sh.shape_type == 1 and sh.left.inches <= 0.1 and sh.top.inches <= 0.1 and sh.width.inches >= 13.0:
        try:
          col = sh.fill.fore_color.rgb
          deck_bg_rgb = col
          is_dark = (col[0] * 0.299 + col[1] * 0.587 + col[2] * 0.114) < 128
          break
        except Exception:
          pass

  for slide in prs.slides:
    slide_bg_rgb = deck_bg_rgb
    slide_is_dark = is_dark
    for sh in slide.shapes:
      if sh.shape_type == 1 and sh.left.inches <= 0.1 and sh.top.inches <= 0.1 and sh.width.inches >= 13.0 and sh.height.inches >= 7.0:
        try:
          col = sh.fill.fore_color.rgb
          if col is not None:
            slide_bg_rgb = col
            slide_is_dark = (col[0] * 0.299 + col[1] * 0.587 + col[2] * 0.114) < 128
            break
        except Exception:
          pass

    default_color = RGBColor(241, 245, 249) if slide_is_dark else RGBColor(17, 17, 21)
    ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=slide_is_dark)
    ensure_slide_typography_consistency(slide, default_font=default_font, default_color=default_color, is_dark=slide_is_dark)
    refine_card_accent_bars(slide)
    audit_and_resolve_slide_collisions(slide)

    for s in slide.shapes:
      if s.has_text_frame and not s.has_table:
        for p in s.text_frame.paragraphs:
          if not p.font.name:
            p.font.name = default_font
          if not p.font.size:
            p.font.size = Pt(10.5)
          for r in p.runs:
            if not r.font.name:
              r.font.name = default_font
            if not r.font.size:
              r.font.size = p.font.size


def build_native_slide_from_code(
    slide,
    prs: Presentation,
    code_str: Optional[str],
    screenshot_source: Optional[Union[bytes, str]] = None,
    design_tokens: Optional[Dict[str, Any]] = None,
    slide_dom_data: Optional[Dict[str, Any]] = None,
    slide_geometry: Optional[Dict[str, Any]] = None,
    slide_index: int = 1,
) -> None:
  font_family = (design_tokens or {}).get("font_name", "Pretendard")
  text_color_hex = (design_tokens or {}).get("text_main_hex", "#111115")
  default_text_color = _hex_to_pptx_color(text_color_hex, default=(17, 17, 21))
  brand_color = _hex_to_pptx_color((design_tokens or {}).get("brand_color_hex"), default=(37, 99, 235))

  geom_bg = (slide_geometry or {}).get("slideBgColor")
  token_bg = (design_tokens or {}).get("bg_color_hex")
  slide_bg_rgb = parse_color_value(geom_bg or token_bg, default=RGBColor(255, 255, 255))
  is_dark = (slide_bg_rgb[0] * 0.299 + slide_bg_rgb[1] * 0.587 + slide_bg_rgb[2] * 0.114) < 128
  if default_text_color == RGBColor(17, 17, 21) and is_dark:
    default_text_color = RGBColor(241, 245, 249)

  ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)

  def _build_deterministic_fallback() -> None:
    if slide_geometry:
      build_slide_from_geometry(
          slide,
          slide_geometry,
          font_name=font_family,
          brand_color_rgb=brand_color,
          screenshot_source=screenshot_source,
      )
      ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
      ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
      refine_card_accent_bars(slide)
      audit_and_resolve_slide_collisions(slide)
      return
    if slide_dom_data and slide_dom_data.get("tables"):
      t_data = slide_dom_data["tables"][0]
      build_styled_native_table(slide, t_data, Inches(0.8), Inches(1.95), Inches(11.733), Inches(4.5), font_name=font_family)
      ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
      ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)

  if not code_str:
    _build_deterministic_fallback()
    return

  clean_code = code_str
  if "```python" in clean_code:
    clean_code = clean_code.split("```python", 1)[1].split("```", 1)[0].strip()
  elif "```" in clean_code:
    clean_code = clean_code.split("```", 1)[1].split("```", 1)[0].strip()

  tbl_dom = (
      slide_geometry.get("tables", [{}])[0]
      if slide_geometry and slide_geometry.get("tables")
      else (slide_dom_data.get("tables", [{}])[0] if slide_dom_data and slide_dom_data.get("tables") else None)
  )

  scope = {
      "prs": prs,
      "slide": slide,
      "slide_index": slide_index,
      "slide_geometry": slide_geometry,
      "slide_dom_data": slide_dom_data,
      "design_tokens": design_tokens,
      "table_data": tbl_dom,
      "charts_data": (slide_geometry or {}).get("charts", []),
      "Presentation": Presentation,
      "Inches": Inches,
      "Pt": Pt,
      "RGBColor": RGBColor,
      "MSO_SHAPE": MSO_SHAPE,
      "PP_ALIGN": PP_ALIGN,
      "MSO_ANCHOR": MSO_ANCHOR,
      "build_styled_native_table": build_styled_native_table,
      "build_styled_native_chart": build_styled_native_chart,
      "apply_semantic_styles_to_table": apply_semantic_styles_to_table,
      "build_slide_from_geometry": build_slide_from_geometry,
  }

  try:
    exec(clean_code, scope)
    idx_func_name = f"build_slide_{slide_index}"
    func = scope.get(idx_func_name) or scope.get("build_slide")
    if callable(func):
      sig = inspect.signature(func)
      if len(sig.parameters) == 1:
        func(slide)
      elif len(sig.parameters) >= 3:
        func(prs, slide, slide_index)
      else:
        func(prs, slide)

      ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)

      if tbl_dom:
        has_table_shape = any(s.has_table for s in slide.shapes)
        if not has_table_shape and tbl_dom.get("rows"):
          tbl_r = tbl_dom.get("rect", {})
          t_left = Inches(tbl_r.get("left", 0.6))
          t_top = Inches(tbl_r.get("top", 1.5))
          t_w = Inches(min(tbl_r.get("width", 12.0), 12.2))
          t_h = Inches(tbl_r.get("height", 3.0))
          tbl_copy = dict(tbl_dom)
          tbl_copy["is_dark"] = is_dark
          tbl_copy["slide_bg_rgb"] = (slide_bg_rgb[0], slide_bg_rgb[1], slide_bg_rgb[2])
          build_styled_native_table(slide, tbl_copy, t_left, t_top, t_w, t_h, font_name=font_family)
        else:
          for s in slide.shapes:
            if s.has_table:
              apply_semantic_styles_to_table(
                  s,
                  tbl_dom,
                  font_name=font_family,
                  is_dark=is_dark,
                  slide_bg_rgb=(slide_bg_rgb[0], slide_bg_rgb[1], slide_bg_rgb[2]),
              )

      if slide_geometry and slide_geometry.get("charts"):
        has_chart_shape = any(getattr(s, "has_chart", False) for s in slide.shapes)
        if not has_chart_shape:
          for ch in slide_geometry["charts"]:
            build_styled_native_chart(
                slide,
                ch,
                font_name=font_family,
                is_dark=is_dark,
                slide_bg_rgb=(slide_bg_rgb[0], slide_bg_rgb[1], slide_bg_rgb[2]),
            )

      if slide_geometry:
        embed_extracted_slide_images(slide, slide_geometry, screenshot_source)

      ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
      refine_card_accent_bars(slide)
      audit_and_resolve_slide_collisions(slide)
    elif len(slide.shapes) > 1:
      if slide_geometry:
        embed_extracted_slide_images(slide, slide_geometry, screenshot_source)
      ensure_slide_canvas_background(slide, prs=prs, bg_color=slide_bg_rgb, is_dark=is_dark)
      ensure_slide_typography_consistency(slide, default_font=font_family, default_color=default_text_color, is_dark=is_dark)
      refine_card_accent_bars(slide)
      audit_and_resolve_slide_collisions(slide)
    else:
      _build_deterministic_fallback()
  except Exception as e:
    logger.warning("[Execution Error in custom slide builder code]: %s", e)
    _build_deterministic_fallback()
