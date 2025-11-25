import json
import yaml

from pyhwpx import Hwp
import re

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    red = int(hex_color[0:2], 16)
    green = int(hex_color[2:4], 16)
    blue = int(hex_color[4:6], 16)
    return (red, green, blue)

def set_alignment(hwp, align):
    if align == "center":
        hwp.ParagraphShapeAlignCenter()
    elif align == "left":
        hwp.ParagraphShapeAlignLeft()
    elif align == "right":
        hwp.ParagraphShapeAlignRight()
    elif align == "justify":
        hwp.ParagraphShapeAlignJustify()

def parse_segments(text):
    patterns = [
        (r'\*\*(.+?)\*\*', 'bold'),
        (r'_(.+?)_', 'italic'),
        (r'\*(.+?)\*', 'italic'),
        (r'<u>(.+?)</u>', 'underline')
    ]
    segments = []
    i = 0
    while i < len(text):
        match = None
        for pat, typ in patterns:
            m = re.match(pat, text[i:])
            if m:
                match = (m, typ)
                break
        if match:
            m, typ = match
            val = m.group(1)
            pre = text[:i]
            if pre:
                segments.append({"text": pre})
            segment = {"text": val}
            if typ == "bold":
                segment["bold"] = True
            elif typ == "italic":
                segment["italic"] = True
            segments.append(segment)
            text = text[i+len(m.group(0)):]
            i = 0
        else:
            i += 1
    if text:
        segments.append({"text": text})
    return segments

def insert_role_and_style(hwp, content_dict, styles, role):
    base_opts = {
        "FaceName": styles[role]["FaceName"],
        "Height": styles[role]["Height"],
        "Bold": styles[role]["Bold"],
        "Italic": False,
    }
    set_alignment(hwp, styles[role]["Align"])
    segments = parse_segments(content_dict[role])
    for seg in segments:
        opts = base_opts.copy()
        opts["Bold"]   = seg.get("bold", base_opts["Bold"])
        opts["Italic"] = seg.get("italic", base_opts["Italic"])
        hwp.set_font(**opts)
        hwp.insert_text(seg["text"] if isinstance(seg, dict) else seg)
        hwp.set_font(**base_opts)
    hwp.insert_text("\r\n")

def insert_table_and_style(hwp, table_data, cell_styles=None, cell_bg_colors=None, col_aligns=None):
    rows, cols = len(table_data), len(table_data[0])
    hwp.create_table(rows, cols, treat_as_char=True)
    
    for r_idx, row in enumerate(table_data):
        for c_idx, val in enumerate(row):
            style = cell_styles[r_idx][c_idx] if cell_styles else {}
            bg    = cell_bg_colors[r_idx][c_idx] if cell_bg_colors else None
            align = col_aligns[c_idx] if col_aligns else "left"
            hwp.set_font(
                FaceName=style.get("FaceName", "바탕체"),
                Height=style.get("Height", 11),
                Bold=style.get("Bold", False)
            )
            if align=="center": hwp.TableCellAlignCenterCenter()
            elif align=="right": hwp.TableCellAlignRightCenter()
            else: hwp.TableCellAlignLeftCenter()
            if bg:
                if isinstance(bg, str) and bg.startswith('#'):
                    red, green, blue = hex_to_rgb(bg)
                    hwp.gradation_on_cell([(red, green, blue)])
                else:
                    hwp.gradation_on_cell([bg])
            hwp.insert_text(str(val))
            if c_idx < cols - 1:
                hwp.TableRightCell()
        if r_idx < rows - 1:
            hwp.TableLowerCell()
            for _ in range(cols - 1):
                hwp.TableLeftCell()
    hwp.MoveDown()

def generate_hwp_from_spec(spec, filename="output.hwpx"):
    hwp = Hwp()
    doc = spec["document"] if "document" in spec else spec

    for key, value in doc.items():
        # 텍스트/문단류
        if isinstance(value, str) or (isinstance(value, dict) and "content" in value):
            # 스타일 추출 - 필요하면 value.get("style") 등
            text = value if isinstance(value, str) else value["content"]
            # 기본 스타일 또는 key별 heuristic (원하면 key마다 스타일 분기 가능)
            style = heuristic_style_for_key(key)
            insert_role_and_style(hwp, {key: text}, style, key)
        # 표류
        elif isinstance(value, dict) and "data" in value:
            table_style = value.get("style", {})
            cell_aligns = table_style.get("cell_align")
            # cell_styles, bg_colors 등 확장
            insert_table_and_style(
                hwp,
                value["data"],
                cell_styles=None,  # 추가할 부분
                cell_bg_colors=[[table_style.get("header_bg")]*len(value["data"][0])] + [[None]*len(value["data"][0])]*(len(value["data"])-1),
                col_aligns=cell_aligns
            )
        # 기타(이미지, 구분선, 추가 미래기능)
        else:
            pass  # type에 따라 확장 가능

    hwp.save_as(filename)
    hwp.quit()

def heuristic_style_for_key(key):
    # 예시: key가 'title', 'header'면 굵게, 크게 등
    style_map = {
        "title": {"FaceName":"돋움", "Height":20, "Bold":True, "Align":"center"},
        "footer": {"FaceName":"바탕체", "Height":10, "Bold":False, "Align":"center"},
        "body": {"FaceName":"바탕체", "Height":11, "Bold":False, "Align":"left"},
        # 기타 key/styles 필요시 계속 추가
    }
    # 값 없을 경우 기본값
    return {key: style_map.get(key, {"FaceName":"바탕체", "Height":11, "Bold":False, "Align":"left"})}
