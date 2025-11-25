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
    time.sleep(1)
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
    doc = spec["document"]  # 반드시 document root dict

    # Title
    if "title" in doc:
        insert_role_and_style(hwp, {"title": doc["title"]}, TITLE_STYLE, "title")
    # Recipient
    if "recipient" in doc:
        insert_role_and_style(hwp, {"recipient": doc["recipient"]}, RECIPIENT_STYLE, "recipient")
    # Body
    if "body" in doc:
        insert_role_and_style(hwp, {"body": doc["body"]["content"]}, BODY_STYLE, "body")
    # Table
    if "table" in doc:
        table_info = doc["table"]
        insert_table_and_style(
            hwp,
            table_info["data"],
            cell_styles=table_info.get("cell_styles"),   # style 확장 적용
            cell_bg_colors=table_info.get("cell_bg_colors"),
            col_aligns=table_info["style"].get("cell_align"),
        )
    # Footer
    if "footer" in doc:
        insert_role_and_style(hwp, {"footer": doc["footer"]["content"]}, FOOTER_STYLE, "footer")
    hwp.save_as(filename)
    hwp.quit()
