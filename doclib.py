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

def insert_table_and_style(
    hwp,
    table_data,
    cell_styles=None,
    cell_bg_colors=None,
    col_aligns=None,
    cell_segments=None,
    cell_merges=None,
    cell_nested=None,   # [r][c] -> nested table list
    nested=False        # 셀 안에 들어가는 표면 True
):
    rows, cols = len(table_data), len(table_data[0])
    hwp.create_table(rows, cols, treat_as_char=True)  # 커서 위치에 표 컨트롤 삽입[web:100]

    for r_idx, row in enumerate(table_data):
        for c_idx, _ in enumerate(row):
            base_style = cell_styles[r_idx][c_idx] if cell_styles else {}
            segs = (cell_segments and cell_segments[r_idx][c_idx]) or None
            merge_info = (cell_merges and cell_merges[r_idx][c_idx]) or {}
            nested_tbls = (cell_nested and cell_nested[r_idx][c_idx]) or []

            # 정렬
            align = col_aligns[c_idx] if col_aligns else "left"
            if align == "center":
                hwp.TableCellAlignCenterCenter()
            elif align == "right":
                hwp.TableCellAlignRightCenter()
            else:
                hwp.TableCellAlignLeftCenter()

            # 배경색 (parser에서 온 bgColor 우선)
            bg = merge_info.get("bgColor")
            if bg is None and cell_bg_colors:
                bg = cell_bg_colors[r_idx][c_idx]
            if bg:
                if isinstance(bg, str) and bg.startswith("#"):
                    red, green, blue = hex_to_rgb(bg)
                    hwp.gradation_on_cell([(red, green, blue)])  # 현재 셀 배경[web:281]
                else:
                    hwp.gradation_on_cell([bg])

            # 셀 텍스트 (segment 단위 적용)
            if segs:
                for seg in segs:
                    s = seg.get("style", {})
                    font_opts = {
                        "FaceName": s.get("FaceName", base_style.get("FaceName", "바탕체")),
                        "Height":  s.get("Height",  base_style.get("Height", 11)),
                        "Bold":    s.get("Bold",    base_style.get("Bold", False)),
                    }
                    hwp.set_font(**font_opts)
                    hwp.insert_text(seg.get("text", ""))
            else:
                hwp.set_font(
                    FaceName=base_style.get("FaceName", "바탕체"),
                    Height=base_style.get("Height", 11),
                    Bold=base_style.get("Bold", False),
                )
                hwp.insert_text(str(table_data[r_idx][c_idx]))

            # 셀 안에 중첩 표들 생성
            for inner in nested_tbls:
                # 셀 안에서 줄바꿈 후, 그 위치에 create_table 호출[web:89][web:100]
                hwp.insert_text("\r\n")
                insert_table_and_style(
                    hwp,
                    inner["data"],
                    cell_styles=inner.get("cell_styles"),
                    cell_bg_colors=None,
                    col_aligns=None,
                    cell_segments=inner.get("cell_segments"),
                    cell_merges=inner.get("cell_merges"),
                    cell_nested=inner.get("cell_nested"),
                    nested=True,       # ← 중요: 중첩 표
                )
                # nested=True 이므로 내부 호출은 TableOutCell까지만 하고 빠져나옴

            if c_idx < cols - 1:
                hwp.TableRightCell()
        if r_idx < rows - 1:
            hwp.TableLowerCell()
            for _ in range(cols - 1):
                hwp.TableLeftCell()

    # 표 종료 시 커서 정리
    if not nested:
        # 셀 안에 만든 표이므로, 표 컨트롤 밖(같은 셀 텍스트 위치)으로만 빠져나온다.[web:76]
        hwp.MoveDown()


def set_current_cell_size(hwp, width_hu, height_hu):
    # 셀 블록 선택
    hwp.HAction.Run("TableCellBlock")
    
    # 파라미터 셋업 (TablePropertyDialog)
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    
    if width_hu:
        pset.HSet.Item("ShapeTableCell").SetItem("Width", int(width_hu))
    if height_hu:
        pset.HSet.Item("ShapeTableCell").SetItem("Height", int(height_hu))
        
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
    hwp.HAction.Run("Cancel")




def generate_hwp_from_spec(spec, filename="output.hwpx"):
    """
    spec 이 parsed_spec(json) 형식이면 `insert_paragraph_from_node` /
    `insert_table_and_style`를 사용하고,
    옛 role+styles 형식이면 insert_role_and_style을 사용한다.
    """
    hwp = Hwp()
    doc = spec.get("document", spec)

    # 첫 항목으로 스펙 형태 판별
    first_val = next(iter(doc.values()))
    is_parsed = isinstance(first_val, dict) and (
        "content" in first_val or "data" in first_val
    )

    if is_parsed:
        # === 새 스펙 경로 (parser 출력) ===
        for _, node in doc.items():
            # 표 블록
            if isinstance(node, dict) and isinstance(node.get("data"), list):
                insert_table_and_style(
                    hwp,
                    node["data"],
                    cell_styles=node.get("cell_styles"),
                    cell_bg_colors=None,
                    col_aligns=node.get("style", {}).get("cell_align"),
                    cell_segments=node.get("cell_segments"),
                    cell_merges=node.get("cell_merges"),
                    cell_nested=node.get("cell_nested"),
                    nested=False,
                )
            # 문단 블록
            elif isinstance(node, dict) and ("content" in node or "segments" in node):
                insert_paragraph_from_node(hwp, node)
            # 기타 타입은 필요시 확장
    else:
        # === 옛 role+styles 경로 (기존 코드 유지) ===
        for key, value in doc.items():
            if isinstance(value, str) or (isinstance(value, dict) and "content" in value):
                text = value if isinstance(value, str) else value["content"]
                styles = heuristic_style_for_key(key)
                insert_role_and_style(hwp, {key: text}, styles, key)
            elif isinstance(value, dict) and "data" in value:
                table_style = value.get("style", {})
                cell_aligns = table_style.get("cell_align")
                insert_table_and_style(
                    hwp,
                    value["data"],
                    cell_styles=None,
                    cell_bg_colors=[[table_style.get("header_bg")] * len(value["data"][0])]
                                   + [[None] * len(value["data"][0])] * (len(value["data"]) - 1),
                    col_aligns=cell_aligns,
                )

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


from pyhwpx import Hwp

def insert_paragraph_from_node(hwp, node):
    """
    node: {"content": str, "style": {...}, "segments": [...]}
    segments가 있으면 segment 스타일 기준으로, 없으면 style 하나만으로 출력.
    """
    content = node.get("content", "")
    segments = node.get("segments") or []

    # segments가 있으면 run 단위로 적용
    if segments:
        for seg in segments:
            s_style = seg.get("style", {})
            hwp.set_font(
                FaceName=s_style.get("FaceName", node["style"].get("FaceName", "바탕체")),
                Height=s_style.get("Height",  node["style"].get("Height", 11)),
                Bold=s_style.get("Bold",      node["style"].get("Bold", False))
            )
            hwp.insert_text(seg.get("text", ""))
    else:
        base = node.get("style", {})
        hwp.set_font(
            FaceName=base.get("FaceName", "바탕체"),
            Height=base.get("Height", 11),
            Bold=base.get("Bold", False)
        )
        hwp.insert_text(content)

    # 문단 정렬
    align = node.get("style", {}).get("Align", "left")
    if align == "center":
        hwp.ParagraphShapeAlignCenter()
    elif align == "right":
        hwp.ParagraphShapeAlignRight()
    elif align == "justify":
        hwp.ParagraphShapeAlignJustify()
    else:
        hwp.ParagraphShapeAlignLeft()

    hwp.insert_text("\r\n")


def generate_hwp_from_parsed_spec(spec, filename="output.hwpx"):
    from pyhwpx import Hwp
    hwp = Hwp()
    doc = spec["document"]

    for _, node in doc.items():
        if isinstance(node, dict) and isinstance(node.get("data"), list):
            # 표 블록
            insert_table_and_style(
                hwp,
                node["data"],
                cell_styles=node.get("cell_styles"),
                cell_bg_colors=None,
                col_aligns=node.get("style", {}).get("cell_align"),
                cell_segments=node.get("cell_segments"),
                cell_merges=node.get("cell_merges"),
                cell_nested=node.get("cell_nested"),  # ← 여기
                nested=False,
            )
        elif isinstance(node, dict) and ("content" in node or "segments" in node):
            insert_paragraph_from_node(hwp, node)
        # 기타 타입은 필요시 확장

    hwp.save_as(filename)
    hwp.quit()
