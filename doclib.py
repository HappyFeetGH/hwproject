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
    cell_styles=None,        # [r][c] 기본 폰트/크기/Bold 등
    cell_bg_colors=None,     # [r][c] "#RRGGBB" 또는 None (옵션)
    col_aligns=None,         # [c] "left"/"center"/"right"
    cell_segments=None,      # [r][c] -> [{"text","style"}, ...]
    cell_merges=None         # [r][c] -> {"colSpan","rowSpan","bgColor","width","height"}
):
    rows, cols = len(table_data), len(table_data[0])
    hwp.create_table(rows, cols, treat_as_char=True)

    for r_idx, row in enumerate(table_data):
        for c_idx, val in enumerate(row):
            base_style = cell_styles[r_idx][c_idx] if cell_styles else {}
            merge_info = cell_merges[r_idx][c_idx] if cell_merges else {}
            segs = cell_segments[r_idx][c_idx] if (cell_segments and cell_segments[r_idx][c_idx]) else None

            # 1) 정렬
            align = col_aligns[c_idx] if col_aligns else "left"
            if align == "center":
                hwp.TableCellAlignCenterCenter()
            elif align == "right":
                hwp.TableCellAlignRightCenter()
            else:
                hwp.TableCellAlignLeftCenter()

            # 2) 배경색: cell_bg_colors보다 parser에서 온 bgColor 우선
            bg = merge_info.get("bgColor")
            if bg is None and cell_bg_colors:
                bg = cell_bg_colors[r_idx][c_idx]
            if bg:
                if isinstance(bg, str) and bg.startswith("#"):
                    red, green, blue = hex_to_rgb(bg)
                    hwp.gradation_on_cell([(red, green, blue)])  # [web:281]
                else:
                    hwp.gradation_on_cell([bg])

            # 3) 셀 크기: parser에서 받은 width/height(HwpUnit 기준)를 훅으로 전달
            cell_w = merge_info.get("width")
            cell_h = merge_info.get("height")
            if cell_w or cell_h:
                set_current_cell_size(hwp, cell_w, cell_h)   # 아래 helper에서 구현 훅

            # 4) 셀 내용: segment 단위 스타일 적용
            if segs:
                for seg in segs:
                    s = seg.get("style", {})
                    # segment style = 기본 셀 스타일 + run 스타일 override
                    font_opts = {
                        "FaceName": s.get("FaceName", base_style.get("FaceName", "바탕체")),
                        "Height":  s.get("Height",  base_style.get("Height", 11)),
                        "Bold":    s.get("Bold",    base_style.get("Bold", False)),
                    }
                    hwp.set_font(**font_opts)
                    hwp.insert_text(seg.get("text", ""))
            else:
                # segment 정보가 없으면 기존처럼 한 번에 입력
                hwp.set_font(
                    FaceName=base_style.get("FaceName", "바탕체"),
                    Height=base_style.get("Height", 11),
                    Bold=base_style.get("Bold", False)
                )
                hwp.insert_text(str(val))

            # 5) 다음 셀로 이동
            if c_idx < cols - 1:
                hwp.TableRightCell()

        if r_idx < rows - 1:
            hwp.TableLowerCell()
            for _ in range(cols - 1):
                hwp.TableLeftCell()

    hwp.MoveDown()


def set_current_cell_size(hwp, width_hu=None, height_hu=None):
    """
    현재 커서가 위치한 셀의 크기를 HwpUnit 기준으로 맞추는 훅.
    width_hu, height_hu 는 hwpx parser에서 뽑은 cellSz width/height (HwpUnit).
    둘 중 None 인 값은 변경하지 않는다.
    """
    if width_hu is None and height_hu is None:
        return

    act = hwp.hwp.CreateAction("TablePropertyDialog")
    pset = act.CreateSet()
    cell_set = pset.CreateItemSet("ShapeTableCell", "Cell")

    # 현재 셀의 기본값 로드
    act.GetDefault(pset)

    if width_hu is not None:
        cell_set.SetItem("Width", width_hu)   # 이미 HwpUnit이면 그대로
    if height_hu is not None:
        cell_set.SetItem("Height", height_hu)

    try:
        act.Execute(pset)
    except Exception:
        # 셀 크기 변경 실패시 조용히 패스
        pass


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
    """
    spec: parsed_spec.json을 그대로 로드한 dict
    document 안의 항목을 '순서대로' 읽어서 문단/표를 생성한다.
    """
    hwp = Hwp()
    doc = spec["document"]

    for key, node in doc.items():
        # 표 노드: data가 있고 list면 table
        if isinstance(node, dict) and isinstance(node.get("data"), list):
            insert_table_and_style(
                hwp,
                node["data"],
                cell_styles=node.get("cell_styles"),
                cell_bg_colors=None,                 # 필요시 추가
                col_aligns=None,                     # 필요시 style에서 파생
                cell_segments=node.get("cell_segments"),
                cell_merges=node.get("cell_merges"),
            )
        # 문단 노드
        elif isinstance(node, dict) and ("content" in node or "segments" in node):
            insert_paragraph_from_node(hwp, node)
        # 그 외 타입(미래 확장)은 일단 스킵 혹은 로그

    hwp.save_as(filename)
    hwp.quit()