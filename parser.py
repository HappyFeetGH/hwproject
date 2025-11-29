import zipfile
import xml.etree.ElementTree as ET
import json

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head"
}
HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"

def parse_styles_from_header(zf: zipfile.ZipFile):
    """
    header.xml의 refList에서
    - fontfaces: 언어별 폰트 id -> faceName
    - charPr: charShapeID -> (Height, FaceName)
    - paraPr: paraShapeID -> Align
    을 추출한다.
    """
    para_shapes = {}   # id -> {"Align": ...}
    char_shapes = {}   # id -> {"Height": ..., "FaceName": ...}

    try:
        with zf.open("Contents/header.xml") as f:
            tree = ET.parse(f)
    except KeyError:
        return para_shapes, char_shapes

    root = tree.getroot()
    ref_list = root.find("hh:refList", NS)
    if ref_list is None:
        return para_shapes, char_shapes

    # 1) fontfaces: 언어(HANGUL/LATIN 등)별 id->face 맵 만들기 [web:150]
    font_map = {}  # {"HANGUL": {0:"함초롬돋움", 1:"함초롬바탕"}, ...}
    fontfaces = ref_list.find("hh:fontfaces", NS)
    if fontfaces is not None:
        for ff in fontfaces.findall("hh:fontface", NS):
            lang = ff.get("lang")  # ex) "HANGUL"
            lang_fonts = {}
            for font in ff.findall("hh:font", NS):
                fid = int(font.get("id"))
                face = font.get("face")
                lang_fonts[fid] = face
            font_map[lang] = lang_fonts

    # 2) charPr: 높이 + fontRef.hangul -> FaceName 매핑 [web:157]
    char_props = ref_list.find("hh:charProperties", NS)
    if char_props is not None:
        for char_pr in char_props.findall("hh:charPr", NS):
            cid = int(char_pr.get("id"))
            height_raw = char_pr.get("height")
            height_pt = int(height_raw) / 100.0 if height_raw and height_raw.isdigit() else None

            # 기본값: bold=False
            is_bold = False
            # 하위 엘리먼트 중에 <hh:bold>가 있으면 True [web:163]
            for child in char_pr:
                if child.tag == "{" + NS["hh"] + "}bold":
                    is_bold = True
                    break

            # fontRef 하위 요소에서 hangul 폰트 id 가져오기
            face_name = None
            font_ref = char_pr.find("hh:fontRef", NS)
            if font_ref is not None:
                hangul_id = font_ref.get("hangul")
                if hangul_id is not None and hangul_id.isdigit():
                    hid = int(hangul_id)
                    hangul_fonts = font_map.get("HANGUL", {})
                    face_name = hangul_fonts.get(hid)

            char_shapes[cid] = {
                "Height": height_pt,
                "FaceName": face_name,
                "Bold": is_bold
            }

    # 3) paraPr: 문단 정렬 추출 [web:157]
    para_props = ref_list.find("hh:paraProperties", NS)
    if para_props is not None:
        for para_pr in para_props.findall("hh:paraPr", NS):
            pid = int(para_pr.get("id"))
            align_el = para_pr.find("hh:align", NS)
            if align_el is not None:
                horiz = align_el.get("horizontal", "LEFT").upper()
                align_map = {
                    "LEFT": "left",
                    "RIGHT": "right",
                    "CENTER": "center",
                    "JUSTIFY": "justify",
                    "BOTH": "justify",
                }
                align = align_map.get(horiz, "left")
            else:
                align = "left"
            para_shapes[pid] = {"Align": align}

    return para_shapes, char_shapes

def paragraph_to_markdown(p_el: ET.Element, char_shapes: dict):
    """
    각 run의 charPrIDRef를 보고 Bold 여부를 판단,
    Bold 구간을 **로 감싸 Markdown 문자열로 만든다.
    """
    runs = p_el.findall("hp:run", NS)
    if not runs:
        # run이 없다면 예전 방식으로
        return extract_text_runs(p_el)

    result = []
    prev_bold = False

    for run in runs:
        run_bold = False
        cid_ref = run.get("charPrIDRef")
        if cid_ref and cid_ref.isdigit():
            cs = char_shapes.get(int(cid_ref), {})
            run_bold = cs.get("Bold", False)

        # 상태 전이 시점에 ** 열고/닫기
        if run_bold and not prev_bold:
            result.append("**")
        elif not run_bold and prev_bold:
            result.append("**")

        # run 내부 모든 t 텍스트 합치기
        txt_parts = []
        for t in run.findall("hp:t", NS):
            if t.text:
                txt_parts.append(t.text)
        result.append("".join(txt_parts))

        prev_bold = run_bold

    # 마지막이 bold 상태면 닫기
    if prev_bold:
        result.append("**")

    return "".join(result).strip()

def paragraph_to_segments(p_el, para_shapes, char_shapes):
    """
    하나의 <hp:p>를 run 단위로 잘라, style이 바뀔 때마다 새 segment를 만드는 함수.
    segment = {"text": "...", "style": {...}}
    """
    segments = []

    # 문단 레벨 정렬
    para_id_ref = p_el.get("paraPrIDRef")
    align = "left"
    if para_id_ref and para_id_ref.isdigit():
        pid = int(para_id_ref)
        align = para_shapes.get(pid, {}).get("Align", "left")

    # run 순회
    for run in p_el.findall("hp:run", NS):
        txt_parts = []
        for t in run.findall("hp:t", NS):
            if t.text:
                txt_parts.append(t.text)
        text = "".join(txt_parts)
        if not text:
            continue

        # run별 charShape
        style = {"Align": align}
        cid_ref = run.get("charPrIDRef")
        if cid_ref and cid_ref.isdigit():
            cs = char_shapes.get(int(cid_ref), {})
            if cs.get("FaceName"):
                style["FaceName"] = cs["FaceName"]
            if cs.get("Height"):
                style["Height"] = cs["Height"]
            if cs.get("Bold") is not None:
                style["Bold"] = cs["Bold"]

        # 이전 segment와 style이 같으면 텍스트만 이어붙이고, 다르면 새 segment 생성
        if segments and segments[-1]["style"] == style:
            segments[-1]["text"] += text
        else:
            segments.append({"text": text, "style": style})

    return segments


# 2) section*.xml 에서 문단 / 표 추출 ----------------------------------------

def extract_text_runs(p_el: ET.Element):
    parts = []
    for run in p_el.findall("hp:run", NS):
        for t in run.findall("hp:t", NS):
            if t.text:
                parts.append(t.text)
    for t in p_el.findall("hp:t", NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts).strip()

def parse_sections_to_blocks(zf, para_shapes, char_shapes, border_fills):
    blocks = []
    section_files = sorted(
        [n for n in zf.namelist()
         if n.startswith("Contents/section") and n.endswith(".xml")]
    )

    for sec in section_files:
        with zf.open(sec) as f:
            tree = ET.parse(f)
        root = tree.getroot()

        section_el = root.find("hp:section", NS)
        if section_el is None:
            section_el = root

        def walk(node, in_table=False):
            tag = node.tag

            if tag == f"{HP}tbl":
                data = []
                cell_styles = []
                cell_segments = []
                cell_merges = []
                cell_nested = []  # ← 새로 추가

                for tr in node.findall("hp:tr", NS):
                    row_texts = []
                    row_styles = []
                    row_seglist = []
                    row_merge = []
                    row_nested = []

                    for tc in tr.findall("hp:tc", NS):
                        col_span, row_span, bg_color, w, h = parse_tc_props(tc, border_fills)
                        cell_text, cell_style, segs_merged, nested_tables = parse_tc_contents(
                            tc, para_shapes, char_shapes, border_fills
                        )

                        row_texts.append(cell_text)
                        row_styles.append(cell_style)
                        row_seglist.append(segs_merged)
                        row_merge.append({
                            "colSpan": col_span,
                            "rowSpan": row_span,
                            "bgColor": bg_color,
                            "width": w,
                            "height": h,
                        })
                        row_nested.append(nested_tables)

                    if row_texts:
                        data.append(row_texts)
                        cell_styles.append(row_styles)
                        cell_segments.append(row_seglist)
                        cell_merges.append(row_merge)
                        cell_nested.append(row_nested)

                if data:
                    blocks.append({
                        "type": "table",
                        "data": data,
                        "style": {},
                        "cell_styles": cell_styles,
                        "cell_segments": cell_segments,
                        "cell_merges": cell_merges,
                        "cell_nested": cell_nested,   # ← 추가
                    })
                return

            if tag == f"{HP}p" and not in_table:
                segs = paragraph_to_segments(node, para_shapes, char_shapes)
                full_text = "".join(s["text"] for s in segs).strip()
                if full_text:
                    blocks.append({
                        "type": "paragraph",
                        "content": full_text,
                        "segments": segs,
                    })

            for child in list(node):
                walk(child, in_table or tag == f"{HP}tbl")

        walk(section_el, False)

    return blocks


def parse_single_table(tbl_el, para_shapes, char_shapes, border_fills):
    """
    <hp:tbl> 요소 하나를 파싱해 table block(dict) 반환.
    (기존 parse_sections_to_blocks 안의 표 처리 로직을 그대로 여기로 옮겼다고 보면 됨)
    """
    data = []
    cell_styles = []
    cell_segments = []
    cell_merges = []

    for tr in tbl_el.findall("hp:tr", NS):
        row_texts = []
        row_styles = []
        row_seglist = []
        row_merge = []

        for tc in tr.findall("hp:tc", NS):
            # 크기/배경/병합 정보
            col_span, row_span, bg_color, w, h = parse_tc_props(tc, border_fills)

            cell_text_chunks = []
            segs_merged = []
            cell_style = None

            # 이 시점에서는 "텍스트용 p"만 처리. (중첩 표는 밖에서 처리)
            for p in tc.findall(".//hp:p", NS):
                segs = paragraph_to_segments(p, para_shapes, char_shapes)
                if segs:
                    cell_text_chunks.append("".join(s["text"] for s in segs))
                    if cell_style is None:
                        cell_style = segs[0]["style"].copy()
                    segs_merged.extend(segs)

            row_texts.append(" ".join([t for t in cell_text_chunks if t]))
            row_styles.append(cell_style or {})
            row_seglist.append(segs_merged)
            row_merge.append({
                "colSpan": col_span,
                "rowSpan": row_span,
                "bgColor": bg_color,
                "width": w,
                "height": h,
            })

        if row_texts:
            data.append(row_texts)
            cell_styles.append(row_styles)
            cell_segments.append(row_seglist)
            cell_merges.append(row_merge)

    return {
        "data": data,
        "cell_styles": cell_styles,
        "cell_segments": cell_segments,
        "cell_merges": cell_merges,
    }

def parse_tc_contents(tc, para_shapes, char_shapes, border_fills):
    """
    하나의 <hp:tc> 안에서
      - 중첩 <hp:tbl> 들을 table spec으로 추출하고
      - 그 표들 안에 속하지 않는 p 들만 외부 셀 텍스트로 사용한다.
    반환: (cell_text:str, cell_style:dict, segs_merged:list, nested_tables:list)
    """
    cell_text_chunks = []
    segs_merged = []
    cell_style = None
    nested_tables = []

    # 1) tc 아래 모든 중첩 tbl 수집
    nested_tbl_elems = list(tc.findall(".//hp:tbl", NS))

    # 1-1) 중첩 tbl 안에 포함된 모든 p를 set에 모아둠
    p_in_nested = set()
    for tbl in nested_tbl_elems:
        for p in tbl.findall(".//hp:p", NS):
            p_in_nested.add(p)

    # 1-2) 중첩 tbl 자체를 table spec으로 파싱
    for tbl in nested_tbl_elems:
        tbl_block = parse_single_table(tbl, para_shapes, char_shapes, border_fills)
        nested_tables.append({
            "type": "table",
            **tbl_block
        })

    # 2) tc 아래의 모든 p 중에서, nested tbl 안에 속하지 않는 것만 외부 텍스트로 사용
    for p in tc.findall(".//hp:p", NS):
        if p in p_in_nested:
            continue  # 중첩 표 안의 문단은 셀 텍스트에서 제외

        segs = paragraph_to_segments(p, para_shapes, char_shapes)
        if segs:
            cell_text_chunks.append("".join(s["text"] for s in segs))
            if cell_style is None:
                cell_style = segs[0]["style"].copy()
            segs_merged.extend(segs)

    cell_text = " ".join([t for t in cell_text_chunks if t])
    return cell_text, (cell_style or {}), segs_merged, nested_tables




def parse_table_styles_from_header(zf: zipfile.ZipFile):
    border_fills = {}
    with zf.open("Contents/header.xml") as f:
        tree = ET.parse(f)
    root = tree.getroot()
    ref_list = root.find("hh:refList", NS)
    if ref_list is None:
        return border_fills

    bf_parent = ref_list.find("hh:borderFills", NS)
    if bf_parent is None:
        return border_fills

    for bf in bf_parent.findall("hh:borderFill", NS):
        bid = bf.get("id")
        if not bid or not bid.isdigit():
            continue
        bid = int(bid)

        color = pick_fill_color_from_borderfill(bf)
        border_fills[bid] = {"fillColor": color}
    return border_fills



def parse_tc_props(tc, border_fills):
    col_span = 1
    row_span = 1
    bg_color = None
    width = height = None

    bf_ref = tc.get("borderFillIDRef")
    if bf_ref and bf_ref.isdigit():
        bf = border_fills.get(int(bf_ref), {})
        bg_color = bf.get("fillColor")

    cell_span_el = tc.find("hp:cellSpan", NS)
    if cell_span_el is not None:
        col_attr = cell_span_el.get("colSpan")
        row_attr = cell_span_el.get("rowSpan")
        if col_attr and col_attr.isdigit():
            col_span = int(col_attr)
        if row_attr and row_attr.isdigit():
            row_span = int(row_attr)

    cell_sz_el = tc.find("hp:cellSz", NS)
    if cell_sz_el is not None:
        w = cell_sz_el.get("width")
        h = cell_sz_el.get("height")
        width = int(w) if w and w.isdigit() else None
        height = int(h) if h and h.isdigit() else None

    return col_span, row_span, bg_color, width, height


def pick_fill_color_from_borderfill(bf_el: ET.Element) -> str | None:
    """
    <borderFill> 요소 하나에서 '배경색'으로 쓸 #RRGGBB를 한 개 골라낸다.
    - 우선순위:
      1) winBrush.faceColor (none이 아닌 경우)
      2) gradation.color/@value 첫 번째
      3) winBrush.hatchColor
    - 모두 없으면 None.
    """
    # ns1: (core 네임스페이스) 접두사까지 포함 가능성 감안해서 로컬 태그/속성이름으로 처리
    # 1) winBrush.faceColor
    for el in bf_el.iter():
        lname = el.tag.split('}')[-1]
        if lname == "winBrush":
            face = el.attrib.get("faceColor")
            if face and face.lower() != "none":
                return normalize_color(face)
    # 2) gradation.color/@value
    for el in bf_el.iter():
        lname = el.tag.split('}')[-1]
        if lname == "color":
            val = el.attrib.get("value")
            if val:
                return normalize_color(val)
    # 3) winBrush.hatchColor (백색이 아닌 경우에 한해)
    for el in bf_el.iter():
        lname = el.tag.split('}')[-1]
        if lname == "winBrush":
            hatch = el.attrib.get("hatchColor")
            if hatch:
                return normalize_color(hatch)
    return None

def normalize_color(val: str) -> str:
    """
    '#RRGGBB' 또는 '#AARRGGBB' 형식 등을 '#RRGGBB'로 정규화.
    """
    if not val.startswith("#"):
        return None
    if len(val) == 7:
        return val
    if len(val) == 9:  # #AARRGGBB → 뒤 6자리만 사용
        return "#" + val[-6:]
    # 그 외 포맷은 일단 무시
    return None


# 3) blocks -> doclib용 document spec 변환 ----------------------------------

def blocks_to_document_spec(blocks):
    doc = {}
    p_idx, t_idx = 1, 1

    for b in blocks:
        if b["type"] == "paragraph":
            key = f"문단{p_idx}"
            base = b["segments"][0]["style"].copy() if b.get("segments") else {}
            doc[key] = {
                "content": b["content"],
                "style": {
                    "FaceName": base.get("FaceName", "바탕체"),
                    "Height": base.get("Height", 11),
                    "Bold": base.get("Bold", False),
                    "Align": base.get("Align", "left")
                },
                "segments": b.get("segments", [])
            }
            p_idx += 1

        elif b["type"] == "table":
            key = f"표{t_idx}"
            cols = len(b["data"][0])
            doc[key] = {
                "data": b["data"],
                "style": {
                    "cell_font": "바탕체",
                    "cell_size": 11,
                    "cell_align": ["left"] * cols,
                },
                "cell_styles": b.get("cell_styles"),
                "cell_segments": b.get("cell_segments"),
                "cell_merges": b.get("cell_merges"),
                "cell_nested": b.get("cell_nested"),  # ← 추가
            }
            t_idx += 1


    return {"document": doc}


def debug_dump_styles(para_shapes, char_shapes, limit=10):
    print("=== ParaShapes (문단 스타일) ===")
    for i, (pid, ps) in enumerate(sorted(para_shapes.items())):
        if i >= limit:
            print("... (more paraShapes omitted)")
            break
        print(f"  id={pid} -> {ps}")

    print("\n=== CharShapes (글자 스타일) ===")
    for i, (cid, cs) in enumerate(sorted(char_shapes.items())):
        if i >= limit:
            print("... (more charShapes omitted)")
            break
        print(f"  id={cid} -> {cs}")


def debug_tcpr_structure(zf, section_name="Contents/section0.xml"):
    with zf.open(section_name) as f:
        tree = ET.parse(f)
    root = tree.getroot()

    tbl = root.find(".//hp:tbl", NS)
    if tbl is None:
        print("No <hp:tbl> found")
        return

    first_tcpr = tbl.find(".//hp:tcPr", NS)
    if first_tcpr is None:
        print("No <hp:tcPr> found")
        return

    print("=== Raw <hp:tcPr> XML ===")
    print(ET.tostring(first_tcpr, encoding="unicode"))


def debug_tc_structure(zf, section_name="Contents/section0.xml"):
    with zf.open(section_name) as f:
        tree = ET.parse(f)
    root = tree.getroot()

    tbl = root.find(".//hp:tbl", NS)
    if tbl is None:
        print("No <hp:tbl> found")
        return

    first_tc = tbl.find(".//hp:tc", NS)
    if first_tc is None:
        print("No <hp:tc> found")
        return

    print("=== Raw <hp:tc> XML ===")
    print(ET.tostring(first_tc, encoding="unicode"))

    print("\n=== attributes of <hp:tc> ===")
    print(first_tc.attrib)

    print("\n=== direct children tags of <hp:tc> ===")
    for child in list(first_tc):
        print("  child tag:", child.tag, "attrib:", child.attrib)

def debug_borderfill(zf, bid):
    with zf.open("Contents/header.xml") as f:
        tree = ET.parse(f)
    root = tree.getroot()
    ref_list = root.find("hh:refList", NS)
    if ref_list is None:
        print("no refList")
        return
    bf_parent = ref_list.find("hh:borderFills", NS)
    if bf_parent is None:
        print("no borderFills")
        return

    for bf in bf_parent.findall("hh:borderFill", NS):
        if bf.get("id") == str(bid):
            print(f"=== <borderFill id=\"{bid}\"> raw XML ===")
            print(ET.tostring(bf, encoding="unicode"))
            break


# 4) 전체 파이프라인 ----------------------------------------------------------

def parse_hwpx_to_spec(hwpx_path: str, out_json_path: str = "parsed_spec.json"):
    with zipfile.ZipFile(hwpx_path, "r") as zf:
        para_shapes, char_shapes = parse_styles_from_header(zf)
        #debug_dump_styles(para_shapes, char_shapes)
        #debug_tc_structure(zf)        
        border_fills = parse_table_styles_from_header(zf)
        blocks = parse_sections_to_blocks(zf, para_shapes, char_shapes, border_fills)
    spec = blocks_to_document_spec(blocks)
    with open(out_json_path, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)
    return spec

if __name__ == "__main__":
    spec = parse_hwpx_to_spec("input.hwpx", "parsed_spec.json")
    print("parsed_spec.json 생성 완료")
