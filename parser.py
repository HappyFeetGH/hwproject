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

def parse_sections_to_blocks(zf: zipfile.ZipFile, para_shapes, char_shapes):
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
                for tr in node.findall("hp:tr", NS):
                    row = []
                    for tc in tr.findall("hp:tc", NS):
                        cell_text = []
                        # 수정: findall 대신 .//로 descendant 전체 검색 [web:150]
                        for p in tc.findall(".//hp:p", NS):
                            segs = paragraph_to_segments(p, para_shapes, char_shapes)
                            txt = "".join(s["text"] for s in segs).strip()
                            if txt:
                                cell_text.append(txt)
                        row.append(" ".join(cell_text))
                    if row:
                        data.append(row)
                if data:
                    blocks.append({
                        "type": "table",
                        "data": data,
                        "style": {}
                    })
                return

            if tag == f"{HP}p" and not in_table:
                segs = paragraph_to_segments(node, para_shapes, char_shapes)
                full_text = "".join(s["text"] for s in segs).strip()
                if full_text:
                    blocks.append({
                        "type": "paragraph",
                        "content": full_text,
                        "segments": segs
                    })

            for child in list(node):
                walk(child, in_table or tag == f"{HP}tbl")

        walk(section_el, False)

    return blocks




# 3) blocks -> doclib용 document spec 변환 ----------------------------------

def blocks_to_document_spec(blocks):
    doc = {}
    p_idx, t_idx = 1, 1

    for b in blocks:
        if b["type"] == "paragraph":
            key = f"문단{p_idx}"
            # 대표 스타일은 첫 segment 기준으로 뽑되, segments는 그대로 모두 넘긴다
            base_style = b["segments"][0]["style"].copy() if b["segments"] else {}
            doc[key] = {
                "content": b["content"],
                "style": {
                    "FaceName": base_style.get("FaceName", "바탕체"),
                    "Height": base_style.get("Height", 11),
                    "Bold": base_style.get("Bold", False),
                    "Align": base_style.get("Align", "left")
                },
                "segments": b["segments"]
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
                    "cell_align": ["left"] * cols
                }
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



# 4) 전체 파이프라인 ----------------------------------------------------------

def parse_hwpx_to_spec(hwpx_path: str, out_json_path: str = "parsed_spec.json"):
    with zipfile.ZipFile(hwpx_path, "r") as zf:
        para_shapes, char_shapes = parse_styles_from_header(zf)
        #debug_dump_styles(para_shapes, char_shapes)
        blocks = parse_sections_to_blocks(zf, para_shapes, char_shapes)
    spec = blocks_to_document_spec(blocks)
    with open(out_json_path, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)
    return spec

if __name__ == "__main__":
    spec = parse_hwpx_to_spec("input.hwpx", "parsed_spec.json")
    print("parsed_spec.json 생성 완료")
