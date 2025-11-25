from pyhwpx import Hwp

def debug_log(*args):
    print("[DEBUG]", *args)

contents = {
    "title": "2학기 총괄평가 실시 계획 안내",
    "receiver": "안녕하십니까?",
    "opening": "학교 교육에 관심을 가져주셔서 감사드리며, 가정에 건강과 행복이 가득하시길 바랍니다.",
    "body": "본교에서는 **3~6**학년을 대상으로 *2025학년도* 2학기 총괄평가를 실시합니다.\n이번 평가는 한 학기 동안 학습한 내용을 선다형, 단답형, 서·논술형 문항으로 확인합니다.\n학생들은 자신의 학습 수준을 파악하고 학교는 맞춤형 지도를 제공합니다.",
    "table_caption": "< 2학기 총괄평가 실시 계획 안내 >",
    "date": "2025년 11월 18일",
    "school_name": "한솔초등학교",
    "principal": "한솔초등학교장"
}

styles = {
    "title":      {"FaceName": "돋움", "Height": 20, "Bold": True, "Align": "center"},
    "receiver":   {"FaceName": "바탕체", "Height": 11, "Bold": False, "Align": "left"},
    "opening":    {"FaceName": "돋움", "Height": 12, "Bold": False, "Align": "left"},
    "body":       {"FaceName": "바탕체", "Height": 13, "Bold": False, "Align": "left"},
    "table_caption": {"FaceName": "돋움", "Height": 14, "Bold": True, "Align": "center"},
    "date":       {"FaceName": "바탕체", "Height": 15, "Bold": False, "Align": "right"},
    "school_name": {"FaceName": "돋움", "Height": 16, "Bold": False, "Align": "right"},
    "principal":  {"FaceName": "바탕체", "Height": 17, "Bold": True, "Align": "right"},   
}

# 표 데이터 입력 예시 (간단 3x4)
table_data = [
    ["학년", "일자", "교과", "평가범위"],
    ["3", "12.2.(화)", "국어", "경험과 **관련**지으며 이해해요~"],
    ["4", "12.3.(수)", "수학", "곱셈 ~ 분수"]
]


def set_alignment(hwp, align):
    if align == "center":
        hwp.ParagraphShapeAlignCenter()
    elif align == "left":
        hwp.ParagraphShapeAlignLeft()
    elif align == "right":
        hwp.ParagraphShapeAlignRight()
    elif align == "justify":
        hwp.ParagraphShapeAlignJustify()

hwp = Hwp()

import re

def parse_segments(text):
    """
    입력된 텍스트 내에서 
    - **bold**
    - _italic_ 또는 *italic*
    - <u>underline</u>
    을 탐지해 segment list로 분리. 평문, 스타일 조각 모두 dict로 반환.
    """
    # 볼드: **텍스트**
    text = text.replace('\r\n', '\n')  # 줄바꿈 통일 (옵션)

    patterns = [
        (r'\*\*(.+?)\*\*', 'bold'),
        (r'_(.+?)_', 'italic'),
        (r'\*(.+?)\*', 'italic'),
        (r'<u>(.+?)</u>', 'underline')
    ]

    # 모든 스타일 포함된 위치/변화 push
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
            # 직전 평문(스타일 아닌 부분) 추가
            pre = text[:i]
            if pre:
                segments.append({"text": pre})
            # 스타일 부분 추가
            segment = {"text": val}
            if typ == "bold":
                segment["bold"] = True
            elif typ == "italic":
                segment["italic"] = True
            segments.append(segment)
            # 다음 위치로 이동
            text = text[i+len(m.group(0)):]
            i = 0
        else:
            i += 1
    # 남은 평문 추가
    if text:
        segments.append({"text": text})

    return segments


def insert_role_and_style(role):
    debug_log(f"\n==[ROLE: {role}]==")
    base_opts = {
        "FaceName": styles[role]["FaceName"],
        "Height": styles[role]["Height"],
        "Bold": styles[role]["Bold"],
        "Italic": False,        
    }
    set_alignment(hwp, styles[role]["Align"])

    # 예시: Markdown-like parsed segments로 입력
    # segments = [{"text": "일시: ", "bold": False}, {"text": "12:00", "bold": True}]
    segments = parse_segments(contents[role])  # 예시 파서 필요

    for seg in segments:
        # 1. 국소 스타일 적용
        opts = base_opts.copy()
        opts["Bold"]   = seg.get("bold", base_opts["Bold"])
        opts["Italic"] = seg.get("italic", base_opts.get("Italic", False))        

        hwp.set_font(**opts)
        hwp.insert_text(seg["text"] if isinstance(seg, dict) else seg)

        # 2. (중요!) 국소 스타일 이후 반드시 “기본 스타일로 복구”
        hwp.set_font(**base_opts)
    hwp.insert_text("\r\n")

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    red = int(hex_color[0:2], 16)
    green = int(hex_color[2:4], 16)
    blue = int(hex_color[4:6], 16)
    return (red, green, blue)

import time
def insert_table_and_style(table_data, cell_styles=None, cell_bg_colors=None, col_aligns=None):
    """
    table_data      : 2D list, each cell is text. (ex: [["학년", "일자"], ["3", "12.2.(화)"]])
    cell_styles     : 2D list, 각 셀 별 {"FaceName", "Height", "Bold", ...} dict.
    col_widths      : list, 각 열 너비(mm) ex: [30, 50, 40]
    row_heights     : list, 각 행 높이(mm) ex: [10, 12]
    cell_bg_colors  : 2D list, 각 셀별 배경색코드 (ex: [["#FFDDEE", None], ...])
    col_aligns      : list, 각 열 정렬값 ("left", "center", "right")
    """
    rows, cols = len(table_data), len(table_data[0])
    hwp.create_table(rows, cols, treat_as_char=True)
    time.sleep(1)  # 표 생성 대기    
    for r, row in enumerate(table_data):
        for c, val in enumerate(row):
            
            style = cell_styles[r][c] if cell_styles else {}
            bg    = cell_bg_colors[r][c] if cell_bg_colors else None
            align = col_aligns[c] if col_aligns else "left"
            # 셀 스타일 적용
            hwp.set_font(
                FaceName=style.get("FaceName", "바탕체"),
                Height=style.get("Height", 11),
                Bold=style.get("Bold", False))
            # 셀 정렬
            if align=="center": hwp.TableCellAlignCenterCenter()
            elif align=="right": hwp.TableCellAlignRightCenter()
            else: hwp.TableCellAlignLeftCenter()
            # 셀 배경색
            if bg:
                if bg.startswith('#'):  # hex 형식일 때
                    red, green, blue = hex_to_rgb(bg)
                    hwp.gradation_on_cell([(red, green, blue)])
                else:
                    hwp.gradation_on_cell([bg])
            # 셀 내용 입력
            hwp.insert_text(str(val))
            if c < cols - 1:
                hwp.TableRightCell()
        if r < rows - 1:
            hwp.TableLowerCell()
            for _ in range(cols - 1):
                hwp.TableLeftCell()
    hwp.MoveDown()

cell_styles = [
    [ {"FaceName":"돋움","Height":14,"Bold":True} for _ in range(4) ],
    [ {"FaceName":"바탕","Height":11,"Bold":False} for _ in range(4) ],
    [ {"FaceName":"바탕","Height":11,"Bold":False} for _ in range(4) ]
]

cell_bg_colors = [
    ["#FFF1AF","#FFF1AF","#FFF1AF","#FFF1AF"],
    ["#FFF9C5", None,None,None],
    ["#FFF9C5",None,None,None]
]
col_aligns = ["center","center","center","left"]


for role in ["title", "receiver", "opening", "body", "table_caption"]:
    insert_role_and_style(role)


insert_table_and_style(
    table_data,
    cell_styles=cell_styles,
    cell_bg_colors=cell_bg_colors,
    col_aligns=col_aligns
)


for role in ["date", "school_name", "principal"]:
    insert_role_and_style(role)



hwp.save_as("test.hwpx")
hwp.quit()
