## **1. agent.md(LLM용 instruction 문서) 기본 구조 제안**

### **목표 역할**

- LLM에게 실제 예시(pdf/png 등)나 설명(context)을 보여주면,
    - (a) 그와 비슷한 문서/표/레이아웃/json/yaml 템플릿을 생성하게 하거나
    - (b) context(키워드, 규칙, 내용) 입력에 따라 JSON/YAML 명세를 작성함
    - (c) 최종적으로 md→hwpx 코드에 input만 넘기면 자동 변환이 가능하게 함

***

### **agent.md 서술 구조 예시**

```markdown
# LLM Document Agent Instruction

## 목적 요약  
- 한글 문서(pdf, png, md 예시)를 최대한 비슷하게 스타일링된 JSON/YAML 템플릿으로 변환  
- 입력 context에 따라 맞춤 문서/표 구조, 역할, 스타일 정보(json, yaml) 생성  
- 생성된 명세를 기반으로 라이브러리에 전달하여 실제 HWPX 문서(표 포함)로 자동 생성 수행

## 요청 예시 및 task 구조

### 1. 이미지(pdf/png) → template json/yaml
- 첨부된 파일 레이아웃/표/타이틀/폰트/배경색 등 최대한 비슷하게 JSON/YAML template 설계
  - 예시:
    ```
    title_font: "나눔명조"
    title_size: 18
    table: 
      header_bg: "#FFEE99"
      cell_font: "바탕체"
      cell_size: 12
      cell_align: ["center", "left", "right"]
      data: [["학년", "일자", ...], ...]
    page_style:
      margin:[^1]
    ```

### 2. 템플릿 + 문서 context → doc json/yaml
- 템플릿 기반, 문서 내용/표 데이터 context(recipient, date, table_rows 등)에 맞춰 전체 문서 YAML 생성  
  - 예시:
    ```
    title: "2025 총괄평가 시행 안내"
    date: "2025-11-30"
    recipient: "학부모님"
    table:
      data: [["3", "12.1", ...], ...]
    footer: "문의: ... 연락처"
    ```

### 3. context → doc json/yaml  
- 추가적 내용(rule, 키워드, 업무 목적 등) 입력 시 적절한 YAML/JSON 출력  
  - 예시:
    ```
    title: "회의록"
    table:
      header_bg: "#AAFFDD"
      data: [["참석자", "의견", ...], ...]
    etc: ...
    ```

## 형식 요구  
- 모든 출력은 **정형화된 JSON 또는 YAML**  
- 표, 문단, 타이틀, 스타일(폰트, 색, 정렬, 배경 등)은 전부 명확한 key:value로 명시  
- 필요시 md 표 구조를 함께 명세로 제공

## output 예시  
```

{
"title": "...",
"table": {
"header_bg": "\#FFEE99",
"cell_bg": [["\#FFEE99", "\#FFEE99", ...], ...],
"data": [["학년", ...], ...]
},
"footer": "..."
}

```

---

## **2. LLM + 라이브러리 연동 프로세스 개요**

- agent.md 기반 프롬프트:
  - 이미지/pdf 예시 → LLM → 템플릿 YAML/JSON 생성
  - 템플릿+context → LLM → 최종 문서 YAML/JSON 생성
- 내가 만든 라이브러리의 함수:
  - `generate_hwp_from_yaml(yaml_spec)`
  - `generate_hwp_from_json(json_spec)`
- 실제 파이프라인 예시:
  1. LLM이 agent.md instruction대로 템플릿+문서\Json 생성
  2. 라이브러리 함수에 전달
  3. HWPX 파일 생성 및 저장

---

## **3. 구현 가능성/요구사항**

- **pdf/png 파일 → 문서 구조 추출:**  
  - LLM, OCR, Vision API 등과 연동해 구조/스타일/레이아웃 추출해 yaml로 생성 가능
- **템플릿 기반/Context 기반 생성:**  
  - 역할, 스타일, 테이블, 폰트 등 json/yaml 형식으로 받으면,  
    이미 구현된 라이브러리 함수를 통해 hwpx/공문/보고서 등 즉시 자동 생성 가능

- **확장성:**  
  - 표/문단/스타일/캡션/링크 등 지원 항목을 계속 추가할 수 있음

---