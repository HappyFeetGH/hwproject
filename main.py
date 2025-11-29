import sys
import json
import yaml
from doclib import generate_hwp_from_spec, generate_hwp_from_parsed_spec

def load_spec(path):
    if path.endswith('.json'):
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    elif path.endswith('.yaml') or path.endswith('.yml'):
        with open(path, encoding='utf-8') as f:
            return yaml.safe_load(f)
    else:
        raise ValueError("지원하지 않는 파일 형식")

def main():
    if len(sys.argv) < 2:
        print("사용법: python main.py parsed_spec.json [output.hwpx]")
        sys.exit(1)

    spec_path = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) >= 3 else "output.hwpx"

    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)

    generate_hwp_from_parsed_spec(spec, filename=output)
    print(f"완료: {output}")

if __name__ == "__main__":
    main()