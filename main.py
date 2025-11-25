import sys
import json
import yaml
from doclib import generate_hwp_from_spec

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
        print("사용법: python main.py [spec.json 또는 spec.yaml]")
        sys.exit(1)
    spec_file = sys.argv[1]
    spec = load_spec(spec_file)
    filename = spec.get("output", "output.hwpx")
    generate_hwp_from_spec(spec, filename=filename)
    print(f"HWPX 생성 완료: {filename}")

if __name__ == '__main__':
    main()
