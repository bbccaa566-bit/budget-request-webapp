"""예산 요구안 JSON -> 엑셀(요약·산출내역·확인 필요). 금액 계산은 엑셀 수식으로 넣는다.

사용법: python -I make_budget_request.py <입력.json> <출력.xlsx> [--source <올해 예산집행 xlsx>]
--source: 올해 엑셀의 세부사업이 모두 있는지, 배정·집행액이 같은지 검사. 문제가 있으면 종료 코드 1.

계산·검사·엑셀 생성은 웹앱과 같은 코어(budget_app/budget_core.py)를 쓴다.
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]  # scripts → skill → skills → .claude → 프로젝트
sys.path.insert(0, str(ROOT / "budget_app"))
import budget_core as core  # noqa: E402


def main(src_json, dst, source_xlsx=None):
    data = json.loads(Path(src_json).read_text(encoding="utf-8"))
    original = None
    if source_xlsx:
        parsed = core.parse_excel(Path(source_xlsx).read_bytes())
        if parsed["errors"]:
            print("원본 엑셀 읽기 오류:", *parsed["errors"], sep="\n - ")
            sys.exit(1)
        original = core.original_from_parse(parsed)
        data["source_verified"] = True
    result = core.check(data, original)
    if original is None:
        result["warnings"] = ["원본 엑셀(--source)을 주지 않아 세부사업 누락·금액 대조를 하지 않았습니다."] + [
            w for w in result["warnings"] if not w.startswith("원본 값 정보가 없어")]

    Path(dst).parent.mkdir(parents=True, exist_ok=True)
    Path(dst).write_bytes(core.workbook_bytes(data, result))

    s = result["summary"]
    unit = data.get("unit", "")
    print(f"저장: {dst}")
    print(f"세부사업 {s['rows']}개 | 산출방식 " + ", ".join(f"{m} {n}" for m, n in s["methods"].items()))
    print(f"(검산) 배정 {s['배정']:,} / 요구 {s['요구']:,} / 증감 {s['증감']:,} {unit}")
    print(f"확인 필요 {s['확인필요']}건")
    for w in result["warnings"]:
        print("경고:", w)
    if result["problems"]:
        print("자동 검사 실패:")
        for p in result["problems"]:
            print(" -", p["msg"])
        sys.exit(1)
    print("자동 검사 통과")


if __name__ == "__main__":
    args = sys.argv[1:]
    src = None
    if "--source" in args:
        k = args.index("--source")
        src = args[k + 1]
        args = args[:k] + args[k + 2:]
    if len(args) != 2:
        sys.exit("사용법: python -I make_budget_request.py <입력.json> <출력.xlsx> [--source <올해 xlsx>]")
    main(args[0], args[1], src)
