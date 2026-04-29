"""룰셋 검증 + 통계 출력."""
import yaml
from collections import Counter
from pathlib import Path

RULES_PATH = Path(__file__).parent / "MVR_rules.yaml"


def load_rules():
    with open(RULES_PATH, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data["rules"]


def validate_rules(rules):
    """필수 필드 누락, ID 중복 검사."""
    errors = []
    seen_ids = set()
    required_fields = {"id", "name", "category", "severity", "check_type", "message"}

    for i, rule in enumerate(rules):
        missing = required_fields - set(rule.keys())
        if missing:
            errors.append(f"[{i}] 필수 필드 누락: {missing}")

        rule_id = rule.get("id")
        if rule_id in seen_ids:
            errors.append(f"[{i}] 중복 ID: {rule_id}")
        seen_ids.add(rule_id)

    return errors


def print_stats(rules):
    print(f"\n총 룰 개수: {len(rules)}")

    enabled = [r for r in rules if r.get("enabled", True)]
    print(f"활성 룰: {len(enabled)} / 비활성 룰: {len(rules) - len(enabled)}")

    print("\n--- Severity별 ---")
    for sev, n in Counter(r["severity"] for r in rules).most_common():
        print(f"  {sev:15s} {n}")

    print("\n--- Category별 ---")
    for cat, n in Counter(r["category"] for r in rules).most_common():
        print(f"  {cat:18s} {n}")

    print("\n--- Check type별 ---")
    for ct, n in Counter(r["check_type"] for r in rules).most_common():
        print(f"  {ct:20s} {n}")


if __name__ == "__main__":
    rules = load_rules()

    print("=" * 50)
    print("MVR 룰셋 검증 결과")
    print("=" * 50)

    errors = validate_rules(rules)
    if errors:
        print("\n[ERROR]")
        for e in errors:
            print(f"  {e}")
    else:
        print("\n[OK] 모든 룰이 스키마를 만족함")

    print_stats(rules)
