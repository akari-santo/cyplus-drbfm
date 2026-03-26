"""
分析結果CSVを展開するスクリプト。

入力CSVの想定列:
  機能, result (JSON)

resultのJSON構造:
  {
    "assumptions": {キー: 値, ...},
    "physical_changes": [
      {"物理量": "...", "変化の内容": "...", "支配方程式": "..."},
      ...
    ],
    "impact_prediction": {
      "最終予測値の変化": "...",
      "総合判定": "..."
    }
  }

出力列:
  機能, 前提, 対象機能への影響予測, 変化する物理量, 予測の根拠
"""

import csv
import json
import sys


def format_assumptions(assumptions: dict) -> str:
    """前提条件をセミコロン区切りの1文字列に結合する。"""
    return " | ".join(f"{k}: {v}" for k, v in assumptions.items())


def format_impact(impact: dict) -> str:
    """影響予測を1文字列にまとめる。"""
    parts = []
    if impact.get("最終予測値の変化"):
        parts.append(impact["最終予測値の変化"])
    if impact.get("総合判定"):
        parts.append(f"【総合判定】{impact['総合判定']}")
    return " / ".join(parts)


def format_physical_quantity(change: dict) -> str:
    """物理量と変化内容を結合する。"""
    quantity = change.get("物理量", "")
    content = change.get("変化の内容", "") or change.get("変化の内容（定量）", "")
    if content:
        return f"{quantity}: {content}"
    return quantity


def format_basis(change: dict) -> str:
    """予測の根拠（支配方程式）を返す。"""
    return (
        change.get("支配方程式", "")
        or change.get("支配方程式・物理的根拠", "")
    )


def expand_analysis(input_path: str, output_path: str) -> None:
    rows = []

    with open(input_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            result_str = row.get("result", "").strip()
            if not result_str:
                continue
            try:
                data = json.loads(result_str)
            except json.JSONDecodeError as e:
                print(f"JSON parse error: {e}", file=sys.stderr)
                continue

            機能 = row.get("機能", "")
            前提 = format_assumptions(data.get("assumptions", {}))
            影響予測 = format_impact(data.get("impact_prediction", {}))

            physical_changes = data.get("physical_changes", [])
            if not physical_changes:
                # 物理量なしでも1行出力
                rows.append({
                    "機能": 機能,
                    "前提": 前提,
                    "対象機能への影響予測": 影響予測,
                    "変化する物理量": "",
                    "予測の根拠": "",
                })
            else:
                for change in physical_changes:
                    rows.append({
                        "機能": 機能,
                        "前提": 前提,
                        "対象機能への影響予測": 影響予測,
                        "変化する物理量": format_physical_quantity(change),
                        "予測の根拠": format_basis(change),
                    })

    fieldnames = ["機能", "前提", "対象機能への影響予測", "変化する物理量", "予測の根拠"]
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Done: {len(rows)} rows written to {output_path}")


if __name__ == "__main__":
    input_path = sys.argv[1] if len(sys.argv) > 1 else "analysis_input.csv"
    output_path = sys.argv[2] if len(sys.argv) > 2 else "analysis_expanded.csv"
    expand_analysis(input_path, output_path)
