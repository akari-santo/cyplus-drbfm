import csv
import json
import sys

def expand_concerns(input_path: str, output_path: str) -> None:
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

            for concern in data.get("concerns", []):
                rows.append({
                    "物性": row["物性"],
                    "機能": row["機能"],
                    "concern_content": concern.get("concern_content", ""),
                    "location": concern.get("location", ""),
                    "mechanism": concern.get("mechanism", ""),
                    "affected_parameters": concern.get("affected_parameters", ""),
                    "affected_function": concern.get("affected_function", ""),
                })

    fieldnames = ["物性", "機能", "concern_content", "location", "mechanism", "affected_parameters", "affected_function"]
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Done: {len(rows)} rows written to {output_path}")


if __name__ == "__main__":
    input_path = sys.argv[1] if len(sys.argv) > 1 else "TableJsonParseTask_result_03220023.csv"
    output_path = sys.argv[2] if len(sys.argv) > 2 else "expanded_concerns.csv"
    expand_concerns(input_path, output_path)
