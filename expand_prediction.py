"""
予測結果CSV（機能, 内容）を展開するスクリプト。

内容フィールドの構造（マークダウンテキスト）:
  1. 解析の前提条件（Assumptions）
     - key: value  ...
  2. 物理量の変化（定量）
     | 物理量 | 変化の内容（定量） | 支配方程式・物理的根拠 |
     | ...   | ...             | ...                |
  3. 対象機能への影響予測（定量統合）
     - 最終予測値の変化: ...
     - 総合判定: ...

出力列:
  機能, 前提, 対象機能への影響予測, 変化する物理量, 予測の根拠
"""

import csv
import re
import sys


def clean(text: str) -> str:
    """マークダウン記法（**bold**, \\[...\\]）を除去して整形する。"""
    text = re.sub(r'\\\[|\\\]', '', text)   # \[ \] 除去
    text = re.sub(r'\*\*', '', text)         # ** 除去
    text = re.sub(r'\\_', '_', text)         # \_ → _
    return text.strip()


def find_bullets(text: str) -> list:
    """
    セクション内の箇条書きを抽出する。
    ヘッダーと最初の箇条書きが同じ行に2スペースで繋がっているケースに対応。
    例: '1. 解析の前提条件  - 運転温度: 80℃\n  - ...'
    """
    # 2スペース以上 + ダッシュ を改行 + ダッシュ に正規化
    normalized = re.sub(r' {2,}-', '\n-', text)
    bullets = []
    for line in normalized.split('\n'):
        line = line.strip()
        if line.startswith('- '):
            bullets.append(line[2:].strip())
    return bullets


def extract_assumptions(section1: str) -> str:
    """Section 1 の箇条書きから前提条件を抽出して連結する。"""
    items = []
    for b in find_bullets(section1):
        cleaned = clean(b)
        if cleaned:
            items.append(cleaned)
    return ' | '.join(items)


def parse_table(section2: str) -> list:
    """Section 2 のマークダウンテーブルを行リストとして返す。"""
    rows = []
    header_passed = False
    for line in section2.splitlines():
        line = line.strip()
        if not line.startswith('|'):
            continue
        cells = [c.strip() for c in line.split('|')[1:-1]]
        if not cells:
            continue
        # セパレータ行 (| :-- | :-- |) をスキップ
        if all(re.match(r'^:?-+:?$', c) for c in cells):
            header_passed = True
            continue
        if not header_passed:
            # ヘッダ行をスキップ
            header_passed = True
            continue
        if len(cells) >= 3:
            rows.append({
                '物理量': clean(cells[0]),
                '変化の内容': clean(cells[1]),
                '支配方程式': clean(cells[2]),
            })
    return rows


def extract_prediction(section3: str) -> str:
    """Section 3 から最終予測値と総合判定を抽出して連結する。"""
    final_val = ""
    judgment = ""

    for b in find_bullets(section3):
        # 最終予測値の変化
        if '最終予測値の変化' in b:
            m = re.search(r'最終予測値の変化[^:：]*[:：]\s*(.+)', b)
            if m:
                final_val = clean(m.group(1))
        # 総合判定
        elif '総合判定' in b:
            m = re.search(r'総合判定[^:：]*[:：]\s*(.+)', b)
            if m:
                judgment = clean(m.group(1))

    parts = []
    if final_val:
        parts.append(final_val)
    if judgment:
        parts.append(f"【総合判定】{judgment}")
    return ' / '.join(parts)


def split_sections(content: str):
    """内容テキストを3つのセクションに分割する。"""
    s2 = content.find('2. 物理量の変化')
    s3 = content.find('3. 対象機能への影響予測')

    if s2 == -1 or s3 == -1:
        return content, '', ''

    return content[:s2], content[s2:s3], content[s3:]


def expand_prediction(input_path: str, output_path: str) -> None:
    rows_out = []

    with open(input_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            機能 = row.get('機能', '')
            content = row.get('内容', '')

            sec1, sec2, sec3 = split_sections(content)

            前提 = extract_assumptions(sec1)
            影響予測 = extract_prediction(sec3)
            table_rows = parse_table(sec2)

            if not table_rows:
                rows_out.append({
                    '機能': 機能,
                    '前提': 前提,
                    '対象機能への影響予測': 影響予測,
                    '変化する物理量': '',
                    '予測の根拠': '',
                })
            else:
                for tr in table_rows:
                    変化する物理量 = f"{tr['物理量']}: {tr['変化の内容']}" if tr['変化の内容'] else tr['物理量']
                    rows_out.append({
                        '機能': 機能,
                        '前提': 前提,
                        '対象機能への影響予測': 影響予測,
                        '変化する物理量': 変化する物理量,
                        '予測の根拠': tr['支配方程式'],
                    })

    fieldnames = ['機能', '前提', '対象機能への影響予測', '変化する物理量', '予測の根拠']
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows_out)

    print(f"Done: {len(rows_out)} rows written to {output_path}")


if __name__ == '__main__':
    input_path = sys.argv[1] if len(sys.argv) > 1 else 'prediction_input.csv'
    output_path = sys.argv[2] if len(sys.argv) > 2 else 'prediction_expanded.csv'
    expand_prediction(input_path, output_path)
