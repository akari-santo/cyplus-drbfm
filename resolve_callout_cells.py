"""
吹き出し図形（wedgeRectCallout）の先端セルを特定するスクリプト

原理：
  吹き出しは「長方形 + 三角形」で構成され、三角形の先端頂点を特定することで
  コメントが紐づくセルを判定する。

  先端の座標計算：
    アンカー(col, colOff, row, rowOff)を基準として、
      dx = colOff + adj1/100000 * cx   (アンカー列左端からのEMUオフセット)
      dy = rowOff + adj2/100000 * cy   (アンカー行上端からのEMUオフセット)
    dx/dy が正なら右・下方向、負なら左・上方向へ進む。

  列幅・行高さのEMU変換：
    列幅(文字数) → ピクセル: trunc(width_chars * MDW + 5)  [MDWは最大桁幅]
    行高さ(pt)   → EMU:     height_pt * 12700

    MDWは描画XMLのxfrm.xとアンカーcolOffから自動キャリブレーションする。
    複数シェイプのアンカーデータを使い、隣接列の実測幅から最適なMDWを逆算。

  注意：非表示列・行は幅0として扱う。
"""

import zipfile
import xml.etree.ElementTree as ET
from math import trunc

# ---- 定数 ----
MDW_DEFAULT = 7.65  # MaxDigitWidth デフォルト値 (キャリブレーション失敗時のフォールバック)
EMU_PER_PX  = 9525  # 1px = 9525 EMU (96dpi)
EMU_PER_PT  = 12700 # 1pt = 12700 EMU

# ---- 名前空間 ----
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
X   = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def calibrate_mdw(draw_root: ET.Element, col_widths_chars: dict) -> float:
    """
    描画XMLのアンカー(col, colOff)とxfrm.xから、
    実測列幅(EMU)を収集してMDWを逆算する。

    Args:
        draw_root:        drawing XML のルート要素
        col_widths_chars: {1-indexed 列番号: 文字幅} の辞書

    Returns:
        推定MDW値。サンプル不足時は MDW_DEFAULT を返す。
    """
    # アンカーごとに col_idx → xfrm.x - colOff = col_start_emu を収集
    col_start_samples: dict = {}
    for anchor in draw_root:
        tag = anchor.tag.split('}')[-1]
        if tag not in ('twoCellAnchor', 'oneCellAnchor'):
            continue
        sp = anchor.find(f'{{{XDR}}}sp')
        if sp is None:
            continue
        spPr = sp.find(f'{{{XDR}}}spPr')
        if spPr is None:
            continue
        xfrm = spPr.find(f'{{{A}}}xfrm')
        if xfrm is None:
            continue
        off = xfrm.find(f'{{{A}}}off')
        if off is None:
            continue
        frm = anchor.find(f'{{{XDR}}}from')
        col_0idx = int(frm.find(f'{{{XDR}}}col').text)
        col_off  = int(frm.find(f'{{{XDR}}}colOff').text)
        x        = int(off.get('x'))
        col_start_samples.setdefault(col_0idx, []).append(x - col_off)

    # 中央値でノイズ除去
    col_starts = {
        c: sorted(v)[len(v) // 2]
        for c, v in col_start_samples.items()
    }

    # 隣接する列ペアから実測幅を収集
    mdw_candidates = []
    sorted_cols = sorted(col_starts.keys())
    for i in range(len(sorted_cols) - 1):
        c1, c2 = sorted_cols[i], sorted_cols[i + 1]
        n_cols = c2 - c1
        if n_cols > 3:   # 離れすぎは精度低下するのでスキップ
            continue
        emu_span = col_starts[c2] - col_starts[c1]
        if emu_span <= 0:
            continue

        # c1+1 〜 c2 の文字幅が全て同じか確認（異なる場合はスキップ）
        widths = set()
        for c in range(c1 + 1, c2 + 1):
            col_1idx = c + 1   # 0-indexed → 1-indexed
            widths.add(col_widths_chars.get(col_1idx))
        if len(widths) != 1 or None in widths:
            continue
        char_width = widths.pop()

        # 実測1列EMU → MDWを逆算
        # emu_per_col = trunc(char_width * MDW + 5) * EMU_PER_PX
        # emu_per_col / EMU_PER_PX = trunc(char_width * MDW + 5)
        emu_per_col = emu_span / n_cols
        px = emu_per_col / EMU_PER_PX
        mdw = (px - 5) / char_width
        if 4.0 < mdw < 12.0:   # 非現実的な値は除外
            mdw_candidates.append(mdw)

    if not mdw_candidates:
        return MDW_DEFAULT

    mdw_candidates.sort()
    return mdw_candidates[len(mdw_candidates) // 2]   # 中央値


def col_width_emu(width_chars: float, hidden: bool = False, mdw: float = MDW_DEFAULT) -> int:
    """列幅(文字数) → EMU。非表示列は 0 を返す。"""
    if hidden:
        return 0
    px = trunc(width_chars * mdw + 5)
    return px * EMU_PER_PX


def row_height_emu(height_pt: float, hidden: bool = False) -> int:
    """行高さ(pt) → EMU。非表示行は 0 を返す。"""
    if hidden:
        return 0
    return round(height_pt * EMU_PER_PT)


def build_col_widths_chars(sheet_root: ET.Element) -> dict:
    """
    シートXMLから列幅(文字数)と非表示フラグの辞書を構築。
    キー: 1-indexed 列番号
    値:   (文字幅 float, hidden bool)
    """
    sfp = sheet_root.find(f'{{{X}}}sheetFormatPr')
    default_width = float(sfp.get('defaultColWidth', '8.38')) if sfp is not None else 8.38

    col_chars = {}
    for col_el in sheet_root.findall(f'{{{X}}}cols/{{{X}}}col'):
        min_c  = int(col_el.get('min'))
        max_c  = int(col_el.get('max'))
        width  = float(col_el.get('width', default_width))
        hidden = col_el.get('hidden', '0') == '1'
        for c in range(min_c, max_c + 1):
            col_chars[c] = (width, hidden)

    return col_chars


def build_col_widths(sheet_root: ET.Element, mdw: float) -> dict:
    """
    シートXMLから列幅(EMU)の辞書を構築。
    キー: 1-indexed 列番号
    値:   EMU幅
    """
    sfp = sheet_root.find(f'{{{X}}}sheetFormatPr')
    default_width = float(sfp.get('defaultColWidth', '8.38')) if sfp is not None else 8.38

    col_widths = {}
    for col_el in sheet_root.findall(f'{{{X}}}cols/{{{X}}}col'):
        min_c  = int(col_el.get('min'))
        max_c  = int(col_el.get('max'))
        width  = float(col_el.get('width', default_width))
        hidden = col_el.get('hidden', '0') == '1'
        emu    = col_width_emu(width, hidden, mdw)
        for c in range(min_c, max_c + 1):
            col_widths[c] = emu

    return col_widths


def build_row_heights(sheet_root: ET.Element) -> dict:
    """
    シートXMLから行高さ(EMU)の辞書を構築。
    キー: 1-indexed 行番号
    値:   EMU高さ
    """
    sfp = sheet_root.find(f'{{{X}}}sheetFormatPr')
    default_height = float(sfp.get('defaultRowHeight', '14.0')) if sfp is not None else 14.0

    row_heights = {}
    for row_el in sheet_root.findall(f'{{{X}}}sheetData/{{{X}}}row'):
        r = int(row_el.get('r'))
        ht = float(row_el.get('ht', default_height))
        hidden = row_el.get('hidden', '0') == '1'
        row_heights[r] = row_height_emu(ht, hidden)

    return row_heights


def emu_offset_to_col(anchor_col_0idx: int, dx_emu: int,
                      col_widths: dict, default_col_emu: int) -> int:
    """
    アンカー列(0-indexed)から dx_emu 進んだときの列番号(1-indexed)を返す。
    dx_emu が負の場合は左方向に進む。
    """
    col_1idx = anchor_col_0idx + 1  # 1-indexed に変換
    remaining = dx_emu

    if remaining >= 0:
        # 右方向
        while True:
            w = col_widths.get(col_1idx, default_col_emu)
            if remaining < w or w == 0:
                return col_1idx
            remaining -= w
            col_1idx += 1
    else:
        # 左方向
        col_1idx -= 1
        while col_1idx >= 1:
            w = col_widths.get(col_1idx, default_col_emu)
            remaining += w
            if remaining >= 0:
                return col_1idx
            col_1idx -= 1
        return 1  # 左端を超えた場合


def emu_offset_to_row(anchor_row_0idx: int, dy_emu: int,
                      row_heights: dict, default_row_emu: int) -> int:
    """
    アンカー行(0-indexed)から dy_emu 進んだときの行番号(1-indexed)を返す。
    dy_emu が負の場合は上方向に進む。
    """
    row_1idx = anchor_row_0idx + 1  # 1-indexed に変換
    remaining = dy_emu

    if remaining >= 0:
        # 下方向
        while True:
            h = row_heights.get(row_1idx, default_row_emu)
            if remaining < h or h == 0:
                return row_1idx
            remaining -= h
            row_1idx += 1
    else:
        # 上方向
        row_1idx -= 1
        while row_1idx >= 1:
            h = row_heights.get(row_1idx, default_row_emu)
            remaining += h
            if remaining >= 0:
                return row_1idx
            row_1idx -= 1
        return 1  # 上端を超えた場合


def col_index_to_letter(col_1idx: int) -> str:
    """1-indexed 列番号 → Excel列アルファベット (例: 1→A, 27→AA)"""
    result = ''
    while col_1idx > 0:
        col_1idx, remainder = divmod(col_1idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def resolve_callout_cells(xlsx_path: str) -> dict:
    """
    xlsm/xlsxファイルを読み込み、wedgeRectCallout の先端セルを解決する。

    Returns:
        dict: { shape_id_text: 'A1' 形式のセル参照 }
    """
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        sheet_xml  = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
        draw_xml   = z.read('xl/drawings/drawing1.xml').decode('utf-8')

    sheet_root = ET.fromstring(sheet_xml)
    draw_root  = ET.fromstring(draw_xml)

    # MDWをキャリブレーション
    col_chars = build_col_widths_chars(sheet_root)
    col_chars_only = {c: w for c, (w, _) in col_chars.items()}
    mdw = calibrate_mdw(draw_root, col_chars_only)

    col_widths  = build_col_widths(sheet_root, mdw)
    row_heights = build_row_heights(sheet_root)

    sfp = sheet_root.find(f'{{{X}}}sheetFormatPr')
    default_col_emu = col_width_emu(
        float(sfp.get('defaultColWidth', '8.38')) if sfp is not None else 8.38,
        mdw=mdw,
    )
    default_row_emu = row_height_emu(
        float(sfp.get('defaultRowHeight', '14.0')) if sfp is not None else 14.0
    )

    results = {}

    for anchor in draw_root:
        tag = anchor.tag.split('}')[-1]
        if tag not in ('twoCellAnchor', 'oneCellAnchor'):
            continue

        sp = anchor.find(f'{{{XDR}}}sp')
        if sp is None:
            continue

        # 吹き出し形状かどうか確認
        spPr = sp.find(f'{{{XDR}}}spPr')
        if spPr is None:
            continue
        prstGeom = spPr.find(f'{{{A}}}prstGeom')
        if prstGeom is None or prstGeom.get('prst') != 'wedgeRectCallout':
            continue

        # アンカー位置を取得
        frm = anchor.find(f'{{{XDR}}}from')
        anchor_col = int(frm.find(f'{{{XDR}}}col').text)    # 0-indexed
        col_off    = int(frm.find(f'{{{XDR}}}colOff').text) # EMU
        anchor_row = int(frm.find(f'{{{XDR}}}row').text)    # 0-indexed
        row_off    = int(frm.find(f'{{{XDR}}}rowOff').text) # EMU

        # 図形サイズを取得 (oneCellAnchor は ext から、twoCellAnchor は xfrm から)
        ext_el = anchor.find(f'{{{XDR}}}ext')
        if ext_el is not None:
            cx = int(ext_el.get('cx'))
            cy = int(ext_el.get('cy'))
        else:
            xfrm = spPr.find(f'{{{A}}}xfrm')
            ext2 = xfrm.find(f'{{{A}}}ext')
            cx = int(ext2.get('cx'))
            cy = int(ext2.get('cy'))

        # adj1, adj2 を取得 (デフォルト: 先端が左上方向)
        avLst = prstGeom.find(f'{{{A}}}avLst')
        adj = {}
        if avLst is not None:
            for gd in avLst.findall(f'{{{A}}}gd'):
                fmla = gd.get('fmla', '')
                if fmla.startswith('val '):
                    adj[gd.get('name')] = int(fmla.split()[1])
        adj1 = adj.get('adj1', 18750)   # デフォルト: 18.75% 右
        adj2 = adj.get('adj2', -8333)   # デフォルト: -8.333% 上 (上方向)

        # 先端のアンカー基準オフセット (EMU)
        dx = col_off + round(adj1 / 100000 * cx)
        dy = row_off + round(adj2 / 100000 * cy)

        # EMUオフセット → セル番号
        tip_col = emu_offset_to_col(anchor_col, dx, col_widths, default_col_emu)
        tip_row = emu_offset_to_row(anchor_row, dy, row_heights, default_row_emu)

        # テキストからIDを抽出
        texts = [t.text or '' for t in sp.findall(f'.//{{{A}}}t')]
        shape_text = ''.join(texts).strip()
        shape_id = shape_text.split(':')[0].strip() if ':' in shape_text else shape_text[:20]

        cell_ref = f'{col_index_to_letter(tip_col)}{tip_row}'
        results[shape_id] = {
            'cell': cell_ref,
            'text': shape_text[:60],
        }

    return results


if __name__ == '__main__':
    import sys
    filename = sys.argv[1] if len(sys.argv) > 1 else '社外共有_DRBFM資料_第2世代MEGA_260204_展開版.xlsm'

    print(f"ファイル: {filename}")
    print("=" * 60)

    # キャリブレーション済みMDWを表示
    with zipfile.ZipFile(filename, 'r') as z:
        sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
        draw_xml  = z.read('xl/drawings/drawing1.xml').decode('utf-8')
    sheet_root = ET.fromstring(sheet_xml)
    draw_root  = ET.fromstring(draw_xml)
    col_chars  = {c: w for c, (w, _) in build_col_widths_chars(sheet_root).items()}
    mdw = calibrate_mdw(draw_root, col_chars)
    print(f"キャリブレーション済み MDW = {mdw:.3f} px/char\n")

    results = resolve_callout_cells(filename)
    for shape_id, info in sorted(results.items()):
        print(f"  {shape_id:8s} → {info['cell']:6s}  [{info['text'][:40]}]")
