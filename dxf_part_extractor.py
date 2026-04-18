import ezdxf
from pathlib import Path
import re
import sys
import subprocess
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy


# ============================
#  除外語（部品ではないもの）
# ============================

EXCLUDE_WORDS = {
    # プレート名称
    "HOL-B", "STO-B", "DIE-B",

    # 断面図注記
    "A-A", "B-B", "C-C", "D-D", "E-E", "F-F", "G-G",

    # ステーション名
    "1ST-UPR", "2ST-UPR", "3ST-UPR",

    # 工程ラベル
    "ID", "ROT", "FL", "PI-CUT", "SEP",

    # 上下型識別
    "UPR", "LWR",

}

# ============================
#  ハイフン無しの例外部品
# ============================

EXCEPTION_PARTS = {
    "MCPM20",
    "OKJS20",
    "MSW10",
}

# ============================
#  日本語を含む例外部品（完全一致）
# ============================

JAPANESE_EXCEPTION_PARTS = {
    "フタバ P-ISH38-60-22-30",
}

# ============================
#  正規表現パターン
# ============================

# ハイフンを含む一般的な部品番号
part_pattern = re.compile(
    r'^(?=.*[A-Za-z])[A-Za-z0-9]+(?:-[A-Za-z0-9\.\(\)]+)+$'
)

# T系・J系（1文字＋数字）
t_or_j_pattern = re.compile(r'^[TJ]\d+$')

# 日本語判定
def contains_japanese(s: str) -> bool:
    return re.search(r'[\u3040-\u30FF\u4E00-\u9FFF]', s) is not None


# ============================
#  部品番号判定ロジック（最終版）
# ============================

def is_part_number(s: str) -> bool:
    s = s.strip() # 前後の空白・改行を除去
    s = s.replace('*', '')  # 文字列中のアスタリスクをすべて除去

    if not s:
        return False

    # 日本語を含む例外部品（除外ルールより優先）
    if s in JAPANESE_EXCEPTION_PARTS:
        return True

    # 日本語が含まれる行は除外
    if contains_japanese(s):
        return False

    # 除外語（最優先）
    if s in EXCLUDE_WORDS:
        return False

    # T系・J系（1文字＋数字）
    if t_or_j_pattern.match(s):
        return True

    # ハイフン無しの例外部品
    if s in EXCEPTION_PARTS:
        return True

    # ハイフンを含む部品番号
    if part_pattern.match(s):
        return True

    return False


# ============================
#  DXF テキスト抽出ロジック
# ============================

def collect_text_in_entity(e, doc) -> list[str]:
    texts: list[str] = []
    dxftype = e.dxftype()

    if dxftype == "TEXT":
        texts.append(e.dxf.text)

    elif dxftype == "MTEXT":
        texts.append(e.plain_text())

    elif dxftype == "ATTRIB":
        texts.append(e.dxf.text)

    elif dxftype == "INSERT":
        block_name = e.dxf.name
        if block_name in doc.blocks:
            block = doc.blocks[block_name]
            for be in block:
                texts.extend(collect_text_in_entity(be, doc))

    return texts


def extract_all_text(dxf_path: Path) -> list[str]:
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    texts: list[str] = []
    for e in msp:
        texts.extend(collect_text_in_entity(e, doc))

    return texts


# ============================
#  メイン処理（複数 DXF 一括）
# ============================

def write_parts_to_xls(xls_path: Path, part_numbers: list[str]) -> None:
    """D列の最初の空欄セルからpart_numbersを追記してxlsを上書き保存する"""
    D_COL = 3  # 0-indexed

    if xls_path.exists():
        rb = xlrd.open_workbook(str(xls_path), formatting_info=True)
        wb = xl_copy(rb)
        ws = wb.get_sheet(0)
        rs = rb.sheets()[0]

        # D列の最初の空欄行を探す
        start_row = rs.nrows
        for r in range(rs.nrows):
            cell = rs.cell(r, D_COL)
            if cell.ctype == xlrd.XL_CELL_EMPTY or str(cell.value).strip() == "":
                start_row = r
                break
    else:
        # ファイルが存在しない場合は新規作成
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")
        start_row = 0

    for i, pn in enumerate(part_numbers):
        ws.write(start_row + i, D_COL, pn)

    wb.save(str(xls_path))
    print(f"Excelに追記完了: {xls_path}  (D{start_row + 1}行目から {len(part_numbers)}件)")


def open_file(path: Path) -> None:
    """OSに応じてファイルを開く"""
    if sys.platform == "win32":
        import os
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def main():
    if len(sys.argv) < 2:
        print("エラー：DXF ファイルをドロップまたは引数で指定してください。")
        return

    # output_dir = Path(r"C:\python\excess_text_extractor\output_text")
    # Macのホームディレクトリの中の dev/new_project/output_text を指定
    if sys.platform == "win32":
        output_dir = Path(r"C:\python\excess_text_extractor\output_text")
    else:
        output_dir = Path.home() / "dev" / "new_project" / "output_text"
        
    output_dir.mkdir(exist_ok=True)

    dxf_files = [Path(p) for p in sys.argv[1:]]

    for dxf_path in dxf_files:
        print(f"処理中: {dxf_path.name}")

        # DXF → テキスト抽出
        texts = extract_all_text(dxf_path)

        # 部品番号だけ抽出
        part_numbers = [t.strip().replace('*', '') for t in texts if is_part_number(t)]

        # TXT 出力
        out_txt = output_dir / (dxf_path.stem + "_parts.txt")
        with out_txt.open("w", encoding="utf-8") as f:
            for pn in part_numbers:
                f.write(pn + "\n")
        print(f"部品番号抽出完了: {out_txt}")

        # XLS への追記（DXFと同じフォルダの同名.xlsファイル）
        xls_path = dxf_path.with_suffix(".xls")
        write_parts_to_xls(xls_path, part_numbers)

        # Excelファイルを開く
        open_file(xls_path)
        print(f"ファイルを開きました: {xls_path}")

    print("=== 全 DXF の部品番号抽出が完了しました ===")


if __name__ == "__main__":
    main()