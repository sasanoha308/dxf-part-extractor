# dxf-part-extractor

DXFファイルから部品番号を自動抽出してExcel・テキストファイルに出力するPythonスクリプト

## 概要

製造業向けのDXFファイルを解析し、部品番号を自動抽出します。
抽出結果はDXFと同じフォルダにXLS形式で保存され、自動的に開きます。
TXTファイルにも同時に出力されます。

手作業によるデータ入力を自動化し、業務効率を改善します。

## 動作環境

- Python 3.13.3
- macOS（動作確認済み）
- Windows（コード上は対応、動作未確認）

## 必要なライブラリ

```bash
pip install ezdxf xlrd xlwt xlutils
```

## 使い方

ターミナルで以下のコマンドを実行します。

```bash
python dxf_part_extractor.py 型図.dxf
```

複数ファイルを一括処理する場合（ファイル名の間は半角スペース）：

```bash
python dxf_part_extractor.py 型図1.dxf 型図2.dxf 型図3.dxf
```

### 出力ファイル

| ファイル | 内容 |
|----------|------|
| `図番.xls` | DXFと同じフォルダに保存。D列に部品番号を追記 |
| `図番_parts.txt` | 部品番号の一覧テキスト |
