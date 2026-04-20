# dxf-part-extractor

DXFファイルから部品番号を自動抽出してExcel・テキストファイルに出力するPythonスクリプト

## 概要

金型設計・プレス金型の部品表作成を効率化するために設計されており、現場で使用される部品番号体系（T系・J系・ハイフン構造・例外部品）に最適化されています。

## Features

- DXF内のTEXT / MTEXT / ATTRIB / ブロック内部文字を完全抽出
- 抽出した文字列から**部品番号だけをフィルタリング**
- 部品番号の判定は以下に対応
  - T系部品（T8 / T10 / T13 / T20 …）
  - J系部品（J16など）
  - 英数字＋ハイフン構造の部品番号
  - ハイフン無しの例外部品（MCPM20など）
- 日本語を含む行は自動除外
- 図面特有の注記（A-A / HOL-B / UPRなど）を除外する辞書方式

## 動作環境

- Python 3.13.3
- macOS（動作確認済み）
- Windows（仮想環境で動作確認済み）

## 必要なライブラリ

```bash
pip install ezdxf xlrd xlwt xlutils
```

## 使い方

### macOS

ターミナルで以下のコマンドを実行します。

```bash
python dxf_part_extractor.py 型図.dxf
```

複数ファイルを一括処理する場合（ファイル名の間は半角スペース）：

```bash
python dxf_part_extractor.py 型図1.dxf 型図2.dxf 型図3.dxf
```

### Windows

仮想環境を有効化後、以下のコマンドを実行します。

```bash
python dxf_part_extractor.py aaa.dxf bbb.dxf
```

## 出力ファイル

| ファイル | 内容 |
|----------|------|
| `図番.xls` | DXFと同じフォルダに保存。D列に部品番号を追記 |
| `図番_parts.txt` | 部品番号の一覧テキスト |
