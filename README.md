# resume-format-converter

Word職務経歴書の表から案件情報を抽出し、転職サイト向けテキストへ変換するPythonツールです。

## セットアップ（Windows / PowerShell）

Python 3.11以上で動作します。以下はPython 3.11を使う手順です。
変換の実行には追加ライブラリのインストールは不要です。

### 1. Pythonをインストール（初回のみ）

Pythonが未導入の場合、PowerShellで公式のPython Install Managerをインストールします。
既にPython 3.11以上を利用できる場合は、この手順を省略できます。

```powershell
winget install --id 9NQ7512CXL7T --exact --source msstore
```

完了したらPowerShellを開き直し、Python 3.11のインストールと確認を行います。
`pymanager` はPython Install Managerのコマンドです。

```powershell
pymanager install 3.11
pymanager exec -V:3.11 --version
```

`winget` が使えない場合は、[Python公式サイト](https://www.python.org/downloads/windows/)から
Python Install Managerをインストールしてください。
詳細は [Python公式のWindows向け手順](https://docs.python.org/3/using/windows.html) を参照してください。

### 2. プロジェクトを取得

GitHubの「Code」→「Download ZIP」からダウンロードして展開します（Gitのインストールは不要）。
取得したフォルダのうち、README.mdがある場所をエクスプローラーで開き、
アドレスバーに `powershell` と入力してEnterを押すと、その場所でPowerShellを開けます。

取得元URL → https://github.com/Takayuki-Isrg/resume-format-converter

Gitを利用する場合は、GitHubの「Code」からコピーした実際のURLを使ってクローンしても構いません。
```powershell
git clone https://github.com/Takayuki-Isrg/resume-format-converter.git
cd resume-format-converter
```

### 3. 仮想環境を作成して実行

```powershell
pymanager exec -V:3.11 -m venv .venv
.\.venv\Scripts\python.exe -m src.transfer "input.docx"
```

既存のPythonを使う場合は、最初の行を `python -m venv .venv` に置き換えてください。
`input.docx` は変換したいWordファイルのパスに置き換えます。空白を含むパスも二重引用符で囲んでください。
仮想環境の有効化やPowerShellの実行ポリシー変更は不要です。

2回目以降は、プロジェクトのフォルダで次のコマンドだけを実行します。

```powershell
.\.venv\Scripts\python.exe -m src.transfer "input.docx"
```

### 開発・テストする場合のみ

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m pytest -q
```

元プロジェクトの仮想環境や設定ファイルは不要です。実物の職務経歴書はGitに追加しないでください。

## Word職務経歴書 → 転職サイト向けテキスト変換（ローカルMVP）

表形式の `.docx` を指定フォーマットのUTF-8テキストへ変換します。
このコマンドはPython標準ライブラリだけで動作し、外部通信・OCR・LLM・APIキーを使用しません。

```powershell
.\.venv\Scripts\python -m src.transfer "input.docx"
.\.venv\Scripts\python -m src.transfer "input.docx" --mask-name "山田 太郎" --mask-name "佐藤 花子"
```

入力と同じフォルダへ `input_converted_YYYYMMDD_HHMMSS_ffffff.txt` を作成します。
タイムスタンプは日本時間（JST）、末尾6桁はマイクロ秒です。実行ごとに別ファイルへ保存します。同名ファイルが存在すると
エラーで終了し、上書きしません。入力ファイルは読み取りのみで開きます。
成功時の終了コードは0、読み取り・変換・保存失敗時は1です。警告は標準エラーに表示します。

対応する表:

- 見出しに「期間」「プロジェクト名・内容」（「プロジェクト名」「案件名」も可）がある、1行1案件の表。
- 「プロジェクト名・内容」の先頭を案件名、`【概要】` と `【担当業務】` の後を各項目として抽出。
- 別列の「担当業務」は「担当工程」として分離し、説明文は「担当業務」に保持。「環境・言語等」「技術要素」は技術要素に転記。
- 複数表と繰り返し見出しに対応。抽出時は原文の順序・段落・セル内改行を保持し、不明項目は空欄。期間・技術要素は後述の表示規則で1行化します。
- 会社見出しなど表外の文章や「役割・規模」は出力項目に含めません。期間の修正・成果の補完・並べ替えはしません。
- 同一案件や担当業務内の同一行は削除せず警告。表をまたぐ類似文や部分一致の重複判定は対象外。
- 見出しのない表は、直前の表と直接隣接し、列数が一致して各行に明示的な期間と概要／担当業務の見出しがある場合のみ継続表として抽出（警告付き）。それ以外は警告して対象外。案件行の結合セル・入れ子表・列数不一致は、案件境界が不明なため出力を中止。
  1行1案件に整えた入力のコピーを使用してください。図形・画像内文字・変更履歴の抽出は対象外です。

個人情報のマスクと保存前確認:

- 所属企業名・顧客企業名・プロジェクト名を自動判定でマスクしません。設定のマスク値に登録した場合は置換されます。
- 既存の氏名ラベル・メール・日本の電話番号の検出に加え、行頭の「住所：」「現住所：」「生年月日：」「誕生日：」の値をマスクします。
  住所・生年月日は同じ行の末尾までが対象です。ラベルなし・複数行の住所や日付だけの生年月日は設定に登録してください。
- 氏名は既存の氏名ラベルから取得し、`--mask-name` と設定の `names` でも追加できます。氏名内の空白差は許容します。
- JSON設定を編集すると、対象文字列を追加・変更できます。各項目は文字列の配列です。自動検出は常に有効で設定値と併用します。
- 検出した対象とマスク後の全文をターミナルに表示し、`yes` を入力した場合のみ保存します。
  検出0件でも確認します。空入力・拒否・Ctrl+C・入力終了では出力を作成しません。修正が必要なら中止して設定を直し、再実行してください。
- 検出値は確認のため端末上に表示します。未マスクの中間ファイルや確認ログは作成しません。
- 自動マスクは完全ではありません。保存予定の全文を確認してください。元のWordは変更しません。

設定ファイルの作成・使用例:

```powershell
Copy-Item mask-config.example.json mask-config.local.json
# mask-config.local.json の配列へマスクする値を登録
python -m src.transfer "input.docx"
# 別の設定を使用する場合
python -m src.transfer "input.docx" --config "private/mask-config.json"
```

```json
{
  "names": ["架空 太郎"],
  "emails": ["非標準表記のメール文字列"],
  "phones": ["内線1234"],
  "addresses": ["架空県架空市1-2-3"],
  "birth_dates": ["1990/01/02"]
}
```

省略時はプロジェクト直下の `mask-config.local.json` を読み込みます。存在しなければ自動検出のみで確認へ進みます。
`--config` で明示したファイルが存在しない・形式が不正な場合は保存せずエラーにします。
設定JSONは `.gitignore` で除外します。公開可能な空の `mask-config.example.json` のみ例外です。
既にGit管理中のファイルはignoreでは除外されないため、個人情報を既存の管理ファイルへ記入しないでください。

実装は `src/readfiles.py`（抽出・案件データ・構造警告）、`src/formatting.py`（表示形式）、`src/transfer.py`（マスク・排他的保存・CLI）です。
事前調査・判断は [設計メモ](docs/docx_conversion.md) を参照してください。

```powershell
python -m pytest tests/test_docx_transfer.py -q
```

自動テストには架空のWord内部XMLを使用します。実物のWordと既存出力はリポジトリ外で読み取り専用の比較を実施し、抽出件数は15件から17件へ改善しました。
この結果は特定のWord文書の本文直下の表にある職歴行を対象としたもので、汎用的な抽出精度を示すものではありません。詳細は [出力改善の検証メモ](docs/conversion_improvements.md) を参照してください。

## 導入効果

自分のWord職務経歴書に記載された17案件を対象に、転職サイト向けのテキスト整形・更新時間を比較しました。

| 項目 | 所要時間 |
|---|---:|
| Before：手作業（1案件約5分 × 17案件、**単純計算**） | 約85分 |
| After：ツール利用（17案件全体、実測） | 約15分 |
| 削減時間 | 約70分 |
| 削減率 | 約82.4% |

ツールの利用により、作業時間を約82%削減できました。

※個人の職務経歴書を対象とした結果です。Beforeは1案件の作業時間から算出した推計値であり、他の文書や利用環境で同様の効果が得られることを保証するものではありません。

## 期間・担当工程・技術要素の表示

- 原文が「2025年8月／～／2025年11月／（4ヶ月）」なら、`2025年8月～2025年11月（4ヶ月）` の1行へ整形します。
- 月数は原文にある場合のみ表示し、計算や補正をしません。注記は同じ行の後ろへ保持します。
- 現在までの期間など、通常の年月範囲として解析できない場合は、各行を空白でつなぎます。
- 「担当業務」は説明文、「担当工程」は別列の原文を元の順序・重複を保って「、」区切りの1行にしたものです。両方に同じ工程名があっても削除・要約しません。原文の【成果】も削除しません。
- 技術要素は空行を除き、元の順序・重複・行内の文字列を保持して「、」でつなぎます。
- 抽出件数を確認画面に表示します。対象外の表・行と継続表の引き継ぎは警告で確認できます。
- 職種による自動除外はしません。経理事務・農業・総務人事も対象です。
  将来の出力対象選択は「表番号・行番号で明示選択」を候補とし、今回の変更では全件出力を維持します。

```text
【案件／プロジェクト名】
（原文）

期間：2025年8月～2025年11月（4ヶ月）

【概要】
（原文）

【担当業務】
（説明文の原文）

【担当工程】
（工程列の原文）

【技術要素】
Java、JavaScript、SQL

---
```
