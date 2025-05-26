# MD-WBS Markwhen/Excel Generator

Markmap形式に対応した独自のWBS書式（MD-WBS）のWBSをMarkwhen(.mw)またはExcelに変換したり、ExcelからMD-WBSを生成したりするPowerShellスクリプト群です。

## 1. 概要

このプロジェクトは、特定のMarkdown形式で記述されたWBS（作業分解構成図）ファイルと祝日リストCSVから、Markwhen 形式のタイムラインデータやExcel形式のガントチャートデータを生成するためのPowerShellスクリプト群を提供します。また、ExcelファイルからMD-WBSファイルを生成する機能も備えています。

このツールは、以下の特徴を持つプロジェクト計画の作成と視覚化を支援します。

* **双方向Excel連携:** MD-WBSからExcelへのエクスポート、ExcelからMD-WBSへのインポートが可能です。
* **ローカル環境完結:** 社外秘情報も安全に取り扱えます。
* **テキストベース管理:** MD-WBSファイル（Markdown）と祝日リスト（CSV）はプレーンテキストで管理でき、Gitなどのバージョン管理システムとの親和性が高いです。
* **逆線表対応:** 各タスクの締切日（`deadline`）と所要営業日数（`duration`）から、土日祝を除外して自動的に開始日と終了日を計算します。
* **階層構造の視覚化:** Markdownの見出し構造をMarkwhenのネストされたセクションやグループにマッピングし、プロジェクトの階層を直感的に表示します。
* **Markmap連携:** MD-WBSファイルはMarkdownの見出しベースのため、[Markmap](https://markmap.js.org/) ツール（VS Code拡張機能など）でWBSの構造をマインドマップとして視覚化できます。
* **VS Code中心のワークフロー:** テキスト編集、スクリプト実行、Markwhen/Markmapでの表示確認など、多くの作業をVisual Studio Code内で完結できます。

## 2. プロジェクトフォルダ構成

```plaintext
md-wbs-markwhen-generator/
├── .gitattributes
├── .gitignore
├── README.md                       # このファイル
├── LICENSE                         # (設定されていれば)
│
├── docs/                           # 設計書やマニュアル
│   ├── 01_requirements_definition_v5.yaml
│   ├── 02_script_design_spec.yaml
│   └── user_manual/
│       └── mdwbs_usage_guide.md
│
├── src/                            # スクリプト本体
│   └── powershell/
│       └── md-wbs2markwhen.ps1     # メインスクリプト
│
├── samples/                        # 利用例やテンプレート
│   ├── mdwbs/
│   │   └── ServiceDev_Project_v5.md # MD-WBSサンプル
│   ├── holiday_lists/
│   │   ├── official_holidays_jp_example.csv
│   │   └── company_holidays_example.csv
│   └── markwhen_outputs/
│       └── ServiceDev_Project_v5_output.mw # Markwhen出力サンプル
│
└── tests/                          # テスト関連
    └── powershell/
        └── test-mdwbs2markwhen.ps1 # テスト実行スクリプト
        └── data/                     # テスト用データ```
```

## 3. セットアップ

### 3.1. 前提条件

* **PowerShell:** バージョン7.0以上を推奨。
  * YAMLフロントマターの厳密なパースにはPowerShell 7.3以上、または `powershell-yaml` モジュールのインストールが望ましいです。これがない場合は、スクリプトは簡易的なキーバリュー形式でYAMLフロントマターをパースします。
* **Visual Studio Code (推奨):** 以下の拡張機能の利用を推奨します。
  * Markmap拡張（例: `vscode-markmap` など）: MD-WBSの構造をマインドマップで表示するため。
  * Markwhen拡張（`markwhen.markwhen`）: 生成された `.mw` ファイルをタイムラインとして表示するため。
  * PowerShell拡張（`ms-vscode.powershell`）: スクリプトの編集と実行のため。

### 3.2. スクリプトの配置

1. このリポジトリをクローンまたはダウンロードします。
2. メインスクリプトは `src/powershell/md-wbs2markwhen.ps1` にあります。

## 4. MD-WBSの記述方法

詳細は `docs/user_manual/mdwbs_usage_guide.md` を参照してください。
以下は主要なポイントです。

### 4.1. YAMLフロントマター

ファイルの先頭に `---` で区切られたYAMLブロックでプロジェクト全体のメタ情報を記述します。

**例:**

```yaml
---
title: プロジェクト名 (バージョンなど)
description: プロジェクトの簡単な説明
date: YYYY-MM-DD # ドキュメント作成日など
projectstartdate: YYYY-MM-DD # プロジェクト全体の計画開始日 (任意)
projectoveralldeadline: YYYY-MM-DD # プロジェクト全体の最終締切日 (任意)
view: month # Markwhenのデフォルト表示 (任意、例: month, week)
---
```

### 4.2. WBS要素 (見出し)

すべてのWBS要素（セクション、グループ、タスク、サブタスク）はMarkdownの見出しで表現します。
各見出しの**先頭に階層的なID**を記述します。

**例:**

```markdown
## 1. セクション名
### 1.1. グループ名
#### 1.1.1. タスク名
##### 1.1.1.1. サブタスク名
```

### 4.3. 属性リスト

各見出し要素の属性は、その見出し行の**直後に1段インデント（スペース4つ推奨）されたMarkdownリスト**で記述します。
リストアイテムの形式は `- キーワード: 値` です。値の `#` 以降はコメントとして扱われます。

**利用可能なキーワード:**

* `deadline: YYYY-MM-DD`（タスクの締切日、必須）
* `duration: Xd`（タスクの所要営業日数、必須）
* `status: 値`（例: `pending`, `inprogress`, `completed`）
* `progress: X%` または `X`（タスクの進捗率）
* `assignee: 名前`（担当者）
* `depends: 先行タスクID`（例: `1.1.1`）

### 4.4. 詳細説明

各見出し要素の詳細な説明は、属性リストの後（または属性リストがない場合は見出しの直後）に通常のMarkdown段落として記述できます。

### 4.5. リスト前後の空行

属性リストや、見出しと本文の間などには、適度に空行を入れると可読性が向上します。

## 5. スクリプトの実行方法 (`md-wbs2markwhen.ps1`)

PowerShellターミナルからスクリプトを実行します。

**基本コマンド:**

```powershell
.\src\powershell\md-wbs2markwhen.ps1 -WbsFilePath "パス\to\your\wbs.md" `
                                     -OfficialHolidayFilePath "パス\to\official_holidays.csv" `
                                     -OutputFilePath "パス\to\output\gantt.mw"
```

**`主要なパラメータ`:**

* `-WbsFilePath`（必須）: 入力するMD-WBSファイルのパス。
* `-OfficialHolidayFilePath`（必須）: 国民の祝日などが記載されたCSVファイルのパス。
* `-CompanyHolidayFilePath`（任意）: 会社独自の休日が記載されたCSVファイルのパス。
* `-OutputFilePath`（必須）: 生成されるMarkwhenデータを出力するファイルのパス。
* `-OutputFormat`（必須、デフォルト: "Markwhen"）: 出力形式。現在は "Markwhen" のみサポート。
* `-DateFormatPattern`（任意、デフォルト: "yyyy-MM-dd"）: Markwhen内で使用する日付の書式。
* `-DefaultEncoding`（任意、デフォルト: "UTF8"）: 入出力ファイルのエンコーディング。祝日CSVのエンコーディングに合わせて調整が必要な場合があります（例: 内閣府CSVはShift_JIS）。スクリプトはShift_JISからのフォールバックを試みます。出力はUTF8NoBOMを推奨。
* `-MarkwhenHolidayDisplayMode`（任意、デフォルト: "InRange"）: Markwhen出力時の祝日表示モード。
  * `All`: すべての祝日を表示。
  * `InRange`: プロジェクト期間内の祝日のみ表示。
  * `None`: 祝日を表示しない。

**実行例 (テストスクリプト `tests/powershell/test-mdwbs2markwhen.ps1` も参照):**

```powershell
.\src\powershell\md-wbs2markwhen.ps1 -WbsFilePath ".\samples\mdwbs\ServiceDev_Project_v5.md" `
                                     -OfficialHolidayFilePath ".\samples\holiday_lists\official_holidays_jp_example.csv" `
                                     -CompanyHolidayFilePath ".\samples\holiday_lists\company_holidays_example.csv" `
                                     -OutputFilePath ".\my_project_timeline.mw" `
                                     -Verbose
```

`-Verbose` スイッチを付けると、処理の詳細なログが出力されます。

## 6. 生成されたMarkwhenファイルの利用

生成された `.mw` ファイルは、VS CodeのMarkwhen拡張機能などで開くことで、タイムラインとして視覚的に確認できます。

## 7. 依存関係の警告について

スクリプトは、タスク間の依存関係に論理的な矛盾（例: 先行タスクが終わる前に後続タスクが開始されている）がある場合に、コンソールに警告メッセージを出力します。
これらの警告が出た場合は、MD-WBSファイル内の該当タスクの `deadline`, `duration`, `depends` の設定を見直し、計画の矛盾を解消してください。

## 8. 既知の制約・注意事項

* **YAMLフロントマターのパース:** PowerShell 7.3未満の環境で `powershell-yaml` モジュールがない場合、YAMLフロントマターは非常に簡易的なキーバリュー形式でのみパースされます。複雑なYAML構造はサポートされません。
* **MD-WBSの属性リストのインデント:** 各見出し要素の属性リストは、親の見出しより**1段インデント（スペース4つ推奨）**して記述する必要があります。
* **依存関係:** 現在サポートしているのは、先行タスクが完了したら後続タスクが開始可能になる「`終了-開始`」(FS) 関係のみです。また、単一の先行タスクIDの指定を推奨します。

## 9. 今後の改善可能性

* より堅牢なYAMLフロントマターパーサーの組み込み（モジュール依存なしで）。
* MD-WBSの構文チェック機能。
* Mermaid形式出力の再サポート（限定的機能で）。
* リソース（担当者）の負荷状況の簡易表示。

## 10. ライセンス

（未設定）
