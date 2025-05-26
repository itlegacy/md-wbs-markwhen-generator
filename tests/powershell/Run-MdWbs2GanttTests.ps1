<#
.SYNOPSIS
    md-wbs2gantt.ps1 の手動テストケースを実行し、出力を整理します。
.DESCRIPTION
    TestCases-MdWbs2Gantt.csv (オプション) またはスクリプト内の定義に基づいて、
    md-wbs2gantt.ps1 を様々なパラメータで実行します。
    結果の検証は手動で行います。
#>
param (
    [string]$TestIdToRun, # 特定のテストIDのみを実行する場合
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path, # プロジェクトルートパス
    [switch]$ForceCleanOutput, # 実行前に既存の出力ファイルを削除するか
    [switch]$RunExcelOnly # Excel出力テストのみを実行する場合
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- スクリプトパス設定 ---
$mainScriptPath = Join-Path $ProjectRoot "src\powershell\md-wbs2gantt.ps1"
$testDataRoot = Join-Path $ProjectRoot "samples\mdwbs2gantt_test_data"
$officialHolidays = Join-Path $ProjectRoot "samples\holiday_lists\official_holidays_jp_example.csv"
$companyHolidays = Join-Path $ProjectRoot "samples\holiday_lists\company_holidays_example.csv"
$excelTemplate = Join-Path $ProjectRoot "samples\excel\wbs-gantt-template.xlsx"

# --- 出力ディレクトリ設定 ---
$outputBaseDir = Join-Path $ProjectRoot "test_outputs\md-wbs2gantt" # テストごとの出力先
if (-not (Test-Path $outputBaseDir)) {
    New-Item -ItemType Directory -Path $outputBaseDir -Force | Out-Null
}

# --- テストケース定義 ---
# ここに直接テストケースを定義するか、外部CSVファイルから読み込む
# 例: ハッシュテーブルの配列として定義
$testCases = @(
    @{
        TestId          = "MD2G-MW-001"
        Description     = "基本的なMarkwhen出力 (正常系)"
        WbsFile         = Join-Path $testDataRoot "MD2G-MW-001_Normal\input.md"
        OutputFormat    = "Markwhen"
        OutputFileSubPath = "MD2G-MW-001\output.mw"
        ExtraParams     = @{} # ExtraParams は空でも定義しておく方が良い
    },
    @{
        TestId          = "MD2G-MW-002"
        Description     = "祝日表示モード: InRange"
        WbsFile         = Join-Path $testDataRoot "MD2G-MW-002_HolidayInRange\input_with_project_dates.md"
        OutputFormat    = "Markwhen"
        OutputFileSubPath = "MD2G-MW-002\output_inrange.mw"
        ExtraParams     = @{ MarkwhenHolidayDisplayMode = "InRange" }
    },
    @{
        TestId          = "MD2G-MW-003"
        Description     = "祝日表示モード: All"
        WbsFile         = Join-Path $testDataRoot "MD2G-MW-003_HolidayAll\input_with_project_dates.md" # MD2G-MW-002と同じ入力で可
        OutputFormat    = "Markwhen"
        OutputFileSubPath = "MD2G-MW-003\output_all_holidays.mw"
        ExtraParams     = @{ MarkwhenHolidayDisplayMode = "All" }
    },
    @{
        TestId          = "MD2G-MW-004"
        Description     = "祝日表示モード: None"
        WbsFile         = Join-Path $testDataRoot "MD2G-MW-004_HolidayNone\input_with_project_dates.md" # MD2G-MW-002と同じ入力で可
        OutputFormat    = "Markwhen"
        OutputFileSubPath = "MD2G-MW-004\output_no_holidays.mw"
        ExtraParams     = @{ MarkwhenHolidayDisplayMode = "None" }
    },
    @{
        TestId          = "MD2G-MW-005"
        Description     = "準正常系 (deadline/durationなし)"
        WbsFile         = Join-Path $testDataRoot "MD2G-MW-005_NoDeadlineDuration\input_missing_attrs.md"
        OutputFormat    = "Markwhen"
        OutputFileSubPath = "MD2G-MW-005\output_missing_attrs.mw"
        ExtraParams     = @{
        }
    },
    @{
        TestId          = "MD2G-MW-006"
        Description     = "依存関係警告 (正常)"
        WbsFile         = Join-Path $testDataRoot "MD2G-MW-006_NormalDeps\input_normal_deps.md" # 正常な依存関係のファイル
        OutputFormat    = "Markwhen"
        OutputFileSubPath = "MD2G-MW-006\output_normal_deps.mw"
        ExtraParams     = @{
        }
    },
    @{
        TestId          = "MD2G-MW-007"
        Description     = "依存関係警告 (矛盾あり)"
        WbsFile         = Join-Path $testDataRoot "MD2G-MW-007_ConflictingDeps\input_conflicting_deps.md"
        OutputFormat    = "Markwhen"
        OutputFileSubPath = "MD2G-MW-007\output_conflicting_deps.mw"
        ExtraParams     = @{
        }
    },
    @{
        TestId          = "MD2G-MW-009"
        Description     = "異常系 (WBSファイルなし)"
        WbsFile         = Join-Path $testDataRoot "NON_EXISTENT_FILE.md" # 存在しないファイルを指定
        OutputFormat    = "Markwhen"
        OutputFileSubPath = "MD2G-MW-009\output_non_existent.mw" # 実際には生成されないはず
        ExpectError     = $true # このテストはエラー終了を期待
        ExtraParams     = @{
        }
    },
    @{
        TestId          = "MD2G-MW-010"
        Description     = "異常系 (公式祝日CSVファイルなし)"
        WbsFile         = Join-Path $testDataRoot "MD2G-MW-001_Normal\input.md"
        OutputFormat    = "Markwhen"
        OutputFileSubPath = "MD2G-MW-010\output_no_official_holiday.mw"
        ExpectError     = $true
        OverrideOfficialHolidaysPath = Join-Path $testDataRoot "NON_EXISTENT_HOLIDAYS.csv"
        ExtraParams     = @{}
    },
    @{
        TestId          = "MD2G-EX-001"
        Description     = "基本的なExcelDirect出力 (正常系)"
        WbsFile         = Join-Path $testDataRoot "MD2G-EX-001_NormalExcel\input.md"
        OutputFormat    = "ExcelDirect"
        OutputFileSubPath = "MD2G-EX-001\output.xlsx"
        PrepareExcel    = $true
        ExtraParams     = @{ ExcelSheetNameOrIndex = 1; ExcelStartRow = 5 }
    },
    @{
        TestId          = "MD2G-EX-002"
        Description     = "ExcelDirect出力 (既存データ上書き)"
        WbsFile         = Join-Path $testDataRoot "MD2G-EX-002_ExcelOverwrite\input_for_overwrite.md"
        OutputFormat    = "ExcelDirect"
        OutputFileSubPath = "MD2G-EX-002\output_overwritten.xlsx"
        PrepareExcel    = $true
        ExtraParams     = @{ ExcelSheetNameOrIndex = "Sheet1"; ExcelStartRow = 3 }
    },
    @{
        TestId          = "MD2G-EX-003"
        Description     = "準正常系 (出力Excelファイルなしで実行、新規作成される)" # 説明を修正
        WbsFile         = Join-Path $testDataRoot "MD2G-EX-001_NormalExcel\input.md"
        OutputFormat    = "ExcelDirect"
        OutputFileSubPath = "MD2G-EX-003\NEW_output.xlsx" # 出力ファイル名を変更し、ディレクトリも分ける
        ExpectError     = $false # ★★★ 修正 ★★★
        PrepareExcel    = $false # この場合はテンプレートコピーは不要
        ExtraParams     = @{ ExcelSheetNameOrIndex = 1 }
    },
    @{
        TestId          = "MD2G-EX-004"
        Description     = "準正常系 (出力Excelシートなし、新規作成される)" # 説明を修正
        WbsFile         = Join-Path $testDataRoot "MD2G-EX-001_NormalExcel\input.md"
        OutputFormat    = "ExcelDirect"
        OutputFileSubPath = "MD2G-EX-004\output_new_sheet.xlsx" # 出力ファイル名を変更
        PrepareExcel    = $true # 既存ファイルに新規シート追加なのでテンプレートはコピー
        ExpectError     = $false # ★★★ 修正 ★★★
        ExtraParams     = @{ ExcelSheetNameOrIndex = "NonExistentSheetName"; ExcelStartRow = 5 }
    }
    # --- ↓↓↓ 以下の重複した不正な定義は削除 ↓↓↓ ---
    # @{
    #     TestId = "MD2G-EX-003"
    #     # ...
    # },
    # @{
    #     TestId = "MD2G-EX-004"
    #     # ...
    # }
)

# --- テスト実行ループ ---
Write-Host "--- md-wbs2gantt.ps1 テスト実行開始 ---" -ForegroundColor Cyan

# Run-MdWbs2GanttTests.ps1 のテスト実行ループ前
$filteredTestCases = $null
if ($TestIdToRun) {
    $selectedTest = $testCases | Where-Object { $_.TestId -eq $TestIdToRun }
    if (-not $selectedTest) {
        Write-Error "指定されたテストID '$TestIdToRun' が見つかりません。"
        exit 1
    }
    $filteredTestCases = @($selectedTest) # 配列として扱う
} elseif ($RunExcelOnly) {
    $filteredTestCases = $testCases | Where-Object { $_.OutputFormat -eq "ExcelDirect" }
    if (-not $filteredTestCases -or $filteredTestCases.Count -eq 0) {
        Write-Warning "実行対象のExcel出力テストケースが見つかりません。"
        # この場合、後続のチェックで「実行対象のテストケースが見つかりません。」と表示され終了する
    }
} else {
    $filteredTestCases = $testCases
}

# 実行対象のテストケースが存在するか最終確認
if (-not $filteredTestCases -or $filteredTestCases.Count -eq 0) {
    Write-Warning "実行対象のテストケースが見つかりません。"
    exit 0
}

foreach ($testCase in $filteredTestCases) {
    if (-not $testCase) {
        Write-Warning "テストケースの定義に問題がある可能性があります (nullまたは空の要素)。"
        continue
    }
    # TestId が定義されているか、WbsFile キーが存在するかなどをチェック
    if (-not $testCase.PSObject.Properties.Name -contains 'TestId' -or -not $testCase.TestId) {
        Write-Warning "TestIdが未定義のテストケース定義をスキップします。"
        continue
    }
    if (-not $testCase.PSObject.Properties.Name -contains 'WbsFile' -or -not $testCase.WbsFile) {
        Write-Warning "テストケース $($testCase.TestId) の WbsFile が未定義です。スキップします。"
        continue
    }


    Write-Host "`n--- テストケース: $($testCase.TestId) - $($testCase.Description) ---" -ForegroundColor Yellow # TestId と Description を表示するように変更

    $Error.Clear() # ★★★ エラー状態をリセット (正しい位置) ★★★

    try {
        $outputDir = Join-Path $outputBaseDir (Split-Path $testCase.OutputFileSubPath -Parent)
        $outputFilePath = Join-Path $outputBaseDir $testCase.OutputFileSubPath

        if (-not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }

        if ($ForceCleanOutput -and (Test-Path $outputFilePath)) {
            Write-Verbose "既存の出力ファイルを削除: $outputFilePath"
            Remove-Item $outputFilePath -Force
        }

        # ExcelDirect出力の場合、テンプレートをコピー
        if ($testCase.ContainsKey('PrepareExcel') -and $testCase['PrepareExcel'] -eq $true) {
            if (Test-Path $excelTemplate) {
                Copy-Item $excelTemplate $outputFilePath -Force
                Write-Verbose "Excelテンプレートをコピー: $excelTemplate -> $outputFilePath"
                Start-Sleep -Seconds 3 # ★★★ コピー後、ファイルが解放されるのを待つために3秒待機 ★★★
            } else {
                Write-Error "FATAL: Excelテンプレートが見つかりません: $excelTemplate。テストケース $($testCase.TestId) は実行できません。"
                continue # このテストケースをスキップ
            }
        }

        # $params の構築 (変更なし)
        $params = @{
            WbsFilePath             = $testCase.WbsFile
            # ... (他の必須パラメータ) ...
            OutputFilePath          = $outputFilePath # $outputFilePath を使用
            OutputFormat            = $testCase.OutputFormat
            Verbose                 = $true
        }
        # CompanyHolidayFilePath は任意なので、キーが存在すれば追加
        if ($testCase.ContainsKey('CompanyHolidayFilePath') -and $testCase.CompanyHolidayFilePath) {
            $params.CompanyHolidayFilePath = $testCase.CompanyHolidayFilePath
        } elseif ($companyHolidays) { # スクリプト全体のデフォルトを使う場合
             $params.CompanyHolidayFilePath = $companyHolidays
        }
        # OverrideOfficialHolidaysPath の処理
        if ($testCase.ContainsKey('OverrideOfficialHolidaysPath') -and $testCase.OverrideOfficialHolidaysPath) {
            $params.OfficialHolidayFilePath = $testCase.OverrideOfficialHolidaysPath
        } else {
            $params.OfficialHolidayFilePath = $officialHolidays
        }


        if ($testCase.PSObject.Properties["ExtraParams"]) {
           if ($testCase.ExtraParams) {
            $testCase.ExtraParams.GetEnumerator() | ForEach-Object {
                $params[$_.Name] = $_.Value
            }
           }
        }

        Write-Host "コマンド実行: $mainScriptPath"
        $paramString = $params.GetEnumerator() | ForEach-Object {"-$($_.Name) '$($_.Value)'"} | Join-String -Separator " "
        Write-Host "実行コマンド: & `"$mainScriptPath`" $paramString"


        # ... (try-catch でのスクリプト実行と結果判定は変更なし) ...
        try {
            Start-Sleep -Seconds 1
            & $mainScriptPath @params -ErrorAction Stop
            if ($testCase.ContainsKey('ExpectError') -and $testCase['ExpectError'] -eq $true) {
                Write-Error "テストケース $($testCase.TestId) はエラーを期待していましたが、正常に終了しました。"
            } else {
                Write-Host "スクリプト実行完了。出力ファイル: $outputFilePath" -ForegroundColor Green
            }
        } catch {
            if ($testCase.ContainsKey('ExpectError') -and $testCase['ExpectError'] -eq $true) {
                Write-Host "テストケース $($testCase.TestId) 実行中にエラーが発生しました (期待通り): $($_.Exception.Message)" -ForegroundColor Green
            } else {
                Write-Error "テストケース $($testCase.TestId) 実行中に予期せぬエラーが発生しました: $($_.Exception.Message)"
                # エラーが発生し、かつ期待していなかった場合は次のテストケースに進まない方が良いかもしれない
                # throw # またはフラグを立てて最後にまとめてエラーを報告する
                continue # 現状は次のテストに進む
            }
        }
        # ... (手動確認項目の表示) ...
    } catch { # foreach ループ自体の包括的なエラーハンドリング
        Write-Error "テストケース $($testCase.TestId) の準備または実行中に致命的なエラーが発生しました: $($_.Exception.Message)"
    }
}

Write-Host "`n--- 全テストケースの実行指示完了 ---" -ForegroundColor Cyan