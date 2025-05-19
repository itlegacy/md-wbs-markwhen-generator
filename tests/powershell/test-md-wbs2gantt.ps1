# md-wbs2ganttのテストスクリプト
# このスクリプトは'tests/powershell'ディレクトリに配置され、
# メインスクリプトはプロジェクトルートからの'src/powershell/'にあります。
# サンプルファイルは'samples/'にあります。

# --- 設定 ---
$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")     # 'tests\powershell'はプロジェクトルート直下にあると仮定
$SrcDir = Join-Path $ProjectRoot "src\powershell"
$TestSamplesDir = Join-Path $ProjectRoot "samples" # サンプルファイルはプロジェクトルート/samples/にあります

$MainScriptPath = Join-Path $SrcDir "md-wbs2gantt.ps1"

# テスト用WBSファイル（samples/から）
$TestWbsFile = Join-Path $TestSamplesDir "mdwbs\ServiceDev_Project_v5.md"

# 祝日ファイル（samples/から）
# 利用可能で推奨される場合はUTF-8版の祝日ファイルを使用
$OfficialHolidays = Join-Path $TestSamplesDir "holiday_lists\official_holidays_jp_example.csv"
# Shift_JIS版でテストしたい場合（スクリプトがエンコーディングを正しく処理する場合）:
# $OfficialHolidays = Join-Path $TestSamplesDir "syukujitsu.csv"

$CompanyHolidays = Join-Path $TestSamplesDir "holiday_lists\company_holidays_example.csv"

# 出力ファイル
$TestOutputFileDir = Join-Path $TestSamplesDir "excel_outputs" # 出力はsamples/excel_outputs/に
if (-not (Test-Path $TestOutputFileDir)) {
    $null = New-Item -ItemType Directory -Path $TestOutputFileDir -Force
}
# WBSファイル名のベース名を出力ファイルに使用
$OutputFileNameBase = [System.IO.Path]::GetFileNameWithoutExtension($TestWbsFile)
$TestOutputFile = Join-Path $TestOutputFileDir "${OutputFileNameBase}_excel_test.xlsx" # Visual Studio Code用の機能拡張は`.mw`のみを許容

# 出力形式
$TestOutputFormat = "ExcelDirect" # または "Mermaid" : Mermaid記法の出力は実装されていません

# Copy the Excel template file to the output folder
$TemplateFile = Join-Path $TestSamplesDir "excel\wbs-gantt-template.xlsx"
$OutputTemplateFile = Join-Path $TestOutputFileDir "${OutputFileNameBase}_excel_test.xlsx"

# 既存のファイルを削除（存在する場合）
if (Test-Path $OutputTemplateFile) {
    Remove-Item $OutputTemplateFile -Force
}

# テンプレートファイルをコピー
Copy-Item -Path $TemplateFile -Destination $OutputTemplateFile -Force

# ファイルが利用可能になるまで待機
$maxAttempts = 10
$attempt = 0
$fileReady = $false

while (-not $fileReady -and $attempt -lt $maxAttempts) {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($OutputTemplateFile)
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        $fileReady = $true
    }
    catch {
        $attempt++
        Start-Sleep -Seconds 1
    }
}

if (-not $fileReady) {
    Write-Error "出力ファイルの準備ができませんでした: $OutputTemplateFile"
    exit 1
}

# --- メインスクリプトの存在確認 ---
if (-not (Test-Path $MainScriptPath -PathType Leaf)) {
    Write-Error "メインスクリプトが見つかりません: $MainScriptPath"
    exit 1
}

# Write-Host "---"
# Get-Content $MainScriptPath | Select-Object -First 10
# Write-Host "---"

# --- 入力ファイルの存在確認 ---
if (-not (Test-Path $TestWbsFile -PathType Leaf)) {
    Write-Error "テスト用WBSファイルが見つかりません: $TestWbsFile"
    exit 1
}
if (-not (Test-Path $OfficialHolidays -PathType Leaf)) {
    Write-Error "祝日ファイルが見つかりません: $OfficialHolidays"
    exit 1
}
if (-not [string]::IsNullOrEmpty($CompanyHolidays) -and -not (Test-Path $CompanyHolidays -PathType Leaf)) {
    Write-Warning "会社休日ファイルが指定されていますが見つかりません: $CompanyHolidays。続行します。"
    $CompanyHolidays = $null # 見つからない場合は明示的にnullに設定
}

# --- テンプレートファイルの確認 ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($TemplateFile)
Write-Host "テンプレートファイルのワークシート:"
foreach ($sheet in $workbook.Worksheets) {
    Write-Host "シート名: $($sheet.Name), インデックス: $($sheet.Index)"
}
$workbook.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# --- スクリプトの実行 ---
Write-Host "md-wbs2excel.ps1を更新されたパラメータで実行中..."
Write-Host "メインスクリプト: $MainScriptPath"
Write-Host "WBSファイル: $TestWbsFile"
Write-Host "祝日ファイル: $OfficialHolidays"
if (-not [string]::IsNullOrEmpty($CompanyHolidays)) {
    Write-Host "会社休日ファイル: $CompanyHolidays"
} else {
    Write-Host "会社休日ファイル: 指定なしまたは見つかりません"
}
Write-Host "出力ファイル: $TestOutputFile"
Write-Host "出力形式: $TestOutputFormat"
Write-Host "---"

# テストスクリプトを修正
$ExcelSheetNameOrIndex = "Sheet1"  # インデックスの代わりにシート名を指定

# スクリプトの実行
try {
    & $MainScriptPath -WbsFilePath $TestWbsFile `
                       -OfficialHolidayFilePath $OfficialHolidays `
                       -CompanyHolidayFilePath $CompanyHolidays `
                       -OutputFilePath $TestOutputFile `
                       -OutputFormat $TestOutputFormat `
                       -Verbose
    Write-Host "---"
    Write-Host "スクリプトの実行が完了しました。"

    if (Test-Path $TestOutputFile -PathType Leaf) {
        Write-Host "出力ファイルが生成されました: $TestOutputFile"
        Write-Host "出力ファイルの内容を確認してください。"
    } else {
        Write-Warning "出力ファイルが生成されませんでした: $TestOutputFile"
    }

    # 大量のデータ処理時のメモリ使用量を最適化
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # 処理の進捗状況をより詳細に表示
    Write-Verbose "処理開始: $($Element.Id) $($Element.Name)"
    Write-Verbose "処理完了: $($Element.Id) $($Element.Name)"

} catch {
    Write-Error "スクリプト実行中にエラーが発生しました:"
    Write-Error $_.Exception.Message
}

Write-Host "テストスクリプトが終了しました。"

# テストスクリプトに以下のケースを追加
# - 不正な日付形式のテスト
# - 大量データの処理テスト
# - エンコーディング変換のテスト

# 関数のヘルプコメントを追加
<#
.SYNOPSIS
    タスクの開始日を計算します。
.DESCRIPTION
    期限日と期間から、土日祝日を除外して開始日を計算します。
.PARAMETER DeadlineDate
    タスクの期限日
.PARAMETER DurationBusinessDays
    タスクの所要日数（営業日）
.PARAMETER Holidays
    祝日リスト
#>