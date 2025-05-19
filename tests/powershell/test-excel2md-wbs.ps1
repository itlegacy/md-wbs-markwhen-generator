# test-excel2md-wbs.ps1 - Test script for excel2md-wbs.ps1

# --- 設定 ---
$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")     # 'tests/' はプロジェクトルート直下にあると仮定
$SrcDir = Join-Path $ProjectRoot "src\powershell"
$TestSamplesDir = Join-Path $PSScriptRoot "samples" # テスト用サンプルファイルは tests/samples/ に配置想定
$OutputMdWbsDir = Join-Path $PSScriptRoot "output_mdwbs" # MD-WBS出力用ディレクトリ

$MainScriptPath = Join-Path $SrcDir "excel2md-wbs.ps1" # メインスクリプト名

# 入力Excelファイル (テスト用データを含む)
$TestExcelFile = Join-Path $TestSamplesDir "excel_inputs\wbs_input_example.xlsx" # ★テスト用の入力Excelファイルパス

# 祝日ファイル (MD-WBSのduration計算に必要)
$OfficialHolidays = Join-Path $TestSamplesDir "holiday_lists\official_holidays_jp_example.csv"
$CompanyHolidays = Join-Path $TestSamplesDir "holiday_lists\company_holidays_example.csv"

# 出力MD-WBSファイル
if (-not (Test-Path $OutputMdWbsDir)) {
    $null = New-Item -ItemType Directory -Path $OutputMdWbsDir -Force
}
$ExcelFileNameBase = [System.IO.Path]::GetFileNameWithoutExtension($TestExcelFile)
$OutputMdWbsFile = Join-Path $OutputMdWbsDir "${ExcelFileNameBase}_output.md"

# Excelシート指定 (メインスクリプトのデフォルト値と同じか、テスト用に指定)
$TestExcelSheetIdentifier = 1 # または "シート名"
$TestDataStartRow = 5

# Excelのプロジェクトメタ情報セル (メインスクリプトのデフォルト値と同じか、テスト用に指定)
$TestExcelProjectNameCell = "D1"
$TestExcelProjectOverallStartDateCell = "O1"


# --- メインスクリプトの存在確認 ---
if (-not (Test-Path $MainScriptPath -PathType Leaf)) {
    Write-Error "メインスクリプトが見つかりません: $MainScriptPath"
    exit 1
}

# --- 入力ファイルの存在確認 ---
if (-not (Test-Path $TestExcelFile -PathType Leaf)) {
    Write-Error "テスト用Excelファイルが見つかりません: $TestExcelFile"
    # ★事前にテスト用の入力Excelファイル (wbs_input_example.xlsxなど) を
    # tests/samples/excel_inputs/ ディレクトリに配置してください。
    # このExcelには、A列のID、B/C/D列の階層名、R/S/T列の日付/期間情報などが入力されている必要があります。
    exit 1
}
if (-not (Test-Path $OfficialHolidays -PathType Leaf)) {
    Write-Error "祝日ファイルが見つかりません: $OfficialHolidays"
    exit 1
}
if (-not [string]::IsNullOrEmpty($CompanyHolidays) -and -not (Test-Path $CompanyHolidays -PathType Leaf)) {
    Write-Warning "会社休日ファイルが指定されていますが見つかりません: $CompanyHolidays。続行します。"
    $CompanyHolidays = $null
}

# --- スクリプトの実行 ---
Write-Host "Executing excel2md-wbs.ps1 ..."
Write-Host "Main Script: $MainScriptPath"
Write-Host "Input Excel File: $TestExcelFile"
Write-Host "Official Holidays: $OfficialHolidays"
if (-not [string]::IsNullOrEmpty($CompanyHolidays)) {
    Write-Host "Company Holidays: $CompanyHolidays"
} else {
    Write-Host "Company Holidays: Not specified or not found"
}
Write-Host "Output MD-WBS File: $OutputMdWbsFile"
Write-Host "Target Excel Sheet: $TestExcelSheetIdentifier"
Write-Host "Excel Data Start Row: $TestDataStartRow"
Write-Host "Excel Project Name Cell: $TestExcelProjectNameCell"
Write-Host "Excel Project Overall Start Date Cell: $TestExcelProjectOverallStartDateCell"
Write-Host "---"

# スクリプト実行パラメータの構築
$scriptArgs = @{
    ExcelFilePath                     = $TestExcelFile
    OfficialHolidayFilePath           = $OfficialHolidays
    OutputMdWbsFilePath               = $OutputMdWbsFile
    SheetIdentifier                   = $TestExcelSheetIdentifier
    DataStartRow                      = $TestDataStartRow
    ExcelProjectNameCell              = $TestExcelProjectNameCell
    ExcelProjectOverallStartDateCell  = $TestExcelProjectOverallStartDateCell
    # DefaultEncoding                 = "UTF8NoBOM" # 必要に応じて指定
    Verbose                           = $true
}
if (-not [string]::IsNullOrEmpty($CompanyHolidays)) {
    $scriptArgs.CompanyHolidayFilePath = $CompanyHolidays
}

# 数値を文字列に変換する際のヘルパー関数
function ConvertTo-ExcelValue {
    param($Value)
    if ($Value -is [int] -or $Value -is [double]) {
        return [string]$Value
    }
    return $Value
}

try {
    # メインスクリプトの実行
    & $MainScriptPath @scriptArgs -ErrorAction Stop

    Write-Host "---"
    Write-Host "Script execution completed."

    if (Test-Path $OutputMdWbsFile -PathType Leaf) {
        Write-Host "Output MD-WBS file generated at: $OutputMdWbsFile"
        Write-Host "Please verify the content of the output file."
        # 生成されたMD-WBSファイルの内容を一部表示 (オプション)
        # Get-Content $OutputMdWbsFile | Select-Object -First 30
    } else {
        Write-Warning "Output MD-WBS file was NOT generated at: $OutputMdWbsFile"
    }

} catch {
    Write-Error "An error occurred during script execution:"
    Write-Error $_.Exception.Message
    # For more details during debugging:
    # Write-Host ($_.ErrorDetails | Format-List * | Out-String)
    # Write-Host ($_.ScriptStackTrace | Format-List * | Out-String)
}

Write-Host "Test script for excel2md-wbs.ps1 finished."

# --- 今後のテストケースのアイデア ---
# - 様々なExcel入力パターン（R/S/T列の組み合わせ、空白行、特殊文字など）
# - 祝日をまたぐ期間のduration計算テスト
# - 階層構造が正しくMD-WBSに変換されるかのテスト
# - 不正なExcelファイルパスやシート名を指定した場合のエラーハンドリングテスト
# - エンコーディングのテスト (特に日本語ファイル名や内容)