# md-wbs2ganttのテストスクリプト
# このスクリプトはプロジェクトルートの 'tests/powershell/' ディレクトリに配置されていることを想定。
# メインスクリプトはプロジェクトルートからの 'src/powershell/' にあることを想定。
# サンプルファイルはプロジェクトルートの 'samples/' ディレクトリ以下にあることを想定。

# --- 設定 ---
try {
    $PSScriptRootResolved = Resolve-Path $PSScriptRoot # スクリプトの実際のパスを取得
} catch {
    Write-Error "テストスクリプトの場所を特定できません。スクリプトを保存してから実行してください。"
    exit 1
}

# プロジェクトルートの決定 (テストスクリプトが tests/powershell/ にある場合)
$ProjectRoot = Resolve-Path (Join-Path $PSScriptRootResolved "..\..")
# もしテストスクリプトが tests/ 直下なら:
# $ProjectRoot = Resolve-Path (Join-Path $PSScriptRootResolved "..")

$SrcDir = Join-Path $ProjectRoot "src\powershell"
$SamplesBaseDir = Join-Path $ProjectRoot "samples" # プロジェクトルート直下の samples ディレクトリ

$MainScriptPath = Join-Path $SrcDir "md-wbs2gantt.ps1" # メインスクリプト名を md-wbs2gantt.ps1 に戻す

# テスト用WBSファイル
$TestWbsFile = Join-Path $SamplesBaseDir "mdwbs\ServiceDev_Project_v5.md"

# 祝日ファイル
$OfficialHolidays = Join-Path $SamplesBaseDir "holiday_lists\official_holidays_jp_example.csv"
$CompanyHolidays = Join-Path $SamplesBaseDir "holiday_lists\company_holidays_example.csv"

# 出力設定
$TestOutputFormat = "ExcelDirect" # "Markwhen" または "ExcelDirect"
$OutputFileNameBase = [System.IO.Path]::GetFileNameWithoutExtension($TestWbsFile)
$TestOutputFile = ""
$ExcelTemplateFile = ""

if ($TestOutputFormat -eq "Markwhen") {
    $TestOutputFileDir = Join-Path $SamplesBaseDir "markwhen_outputs"
    if (-not (Test-Path $TestOutputFileDir)) {
        $null = New-Item -ItemType Directory -Path $TestOutputFileDir -Force
    }
    $TestOutputFile = Join-Path $TestOutputFileDir "${OutputFileNameBase}_output.mw"
} elseif ($TestOutputFormat -eq "ExcelDirect") {
    $TestOutputFileDir = Join-Path $SamplesBaseDir "excel_outputs"
    if (-not (Test-Path $TestOutputFileDir)) {
        $null = New-Item -ItemType Directory -Path $TestOutputFileDir -Force
    }
    $ExcelTemplateFile = Join-Path $SamplesBaseDir "excel\wbs-gantt-template.xlsx" # Excelテンプレートのパス
    $TestOutputFile = Join-Path $TestOutputFileDir "${OutputFileNameBase}_excel_output.xlsx" # 出力ファイル名

    if (-not (Test-Path $ExcelTemplateFile -PathType Leaf)) {
        Write-Error "Excelテンプレートファイルが見つかりません: $ExcelTemplateFile"
        exit 1
    }
    # 既存の出力ファイルを削除（存在する場合）
    if (Test-Path $TestOutputFile) {
        Write-Verbose "既存の出力ファイルを削除します: $TestOutputFile"
        Remove-Item $TestOutputFile -Force -ErrorAction SilentlyContinue
    }
    # テンプレートファイルをコピーして出力ファイルとして使用
    Write-Verbose "テンプレートファイルをコピーしています: $ExcelTemplateFile -> $TestOutputFile"
    Copy-Item -Path $ExcelTemplateFile -Destination $TestOutputFile -Force
    Start-Sleep -Seconds 3 # コピー直後のファイルロックを避けるための待機時間を少し増やす
}

# --- スクリプトと入力ファイルの存在確認 ---
if (-not (Test-Path $MainScriptPath -PathType Leaf)) { Write-Error "メインスクリプトが見つかりません: $MainScriptPath"; exit 1 }
if (-not (Test-Path $TestWbsFile -PathType Leaf)) { Write-Error "テスト用WBSファイルが見つかりません: $TestWbsFile"; exit 1 }
if (-not (Test-Path $OfficialHolidays -PathType Leaf)) { Write-Error "祝日ファイルが見つかりません: $OfficialHolidays"; exit 1 }
if (-not [string]::IsNullOrEmpty($CompanyHolidays) -and -not (Test-Path $CompanyHolidays -PathType Leaf)) {
    Write-Warning "会社休日ファイルが指定されていますが見つかりません: $CompanyHolidays。処理は続行しますが、このファイルは無視されます。"
    $CompanyHolidays = $null
}

# --- (テンプレートファイルのワークシート名確認のブロックは、エラーの原因になるため一旦削除またはコメントアウト) ---
# if ($TestOutputFormat -eq "ExcelDirect") {
#     Write-Host "テンプレートファイルのワークシート (確認用):"
#     try {
#         $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false
#         $workbook = $excel.Workbooks.Open($TestOutputFile)
#         foreach ($sheet in $workbook.Worksheets) {
#             Write-Host "  シート名: $($sheet.Name), インデックス: $($sheet.Index)"
#             # 開始行の確認を追加
#             $usedRange = $sheet.UsedRange
#             if ($usedRange) {
#                 Write-Host "  使用中の開始行: $($usedRange.Row)"
#                 # $scriptParams.ExcelStartRow = $usedRange.Row # ここで設定すると、以下の明示的な設定を上書きしてしまう可能性がある
#             }
#         }
#         $workbook.Close($false); $excel.Quit()
#         $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
#         $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
#         [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
#     } catch {
#         Write-Warning "Excelテンプレートのシート名確認中にエラー: $($_.Exception.Message)"
#     }
# }


# --- スクリプトの実行 ---
Write-Host "メインスクリプト ($($MainScriptPath)) を更新されたパラメータで実行中..."
Write-Host "WBSファイル: $TestWbsFile"
Write-Host "祝日ファイル: $OfficialHolidays"
if (-not [string]::IsNullOrEmpty($CompanyHolidays)) { Write-Host "会社休日ファイル: $CompanyHolidays" } else { Write-Host "会社休日ファイル: 指定なし" }
Write-Host "出力ファイル: $TestOutputFile"
Write-Host "出力形式: $TestOutputFormat"
Write-Host "---"

# メインスクリプトへのパラメータ準備
$scriptParams = @{
    WbsFilePath             = $TestWbsFile
    OfficialHolidayFilePath = $OfficialHolidays
    OutputFilePath          = $TestOutputFile
    OutputFormat            = $TestOutputFormat
    DateFormatPattern       = "yyyy/MM/dd" # ExcelDirect 時のメインスクリプトDateFormatPatternのデフォルト上書き
    DefaultEncoding         = "UTF8"       # メインスクリプトのDefaultEncoding
    Verbose                 = $true        # 詳細ログを有効化
}
if (-not [string]::IsNullOrEmpty($CompanyHolidays)) {
    $scriptParams.CompanyHolidayFilePath = $CompanyHolidays
}
# MarkwhenHolidayDisplayMode は $TestOutputFormat が "Markwhen" の場合のみ設定
if ($TestOutputFormat -eq "Markwhen") {
    $scriptParams.MarkwhenHolidayDisplayMode = "InRange" # Markwhen専用パラメータ
}
# ExcelDirect の場合のパラメータはメインスクリプトのparamブロックでデフォルト値が設定されているので、
# テストスクリプト側で上書きしたい場合のみ設定する。
# 今回はメインスクリプトのデフォルト値 (SheetIdentifier=1, ExcelStartRow=5) を使う想定。
# もしテストスクリプトでこれらを変更したい場合は、以下のように追加する。
# elseif ($TestOutputFormat -eq "ExcelDirect") {
#     $scriptParams.ExcelSheetNameOrIndex = "Sheet1" # または 1
#     $scriptParams.ExcelStartRow = 5
# }

# ExcelDirectの場合、メインスクリプトのデフォルト値を使用するため、ここではExcel関連のパラメータは設定しない。
# Write-Verbose "ExcelStartRowの設定値: $($scriptParams.ExcelStartRow)" # デフォルト値を使う場合はこのログは意味がなくなる

try {
    # メインスクリプトの呼び出し
    # -ProjectInfo や StartDate関連のハッシュテーブル作成は不要
    & $MainScriptPath @scriptParams -ErrorAction Stop

    Write-Host "---"
    Write-Host "スクリプトの実行が完了しました。"

    if (Test-Path $TestOutputFile -PathType Leaf) {
        Write-Host "出力ファイルが生成されました: $TestOutputFile"
        Write-Host "出力ファイルの内容を確認してください。"
    } else {
        Write-Warning "出力ファイルが生成されませんでした: $TestOutputFile"
    }

} catch {
    Write-Error "スクリプト実行中にエラーが発生しました:"
    Write-Error $_.Exception.ToString() # 詳細なエラー情報を表示
    if ($_.InvocationInfo) {
        Write-Error "エラー発生箇所: $($_.InvocationInfo.ScriptName) - Line $($_.InvocationInfo.ScriptLineNumber)"
    }
} finally {
    # 大量のデータ処理時のメモリ使用量を最適化 (メインスクリプト側で解放処理があるので、ここでは不要かも)
    # [System.GC]::Collect()
    # [System.GC]::WaitForPendingFinalizers()
}

Write-Host "テストスクリプトが終了しました。"