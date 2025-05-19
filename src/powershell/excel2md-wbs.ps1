#Requires -Version 7.0

<#
.SYNOPSIS
    Excel WBS/ガントチャートからMD-WBSファイルを生成します。
.DESCRIPTION
    指定されたExcelファイルを読み込み、定義された列マッピングと計算ロジックに基づいて、
    階層ID付き見出しと属性リストを用いたMD-WBS形式のテキストファイルを出力します。
.PARAMETER ExcelFilePath
    入力するExcelファイルのパス。(必須)
.PARAMETER OfficialHolidayFilePath
    国民の祝日などが記載されたCSVファイルのパス。(必須)
.PARAMETER CompanyHolidayFilePath
    会社独自の休日が記載されたCSVファイルのパス。(任意)
.PARAMETER OutputMdWbsFilePath
    生成されるMD-WBSファイルのパス。(必須)
.PARAMETER SheetIdentifier
    処理対象のExcelシート名またはインデックス(1始まり)。(任意、デフォルト1)
.PARAMETER DataStartRow
    Excelシート内のWBSデータが開始する行番号。(任意、デフォルト5)
.PARAMETER ExcelProjectNameCell
    プロジェクト名が記載されているExcelセル番地。(任意、デフォルト "D1")
.PARAMETER ExcelProjectOverallStartDateCell
    プロジェクト全体の開始日が記載されているExcelセル番地。(任意、デフォルト "O1")
.PARAMETER DefaultEncoding
    出力MD-WBSファイルのエンコーディング。(任意、デフォルト "UTF8NoBOM")
.EXAMPLE
    .\excel2md-wbs.ps1 -ExcelFilePath ".\MyProject.xlsx" -OfficialHolidayFilePath ".\holidays_jp.csv" -OutputMdWbsFilePath ".\MyProject.md"
.NOTES
    Version: 1.0
    Author: AI Assistant (Gemini)
    Date: 2025-05-17
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelFilePath,

    [Parameter(Mandatory = $true)]
    [string]$OfficialHolidayFilePath,

    [Parameter(Mandatory = $false)]
    [string]$CompanyHolidayFilePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputMdWbsFilePath,

    [Parameter(Mandatory = $false)]
    [object]$SheetIdentifier = 1,

    [Parameter(Mandatory = $false)]
    [int]$DataStartRow = 5,

    [Parameter(Mandatory = $false)]
    [string]$ExcelProjectNameCell = "D1",

    [Parameter(Mandatory = $false)]
    [string]$ExcelProjectOverallStartDateCell = "O1",

    [Parameter(Mandatory = $false)]
    [string]$DefaultEncoding = "UTF8NoBOM"
)

# --- グローバル変数・初期設定 ---
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$Global:DateFormatPattern = "yyyy-MM-dd" # MD-WBSに出力する日付の書式

# --- クラス定義 (MD-WBS出力用) ---
class MdWbsOutputElement {
    [string]$Id
    [string]$Name
    [int]$HierarchyLevel # Markdownの見出しレベル ## -> 2
    [System.Collections.Hashtable]$Attributes = @{} # deadline, duration, status, progress, assignee, depends
    [string]$DescriptionText = ""
    [System.Collections.Generic.List[MdWbsOutputElement]]$Children = [System.Collections.Generic.List[MdWbsOutputElement]]::new()
}

class MdProjectMetadata {
    [string]$Title
    [string]$Description
    [datetime]$DefinedDate
    [datetime]$ProjectPlanStartDate
    [datetime]$ProjectPlanOverallDeadline
    [string]$View = "month"
}


# --- 主要関数 ---

function Import-HolidayListFromCsv { # 以前のスクリプトから流用・調整
    param(
        [Parameter(Mandatory)] [string]$InputOfficialHolidayFilePath,
        [Parameter(Mandatory=$false)] [string]$InputCompanyHolidayFilePath,
        [Parameter(Mandatory)] [string]$BaseEncoding # CSV読み込み時のエンコーディング
    )
    Write-Verbose "祝日リストの読み込みを開始します。"
    $mergedHolidaysSet = New-Object System.Collections.Generic.HashSet[datetime]
    $filePathsToProcess = @()
    if (-not [string]::IsNullOrEmpty($InputOfficialHolidayFilePath) -and (Test-Path $InputOfficialHolidayFilePath -PathType Leaf)) { $filePathsToProcess += $InputOfficialHolidayFilePath } 
    elseif(-not [string]::IsNullOrEmpty($InputOfficialHolidayFilePath)) { Write-Warning "法定休日ファイル '$InputOfficialHolidayFilePath' が見つかりません。" }
    if (-not [string]::IsNullOrEmpty($InputCompanyHolidayFilePath) -and (Test-Path $InputCompanyHolidayFilePath -PathType Leaf)) { $filePathsToProcess += $InputCompanyHolidayFilePath } 
    elseif(-not [string]::IsNullOrEmpty($InputCompanyHolidayFilePath)) { Write-Warning "会社休日ファイル '$InputCompanyHolidayFilePath' が見つかりません。" }

    if ($filePathsToProcess.Count -eq 0) { Write-Warning "有効な祝日ファイルが指定されていません。"; return ([System.Collections.Generic.List[datetime]]::new()) }

    foreach ($filePath in $filePathsToProcess) {
        Write-Verbose "処理中の祝日ファイル: $filePath"; $csvContent = $null
        try {
            $currentEncoding = $BaseEncoding
            try { $csvContent = Get-Content -Path $filePath -Raw -Encoding $currentEncoding -ErrorAction Stop } 
            catch { Write-Verbose "$filePath $currentEncoding 失敗。Shift_JIS試行"; $currentEncoding = "Shift_JIS"; $csvContent = Get-Content -Path $filePath -Raw -Encoding $currentEncoding -ErrorAction Stop }
            
            if ([string]::IsNullOrEmpty($csvContent)) { Write-Warning "ファイル内容が空: $filePath"; continue }
            
            $csvContentForParsing = $csvContent
            if ($currentEncoding -ieq "UTF8" -and $csvContentForParsing.StartsWith([char]0xEF + [char]0xBB + [char]0xBF)) { $csvContentForParsing = $csvContentForParsing.Substring(3); Write-Verbose "UTF-8 BOM除去: $filePath" }

            $lines = $csvContentForParsing -split '\r?\n' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            if ($lines.Count -lt 2) { Write-Warning "データ行不足(ヘッダ込2行以上必要): $filePath"; continue }
            $dataLines = $lines | Select-Object -Skip 1 
            if ($dataLines -isnot [string[]] -and $null -ne $dataLines) { $dataLines = @($dataLines) } 
            if ($null -eq $dataLines -or $dataLines.Count -eq 0) { Write-Warning "スキップ後のデータ行無: $filePath"; continue }

            $holidaysInFile = $dataLines | ConvertFrom-Csv -Header "日付", "名称" -ErrorAction Stop
            foreach ($holidayEntry in $holidaysInFile) {
                $dateString = $null
                if ($holidayEntry.PSObject.Properties.Name -contains "日付") { $dateString = $holidayEntry."日付" } 
                if (-not [string]::IsNullOrWhiteSpace($dateString)) {
                    try { 
                        $dateValue = [datetime]::Parse($dateString, [System.Globalization.CultureInfo]::InvariantCulture).Date
                        if (-not $mergedHolidaysSet.Contains($dateValue)) { [void]$mergedHolidaysSet.Add($dateValue) }
                    } catch { Write-Warning "無効日付形式: '$dateString' ($filePath)" }
                }
            }
        } catch { Write-Warning "祝日ファイル処理エラー: $filePath. Error: $($_.Exception.Message)" }
    }
    Write-Verbose "祝日リスト読込完了。合計 $($mergedHolidaysSet.Count)件。"
    return [System.Collections.Generic.List[datetime]]::new($mergedHolidaysSet)
}

function Calculate-BusinessDays { # 新規: 開始日と終了日から営業日数を計算
    param(
        [Parameter(Mandatory)] [datetime]$StartDate,
        [Parameter(Mandatory)] [datetime]$EndDate,
        [Parameter(Mandatory)] [System.Collections.Generic.List[datetime]]$Holidays
    )
    if ($StartDate -gt $EndDate) { return 0 } # 開始日が終了日より後なら0日
    
    $businessDays = 0
    $currentDate = $StartDate.Date
    while ($currentDate -le $EndDate.Date) {
        if ($currentDate.DayOfWeek -ne [DayOfWeek]::Saturday -and $currentDate.DayOfWeek -ne [DayOfWeek]::Sunday) {
            if (-not $Holidays.Contains($currentDate)) {
                $businessDays++
            }
        }
        $currentDate = $currentDate.AddDays(1)
    }
    return $businessDays
}

function Calculate-TargetEndDate { # 新規: 開始日と営業日数から終了日を計算
    param(
        [Parameter(Mandatory)] [datetime]$StartDate,
        [Parameter(Mandatory)] [int]$DurationBusinessDays,
        [Parameter(Mandatory)] [System.Collections.Generic.List[datetime]]$Holidays
    )
    if ($DurationBusinessDays -le 0) { return $StartDate }
    $calculatedEndDate = $StartDate
    $daysCounted = 0
    # 開始日自身が営業日なら1日目としてカウント
    if ($StartDate.DayOfWeek -ne [DayOfWeek]::Saturday -and $StartDate.DayOfWeek -ne [DayOfWeek]::Sunday -and (-not $Holidays.Contains($StartDate.Date))) {
        $daysCounted = 1
    }
    while ($daysCounted < $DurationBusinessDays) {
        $calculatedEndDate = $calculatedEndDate.AddDays(1)
        if ($calculatedEndDate.DayOfWeek -ne [DayOfWeek]::Saturday -and $calculatedEndDate.DayOfWeek -ne [DayOfWeek]::Sunday) {
            if (-not $Holidays.Contains($calculatedEndDate.Date)) {
                $daysCounted++
            }
        }
        if ($calculatedEndDate -gt $StartDate.AddYears(2)) { Write-Warning "Calculate-TargetEndDate: 計算が2年以上先に及びました。"; return $StartDate.AddYears(2) }
    }
    return $calculatedEndDate
}


function Convert-ExcelDataToMdWbs {
    param(
        [Parameter(Mandatory)]
        [object]$Worksheet, # ExcelワークシートのCOMオブジェクト
        [Parameter(Mandatory)]
        [int]$StartDataRow,
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[datetime]]$Holidays
    )
    Write-Verbose "ExcelデータのMD-WBS要素への変換を開始します。"
    $rootElements = [System.Collections.Generic.List[MdWbsOutputElement]]::new()
    $parentStack = New-Object System.Collections.Stack
    $currentRow = $StartDataRow

    while ($true) {
        # A列のIDを読み取り、空ならデータ終了とみなす
        $idCellValue = $Worksheet.Cells.Item($currentRow, 1).Value2
        if ([string]::IsNullOrWhiteSpace($idCellValue)) {
            Write-Verbose "A列が空のため、Excelデータ読み取り終了 (行: $currentRow)。"
            break
        }
        $id = $idCellValue -replace "'", "" # 先頭のアポストロフィ除去

        # 階層レベルと要素名を決定 (B, C, D列の値から)
        $name = ""
        $level = 0 # MD-WBSの見出しレベル (## -> 2)
        $elementType = "Task" # デフォルトはタスク

        if (-not [string]::IsNullOrWhiteSpace($Worksheet.Cells.Item($currentRow, 2).Value2)) { # B列: 大分類
            $name = $Worksheet.Cells.Item($currentRow, 2).Value2
            $level = 2
            $elementType = "Section" # または "Group" の最上位
        } elseif (-not [string]::IsNullOrWhiteSpace($Worksheet.Cells.Item($currentRow, 3).Value2)) { # C列: 中分類
            $name = $Worksheet.Cells.Item($currentRow, 3).Value2
            $level = 3
            $elementType = "Group"
        } elseif (-not [string]::IsNullOrWhiteSpace($Worksheet.Cells.Item($currentRow, 4).Value2)) { # D列: タスク名称
            $name = $Worksheet.Cells.Item($currentRow, 4).Value2
            $level = 4 # レベル4以上はタスクと仮定
            # より深い階層はIDのドットの数などで判定も可能だが、ここではExcelの列で簡易判定
        } else {
            Write-Warning "行 $currentRow: B,C,D列に名称が見つかりません。この行をスキップします。"
            $currentRow++; continue
        }
        
        $newElement = New-Object MdWbsOutputElement
        $newElement.Id = $id
        $newElement.Name = $name
        $newElement.HierarchyLevel = $level
        # $newElement.ElementType は現状設定不要 (Format関数側でレベルから判断)

        # 属性の読み取りと設定
        $startDateInputStr = $Worksheet.Cells.Item($currentRow, 18).Text # R列 (Textで文字列として取得)
        $endDateInputStr = $Worksheet.Cells.Item($currentRow, 19).Text   # S列
        $durationInputStr = $Worksheet.Cells.Item($currentRow, 20).Text # T列
        
        $startDateInput = $null
        $endDateInput = $null
        $durationInput = 0

        if (-not [string]::IsNullOrWhiteSpace($startDateInputStr) -and ($startDateInputStr -as [datetime])) { $startDateInput = [datetime]$startDateInputStr }
        if (-not [string]::IsNullOrWhiteSpace($endDateInputStr) -and ($endDateInputStr -as [datetime])) { $endDateInput = [datetime]$endDateInputStr }
        if (-not [string]::IsNullOrWhiteSpace($durationInputStr) -and ($durationInputStr -match "^\d+$")) { $durationInput = [int]$durationInputStr }

        # MD-WBSのdeadlineとdurationを決定
        if ($endDateInput -and $durationInput -gt 0) { # Rule 1 (SとT)
            $newElement.Attributes["deadline"] = $endDateInput.ToString($Global:DateFormatPattern)
            $newElement.Attributes["duration"] = "$($durationInput)d"
            # TODO: 矛盾チェック (S,Tから計算される開始日とR列の比較) をここに入れるか、別途行う
        } elseif ($startDateInput -and $durationInput -gt 0) { # Rule 2 (RとT)
            $calculatedRealEndDate = Calculate-TargetEndDate -StartDate $startDateInput -DurationBusinessDays $durationInput -Holidays $Holidays
            $newElement.Attributes["deadline"] = $calculatedRealEndDate.ToString($Global:DateFormatPattern)
            $newElement.Attributes["duration"] = "$($durationInput)d"
        } elseif ($startDateInput -and $endDateInput) { # Rule 3 (RとS)
            $calculatedBusinessDays = Calculate-BusinessDays -StartDate $startDateInput -EndDate $endDateInput -Holidays $Holidays
            if ($calculatedBusinessDays -gt 0) {
                $newElement.Attributes["duration"] = "$($calculatedBusinessDays)d"
            } else { # 期間0以下は警告し、durationを1dと仮定するなど
                Write-Warning "タスク '$id $name': 開始日と終了日から計算された期間が0日以下です。durationを1dとします。"
                $newElement.Attributes["duration"] = "1d"
            }
            $newElement.Attributes["deadline"] = $endDateInput.ToString($Global:DateFormatPattern)
        } else { # Rule 4 (情報不足)
            Write-Warning "タスク '$id $name': deadlineとdurationを確定できません。S列(終了入力)の値をdeadlineに設定します(もしあれば)。"
            if ($endDateInput) { $newElement.Attributes["deadline"] = $endDateInput.ToString($Global:DateFormatPattern) }
            # duration は空か、デフォルト値 (例: 1d) を入れてコメントで促す
        }

        # その他の属性
        if (-not [string]::IsNullOrWhiteSpace($Worksheet.Cells.Item($currentRow, 24).Value2)) { # X列: 進捗率
            $progressExcel = $Worksheet.Cells.Item($currentRow, 24).Value2
            if ($progressExcel -is [double] -and $progressExcel -ge 0 -and $progressExcel -le 1) {
                $newElement.Attributes["progress"] = "$([int]($progressExcel * 100))%"
            } elseif ($progressExcel -is [string] -and $progressExcel -match "^\d{1,3}%?$") {
                 $newElement.Attributes["progress"] = $progressExcel
            }
        }
        
        $actualStartDateExcel = $Worksheet.Cells.Item($currentRow, 25).Text # Y列
        $actualEndDateExcel = $Worksheet.Cells.Item($currentRow, 26).Text   # Z列
        if (-not [string]::IsNullOrWhiteSpace($actualEndDateExcel) -and ($actualEndDateExcel -as [datetime])) {
            $newElement.Attributes["status"] = "completed"
        } elseif (-not [string]::IsNullOrWhiteSpace($actualStartDateExcel) -and ($actualStartDateExcel -as [datetime])) {
            $newElement.Attributes["status"] = "inprogress"
        } else {
            $newElement.Attributes["status"] = "pending"
        }

        if (-not [string]::IsNullOrWhiteSpace($Worksheet.Cells.Item($currentRow, 16).Value2)) { # P列: 担当者
            $newElement.Attributes["assignee"] = $Worksheet.Cells.Item($currentRow, 16).Value2
        }
        if (-not [string]::IsNullOrEmpty($Worksheet.Cells.Item($currentRow, 6).Value2)) { # F列: 先行タスク番号
             $newElement.Attributes["depends"] = $Worksheet.Cells.Item($currentRow, 6).Value2
        }
        
        $newElement.DescriptionText = if (-not [string]::IsNullOrEmpty($Worksheet.Cells.Item($currentRow, 9).Value2)) { $Worksheet.Cells.Item($currentRow, 9).Value2 } else { "" } # I列: アクションプラン

        # 階層構造の構築
        while ($parentStack.Count -gt 0 -and ($parentStack.Peek()).HierarchyLevel -ge $newElement.HierarchyLevel) {
            [void]$parentStack.Pop()
        }
        if ($parentStack.Count -gt 0) {
            ($parentStack.Peek()).Children.Add($newElement)
        } else {
            $rootElements.Add($newElement)
        }
        $parentStack.Push($newElement)
        
        $currentRow++
        if ($currentRow -gt ($Worksheet.UsedRange.Rows.Count + $DataStartRow)) { # 無限ループ防止 (UsedRangeの行数を超える場合)
            Write-Warning "ExcelシートのUsedRangeを超えたため読み取りを終了します。"
            break
        }
    }
    Write-Verbose "Excelデータ変換完了。トップレベル要素数: $($rootElements.Count)"
    return $rootElements
}


function Format-MdWbsFromElements {
    param(
        [Parameter(Mandatory)]
        [MdProjectMetadata]$Metadata,
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[MdWbsOutputElement]]$RootElements
    )
    Write-Verbose "MD-WBS形式のテキストを生成します。"
    $mdOutputLines = New-Object System.Collections.Generic.List[string]

    # YAMLフロントマター出力
    $mdOutputLines.Add("---")
    if (-not [string]::IsNullOrEmpty($Metadata.Title)) { $mdOutputLines.Add("title: $($Metadata.Title)") }
    if (-not [string]::IsNullOrEmpty($Metadata.Description)) { $mdOutputLines.Add("description: $($Metadata.Description)") }
    if ($Metadata.DefinedDate -ne ([datetime]::MinValue)) { $mdOutputLines.Add("date: $($Metadata.DefinedDate.ToString('yyyy-MM-dd'))") }
    if ($Metadata.ProjectPlanStartDate -ne ([datetime]::MinValue)) { $mdOutputLines.Add("projectstartdate: $($Metadata.ProjectPlanStartDate.ToString('yyyy-MM-dd'))") }
    if ($Metadata.ProjectPlanOverallDeadline -ne ([datetime]::MinValue)) { $mdOutputLines.Add("projectoveralldeadline: $($Metadata.ProjectPlanOverallDeadline.ToString('yyyy-MM-dd'))") }
    if (-not [string]::IsNullOrEmpty($Metadata.View)) { $mdOutputLines.Add("view: $($Metadata.View)") }
    $mdOutputLines.Add("---")
    $mdOutputLines.Add("") # 空行

    # WBS要素を再帰的に処理してMD形式に変換
    function ConvertTo-MdRecursive {
        param(
            [MdWbsOutputElement]$Element,
            [System.Collections.Generic.List[string]]$OutputLinesRef
        )
        $headingMarker = "#" * $Element.HierarchyLevel
        $OutputLinesRef.Add("$headingMarker $($Element.Id) $($Element.Name)")
        
        # 属性リストの出力 (1段インデント)
        if ($Element.Attributes.Count -gt 0) {
            $OutputLinesRef.Add("") # 属性リストの前に空行
            foreach ($key in $Element.Attributes.PSObject.Properties.Name | Sort-Object) {
                if (-not [string]::IsNullOrEmpty($Element.Attributes[$key])) { # 値が空でない属性のみ出力
                    $OutputLinesRef.Add("    - $($key): $($Element.Attributes[$key])")
                }
            }
            $OutputLinesRef.Add("") # 属性リストの後に空行
        }

        # 説明文の出力
        if (-not [string]::IsNullOrEmpty($Element.DescriptionText)) {
            # Markdownとして出力するため、行ごとに処理 (Excelのセル内改行はLFになっている想定)
            foreach($descLine in ($Element.DescriptionText -split "`n")) {
                 $OutputLinesRef.Add($descLine)
            }
            $OutputLinesRef.Add("") # 説明文の後に空行
        }

        # 子要素を再帰処理
        foreach ($child in $Element.Children) {
            ConvertTo-MdRecursive -Element $child -OutputLinesRef $OutputLinesRef
        }
    }

    foreach ($rootElement in $RootElements) {
        ConvertTo-MdRecursive -Element $rootElement -OutputLinesRef $mdOutputLines
    }
    
    Write-Verbose "MD-WBSテキスト生成完了。"
    return $mdOutputLines -join [System.Environment]::NewLine
}

# --- メイン処理 ---
try {
    Write-Host "処理を開始します (Excel to MD-WBS)..." -ForegroundColor Green

    $Global:Holidays = Import-HolidayListFromCsv -InputOfficialHolidayFilePath $OfficialHolidayFilePath -InputCompanyHolidayFilePath $CompanyHolidayFilePath -BaseEncoding $DefaultEncoding
    
    # Excel操作の準備
    $excelCom = New-Object -ComObject Excel.Application
    $excelCom.Visible = $false # バックグラウンドで実行
    $excelCom.DisplayAlerts = $false

    if (-not (Test-Path $ExcelFilePath -PathType Leaf)) { throw "入力Excelファイル '$ExcelFilePath' が見つかりません。" }
    $workbook = $excelCom.Workbooks.Open($ExcelFilePath)
    $worksheet = $null
    try {
        $worksheet = $workbook.Worksheets.Item($SheetIdentifier)
    } catch {
        throw "指定されたワークシート '$SheetIdentifier' がExcelファイル内に見つかりません。"
    }
    if ($null -eq $worksheet) { throw "ワークシート '$SheetIdentifier' の取得に失敗しました。" }

    Write-Verbose "Excelファイルとシートの読み込み成功。"

    # プロジェクトメタデータの読み取り
    $mdProjectMetadata = New-Object MdProjectMetadata
    if (-not [string]::IsNullOrEmpty($worksheet.Cells.Item($ExcelProjectNameCell.Split([char[]]'0..9')[0], [int]($ExcelProjectNameCell -replace '[^0-9]')).Value2)) { # "D1" -> D, 1
        $mdProjectMetadata.Title = $worksheet.Cells.Item([int]($ExcelProjectNameCell -replace '[^0-9]'), $ExcelProjectNameCell.Split([char[]]'0..9')[0]).Value2
    }
    if (-not [string]::IsNullOrEmpty($worksheet.Cells.Item($ExcelProjectOverallStartDateCell.Split([char[]]'0..9')[0], [int]($ExcelProjectOverallStartDateCell -replace '[^0-9]')).Text) -and `
       ($worksheet.Cells.Item([int]($ExcelProjectOverallStartDateCell -replace '[^0-9]'), $ExcelProjectOverallStartDateCell.Split([char[]]'0..9')[0]).Text -as [datetime])) {
        $mdProjectMetadata.ProjectPlanStartDate = [datetime]$worksheet.Cells.Item([int]($ExcelProjectOverallStartDateCell -replace '[^0-9]'), $ExcelProjectOverallStartDateCell.Split([char[]]'0..9')[0]).Text
    }
    $mdProjectMetadata.DefinedDate = (Get-Date) # ドキュメント生成日をセット

    # ExcelからWBS要素ツリーを構築
    $wbsRootElements = Convert-ExcelDataToMdWbs -Worksheet $worksheet -StartDataRow $DataStartRow -Holidays $Global:Holidays
    
    # MD-WBSテキストを生成
    $mdWbsContent = Format-MdWbsFromElements -Metadata $mdProjectMetadata -RootElements $wbsRootElements

    # ファイル出力
    try {
        Set-Content -Path $OutputMdWbsFilePath -Value $mdWbsContent -Encoding $DefaultEncoding -Force
        Write-Host "MD-WBSファイルが正常に出力されました: $OutputMdWbsFilePath (Encoding: $DefaultEncoding)" -ForegroundColor Green
    } catch {
        Write-Error "MD-WBSファイルの書き出しに失敗しました: $OutputMdWbsFilePath. Error: $($_.Exception.Message)"
        exit 1
    }

    Write-Host "処理が正常に完了しました。" -ForegroundColor Green

} catch {
    Write-Host "Debug: CRITICAL ERROR - $($_.Exception.GetType().FullName) - $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Debug: StackTrace - $($_.Exception.StackTrace)" -ForegroundColor DarkYellow
    if ($_.InvocationInfo) { Write-Host "Debug: Error at Line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line)" -ForegroundColor Yellow }
    Write-Error "スクリプト実行中に致命的なエラー: $($_.Exception.Message)"
    exit 1
} finally {
    # Excel COMオブジェクトの解放
    if ($excelCom) {
        if ($workbook) { $workbook.Close($false) } # 保存せずに閉じる
        $excelCom.Quit()
        if ($worksheet) { while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) -gt 0) {} }
        if ($workbook) { while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) -gt 0) {} }
        while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelCom) -gt 0) {}
        Remove-Variable excelCom, workbook, worksheet -ErrorAction SilentlyContinue
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        Write-Verbose "Excel COMオブジェクトを解放しました。"
    }
    Remove-Variable -Name Holidays, Global:DateFormatPattern -ErrorAction SilentlyContinue
}