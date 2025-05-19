#Requires -Version 7.0

<#
.SYNOPSIS
    拡張MD-WBSファイル (新仕様) からMarkwhenタイムラインデータまたはExcelガントチャートシート用データを生成します。
.DESCRIPTION
    YAMLフロントマター (簡易サポート)、階層ID付き見出し、リスト属性を持つ拡張MD-WBSファイルを解析し、
    タスクの開始日を逆算（土日祝除外）して、Markwhen形式のタイムラインデータ、または既存のExcelシートへの
    WBSデータ書き込みを行います。
.NOTES
    Version: 3.7 (YAMLパース修正, コメント行スキップ, WbsElementNode内変数修正)
    Author: AI Assistant (Gemini) based on user requirements
    Date: 2025-05-17
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WbsFilePath,

    [Parameter(Mandatory = $true)]
    [string]$OfficialHolidayFilePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputFilePath,

    [Parameter(Mandatory = $true)]
    [ValidateSet("Markwhen", "ExcelDirect")]
    [string]$OutputFormat = "Markwhen",

    [Parameter(Mandatory = $false)]
    [string]$CompanyHolidayFilePath,

    [Parameter(Mandatory = $false)]
    [string]$DateFormatPattern = "yyyy/MM/dd", # Default changed

    [Parameter(Mandatory = $false)]
    [string]$DefaultEncoding = "UTF8",

    [Parameter(Mandatory = $false)]
    [ValidateSet("InRange", "All", "None")] # Added "None"
    [string]$MarkwhenHolidayDisplayMode = "InRange",

    [Parameter(Mandatory = $false)]
    [object]$ExcelSheetNameOrIndex = 1, # Default changed, type changed to object to allow string or int

    [Parameter(Mandatory = $false)]
    [int]$ExcelStartRow = 5
)

# Add-Type -AssemblyName "Microsoft.Office.Interop.Excel" # この行をコメントアウトまたは削除

# --- グローバル変数・初期設定 ---
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:AllTaskNodesFlatList = @() # scriptスコープに変更
$script:Holidays = @()             # scriptスコープに変更
$script:DateFormatPattern = $DateFormatPattern # scriptスコープに変更 (関数内で $Global より安全)
$script:ExcelRowCounter = 0 # Excel書き込み用のスクリプトスコープ行カウンタ

# --- クラス定義 ---
class WbsElementNode {
    [string]$Id
    [string]$Name
    [int]$HierarchyLevel
    [string]$ElementType
    [System.Collections.Hashtable]$Attributes = @{}
    [string]$DescriptionText = ""
    [datetime]$Deadline
    [int]$DurationDays = 0
    [datetime]$CalculatedStartDate
    [datetime]$CalculatedEndDate
    [System.Collections.Generic.List[WbsElementNode]]$Children = [System.Collections.Generic.List[WbsElementNode]]::new()
    [WbsElementNode]$Parent
    [string[]]$RawAttributeLines = @()

    WbsElementNode([string]$Id, [string]$Name, [int]$Level) {
        $this.Id = $Id
        $this.Name = $Name.Trim()
        $this.HierarchyLevel = $Level
    }

    [void] AddRawAttributeLine([string]$Line) {
        $this.RawAttributeLines += $Line.TrimStart()
    }

    [void] ParseCollectedAttributes() {
        # Write-Verbose "Parsing attributes for $($this.Id) $($this.Name)" # 必要に応じてコメント解除
        foreach ($line in $this.RawAttributeLines) {
            # 許可するキーワードを正規表現に追加 (org)
            if ($line -match "^\s*-\s*(?<Key>deadline|duration|status|progress|org|assignee|depends):\s*(?<Value>.+)") {
                $key = $matches.Key.Trim().ToLower()
                $rawValue = $matches.Value.Trim()
                $value = ($rawValue -split '#')[0].Trim()

                $this.Attributes[$key] = $value
                # Write-Verbose "  Attribute found: $key = $value (Raw: $rawValue)" # 必要に応じてコメント解除
                switch ($key) {
                    "deadline" {
                        try { $this.Deadline = [datetime]::ParseExact($value, "yyyy-MM-dd", $null) }
                        catch { Write-Warning "要素 '$($this.Id) $($this.Name)' deadline形式無効: $value" } # $this.Name に修正
                    }
                    "duration" {
                        $valNum = $value -replace 'd',''
                        if ($valNum -match '^\d+$') { $this.DurationDays = [int]$valNum }
                        else { Write-Warning "要素 '$($this.Id) $($this.Name)' duration形式無効: $value"; $this.DurationDays = 0 } # $this.Name に修正
                    }
                    "depends" { $this.Attributes["depends"] = $value.TrimEnd('.') }
                    # org, assignee, status, progress は $this.Attributes に格納されるので、ここでは特別な処理は不要
                }
            }
        }
    }

    # Excel出力用に階層構造をフラット化するヘルパーメソッド
    [System.Collections.Generic.List[WbsElementNode]] GetAllDescendants() {
        $descendants = [System.Collections.Generic.List[WbsElementNode]]::new()
        foreach ($child in $this.Children) {
            $descendants.Add($child)
            $descendants.AddRange($child.GetAllDescendants())
        }
        return $descendants
    }
}

class ProjectMetadata { # (変更なし)
    [string]$Title; [string]$Description; [datetime]$DefinedDate
    [datetime]$ProjectPlanStartDate; [datetime]$ProjectPlanOverallDeadline; [string]$View = "month"
}

# --- 主要関数 ---
# 祝日リストの読み込み関数 (大きな変更なし)
function Import-HolidayList {
    param(
        [Parameter(Mandatory)] [string]$InputOfficialHolidayFilePath,
        [Parameter(Mandatory=$false)] [string]$InputCompanyHolidayFilePath,
        [Parameter(Mandatory)] [string]$BaseEncoding
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

# Markwhen形式データ生成関数 (大きな変更なし)
function Format-GanttMarkwhen {
    param(
        [Parameter(Mandatory)] [ProjectMetadata]$Metadata,
        [Parameter(Mandatory)] [System.Collections.Generic.List[WbsElementNode]]$RootWbsElements,
        [Parameter(Mandatory)] [string]$DatePattern,
        [Parameter(Mandatory)] [System.Collections.Generic.List[datetime]]$Holidays,
        [Parameter(Mandatory=$false)] [datetime]$ProjectOverallStartDateForHolidayFilter,
        [Parameter(Mandatory=$false)] [datetime]$ProjectOverallEndDateForHolidayFilter,
        [Parameter(Mandatory)] [string]$HolidayDisplayMode
    )
    Write-Verbose "Markwhen形式のガントチャートデータを生成 (新仕様)。祝日モード: $HolidayDisplayMode"
    $outputLines = New-Object System.Collections.Generic.List[string]

    $outputLines.Add("---")
    if ($Metadata -and (-not [string]::IsNullOrEmpty($Metadata.Title))) { $outputLines.Add("title: $($Metadata.Title)") }
    if ($Metadata -and (-not [string]::IsNullOrEmpty($Metadata.Description))) { $outputLines.Add("description: $($Metadata.Description)") }
    if ($Metadata -and $Metadata.DefinedDate -ne ([datetime]::MinValue)) { $outputLines.Add("date: $($Metadata.DefinedDate.ToString('yyyy-MM-dd'))") }
    if ($Metadata -and (-not [string]::IsNullOrEmpty($Metadata.View))) { $outputLines.Add("view: $($Metadata.View)") }
    else { $outputLines.Add("view: month") }
    $outputLines.Add("---")
    $outputLines.Add("")

    function ConvertTo-MarkwhenRecursiveUpdated {
        param (
            [WbsElementNode]$Element,
            [int]$CurrentMarkwhenIndentLevel,
            [System.Collections.Generic.List[string]]$OutputLinesRef,
            [string]$OutputDatePattern
        )
        $indent = "    " * $CurrentMarkwhenIndentLevel

        $isEffectivelyGroup = ($Element.ElementType -eq "Group") -or `
                              ($Element.ElementType -eq "Task" -and $Element.Children.Count -gt 0)

        if ($Element.ElementType -eq "Section") {
            $OutputLinesRef.Add("${indent}section $($Element.Name)")
            foreach ($child in $Element.Children) {
                ConvertTo-MarkwhenRecursiveUpdated -Element $child -CurrentMarkwhenIndentLevel $CurrentMarkwhenIndentLevel -OutputLinesRef $OutputLinesRef -OutputDatePattern $OutputDatePattern
            }
            $OutputLinesRef.Add("${indent}endsection")
        } elseif ($isEffectivelyGroup) {
            $OutputLinesRef.Add("${indent}group `"$($Element.Name)`"")
            if ($Element.ElementType -eq "Task" -and $Element.CalculatedStartDate -ne ([datetime]::MinValue) -and $Element.CalculatedEndDate -ne ([datetime]::MinValue) -and $Element.DurationDays -gt 0) {
                $startDateStr = $Element.CalculatedStartDate.ToString($OutputDatePattern)
                $endDateStr = $Element.CalculatedEndDate.ToString($OutputDatePattern)
                $taskDisplayName = "概要"
                $tags = New-Object System.Collections.Generic.List[string]
                $tags.Add("#$($Element.Id.Replace('.','_'))")
                if ($Element.Attributes.ContainsKey("status")) { $tags.Add("#$($Element.Attributes['status'])") }
                if ($Element.Attributes.ContainsKey("progress")) { $taskDisplayName += " ($($Element.Attributes['progress'] -replace '%',''))%" }

                $groupTaskLine = "$($indent)    ${startDateStr}/${endDateStr}: ${taskDisplayName}"
                if ($tags.Count -gt 0) { $groupTaskLine += " $([string]::Join(' ', $tags))" }
                $OutputLinesRef.Add($groupTaskLine)
            }
            foreach ($child in $Element.Children) {
                ConvertTo-MarkwhenRecursiveUpdated -Element $child -CurrentMarkwhenIndentLevel ($CurrentMarkwhenIndentLevel + 1) -OutputLinesRef $OutputLinesRef -OutputDatePattern $OutputDatePattern
            }
            $OutputLinesRef.Add("${indent}endgroup")
        } elseif ($Element.ElementType -eq "Task") {
            if ($Element.CalculatedStartDate -ne ([datetime]::MinValue) -and $Element.CalculatedEndDate -ne ([datetime]::MinValue) -and $Element.DurationDays -gt 0) {
                $currentIndent = "    " * $CurrentMarkwhenIndentLevel
                $currentSDateStr = $Element.CalculatedStartDate.ToString($OutputDatePattern)
                $currentEDateStr = $Element.CalculatedEndDate.ToString($OutputDatePattern)
                $currentTaskName = $Element.Name
                if ($Element.Attributes.ContainsKey("progress") -and $Element.Attributes["progress"] -match "^\d{1,3}\%?$") {
                    $currentTaskName += " ($($Element.Attributes['progress'] -replace '%',''))%"
                }
                $currentTags = New-Object System.Collections.Generic.List[string]
                $currentTags.Add("#$($Element.Id.Replace('.','_'))")
                if ($Element.Attributes.ContainsKey("status")) { $currentTags.Add("#$($Element.Attributes['status'])") }
                if ($Element.Attributes.ContainsKey("assignee")) { $currentTags.Add("#assignee-$($Element.Attributes['assignee'] -replace ' ','_')") }
                 if ($Element.Attributes.ContainsKey("org")) { $currentTags.Add("#org-$($Element.Attributes['org'] -replace ' ','_')") } # Add org tag
                $tagsString = ""
                if ($currentTags.Count -gt 0) { $tagsString = " $([string]::Join(' ', $currentTags))" }
                $finalTaskLine = "${currentIndent}${currentSDateStr}/${currentEDateStr}: ${currentTaskName}${tagsString}"
                $OutputLinesRef.Add($finalTaskLine)
                if (-not [string]::IsNullOrEmpty($Element.DescriptionText)) {
                    foreach($descLine in ($Element.DescriptionText.TrimEnd("`n") -split "`n")) {
                        $OutputLinesRef.Add("$currentIndent    // $descLine")
                    }
                }
            } else { Write-Warning "タスク '$($Element.Id) $($Element.Name)' の日付/期間未設定、または0のため出力スキップ。" }
        }
    }

    foreach ($rootElement in $RootWbsElements) {
        ConvertTo-MarkwhenRecursiveUpdated -Element $rootElement -CurrentMarkwhenIndentLevel 0 -OutputLinesRef $outputLines -OutputDatePattern $DatePattern
        $outputLines.Add("")
    }

    if ($HolidayDisplayMode -ne "None" -and $Holidays -and $Holidays.Count -gt 0) {
        $holidaysToDisplay = @()
        if ($HolidayDisplayMode -eq "All") { $holidaysToDisplay = $Holidays | Sort-Object }
        elseif ($HolidayDisplayMode -eq "InRange" -and $ProjectOverallStartDateForHolidayFilter -ne $null -and $ProjectOverallEndDateForHolidayFilter -ne $null) {
            $holidaysToDisplay = $Holidays | Where-Object { $_.Date -ge $ProjectOverallStartDateForHolidayFilter.Date -and $_.Date -le $ProjectOverallEndDateForHolidayFilter.Date } | Sort-Object
        }
        if ($holidaysToDisplay.Count -gt 0) {
            $outputLines.Add("section 祝日")
            $outputLines.Add("group Holidays #project_holidays")
            foreach($holiday in $holidaysToDisplay) { $outputLines.Add("    $($holiday.ToString($DatePattern)): Holiday") }
            $outputLines.Add("endgroup")
            $outputLines.Add("endsection")
            $outputLines.Add("")
        }
    }
    Write-Verbose "Markwhenデータ生成完了 (新仕様)。"
    $result = [string]($outputLines -join [System.Environment]::NewLine)
    return $result
}

# データ型変換用のヘルパー関数 (現在はExcel COM関数内で直接処理しているため未使用)
function ConvertTo-ExcelValue {
    param (
        [Parameter(Mandatory=$true)]
        $Value,
        [Parameter(Mandatory=$false)]
        [string]$DateFormat = "yyyy/MM/dd"
    )

    try {
        if ($null -eq $Value) {
            return ""
        }

        # 数値型の場合
        if ($Value -is [double] -or $Value -is [int]) {
            return $Value.ToString()
        }

        # 日付型の場合
        if ($Value -is [DateTime]) {
            return $Value.ToString($DateFormat)
        }

        # 文字列の場合
        if ($Value -is [string]) {
            return $Value
        }

        # その他の型は文字列に変換
        return $Value.ToString()
    }
    catch {
        Write-Warning "値の変換中にエラー: $Value (型: $($Value.GetType().Name)) - $_"
        return ""
    }
}

# Excelへの書き込み関数 (変更なし)
function Write-WbsToExcelCom {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [ProjectMetadata]$ProjectInfo,
        [Parameter(Mandatory)] [System.Collections.Generic.List[WbsElementNode]]$RootWbsElements, # ツリー構造のルート
        [Parameter(Mandatory)] [System.Collections.Generic.List[WbsElementNode]]$AllTaskNodesFlatListForDateLookup, # 先行タスク名検索用 (全ノードリスト)
        [Parameter(Mandatory)] [string]$TargetExcelFilePath,
        [Parameter(Mandatory)] [object]$SheetIdentifier,
        [Parameter(Mandatory)] [int]$StartRow,
        [Parameter(Mandatory)] [string]$ExcelDateFormat
    )

    Write-Verbose "=== Excel COMオブジェクト初期化開始 ==="
    $excelCom = $null; $workbook = $null; $worksheet = $null
    try {
        # Excelアプリケーションの初期化
        $excelCom = New-Object -ComObject Excel.Application
        $excelCom.Visible = $false
        $excelCom.DisplayAlerts = $false
        $excelCom.EnableEvents = $false
        $excelCom.ScreenUpdating = $false

        # ワークブックの取得
        if (-not (Test-Path $TargetExcelFilePath -PathType Leaf)) {
            throw "Excelファイル '$TargetExcelFilePath' が見つかりません。"
        }
        $workbook = $excelCom.Workbooks.Open($TargetExcelFilePath)
        if ($null -eq $workbook) {
            throw "ワークブック '$TargetExcelFilePath' のオープンに失敗しました。"
        }

        # シートの取得
        try {
            $worksheet = $workbook.Worksheets.Item($SheetIdentifier)
            if ($null -eq $worksheet) {
                throw "シート '$SheetIdentifier' の取得に失敗しました。"
            }
        } catch {
            throw "シート '$SheetIdentifier' が見つかりません: $($_.Exception.Message)"
        }

        Write-Verbose "シート '$SheetIdentifier' を取得しました。"

        # プロジェクト情報の書き込み (D1: プロジェクト名, O1: プロジェクト開始日)
        if ($ProjectInfo -and (-not [string]::IsNullOrEmpty($ProjectInfo.Title))) {
            $worksheet.Cells.Item(1, 4).Value2 = $ProjectInfo.Title
            Write-Verbose "D1にプロジェクト名 '$($ProjectInfo.Title)' を書き込みました。"
        }
        # O1にプロジェクト全体の開始日を書き込み (計算されたタスク開始日の最小値)
        $overallProjectStartDate = $null
        if ($AllTaskNodesFlatListForDateLookup.Count -gt 0) {
            # CalculatedStartDateが設定されている要素のみを対象
            $validTasksForStartDate = $AllTaskNodesFlatListForDateLookup | Where-Object {$_.CalculatedStartDate -ne ([datetime]::MinValue)}
            if ($validTasksForStartDate.Count -gt 0) {
                $overallProjectStartDate = ($validTasksForStartDate | Sort-Object CalculatedStartDate | Select-Object -First 1).CalculatedStartDate
            }
        }
        if ($overallProjectStartDate -ne $null -and $overallProjectStartDate -ne ([datetime]::MinValue)) {
             $worksheet.Cells.Item(1, 15).Value2 = $overallProjectStartDate.ToString($ExcelDateFormat)
             Write-Verbose "O1にプロジェクト開始日 '$($overallProjectStartDate.ToString($ExcelDateFormat))' を書き込みました。"
        } else {
             Write-Verbose "O1に書き込むプロジェクト開始日が見つかりませんでした。"
        }


        # データ書き込み開始
        Write-Verbose "=== データ書き込み開始 (開始行: $StartRow) ==="
        $script:ExcelRowCounter = $StartRow # スクリプトスコープの行カウンタを初期化

        function Add-RowToExcelRecursiveDetailed {
            param (
                [WbsElementNode]$Element,
                $CurrentWorksheet,
                [string]$OutputDateFormat,
                [System.Collections.Generic.List[WbsElementNode]]$AllTasksForNameLookup
            )

            Write-Verbose "Excel行 $($script:ExcelRowCounter): ID '$($Element.Id)' 名前 '$($Element.Name)' を書き込み中..."

            # A列(1): No
            $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 1).Value2 = [string]"$($Element.Id)" # Excelで数値として扱われないようにシングルクォートを付加

            # B, C, D列: 要素名 (階層レベルに応じて列を分ける)
            if ($Element.HierarchyLevel -eq 2) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 2).Value2 = $Element.Name }
            elseif ($Element.HierarchyLevel -eq 3) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 3).Value2 = $Element.Name }
            elseif ($Element.HierarchyLevel -ge 4) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 4).Value2 = $Element.Name }

            # E列(5): 先行後続, F列(6): 関連番号(先行ID), G列(7): 先行タスク名
            if ($Element.Attributes.ContainsKey("depends") -and -not [string]::IsNullOrEmpty($Element.Attributes.depends)) {
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 5).Value2 = "先行"
                $dependsOnId = $Element.Attributes.depends.Trim() # depends属性はカンマ区切りだが、ここでは最初のIDのみを想定
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 6).Value2 = $dependsOnId
                # 全ノードリストから先行タスクを検索して名前を取得
                $parentTaskObject = $AllTasksForNameLookup | Where-Object {$_.Id -eq $dependsOnId} | Select-Object -First 1
                if ($parentTaskObject) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 7).Value2 = $parentTaskObject.Name }
                else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 7).Value2 = "" }
            } else {
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 5).Value2 = ""
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 6).Value2 = ""
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 7).Value2 = ""
            }

            # H列は書き込まない (仕様変更により削除)

            # I列(9): アクションプラン (DescriptionText)
            $description = if (-not [string]::IsNullOrEmpty($Element.DescriptionText)) { $Element.DescriptionText.Trim() -replace "`r?`n", [char]10 } else { "" }
            $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 9).Value2 = $description
            if (-not [string]::IsNullOrEmpty($description)) {$CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 9).WrapText = $true}

            # N列(14): 担当組織 (org属性)
            if ($Element.Attributes.ContainsKey("org")) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 14).Value2 = $Element.Attributes.org.Trim() }
            else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 14).Value2 = "" }

            # O列(15): 担当者 (assignee属性)
            if ($Element.Attributes.ContainsKey("assignee")) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 15).Value2 = $Element.Attributes.assignee.Trim() }
            else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 15).Value2 = "" }

            # R列(18): 開始入力 (CalculatedStartDate)
            if ($Element.CalculatedStartDate -ne ([datetime]::MinValue)) {
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 18).Value2 = $Element.CalculatedStartDate.ToString($OutputDateFormat)
            }
            else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 18).Value2 = "" }

            # S列(19): 終了入力 (Deadline)
            if ($Element.Deadline -ne ([datetime]::MinValue)) {
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 19).Value2 = $Element.Deadline.ToString($OutputDateFormat)
            }
            else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 19).Value2 = "" }

            # T列(20): 日数入力 (DurationDays)
            if ($Element.DurationDays -gt 0) { 
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 20).Value2 = [string]$Element.DurationDays 
            }

            # X列(24): 進捗率 (progress属性)
            $progressToSet = "";
            if ($Element.Attributes.ContainsKey("progress")) {
                $pStr = $Element.Attributes.progress.Trim() -replace '%';
                if ($pStr -match '^\d+$') {
                    $progressValue = [double]$pStr;
                    if ($progressValue -ge 0 -and $progressValue -le 100) {
                        $progressToSet = [string]($progressValue / 100);
                    }
                }
            }
            $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 24).Value2 = $progressToSet

            # Y列(25): 開始実績, Z列(26): 終了実績 (status属性に基づく)
            $actualStartDateToSet = ""; $actualEndDateToSet = ""
            if ($Element.Attributes.ContainsKey("status")) {
                $status = $Element.Attributes.status.Trim().ToLower()
                # 開始実績: inprogress または completed で、かつ CalculatedStartDate があれば設定
                if (($status -eq "inprogress" -or $status -eq "completed") -and $Element.CalculatedStartDate -ne ([datetime]::MinValue)) {
                    $actualStartDateToSet = $Element.CalculatedStartDate.ToString($OutputDateFormat)
                }
                # 終了実績: completed で、かつ CalculatedEndDate があれば設定
                if ($status -eq "completed" -and $Element.CalculatedEndDate -ne ([datetime]::MinValue)) {
                    $actualEndDateToSet = $Element.CalculatedEndDate.ToString($OutputDateFormat);
                    # completed なら進捗率を100%にする (既に100%でなければ)
                    if ($CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 24).Value2 -ne 1) {
                         $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 24).Value2 = 1
                    }
                }
            }
            $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 25).Value2 = $actualStartDateToSet
            $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 26).Value2 = $actualEndDateToSet

            $script:ExcelRowCounter++ # 行カウンタをインクリメント

            # 子要素を再帰的に処理
            foreach ($child in $Element.Children) {
                Add-RowToExcelRecursiveDetailed -Element $child `
                                                -CurrentWorksheet $CurrentWorksheet `
                                                -OutputDateFormat $OutputDateFormat `
                                                -AllTasksForNameLookup $AllTasksForNameLookup
            }
        }

        # ルート要素から再帰処理を開始
        foreach ($rootElement in $RootWbsElements) {
            Add-RowToExcelRecursiveDetailed -Element $rootElement `
                                            -CurrentWorksheet $worksheet `
                                            -OutputDateFormat $ExcelDateFormat `
                                            -AllTasksForNameLookup $AllTaskNodesFlatListForDateLookup # 全ノードのリストを渡す
        }

        $workbook.Save()
        Write-Verbose "Excelファイルにデータを書き込み、保存しました。"

    } catch {
        Write-Error "Excel操作中にエラーが発生しました: $($_.Exception.Message)"
        throw
    } finally {
        if ($workbook) {
            $workbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($worksheet) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        }
        if ($excelCom) {
            $excelCom.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelCom) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# MD-WBSファイル解析関数 (YAMLパースとコメント行スキップ修正)
function Parse-WbsMarkdownAndMetadataAdvanced {
    param(
        [Parameter(Mandatory)] [string]$FilePath,
        [Parameter(Mandatory)] [string]$Encoding
    )
    Write-Verbose "拡張MD-WBSファイル(新仕様)の解析開始: $FilePath"
    $projectMetadata = New-Object ProjectMetadata
    $rootElements = [System.Collections.Generic.List[WbsElementNode]]::new()
    if (-not (Test-Path $FilePath -PathType Leaf)) { Write-Error "WBSファイル '$FilePath' が見つかりません。"; return $null }

    $allLines = Get-Content -Path $FilePath -Encoding $Encoding
    $lineIndex = 0

    # 1. YAMLフロントマター解析
    if ($allLines.Count -gt 0 -and $allLines[$lineIndex].Trim() -eq "---") {
        $lineIndex = 1 # --- の次の行から開始
        $yamlLines = New-Object System.Collections.Generic.List[string]
        while ($lineIndex -lt $allLines.Count -and $allLines[$lineIndex].Trim() -ne "---") {
            $yamlLines.Add($allLines[$lineIndex])
            $lineIndex++
        }
        if ($lineIndex -lt $allLines.Count -and $allLines[$lineIndex].Trim() -eq "---") { # 閉じの --- があるか
            $lineIndex++ # --- の次の行からMarkdown本文のパースを開始

            $convertFromYamlAvailable = $false
            try { if (Get-Command ConvertFrom-Yaml -ErrorAction SilentlyContinue) { $convertFromYamlAvailable = $true } } catch {}

            if ($convertFromYamlAvailable) {
                Write-Verbose "YAMLフロントマターを ConvertFrom-Yaml で解析中..."
                try {
                    $parsedYaml = ($yamlLines -join "`n") | ConvertFrom-Yaml
                    if ($parsedYaml) {
                        if ($parsedYaml.PSObject.Properties.Name -contains "title") { $projectMetadata.Title = $parsedYaml.title }
                        if ($parsedYaml.PSObject.Properties.Name -contains "description") { $projectMetadata.Description = $parsedYaml.description }
                        if ($parsedYaml.PSObject.Properties.Name -contains "date") { try { $projectMetadata.DefinedDate = [datetime]$parsedYaml.date } catch {} }
                        if ($parsedYaml.PSObject.Properties.Name -contains "projectstartdate") { try { $projectMetadata.ProjectPlanStartDate = [datetime]$parsedYaml.projectstartdate } catch {} }
                        if ($parsedYaml.PSObject.Properties.Name -contains "projectoveralldeadline") { try { $projectMetadata.ProjectPlanOverallDeadline = [datetime]$parsedYaml.projectoveralldeadline } catch {} }
                        if ($parsedYaml.PSObject.Properties.Name -contains "view") { $projectMetadata.View = $parsedYaml.view }
                    }
                } catch { Write-Warning "YAMLフロントマター解析失敗 (ConvertFrom-Yaml): $($_.Exception.Message)" }
            } else {
                Write-Warning "ConvertFrom-Yaml が利用できません。YAMLフロントマターは簡易キーバリュー形式で解析します。"
                foreach ($yamlLine in $yamlLines) {
                    if ($yamlLine -match "^\s*(?<Key>[^:]+):\s*(?<Value>.+)") {
                        $key = $matches.Key.Trim().ToLower(); $value = $matches.Value.Trim()
                        switch ($key) {
                            "title"       { $projectMetadata.Title = $value }
                            "description" { $projectMetadata.Description = $value }
                            "date"        { try { $projectMetadata.DefinedDate = [datetime]$value } catch {} }
                            "projectstartdate" { try { $projectMetadata.ProjectPlanStartDate = [datetime]$value } catch {} }
                            "projectoveralldeadline" { try { $projectMetadata.ProjectPlanOverallDeadline = [datetime]$value } catch {} }
                            "view"        { $projectMetadata.View = $value }
                        }
                    }
                }
            }
            Write-Verbose "ProjectMetadata Title from YAML: $($projectMetadata.Title)"
        } else {
            Write-Warning "YAMLフロントマターが正しく閉じられていません。最初の'---'以降をMarkdown本文として処理します。"
            $lineIndex = 0 # ファイルの先頭からMarkdown本文としてパースし直す (YAMLなしとみなす)
        }
    } # YAMLフロントマター処理の終わり

    Write-Verbose "Markdown本文の解析中 (開始行: $($lineIndex + 1))..."
    $parentStack = New-Object System.Collections.Stack; $currentWbsElement = $null
    $attributeLinesBuffer = New-Object System.Collections.Generic.List[string]
    $descriptionBuffer = New-Object System.Text.StringBuilder

    # 現在処理中の要素の属性と説明を確定させるヘルパー関数
    function Process-BufferedAttributesAndDescription {
        param ([WbsElementNode]$Element,
               [System.Collections.Generic.List[string]]$AttribBuffer,
               [System.Text.StringBuilder]$DescBuffer)
        if ($Element) {
            if ($AttribBuffer.Count -gt 0) {
                $Element.RawAttributeLines = $AttribBuffer.ToArray()
                $Element.ParseCollectedAttributes() # ここで実際にパース
                $AttribBuffer.Clear()
            }
            if ($DescBuffer.Length -gt 0) {
                $Element.DescriptionText = $DescBuffer.ToString().Trim()
                [void]$DescBuffer.Clear()
            }
        }
    }

    while ($lineIndex -lt $allLines.Count) {
        $line = $allLines[$lineIndex]; $trimmedLineForCheck = $line.Trim() # 判定用にトリム

        if ([string]::IsNullOrWhiteSpace($trimmedLineForCheck)) {
            # 空行の場合、現在処理中の要素の属性と説明を確定させる
            Process-BufferedAttributesAndDescription -Element $currentWbsElement -AttribBuffer $attributeLinesBuffer -DescBuffer $descriptionBuffer
            $lineIndex++; continue
        }

        # コメント行のスキップ (#で始まるが##ではない行)
        if ($trimmedLineForCheck.StartsWith("#") -and -not ($trimmedLineForCheck -match "^#+\s+")) {
            Write-Verbose "コメント行をスキップ: $line"
            $lineIndex++; continue
        }

        # 見出し行か判定
        if ($trimmedLineForCheck -match "^(#+)\s+([\d\.]*[\d])\.?\s*(.+)") {
            # 前の要素の属性と説明をここで処理
            Process-BufferedAttributesAndDescription -Element $currentWbsElement -AttribBuffer $attributeLinesBuffer -DescBuffer $descriptionBuffer
            $currentWbsElement = $null # 新しい見出しなのでリセット (属性と説明は新しい要素に紐づく)

            $level = $matches[1].Length; $id = $matches[2]; $name = $matches[3].Trim()
            $newElement = New-Object WbsElementNode -ArgumentList $id, $name, $level
            Write-Verbose "見出しを検出: Level=$level, ID=$id, Name=$name"

            if ($level -eq 2) { $newElement.ElementType = "Section" }
            elseif ($level -eq 3) { $newElement.ElementType = "Group" }
            elseif ($level -ge 4) { $newElement.ElementType = "Task" }

            while ($parentStack.Count -gt 0 -and ($parentStack.Peek()).HierarchyLevel -ge $level) { [void]$parentStack.Pop() }
            if ($parentStack.Count -gt 0) { ($parentStack.Peek()).Children.Add($newElement); $newElement.Parent = $parentStack.Peek() }
            else { $rootElements.Add($newElement) }
            $parentStack.Push($newElement); $currentWbsElement = $newElement
        }
        # 属性リスト行か判定 (親見出しより1段インデント: スペース4つ)
        elseif ($currentWbsElement -and $line -match "^\s{4}-\s*(?:deadline|duration|status|progress|org|assignee|depends):\s*.+") {
            if ($descriptionBuffer.Length -gt 0) { # 属性の前に本文があった場合は先に確定
                 $currentWbsElement.DescriptionText = $descriptionBuffer.ToString().Trim(); [void]$descriptionBuffer.Clear()
            }
            $attributeLinesBuffer.Add($line) # 元の行(インデント含む)をバッファに追加、ParseCollectedAttributesでTrimStart
            Write-Verbose "属性行を検出: $line" # デバッグ出力
        }
        # 見出し直後の本文か判定 (属性リスト行でも見出し行でもない行)
        elseif ($currentWbsElement -and ($attributeLinesBuffer.Count -eq 0) -and (-not ($line.TrimStart() -match "^\s*-\s*")) -and (-not ($trimmedLineForCheck -match "^#+"))) {
             [void]$descriptionBuffer.AppendLine($line.Trim()) # 行頭のインデントは除去して本文として収集
        }
        # その他の行 (未分類または要素の区切りとみなし、前の要素の属性/説明を確定)
        else {
            Write-Verbose "未分類の行 (または前の要素の終端とみなし属性処理): $line"
            Process-BufferedAttributesAndDescription -Element $currentWbsElement -AttribBuffer $attributeLinesBuffer -DescBuffer $descriptionBuffer
        }
        $lineIndex++
    }
    # ループ終了後、最後の要素の属性と説明を処理
    Process-BufferedAttributesAndDescription -Element $currentWbsElement -AttribBuffer $attributeLinesBuffer -DescBuffer $descriptionBuffer

    Write-Verbose "Markdown本文解析完了。トップレベル要素数: $($rootElements.Count)"
    return @{ ProjectMetadata = $projectMetadata; RootElements = $rootElements }
}

# タスク開始日計算関数 (大きな変更なし)
function Calculate-TaskStartDate {
    param( [Parameter(Mandatory)] [datetime]$DeadlineDate, [Parameter(Mandatory)] [int]$DurationBusinessDays, [Parameter(Mandatory)] [System.Collections.Generic.List[datetime]]$Holidays )
    if ($DurationBusinessDays -le 0) { return $DeadlineDate }
    $calculatedStartDate = $DeadlineDate; $daysToCount = $DurationBusinessDays
    while ($daysToCount > 0) {
        $calculatedStartDate = $calculatedStartDate.AddDays(-1)
        if ($calculatedStartDate.DayOfWeek -ne [DayOfWeek]::Saturday -and $calculatedStartDate.DayOfWeek -ne [DayOfWeek]::Sunday) {
            if (-not $Holidays.Contains($calculatedStartDate.Date)) { $daysToCount-- }
        }
        if ($calculatedStartDate -lt $DeadlineDate.AddYears(-10)) { Write-Warning "開始日計算異常: 10年以上遡りました。"; return $DeadlineDate }
    }
    return $calculatedStartDate
}

# タスク終了日計算関数 (大きな変更なし)
function Calculate-TaskEndDate {
    param( [Parameter(Mandatory)] [datetime]$StartDate, [Parameter(Mandatory)] [int]$DurationBusinessDays, [Parameter(Mandatory)] [System.Collections.Generic.List[datetime]]$Holidays )
    if ($DurationBusinessDays -le 0) { return $StartDate }

    $calculatedEndDate = $StartDate
    $businessDaysCounted = 0

    # Start date counts as the first business day if it's not a weekend/holiday
    if ($StartDate.DayOfWeek -ne [DayOfWeek]::Saturday -and $StartDate.DayOfWeek -ne [DayOfWeek]::Sunday -and (-not $Holidays.Contains($StartDate.Date))) {
        $businessDaysCounted = 1
    }

    # If duration is 1 and start date is a business day, end date is the same as start date
    if ($DurationBusinessDays -eq 1 -and $businessDaysCounted -eq 1) {
        return $StartDate
    }
    # If duration is 1 but start date is NOT a business day (e.g., holiday), we need to find the next business day
    if ($DurationBusinessDays -eq 1 -and $businessDaysCounted -eq 0) {
         $calculatedEndDate = $StartDate.AddDays(1)
         while ($calculatedEndDate.DayOfWeek -eq [DayOfWeek]::Saturday -or $calculatedEndDate.DayOfWeek -eq [DayOfWeek]::Sunday -or $Holidays.Contains($calculatedEndDate.Date)) {
             $calculatedEndDate = $calculatedEndDate.AddDays(1)
         }
         return $calculatedEndDate
    }


    while ($businessDaysCounted -lt $DurationBusinessDays) {
        $calculatedEndDate = $calculatedEndDate.AddDays(1)
        if ($calculatedEndDate.DayOfWeek -ne [DayOfWeek]::Saturday -and $calculatedEndDate.DayOfWeek -ne [DayOfWeek]::Sunday) {
            if (-not $Holidays.Contains($calculatedEndDate.Date)) {
                $businessDaysCounted++
            }
        }
        if ($calculatedEndDate -gt $StartDate.AddYears(2)) { Write-Warning "終了日計算異常: 2年以上先に及びました。"; return $StartDate.AddYears(2) }
    }
    return $calculatedEndDate
}

# Excelプロセスの終了を確実に行う関数
function Stop-ExcelProcesses {
    $excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        Write-Verbose "既存のExcelプロセスを終了中..."
        $excelProcesses | ForEach-Object {
            try {
                $_.CloseMainWindow()
                Start-Sleep -Seconds 1
                if (-not $_.HasExited) {
                    $_.Kill()
                }
            }
            catch {
                Write-Warning "Excelプロセスの終了中にエラー: $_"
            }
        }
        Start-Sleep -Seconds 2
    }
}

# 依存関係チェック関数 (大きな変更なし)
function Validate-TaskDependencies {
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.Generic.List[WbsElementNode]]$TaskNodes # 全ノードリストを想定
    )

    $warnings = [System.Collections.Generic.List[string]]::new()
    # Create a quick lookup map for tasks by ID from the full list
    $taskLookup = @{}
    $TaskNodes | ForEach-Object {$taskLookup[$_.Id] = $_}

    foreach ($task in $TaskNodes) {
        # 依存関係チェックはタスク要素のみを対象とする
        if ($task.ElementType -eq "Task" -and $task.Attributes.ContainsKey("depends")) {
            $dependsIds = $task.Attributes["depends"] -split ',' | ForEach-Object { $_.Trim() }
            foreach ($dependsId in $dependsIds) {
                if ([string]::IsNullOrWhiteSpace($dependsId)) { continue }

                $parentTask = $taskLookup[$dependsId]

                if ($parentTask) {
                    # Check for date conflicts only if both tasks have calculated dates
                    if ($parentTask.CalculatedEndDate -ne ([datetime]::MinValue) -and
                        $task.CalculatedStartDate -ne ([datetime]::MinValue) -and
                        $parentTask.CalculatedEndDate -ge $task.CalculatedStartDate) {

                        $warningMsg = "依存関係警告: タスク '$($task.Id) $($task.Name)' " +
                                    "は先行タスク '$($parentTask.Id) $($parentTask.Name)' " +
                                    "($($parentTask.CalculatedEndDate.ToString($script:DateFormatPattern))) " +
                                    "の終了日より前に開始 ($($task.CalculatedStartDate.ToString($script:DateFormatPattern))) しています。"
                        $warnings.Add($warningMsg) # $dependencyWarnings -> $warnings に修正
                        Write-Warning $warningMsg
                    }
                } else {
                    $errorMsg = "依存関係エラー: タスク '$($task.Id) $($task.Name)' " +
                               "が依存するID '$dependsId' が見つかりません。"
                    Write-Warning $errorMsg
                }
            }
        }
    }
    return $warnings
}

# --- メイン処理 ---
try {
    Write-Host "処理を開始します (MD-WBS仕様)..." -ForegroundColor Green

    # スクリプト開始時に既存のExcelプロセスを確認し、終了を試みる
    Stop-ExcelProcesses # 関数呼び出しに変更

    $script:Holidays = Import-HolidayList -InputOfficialHolidayFilePath $OfficialHolidayFilePath -InputCompanyHolidayFilePath $CompanyHolidayFilePath -BaseEncoding $DefaultEncoding
    $parsedResult = Parse-WbsMarkdownAndMetadataAdvanced -FilePath $WbsFilePath -Encoding $DefaultEncoding
    if (-not $parsedResult) { Write-Error "MD-WBSファイル解析失敗。"; exit 1 }
    $projectMetadata = $parsedResult.ProjectMetadata
    $rootWbsElements = $parsedResult.RootElements

    if ($rootWbsElements.Count -eq 0 -and (-not $projectMetadata -or ([string]::IsNullOrEmpty($projectMetadata.Title)) )) { Write-Error "MD-WBS有効データ無。"; exit 1 }

    # Excel出力のためには、日付計算対象だけでなく、すべてのWBS要素のリストが必要になる場合がある（先行タスク名検索など）
    # ここでは、$script:AllTaskNodesFlatList を日付計算用にフィルタリングする前の全要素リストとして扱う。
    $script:AllTaskNodesFlatList = New-Object System.Collections.Generic.List[WbsElementNode]
    function Flatten-WbsTreeNodesForAll { # 全ノードをフラット化する関数
        param ([System.Collections.Generic.List[WbsElementNode]]$Elements)
        if ($null -eq $Elements) { return }
        foreach ($element in $Elements) {
            $script:AllTaskNodesFlatList.Add($element)
            if ($element.Children.Count -gt 0) { Flatten-WbsTreeNodesForAll -Elements $element.Children }
        }
    }
    Flatten-WbsTreeNodesForAll -Elements $rootWbsElements # 全ノードをフラット化

    Write-Verbose "各タスクの日付を算出中 (期間と期限がある要素のみ)..."
    # Calculate dates only for nodes that are Tasks and have both duration and deadline
    if ($null -ne $script:AllTaskNodesFlatList) {
        foreach ($taskNode in ($script:AllTaskNodesFlatList | Where-Object {$_.ElementType -eq "Task" -and $_.DurationDays -gt 0 -and $_.Deadline -ne ([datetime]::MinValue)})) {
            $taskNode.CalculatedStartDate = Calculate-TaskStartDate `
                -DeadlineDate $taskNode.Deadline `
                -DurationBusinessDays $taskNode.DurationDays `
                -Holidays $script:Holidays
            $taskNode.CalculatedEndDate = Calculate-TaskEndDate `
                -StartDate $taskNode.CalculatedStartDate `
                -DurationBusinessDays $taskNode.DurationDays `
                -Holidays $script:Holidays
            Write-Verbose "タスク '$($taskNode.Id) $($taskNode.Name)' 開始: $($taskNode.CalculatedStartDate.ToString($script:DateFormatPattern)), 終了: $($taskNode.CalculatedEndDate.ToString($script:DateFormatPattern)) (期間: $($taskNode.DurationDays)d)"
        }
    }


    # Determine overall project date range for Markwhen holiday filtering and Excel O1
    # Use calculated dates from all relevant nodes in the flat list
    $projectOverallStartDate = $projectMetadata.ProjectPlanStartDate # YAMLから取得した開始日を優先
    $projectOverallEndDate = $projectMetadata.ProjectPlanOverallDeadline # YAMLから取得した終了日を優先

    if (($null -ne $script:AllTaskNodesFlatList) -and $script:AllTaskNodesFlatList.Count -gt 0) {
        # Find tasks with *any* date information (calculated start/end or deadline)
        $validNodesForDateRange = @($script:AllTaskNodesFlatList | Where-Object {$_.CalculatedStartDate -ne ([datetime]::MinValue) -or $_.CalculatedEndDate -ne ([datetime]::MinValue) -or $_.Deadline -ne ([datetime]::MinValue)})
        if ($validNodesForDateRange.Count -gt 0) {
            # If project start date is not defined in YAML, find the earliest calculated start/deadline among nodes
            if (($projectOverallStartDate -eq $null -or $projectOverallStartDate -eq ([datetime]::MinValue))) {
                 $earliestNodeDate = $validNodesForDateRange | Sort-Object @{Expression={$_.CalculatedStartDate}; Descending=$false}, @{Expression={$_.Deadline}; Descending=$false} | Where-Object {$_.CalculatedStartDate -ne ([datetime]::MinValue) -or $_.Deadline -ne ([datetime]::MinValue)} | Select-Object -First 1
                 if ($earliestNodeDate) {
                     $projectOverallStartDate = if ($earliestNodeDate.CalculatedStartDate -ne ([datetime]::MinValue)) { $earliestNodeDate.CalculatedStartDate } else { $earliestNodeDate.Deadline }
                 }
            }
            # If project end date is not defined in YAML, find the latest calculated end/deadline among nodes
            if (($projectOverallEndDate -eq $null -or $projectOverallEndDate -eq ([datetime]::MinValue))) {
                $latestNodeDate = $validNodesForDateRange | Sort-Object @{Expression={$_.CalculatedEndDate}; Descending=$true}, @{Expression={$_.Deadline}; Descending=$true} | Where-Object {$_.CalculatedEndDate -ne ([datetime]::MinValue) -or $_.Deadline -ne ([datetime]::MinValue)} | Select-Object -First 1
                 if ($latestNodeDate) {
                     $projectOverallEndDate = if ($latestNodeDate.CalculatedEndDate -ne ([datetime]::MinValue)) { $latestNodeDate.CalculatedEndDate } else { $latestNodeDate.Deadline }
                 }
            }
        }
    }

    # Use the determined overall project dates for Markwhen holiday filtering
    $projectOverallStartDateForHolidayFilter = $projectOverallStartDate
    $projectOverallEndDateForHolidayFilter = $projectOverallEndDate


    if($projectOverallStartDate -and $projectOverallEndDate -and $projectOverallStartDate -ne ([datetime]::MinValue) -and $projectOverallEndDate -ne ([datetime]::MinValue)) { Write-Verbose "プロジェクト期間: $($projectOverallStartDate.ToString($script:DateFormatPattern)) - $($projectOverallEndDate.ToString($script:DateFormatPattern))" }
    else { Write-Verbose "プロジェクト期間を特定できませんでした。" }


    # Dependency check (using the full flat list for lookup)
    $dependencyWarnings = Validate-TaskDependencies -TaskNodes $script:AllTaskNodesFlatList


    if ($OutputFormat -eq "Markwhen") {
        $outputContent = Format-GanttMarkwhen -Metadata $projectMetadata `
                                            -RootWbsElements $rootWbsElements `
                                            -DatePattern $script:DateFormatPattern `
                                            -Holidays $script:Holidays `
                                            -ProjectOverallStartDateForHolidayFilter $projectOverallStartDateForHolidayFilter `
                                            -ProjectOverallEndDateForHolidayFilter $projectOverallEndDateForHolidayFilter `
                                            -HolidayDisplayMode $MarkwhenHolidayDisplayMode
        try {
            Set-Content -Path $OutputFilePath -Value $outputContent -Encoding UTF8NoBOM -Force
            Write-Host "Markwhenデータが正常に出力: $OutputFilePath (Encoding: UTF8NoBOM)" -ForegroundColor Green
        } catch { Write-Error "出力ファイル書込失敗: $OutputFilePath. Error: $($_.Exception.Message)"; exit 1 }

    } elseif ($OutputFormat -eq "ExcelDirect") {
        $excelDatePattern = if ([string]::IsNullOrEmpty($script:DateFormatPattern) -or $script:DateFormatPattern -eq "yyyy-MM-DD") { "yyyy/MM/dd" } else { $script:DateFormatPattern }

        Write-Host "DEBUG: OutputFilePath before call = '$($OutputFilePath)'" -ForegroundColor Cyan
        Write-Host "DEBUG: OutputFilePath Type = $($OutputFilePath.GetType().FullName)" -ForegroundColor Cyan
        Write-Host "DEBUG: ProjectMetadata Type = $($projectMetadata.GetType().FullName)" -ForegroundColor Cyan
        Write-Host "DEBUG: RootWbsElements Count = $($rootWbsElements.Count)" -ForegroundColor Cyan

        # Pass the tree structure and the full flat list for lookups to the Excel writing function
        $params = @{
            ProjectInfo = $projectMetadata
            RootWbsElements = $rootWbsElements
            AllTaskNodesFlatListForDateLookup = $script:AllTaskNodesFlatList
            TargetExcelFilePath = $OutputFilePath
            SheetIdentifier = $ExcelSheetNameOrIndex
            StartRow = $ExcelStartRow
            ExcelDateFormat = $excelDatePattern
        }

        Write-WbsToExcelCom @params
        Write-Host "Excelデータが正常に出力されました: $OutputFilePath" -ForegroundColor Green

    } else {
        Write-Error "未対応の出力フォーマットです: $OutputFormat"
        exit 1
    }

    if ($dependencyWarnings.Count -gt 0) { Write-Warning "$($dependencyWarnings.Count) 件の依存関係に関する警告があります。" }
    Write-Host "処理正常完了。" -ForegroundColor Green

} catch {
    Write-Host "Debug: CRITICAL ERROR - $($_.Exception.GetType().FullName) - $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Debug: StackTrace - $($_.Exception.StackTrace)" -ForegroundColor DarkYellow
    if ($_.InvocationInfo) {
        Write-Host "Debug: Error at Line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line)" -ForegroundColor Yellow
    }
    Write-Error "スクリプト実行中致命的エラー: $($_.Exception.Message)"
    exit 1
} finally {
    # クリーンアップ処理
    # Global variables are automatically removed when the script finishes,
    # but explicit removal can be done if needed for clarity or specific scenarios.
    Remove-Variable -Name Holidays, AllTaskNodesFlatList, projectMetadata, rootWbsElements, parsedResult, DateFormatPattern, ExcelRowCounter -ErrorAction SilentlyContinue
}

# Removed unused functions: ConvertTo-ExcelString, Initialize-ProjectMetadata, Initialize-GlobalVariables, Calculate-TaskDates, Process-WbsFile, Get-ExcelCellAddress, Initialize-DependencyMap, Write-SuccessorInfo

# Removed excelColumnSpecs comment block as it describes the old mapping.