#Requires -Version 7.0

<#
.SYNOPSIS
    拡張MD-WBSファイル (新仕様) からMarkwhenタイムラインデータまたはExcelガントチャートシート用データを生成します。
.DESCRIPTION
    YAMLフロントマター (簡易サポート)、階層ID付き見出し、リスト属性を持つ拡張MD-WBSファイルを解析し、
    タスクの開始日を逆算（土日祝除外）して、Markwhen形式のタイムラインデータ、または既存のExcelシートへの
    WBSデータ書き込みを行います。
.NOTES
    Version: 3.6 (Excel書き込みデバッグ強化)
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
    [string]$DateFormatPattern = "yyyy/MM/dd",

    [Parameter(Mandatory = $false)]
    [string]$DefaultEncoding = "UTF8",

    [Parameter(Mandatory = $false)]
    [ValidateSet("InRange", "All", "None")]
    [string]$MarkwhenHolidayDisplayMode = "InRange",

    [Parameter(Mandatory = $false)]
    [object]$ExcelSheetNameOrIndex = 1,

    [Parameter(Mandatory = $false)]
    [int]$ExcelStartRow = 5
)

# --- グローバル変数・初期設定 ---
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:DateFormatPattern = $DateFormatPattern
$script:ExcelRowCounter = 0 # Excel書き込み用のスクリプトスコープ行カウンタ

# --- クラス定義 ---
# (WbsElementNode, ProjectMetadata クラスは変更なしのため、ここでは省略します。前回のコードと同じものを想定)
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
        # Write-Verbose "Parsing attributes for $($this.Id) $($this.Name)" # 詳細すぎるのでコメントアウト
        foreach ($line in $this.RawAttributeLines) {
            if ($line -match "^\s*-\s*(?<Key>deadline|duration|status|progress|org|assignee|depends):\s*(?<Value>.+)") {
                $key = $matches.Key.Trim().ToLower()
                $rawValue = $matches.Value.Trim()
                $value = ($rawValue -split '#')[0].Trim()

                $this.Attributes[$key] = $value
                # Write-Verbose "  Attribute found: $key = $value (Raw: $rawValue)" # 詳細すぎるのでコメントアウト
                switch ($key) {
                    "deadline" {
                        try { $this.Deadline = [datetime]::ParseExact($value, "yyyy-MM-dd", $null) }
                        catch { Write-Warning "要素 '$($this.Id) $($this.Name)' deadline形式無効: $value" }
                    }
                    "duration" {
                        $valNum = $value -replace 'd',''
                        if ($valNum -match '^\d+$') { $this.DurationDays = [int]$valNum }
                        else { Write-Warning "要素 '$($this.Id) $($this.Name)' duration形式無効: $value"; $this.DurationDays = 0 }
                    }
                    "depends" { $this.Attributes["depends"] = $value.TrimEnd('.') }
                }
            }
        }
    }
}

class ProjectMetadata {
    [string]$Title; [string]$Description; [datetime]$DefinedDate
    [datetime]$ProjectPlanStartDate; [datetime]$ProjectPlanOverallDeadline; [string]$View = "month"
}

# --- 主要関数 ---
# (Import-HolidayList, Parse-WbsMarkdownAndMetadataAdvanced, Calculate-TaskStartDate, Calculate-TaskEndDate, Format-GanttMarkwhen は変更なしのため省略)
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

function Parse-WbsMarkdownAndMetadataAdvanced {
    param(
        [Parameter(Mandatory)] [string]$FilePath,
        [Parameter(Mandatory)] [string]$Encoding
    )
    Write-Verbose "拡張MD-WBSファイルの解析開始: $FilePath"
    $projectMetadata = New-Object ProjectMetadata
    $rootElements = [System.Collections.Generic.List[WbsElementNode]]::new()
    if (-not (Test-Path $FilePath -PathType Leaf)) { Write-Error "WBSファイル '$FilePath' が見つかりません。"; return $null }

    $allLines = Get-Content -Path $FilePath -Encoding $Encoding
    $lineIndex = 0

    $convertFromYamlAvailable = $false
    try { if (Get-Command ConvertFrom-Yaml -ErrorAction SilentlyContinue) { $convertFromYamlAvailable = $true } } catch {}

    if ($allLines.Count -gt 0 -and $allLines[$lineIndex].Trim() -eq "---") {
        $lineIndex++; $yamlLines = New-Object System.Collections.Generic.List[string]
        while ($lineIndex -lt $allLines.Count -and $allLines[$lineIndex].Trim() -ne "---") { $yamlLines.Add($allLines[$lineIndex]); $lineIndex++ }
        if ($lineIndex -lt $allLines.Count -and $allLines[$lineIndex].Trim() -eq "---") {
            $lineIndex++
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
        } else { Write-Warning "YAMLフロントマターが正しく閉じられていません。" }
    }
    
    Write-Verbose "Markdown本文の解析中..."
    $parentStack = New-Object System.Collections.Stack; $currentWbsElement = $null 
    $attributeLinesBuffer = New-Object System.Collections.Generic.List[string]
    $descriptionBuffer = New-Object System.Text.StringBuilder

    function Process-BufferedAttributesAndDescriptionLocal { # ローカル関数化
        param ([WbsElementNode]$Element, [System.Collections.Generic.List[string]]$AttribBuffer, [System.Text.StringBuilder]$DescBuffer)
        if ($Element) {
            if ($AttribBuffer.Count -gt 0) { $Element.RawAttributeLines = $AttribBuffer.ToArray(); $Element.ParseCollectedAttributes(); $AttribBuffer.Clear() }
            if ($DescBuffer.Length -gt 0) { $Element.DescriptionText = $DescBuffer.ToString().Trim(); [void]$DescBuffer.Clear() }
        }
    }

    while ($lineIndex -lt $allLines.Count) {
        $line = $allLines[$lineIndex]; $trimmedLineForCheck = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmedLineForCheck)) { 
            Process-BufferedAttributesAndDescriptionLocal -Element $currentWbsElement -AttribBuffer $attributeLinesBuffer -DescBuffer $descriptionBuffer
            $lineIndex++; continue
        }
        if ($trimmedLineForCheck.StartsWith("#") -and -not ($trimmedLineForCheck -match "^#+\s+")) { Write-Verbose "コメント行をスキップ: $line"; $lineIndex++; continue }

        if ($trimmedLineForCheck -match "^(#+)\s+([\d\.]*[\d])\.?\s*(.+)") {
            Process-BufferedAttributesAndDescriptionLocal -Element $currentWbsElement -AttribBuffer $attributeLinesBuffer -DescBuffer $descriptionBuffer
            $currentWbsElement = $null 
            $level = $matches[1].Length; $id = $matches[2]; $name = $matches[3].Trim()
            $newElement = New-Object WbsElementNode -ArgumentList $id, $name, $level
            # Write-Verbose "見出しを検出: Level=$level, ID=$id, Name=$name" # 詳細すぎるのでコメントアウト
            if ($level -eq 2) { $newElement.ElementType = "Section" } elseif ($level -eq 3) { $newElement.ElementType = "Group" } elseif ($level -ge 4) { $newElement.ElementType = "Task" }  
            while ($parentStack.Count -gt 0 -and ($parentStack.Peek()).HierarchyLevel -ge $level) { [void]$parentStack.Pop() }
            if ($parentStack.Count -gt 0) { ($parentStack.Peek()).Children.Add($newElement); $newElement.Parent = $parentStack.Peek() } else { $rootElements.Add($newElement) }
            $parentStack.Push($newElement); $currentWbsElement = $newElement
        }
        elseif ($currentWbsElement -and $line -match "^\s{4}-\s*(?:deadline|duration|status|progress|org|assignee|depends):\s*.+") {
            if ($descriptionBuffer.Length -gt 0) { $currentWbsElement.DescriptionText = $descriptionBuffer.ToString().Trim(); [void]$descriptionBuffer.Clear() }
            $attributeLinesBuffer.Add($line) 
        }
        elseif ($currentWbsElement -and ($attributeLinesBuffer.Count -eq 0) -and (-not ($line.TrimStart() -match "^\s*-\s*")) -and (-not ($trimmedLineForCheck -match "^#+"))) { 
             [void]$descriptionBuffer.AppendLine($line.Trim()) 
        }
        else {
            # Write-Verbose "未分類の行 (または前の要素の終端とみなし属性処理): $line" # 詳細すぎるのでコメントアウト
            Process-BufferedAttributesAndDescriptionLocal -Element $currentWbsElement -AttribBuffer $attributeLinesBuffer -DescBuffer $descriptionBuffer
        }
        $lineIndex++
    }
    Process-BufferedAttributesAndDescriptionLocal -Element $currentWbsElement -AttribBuffer $attributeLinesBuffer -DescBuffer $descriptionBuffer
    Write-Verbose "Markdown本文解析完了。トップレベル要素数: $($rootElements.Count)"
    return @{ ProjectMetadata = $projectMetadata; RootElements = $rootElements }
}

function Calculate-TaskStartDate { # (変更なし)
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

function Calculate-TaskEndDate { # (変更なし - 安定版ロジック)
    param( [Parameter(Mandatory)] [datetime]$StartDate, [Parameter(Mandatory)] [int]$DurationBusinessDays, [Parameter(Mandatory)] [System.Collections.Generic.List[datetime]]$Holidays )
    if ($DurationBusinessDays -le 0) { return $StartDate } 
    $calculatedEndDate = $StartDate; $businessDaysCounted = 0
    if ($StartDate.DayOfWeek -ne [DayOfWeek]::Saturday -and $StartDate.DayOfWeek -ne [DayOfWeek]::Sunday -and (-not $Holidays.Contains($StartDate.Date))) { $businessDaysCounted = 1 }
    if ($DurationBusinessDays -eq 1 -and $businessDaysCounted -eq 1) { return $StartDate }
    if ($DurationBusinessDays -eq 1 -and $businessDaysCounted -eq 0) {
         $tempEndDate = $StartDate.AddDays(1)
         while ($tempEndDate.DayOfWeek -eq [DayOfWeek]::Saturday -or $tempEndDate.DayOfWeek -eq [DayOfWeek]::Sunday -or $Holidays.Contains($tempEndDate.Date)) { $tempEndDate = $tempEndDate.AddDays(1) }
         return $tempEndDate
    }
    while ($businessDaysCounted -lt $DurationBusinessDays) {
        $calculatedEndDate = $calculatedEndDate.AddDays(1)
        if ($calculatedEndDate.DayOfWeek -ne [DayOfWeek]::Saturday -and $calculatedEndDate.DayOfWeek -ne [DayOfWeek]::Sunday) {
            if (-not $Holidays.Contains($calculatedEndDate.Date)) { $businessDaysCounted++ }
        }
        if ($calculatedEndDate -gt $StartDate.AddYears(2)) { Write-Warning "終了日計算異常: 2年以上先に及びました。"; return $StartDate.AddYears(2) }
    }
    return $calculatedEndDate
}

function Format-GanttMarkwhen { # (デバッグ出力削除)
    param(
        [Parameter(Mandatory)] [ProjectMetadata]$Metadata,
        [Parameter(Mandatory)] [System.Collections.Generic.List[WbsElementNode]]$RootWbsElements,
        [Parameter(Mandatory)] [string]$DatePattern, 
        [Parameter(Mandatory)] [System.Collections.Generic.List[datetime]]$Holidays,
        [Parameter(Mandatory=$false)] [datetime]$ProjectOverallStartDateForHolidayFilter,
        [Parameter(Mandatory=$false)] [datetime]$ProjectOverallEndDateForHolidayFilter,
        [Parameter(Mandatory)] [string]$HolidayDisplayMode
    )
    Write-Verbose "Markwhen形式のガントチャートデータを生成。祝日モード: $HolidayDisplayMode"
    $outputLines = New-Object System.Collections.Generic.List[string]
    $outputLines.Add("---")
    if ($Metadata -and (-not [string]::IsNullOrEmpty($Metadata.Title))) { $outputLines.Add("title: $($Metadata.Title)") }
    if ($Metadata -and (-not [string]::IsNullOrEmpty($Metadata.Description))) { $outputLines.Add("description: $($Metadata.Description)") }
    if ($Metadata -and $Metadata.DefinedDate -ne ([datetime]::MinValue)) { $outputLines.Add("date: $($Metadata.DefinedDate.ToString('yyyy-MM-dd'))") }
    if ($Metadata -and (-not [string]::IsNullOrEmpty($Metadata.View))) { $outputLines.Add("view: $($Metadata.View)") } else { $outputLines.Add("view: month") } 
    $outputLines.Add("---"); $outputLines.Add("")

    function ConvertTo-MarkwhenRecursiveUpdated {
        param ([WbsElementNode]$Element, [int]$CurrentMarkwhenIndentLevel, [System.Collections.Generic.List[string]]$OutputLinesRef, [string]$OutputDatePattern)
        $indent = "    " * $CurrentMarkwhenIndentLevel 
        $isEffectivelyGroup = ($Element.ElementType -eq "Group") -or ($Element.ElementType -eq "Task" -and $Element.Children.Count -gt 0)
        if ($Element.ElementType -eq "Section") {
            $OutputLinesRef.Add("${indent}section $($Element.Name)"); foreach ($child in $Element.Children) { ConvertTo-MarkwhenRecursiveUpdated -Element $child -CurrentMarkwhenIndentLevel $CurrentMarkwhenIndentLevel -OutputLinesRef $OutputLinesRef -OutputDatePattern $OutputDatePattern }; $OutputLinesRef.Add("${indent}endsection") 
        } elseif ($isEffectivelyGroup) {
            $OutputLinesRef.Add("${indent}group `"$($Element.Name)`"") 
            if ($Element.ElementType -eq "Task" -and $Element.CalculatedStartDate -ne ([datetime]::MinValue) -and $Element.CalculatedEndDate -ne ([datetime]::MinValue) -and $Element.DurationDays -gt 0) {
                $startDateStr = $Element.CalculatedStartDate.ToString($OutputDatePattern); $endDateStr = $Element.CalculatedEndDate.ToString($OutputDatePattern); $taskDisplayName = "概要"; $tags = @("#$($Element.Id.Replace('.','_'))"); if ($Element.Attributes.ContainsKey("status")) { $tags += "#$($Element.Attributes['status'])" }; if ($Element.Attributes.ContainsKey("progress")) { $taskDisplayName += " ($($Element.Attributes['progress'] -replace '%',''))%" }
                $groupTaskLine = "$($indent)    ${startDateStr}/${endDateStr}: ${taskDisplayName}"; if ($tags.Count -gt 0) { $groupTaskLine += " $([string]::Join(' ', $tags))" }; $OutputLinesRef.Add($groupTaskLine)
            }
            foreach ($child in $Element.Children) { ConvertTo-MarkwhenRecursiveUpdated -Element $child -CurrentMarkwhenIndentLevel ($CurrentMarkwhenIndentLevel + 1) -OutputLinesRef $OutputLinesRef -OutputDatePattern $OutputDatePattern }
            $OutputLinesRef.Add("${indent}endgroup")
        } elseif ($Element.ElementType -eq "Task") { 
            if ($Element.CalculatedStartDate -ne ([datetime]::MinValue) -and $Element.CalculatedEndDate -ne ([datetime]::MinValue) -and $Element.DurationDays -gt 0) { 
                $currentIndent = "    " * $CurrentMarkwhenIndentLevel; $currentSDateStr = $Element.CalculatedStartDate.ToString($OutputDatePattern); $currentEDateStr = $Element.CalculatedEndDate.ToString($OutputDatePattern); $currentTaskName = $Element.Name
                if ($Element.Attributes.ContainsKey("progress")) { 
                    $pStr = $Element.Attributes.progress.Trim() -replace '%'
                    if ($pStr -match '^\d+$') { 
                        $progressVal = [double]$pStr
                        if ($progressVal -ge 0 -and $progressVal -le 100) { 
                            $progressToSet = ($progressVal / 100).ToString() 
                        } 
                    } 
                }
                if ($Element.Attributes.ContainsKey("status")) {
                    $status = $Element.Attributes.status.Trim().ToLower()
                    if (($status -eq "inprogress" -or $status -eq "completed") -and $Element.CalculatedStartDate -ne ([datetime]::MinValue)) { $actualStartDateToSet = $Element.CalculatedStartDate.ToString($OutputDatePattern) }
                    if ($status -eq "completed" -and $Element.CalculatedEndDate -ne ([datetime]::MinValue)) { $actualEndDateToSet = $Element.CalculatedEndDate.ToString($OutputDatePattern); $progressToSet = 1 }
                }
                $currentTags = @("#$($Element.Id.Replace('.','_'))"); if ($Element.Attributes.ContainsKey("status")) { $currentTags += "#$($Element.Attributes['status'])" }; if ($Element.Attributes.ContainsKey("assignee")) { $currentTags += "#assignee-$($Element.Attributes['assignee'] -replace ' ','_')" }; if ($Element.Attributes.ContainsKey("org")) { $currentTags += "#org-$($Element.Attributes['org'] -replace ' ','_')" }
                $tagsString = ""; if ($currentTags.Count -gt 0) { $tagsString = " $([string]::Join(' ', $currentTags))" }
                $finalTaskLine = "${currentIndent}${currentSDateStr}/${currentEDateStr}: ${currentTaskName}${tagsString}"; $OutputLinesRef.Add($finalTaskLine)
                if (-not [string]::IsNullOrEmpty($Element.DescriptionText)) { foreach($descLine in ($Element.DescriptionText.TrimEnd("`n") -split "`n")) { $OutputLinesRef.Add("$currentIndent    // $descLine") } }
            } else { Write-Warning "タスク '$($Element.Id) $($Element.Name)' の日付/期間未設定、または0のため出力スキップ。" }
        }
    }
    foreach ($rootElement in $RootWbsElements) { ConvertTo-MarkwhenRecursiveUpdated -Element $rootElement -CurrentMarkwhenIndentLevel 0 -OutputLinesRef $outputLines -OutputDatePattern $DatePattern; $outputLines.Add("") }
    if ($HolidayDisplayMode -ne "None" -and $Holidays -and $Holidays.Count -gt 0) {
        $holidaysToDisplay = @(); if ($HolidayDisplayMode -eq "All") { $holidaysToDisplay = $Holidays | Sort-Object } elseif ($HolidayDisplayMode -eq "InRange" -and $ProjectOverallStartDateForHolidayFilter -ne $null -and $ProjectOverallEndDateForHolidayFilter -ne $null) { $holidaysToDisplay = $Holidays | Where-Object { $_.Date -ge $ProjectOverallStartDateForHolidayFilter.Date -and $_.Date -le $ProjectOverallEndDateForHolidayFilter.Date } | Sort-Object }
        if ($holidaysToDisplay.Count -gt 0) { $outputLines.Add("section 祝日"); $outputLines.Add("group Holidays #project_holidays"); foreach($holiday in $holidaysToDisplay) { $outputLines.Add("    $($holiday.ToString($DatePattern)): Holiday") }; $outputLines.Add("endgroup"); $outputLines.Add("endsection"); $outputLines.Add("") }
    }
    Write-Verbose "Markwhenデータ生成完了。"
    return $outputLines -join [System.Environment]::NewLine
}

# --- Excelへの書き込み関数 (デバッグ強化) ---
function Write-WbsToExcelCom {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ProjectMetadata]$ProjectMetadata,
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[WbsElementNode]]$RootWbsElements,
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[WbsElementNode]]$AllTaskNodesFlatListForDateLookup,
        [Parameter(Mandatory)]
        [string]$TargetExcelFilePath,
        [Parameter(Mandatory)]
        [object]$SheetIdentifier, # メインスクリプトの型と合わせる
        [Parameter(Mandatory = $false)] # Mandatoryをfalseに変更
        [int]$StartRow = 5,          # デフォルト値を設定
        [Parameter(Mandatory)]
        [string]$ExcelDateFormat
    )

    try {
        Write-Host "DEBUG: Write-WbsToExcelCom - 関数開始" -ForegroundColor Cyan
        Write-Verbose "=== Excel COMオブジェクト初期化開始 ==="
        $excelCom = $null; $workbook = $null; $worksheet = $null
        
        try {
            $excelCom = New-Object -ComObject Excel.Application
            if ($null -eq $excelCom) { 
                throw "Excelアプリケーションの初期化に失敗しました。" 
            }
            $excelCom.Visible = $false
            $excelCom.DisplayAlerts = $false
            $excelCom.EnableEvents = $false # イベントを無効化して高速化・安定化
            $excelCom.ScreenUpdating = $false # 画面描画をオフにして高速化

            Write-Host "DEBUG: Write-WbsToExcelCom - ワークブックオープン試行: $TargetExcelFilePath" -ForegroundColor Cyan
            if (-not (Test-Path $TargetExcelFilePath -PathType Leaf)) { throw "指定されたExcelファイルが見つかりません: $TargetExcelFilePath" }
            $workbook = $excelCom.Workbooks.Open($TargetExcelFilePath)
            if ($null -eq $workbook) { throw "ワークブックのオープンに失敗しました。" }

            Write-Host "DEBUG: Write-WbsToExcelCom - ワークシート取得試行: $SheetIdentifier" -ForegroundColor Cyan
            try {
                $worksheet = $workbook.Worksheets.Item($SheetIdentifier)
                if ($null -eq $worksheet) { throw "シート '$SheetIdentifier' の取得に失敗しました（オブジェクトがnull）。" }
                Write-Host "DEBUG: Write-WbsToExcelCom - ワークシート名: $($worksheet.Name)" -ForegroundColor Cyan
            } catch {
                throw "指定されたワークシート '$SheetIdentifier' が見つかりません。エラー: $($_.Exception.Message)"
            }

            # D1セルにプロジェクト名を書き込む
            if ($ProjectMetadata -and (-not [string]::IsNullOrEmpty($ProjectMetadata.Title))) {
                Write-Verbose "D1セルにプロジェクト名を設定: $($ProjectMetadata.Title)"
                $worksheet.Cells.Item(1, 4).Value2 = $ProjectMetadata.Title
            } else {
                Write-Warning "プロジェクトメタデータまたはタイトルが取得できなかったため、ExcelのD1セルへの書き込みをスキップします。"
            }

            # O1セルにプロジェクト全体の計算上の開始日を書き込む
            $overallProjectStartDate = $null
            if ($AllTaskNodesFlatListForDateLookup.Count -gt 0) {
                $validTasksForStartDate = $AllTaskNodesFlatListForDateLookup | Where-Object {$_.CalculatedStartDate -ne ([datetime]::MinValue)}
                if ($validTasksForStartDate.Count -gt 0) {
                    $overallProjectStartDate = ($validTasksForStartDate | Sort-Object CalculatedStartDate | Select-Object -First 1).CalculatedStartDate
                }
            }
            if ($overallProjectStartDate -ne $null -and $overallProjectStartDate -ne ([datetime]::MinValue)) {
                 Write-Verbose "O1セルにプロジェクト開始日を設定: $($overallProjectStartDate.ToString($ExcelDateFormat))"
                 $worksheet.Cells.Item(1, 15).Value2 = $overallProjectStartDate.ToString($ExcelDateFormat)
            } else {
                 Write-Warning "プロジェクト全体の計算上の開始日が特定できなかったため、ExcelのO1セルへの書き込みをスキップします。"
            }

            Write-Verbose "=== データ書き込み開始 (開始行: $StartRow) ==="
            $script:ExcelRowCounter = $StartRow # スクリプトスコープの行カウンタを初期化

            function Add-RowToExcelRecursiveDetailed {
                param (
                    [WbsElementNode]$Element,
                    $CurrentWorksheet, # Worksheetオブジェクトを渡す
                    [string]$OutputDateFormat,
                    [System.Collections.Generic.List[WbsElementNode]]$AllTasksForNameLookup
                )

                Write-Host "DEBUG Add-Row: ID '$($Element.Id)' Name '$($Element.Name)' to Row $($script:ExcelRowCounter)" -ForegroundColor Green

                # A列(1): No
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 1).Value2 = $Element.Id.ToString()
                # B, C, D列
                if ($Element.HierarchyLevel -eq 2) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 2).Value2 = $Element.Name } # Nameは既にstringだが念のため
                elseif ($Element.HierarchyLevel -eq 3) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 3).Value2 = $Element.Name } # Nameは既にstringだが念のため
                elseif ($Element.HierarchyLevel -ge 4) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 4).Value2 = $Element.Name } # Nameは既にstringだが念のため
                # E, F, G列
                if ($Element.Attributes.ContainsKey("depends") -and -not [string]::IsNullOrEmpty($Element.Attributes.depends)) {
                    $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 5).Value2 = "先行"
                    $dependsOnId = $Element.Attributes.depends.Trim()
                    $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 6).Value2 = $dependsOnId.ToString() # IDも文字列として扱う
                    $parentTaskObject = $AllTasksForNameLookup | Where-Object {$_.Id -eq $dependsOnId} | Select-Object -First 1
                    if ($parentTaskObject) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 7).Value2 = $parentTaskObject.Name.ToString() }
                    else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 7).Value2 = "" } # 先行タスク名が見つからない場合は空
                } else {
                    $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 5).Value2 = ""
                    $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 6).Value2 = ""
                    $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 7).Value2 = ""
                }
                # I列
                if (-not [string]::IsNullOrEmpty($Element.DescriptionText)) {
                    $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 9).Value2 = $Element.DescriptionText -replace "`r?`n", [char]10
                    $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 9).WrapText = $true
                } else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 9).Value2 = "" }
                # N, O列
                if ($Element.Attributes.ContainsKey("org")) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 14).Value2 = $Element.Attributes.org.Trim().ToString() } else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 14).Value2 = "" }
                if ($Element.Attributes.ContainsKey("assignee")) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 15).Value2 = $Element.Attributes.assignee.Trim().ToString() } else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 15).Value2 = "" }
                # R, S, T列
                if ($Element.CalculatedStartDate -ne ([datetime]::MinValue)) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 18).Value2 = $Element.CalculatedStartDate.ToString($OutputDateFormat) } else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 18).Value2 = "" } # R列: 開始入力 (スクリプト計算結果)
                if ($Element.Deadline -ne ([datetime]::MinValue)) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 19).Value2 = $Element.Deadline.ToString($OutputDateFormat) } else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 19).Value2 = "" } # S列: 終了入力 (MD-WBS Deadline)
                if ($Element.DurationDays -gt 0) { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 20).Value2 = $Element.DurationDays.ToString() } else { $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 20).Value2 = "" } # T列: 日数入力
                # X, Y, Z列
                $progressToSet = ""; $actualStartDateToSet = ""; $actualEndDateToSet = ""
                if ($Element.Attributes.ContainsKey("progress")) { 
                    $pStr = $Element.Attributes.progress.Trim() -replace '%'
                    if ($pStr -match '^\d+$') { 
                        $progressVal = [double]$pStr
                        if ($progressVal -ge 0 -and $progressVal -le 100) { 
                            $progressToSet = ($progressVal / 100).ToString() 
                        } 
                    } 
                }
                if ($Element.Attributes.ContainsKey("status")) {
                    $status = $Element.Attributes.status.Trim().ToLower()
                    if (($status -eq "inprogress" -or $status -eq "completed") -and $Element.CalculatedStartDate -ne ([datetime]::MinValue)) { $actualStartDateToSet = $Element.CalculatedStartDate.ToString($OutputDateFormat) }
                    if ($status -eq "completed" -and $Element.CalculatedEndDate -ne ([datetime]::MinValue)) { $actualEndDateToSet = $Element.CalculatedEndDate.ToString($OutputDateFormat); $progressToSet = "1" } # 数値の1を文字列"1"として設定
                }
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 24).Value2 = $progressToSet
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 25).Value2 = $actualStartDateToSet
                $CurrentWorksheet.Cells.Item($script:ExcelRowCounter, 26).Value2 = $actualEndDateToSet
                
                $script:ExcelRowCounter++

                foreach ($child in $Element.Children) {
                    Add-RowToExcelRecursiveDetailed -Element $child -CurrentWorksheet $CurrentWorksheet -OutputDateFormat $OutputDateFormat -AllTasksForNameLookup $AllTasksForNameLookup
                }
            }

            Write-Host "DEBUG: Write-WbsToExcelCom - RootWbsElements Count: $($RootWbsElements.Count)" -ForegroundColor Cyan
            foreach ($rootElement in $RootWbsElements) {
                Add-RowToExcelRecursiveDetailed -Element $rootElement -CurrentWorksheet $worksheet -OutputDateFormat $ExcelDateFormat -AllTasksForNameLookup $AllTaskNodesFlatListForDateLookup
            }

            $workbook.Save()
            Write-Verbose "Excelファイルにデータを書き込み、保存しました。"

        } catch {
            $errMsg = "Excel操作中にエラーが発生しました: "
            if ($PSItem -and $PSItem.Exception) {
                $errMsg += "$($PSItem.Exception.GetType().FullName) - $($PSItem.Exception.Message)"
                Write-Verbose "Original Exception StackTrace: $($PSItem.Exception.StackTrace)"
            }
            Write-Error $errMsg
            throw
        }
    } catch {
        Write-Error "致命的なエラーが発生しました: $($_.Exception.Message)"
        exit 1
    } finally {
        Write-Verbose "=== Excelリソース解放処理 ==="
        if ($excelCom) {
            try {
                # 1. ワークシートの解放
                if ($worksheet -ne $null) {
                    Write-Verbose "ワークシートCOMオブジェクトを解放中..."
                    $worksheet = $null
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
                }

                # 2. ワークブックの解放
                if ($workbook -ne $null) {
                    Write-Verbose "ワークブックを閉じています..."
                    $workbook.Close($false)
                    $workbook = $null
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
                }

                # 3. Excelアプリケーションの解放
                Write-Verbose "Excelアプリケーションを終了しています..."
                $excelCom.Quit()
                $excelCom = $null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelCom) | Out-Null

                # 4. 変数の削除とGC
                Remove-Variable excelCom, workbook, worksheet -ErrorAction SilentlyContinue
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                Start-Sleep -Seconds 1
            } catch {
                Write-Warning "Excelリソース解放中にエラー: $($_.Exception.Message)"
            }
        }
    }
}

function Validate-TaskDependencies {
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[WbsElementNode]]$TaskNodes
    )
    
    # 戻り値の型を明示的に指定
    [System.Collections.Generic.List[string]]$warnings = New-Object System.Collections.Generic.List[string]
    
    # 警告の処理
    foreach ($task in $TaskNodes) {
        if ($task.Attributes.ContainsKey("depends")) {
            # 依存関係のチェックロジック
            # 警告がある場合は $warnings.Add("警告メッセージ") を実行
        }
    }
    
    return $warnings
}

function Stop-ExcelProcesses {
    Write-Verbose "既存のExcelプロセスを終了試行中..."
    $excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        $excelProcesses | ForEach-Object {
            Write-Warning "実行中のExcelプロセス $($_.Id) を強制終了します。"
            try {
                Stop-Process -Id $_.Id -Force -ErrorAction Stop
                Wait-Process -Id $_.Id -Timeout 5 -ErrorAction SilentlyContinue
            } catch {
                Write-Warning "プロセス $($_.Id) の終了に失敗: $($_.Exception.Message)"
            }
        }
        Start-Sleep -Seconds 2
    }
}

# --- メイン処理 ---
try {
    Write-Host "処理を開始します (MD-WBS仕様)..." -ForegroundColor Green

    # メイン処理の最初で呼び出し
    Stop-ExcelProcesses

    $script:Holidays = Import-HolidayList -InputOfficialHolidayFilePath $OfficialHolidayFilePath -InputCompanyHolidayFilePath $CompanyHolidayFilePath -BaseEncoding $DefaultEncoding
    $parsedResult = Parse-WbsMarkdownAndMetadataAdvanced -FilePath $WbsFilePath -Encoding $DefaultEncoding
    if (-not $parsedResult) { Write-Error "MD-WBSファイル解析失敗。"; exit 1 }
    $projectMetadata = $parsedResult.ProjectMetadata
    $rootWbsElements = $parsedResult.RootElements
    if ($rootWbsElements.Count -eq 0 -and (-not $projectMetadata -or ([string]::IsNullOrEmpty($projectMetadata.Title)) )) { Write-Error "MD-WBS有効データ無。"; exit 1 }

    $script:AllTaskNodesFlatList = New-Object System.Collections.Generic.List[WbsElementNode]
    function Flatten-WbsTreeNodes { 
        param ([System.Collections.Generic.List[WbsElementNode]]$Elements)
        if ($null -eq $Elements) { return }
        foreach ($element in $Elements) {
            # 日付計算の対象はDurationDaysとDeadlineの両方を持つもの
            if ($element.DurationDays -gt 0 -and $element.Deadline -ne ([datetime]::MinValue)) {
                 $script:AllTaskNodesFlatList.Add($element)
            }
            # Excel書き込みのためには全ノードのフラットリストが必要な場合があるので、日付計算用とは別に作成も検討
            # ここでは日付計算用のリストをそのままExcelの先行タスク名検索にも使う
            if ($element.Children.Count -gt 0) { Flatten-WbsTreeNodes -Elements $element.Children }
        }
    }
    # Excel書き込みのためには全ノードのフラットリストが必要な場合があるので、日付計算用とは別に作成も検討
    # ここでは日付計算用のリストをそのままExcelの先行タスク名検索にも使う
    Flatten-WbsTreeNodes -Elements $rootWbsElements


    Write-Verbose "各タスクの日付を算出中..."
    if ($null -ne $script:AllTaskNodesFlatList) {
        foreach ($taskNode in $script:AllTaskNodesFlatList) { 
            $taskNode.CalculatedStartDate = Calculate-TaskStartDate -DeadlineDate $taskNode.Deadline -DurationBusinessDays $taskNode.DurationDays -Holidays $script:Holidays
            $taskNode.CalculatedEndDate = Calculate-TaskEndDate -StartDate $taskNode.CalculatedStartDate -DurationBusinessDays $taskNode.DurationDays -Holidays $script:Holidays
            Write-Verbose "タスク '$($taskNode.Id) $($taskNode.Name)' 開始: $($taskNode.CalculatedStartDate.ToString($script:DateFormatPattern)), 終了: $($taskNode.CalculatedEndDate.ToString($script:DateFormatPattern)) (期間: $($taskNode.DurationDays)d)"
        }
    }
    
    # (プロジェクト全体期間特定は変更なしのため省略)
    $projectOverallStartDateForHolidayFilter = $projectMetadata.ProjectPlanStartDate; $projectOverallEndDateForHolidayFilter = $projectMetadata.ProjectPlanOverallDeadline; if (($null -ne $script:AllTaskNodesFlatList) -and $script:AllTaskNodesFlatList.Count -gt 0) {$validTasksForDateRange = $script:AllTaskNodesFlatList | Where-Object {$_.CalculatedStartDate -ne ([datetime]::MinValue) -and $_.CalculatedEndDate -ne ([datetime]::MinValue)}; if ($validTasksForDateRange.Count -gt 0) {if (($projectOverallStartDateForHolidayFilter -eq $null -or $projectOverallStartDateForHolidayFilter -eq ([datetime]::MinValue))) { $projectOverallStartDateForHolidayFilter = ($validTasksForDateRange | Sort-Object CalculatedStartDate | Select-Object -First 1).CalculatedStartDate}; if (($projectOverallEndDateForHolidayFilter -eq $null -or $projectOverallEndDateForHolidayFilter -eq ([datetime]::MinValue))) {$projectOverallEndDateForHolidayFilter = ($validTasksForDateRange | Sort-Object CalculatedEndDate -Descending | Select-Object -First 1).CalculatedEndDate}}}
    if($projectOverallStartDateForHolidayFilter -and $projectOverallEndDateForHolidayFilter) { Write-Verbose "祝日フィルタ用プロジェクト期間: $($projectOverallStartDateForHolidayFilter.ToString($script:DateFormatPattern)) - $($projectOverallEndDateForHolidayFilter.ToString($script:DateFormatPattern))" }

    $dependencyWarnings = Validate-TaskDependencies -TaskNodes $script:AllTaskNodesFlatList # 既存の依存関係チェックを呼び出し

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
        $excelDatePattern = if ([string]::IsNullOrEmpty($script:DateFormatPattern) -or $script:DateFormatPattern -match "yyyy-MM-dd") { "yyyy/MM/dd" } else { $script:DateFormatPattern }
        
        Write-Host "DEBUG: Excel書き込み呼び出し - OutputFilePath: $OutputFilePath" -ForegroundColor Green
        # Add debug lines to check type and value just before the call
        Write-Host "DEBUG: Checking \$OutputFilePath before Write-WbsToExcelCom call." -ForegroundColor Yellow
        Write-Host "DEBUG: \$OutputFilePath Value: '$($OutputFilePath)'" -ForegroundColor Yellow
        Write-Host "DEBUG: \$OutputFilePath Type: $($OutputFilePath.GetType().FullName)" -ForegroundColor Yellow

        # 1. プロパティの存在確認を追加
        if ($projectMetadata.PSObject.Properties.Name -contains 'StartDate') {
            $startDate = $projectMetadata.StartDate
        } else {
            # デフォルト値の設定やエラーハンドリング
            Write-Warning "StartDate プロパティが見つかりません。デフォルト値を使用します。"
            $startDate = Get-Date  # または適切なデフォルト値
        }

        # 2. Hashtable の作成
        $projectInfoHashtable = @{
            Title = $projectMetadata.Title
            ProjectPlanStartDate = $projectMetadata.ProjectPlanStartDate
            ProjectPlanOverallDeadline = $projectMetadata.ProjectPlanOverallDeadline
        }

        # 2. 変換した Hashtable を渡す
        Write-WbsToExcelCom -ProjectMetadata $projectMetadata `
                            -RootWbsElements $rootWbsElements `
                            -AllTaskNodesFlatListForDateLookup $script:AllTaskNodesFlatList `
                            -TargetExcelFilePath $OutputFilePath `
                            -SheetIdentifier $ExcelSheetNameOrIndex `
                            -StartRow $ExcelStartRow `
                            -ExcelDateFormat $excelDatePattern
        Write-Host "Excelデータが正常に出力されました: $OutputFilePath" -ForegroundColor Green
    } else { Write-Error "未対応の出力フォーマットです: $OutputFormat"; exit 1 }

    if ($null -ne $dependencyWarnings -and $dependencyWarnings.Count -gt 0) { Write-Warning "$($dependencyWarnings.Count) 件の依存関係に関する警告があります。MD-WBSの計画を見直してください。" } # メッセージ調整
    Write-Host "処理正常完了。" -ForegroundColor Green

} catch {
    Write-Host "Debug: CRITICAL ERROR - $($_.Exception.GetType().FullName) - $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Debug: StackTrace - $($_.Exception.StackTrace)" -ForegroundColor DarkYellow
    if ($_.InvocationInfo) { Write-Host "Debug: Error at Line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line)" -ForegroundColor Yellow }
    Write-Error "スクリプト実行中致命的エラー: $($_.Exception.Message)"
    exit 1
} finally {
    # スクリプトスコープ変数をクリア
    Remove-Variable -Name Holidays, AllTaskNodesFlatList, DateFormatPattern, ExcelRowCounter -Scope script -ErrorAction SilentlyContinue
    # メイン処理のローカル変数は自動でクリアされる
}

# メインスクリプトへのパラメータ準備
$scriptParams = @{
    WbsFilePath             = $TestWbsFile
    OfficialHolidayFilePath = $OfficialHolidays
    OutputFilePath          = $TestOutputFile
    OutputFormat            = $TestOutputFormat
    DateFormatPattern       = "yyyy/MM/dd"
    DefaultEncoding         = "UTF8"
    Verbose                 = $true
    ExcelStartRow           = 5  # 固定値として設定
}

if ($TestOutputFormat -eq "ExcelDirect") {
    $scriptParams.ExcelSheetNameOrIndex = "Sheet1"
}

# テンプレートファイルの確認をここに移動
if ($TestOutputFormat -eq "ExcelDirect") {
    Write-Host "テンプレートファイルのワークシート (確認用):"
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($TestOutputFile)
        foreach ($sheet in $workbook.Worksheets) { 
            Write-Host "  シート名: $($sheet.Name), インデックス: $($sheet.Index)"
        }
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Remove-Variable excel, workbook -ErrorAction SilentlyContinue
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Start-Sleep -Seconds 2
    } catch { 
        Write-Warning "Excelテンプレートのシート名確認中にエラー: $($_.Exception.Message)"
    }
}