#Requires -Version 7.0
# (ConvertFrom-YamlのRequiresはコメントアウトのまま)

<#
.SYNOPSIS
    拡張MD-WBSファイル (新仕様) からMarkwhenタイムラインデータを生成します。
.DESCRIPTION
    YAMLフロントマター (簡易サポート)、階層ID付き見出し、リスト属性を持つ拡張MD-WBSファイルを解析し、
    タスクの開始日を逆算（土日祝除外）して、Markwhen形式のタイムラインデータを出力します。
.PARAMETER WbsFilePath
    拡張MD-WBSデータが記述されたMarkdownファイルのパス。 (必須)
.PARAMETER OfficialHolidayFilePath
    国民の祝日などが記載されたCSVファイルのパス。 (必須)
.PARAMETER CompanyHolidayFilePath
    会社独自の休日が記載されたCSVファイルのパス。 (任意)
.PARAMETER OutputFilePath
    生成されるMarkwhenデータを出力するファイルのパス。 (必須)
.PARAMETER OutputFormat
    出力形式。"Markwhen" のみをサポート。 (必須、デフォルト "Markwhen")
.PARAMETER DateFormatPattern
    Markwhen内で使用する日付の書式。デフォルトは "yyyy-MM-dd"。 (任意)
.PARAMETER DefaultEncoding
    入出力ファイルのデフォルトエンコーディング。デフォルトは "UTF8"。 (任意)
.PARAMETER MarkwhenHolidayDisplayMode
    Markwhen出力時の祝日表示モード。"All", "InRange", "None"。デフォルトは "InRange"。 (任意)
.EXAMPLE
    .\md-wbs2gantt.ps1 -WbsFilePath ".\ServiceDev_v5.md" -OfficialHolidayFilePath ".\syukujitsu.csv" -CompanyHolidayFilePath ".\company_holidays.csv" -OutputFilePath ".\ServiceDev.mw"
.NOTES
    Version: 3.0 (Cleaned up debug logs and comments)
    Author: AI Assistant (Gemini) based on user requirements
    Date: 2025-05-16
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$WbsFilePath,

    [Parameter(Mandatory=$true, Position=1)]
    [string]$OfficialHolidayFilePath,

    [Parameter(Mandatory=$false, Position=2)]
    [string]$CompanyHolidayFilePath,

    [Parameter(Mandatory=$true, Position=3)]
    [string]$OutputFilePath,

    [Parameter(Mandatory=$true, Position=4)]
    [ValidateSet("Markwhen")]
    [string]$OutputFormat = "Markwhen",

    [Parameter(Mandatory=$false)]
    [string]$DateFormatPattern = "yyyy-MM-dd",

    [Parameter(Mandatory=$false)]
    [string]$DefaultEncoding = "UTF8",

    [Parameter(Mandatory=$false)]
    [ValidateSet("All", "InRange", "None")]
    [string]$MarkwhenHolidayDisplayMode = "InRange"
)

# --- グローバル変数・初期設定 ---
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$Global:DateFormatPattern = $DateFormatPattern

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
        # Write-Verbose "Parsing attributes for $($this.Id) $($this.Name)" # 詳細すぎるためコメントアウト
        foreach ($line in $this.RawAttributeLines) {
            if ($line -match "^\s*-\s*(?<Key>deadline|duration|status|progress|assignee|depends):\s*(?<Value>.+)") {
                $key = $matches.Key.Trim().ToLower()
                $rawValue = $matches.Value.Trim()
                $value = ($rawValue -split '#')[0].Trim()

                $this.Attributes[$key] = $value
                # Write-Verbose "  Attribute found: $key = $value (Raw: $rawValue)" # 詳細すぎるためコメントアウト
                switch ($key) {
                    "deadline" {
                        try {
                            $this.Deadline = [datetime]::ParseExact($value, "yyyy-MM-dd", $null)
                            # Write-Verbose "    Deadline parsed: $($this.Deadline.ToString($Global:DateFormatPattern))"
                        } catch {
                            Write-Warning "要素 '$($this.Id) $($this.Name)' deadline形式無効: $value"
                        }
                    }
                    "duration" {
                        $valNum = $value -replace 'd',''
                        if ($valNum -match '^\d+$') {
                            $this.DurationDays = [int]$valNum
                            # Write-Verbose "    DurationDays parsed: $($this.DurationDays)"
                        } else {
                            Write-Warning "要素 '$($this.Id) $($this.Name)' duration形式無効: $value"
                            $this.DurationDays = 0
                        }
                    }
                    "depends" {
                        $this.Attributes["depends"] = $value.TrimEnd('.')
                        # Write-Verbose "    Depends parsed (trimmed): $($this.Attributes["depends"])"
                    }
                }
            }
        }
    }
}

class ProjectMetadata {
    [string]$Title
    [string]$Description
    [datetime]$DefinedDate
    [datetime]$ProjectPlanStartDate
    [datetime]$ProjectPlanOverallDeadline
    [string]$View = "month"
}

# --- 主要関数 ---
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
        [Parameter(Mandatory)]
        [string]$FilePath,
        [Parameter(Mandatory)]
        [string]$Encoding
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
                            "projectoveralldeadline" { try { $projectMetadata.ProjectPlanOverallDeadline = [datetime]$value } catch {} } # $parsedYaml を $value に修正
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

    while ($lineIndex -lt $allLines.Count) {
        $line = $allLines[$lineIndex]; $trimmedLineCheck = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmedLineCheck)) {
            if ($currentWbsElement -and $attributeLinesBuffer.Count -gt 0) { $currentWbsElement.RawAttributeLines = $attributeLinesBuffer.ToArray(); $currentWbsElement.ParseCollectedAttributes(); $attributeLinesBuffer.Clear() }
            if ($currentWbsElement -and $descriptionBuffer.Length -gt 0) { $currentWbsElement.DescriptionText = $descriptionBuffer.ToString().Trim(); [void]$descriptionBuffer.Clear() }
            $lineIndex++; continue
        }
        if ($trimmedLineCheck -match "^(#+)\s+([\d\.]*[\d])\.?\s*(.+)") {
            if ($currentWbsElement -and $attributeLinesBuffer.Count -gt 0) { $currentWbsElement.RawAttributeLines = $attributeLinesBuffer.ToArray(); $currentWbsElement.ParseCollectedAttributes(); $attributeLinesBuffer.Clear() }
            if ($currentWbsElement -and $descriptionBuffer.Length -gt 0) { $currentWbsElement.DescriptionText = $descriptionBuffer.ToString().Trim(); [void]$descriptionBuffer.Clear() }
            $currentWbsElement = $null
            $level = $matches[1].Length; $id = $matches[2]; $name = $matches[3].Trim()
            $newElement = New-Object WbsElementNode -ArgumentList $id, $name, $level
            if ($level -eq 2) { $newElement.ElementType = "Section" } elseif ($level -eq 3) { $newElement.ElementType = "Group" } elseif ($level -ge 4) { $newElement.ElementType = "Task" }
            while ($parentStack.Count -gt 0 -and ($parentStack.Peek()).HierarchyLevel -ge $level) { [void]$parentStack.Pop() }
            if ($parentStack.Count -gt 0) { ($parentStack.Peek()).Children.Add($newElement); $newElement.Parent = $parentStack.Peek() } else { $rootElements.Add($newElement) }
            $parentStack.Push($newElement); $currentWbsElement = $newElement
        }
        elseif ($currentWbsElement -and $line -match "^\s{4}-\s*(?:deadline|duration|status|progress|assignee|depends):\s*.+") {
            if ($descriptionBuffer.Length -gt 0) { $currentWbsElement.DescriptionText = $descriptionBuffer.ToString().Trim(); [void]$descriptionBuffer.Clear() }
            $attributeLinesBuffer.Add($line)
        }
        elseif ($currentWbsElement -and ($attributeLinesBuffer.Count -eq 0) -and (-not ($line -match "^\s{0,3}-")) -and (-not ($line -match "^#+"))) {
             [void]$descriptionBuffer.AppendLine($line.Trim())
        }
        else {
            # Write-Verbose "未分類の行 (または前の要素の終端): $line" # 詳細すぎるためコメントアウト
            if ($currentWbsElement -and $attributeLinesBuffer.Count -gt 0) { $currentWbsElement.RawAttributeLines = $attributeLinesBuffer.ToArray(); $currentWbsElement.ParseCollectedAttributes(); $attributeLinesBuffer.Clear() }
            if ($currentWbsElement -and $descriptionBuffer.Length -gt 0) { $currentWbsElement.DescriptionText = $descriptionBuffer.ToString().Trim(); [void]$descriptionBuffer.Clear() }
        }
        $lineIndex++
    }
    if ($currentWbsElement -and $attributeLinesBuffer.Count -gt 0) { $currentWbsElement.RawAttributeLines = $attributeLinesBuffer.ToArray(); $currentWbsElement.ParseCollectedAttributes() }
    if ($currentWbsElement -and $descriptionBuffer.Length -gt 0) { $currentWbsElement.DescriptionText = $descriptionBuffer.ToString().Trim() }
    Write-Verbose "Markdown本文解析完了。トップレベル要素数: $($rootElements.Count)"
    return @{ ProjectMetadata = $projectMetadata; RootElements = $rootElements }
}

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

function Calculate-TaskEndDate {
    param( [Parameter(Mandatory)] [datetime]$StartDate, [Parameter(Mandatory)] [int]$DurationBusinessDays, [Parameter(Mandatory)] [System.Collections.Generic.List[datetime]]$Holidays )
    if ($DurationBusinessDays -le 0) { return $StartDate }

    $calculatedEndDate = $StartDate
    $businessDaysCounted = 0

    if ($StartDate.DayOfWeek -ne [DayOfWeek]::Saturday -and $StartDate.DayOfWeek -ne [DayOfWeek]::Sunday -and (-not $Holidays.Contains($StartDate.Date))) {
        $businessDaysCounted = 1
    }

    while ($businessDaysCounted -lt $DurationBusinessDays) { # `<`は`-lt`に修正。修正日時: 2025/05/16 04:04
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
                # Write-Host "DEBUG Step1: ID=$($Element.Id), sDate='$currentSDateStr', eDate='$currentEDateStr', Name='$currentTaskName'" -ForegroundColor Yellow # コメントアウト
                if ($Element.Attributes.ContainsKey("progress") -and $Element.Attributes["progress"] -match "^\d{1,3}\%?$") {
                    $currentTaskName += " ($($Element.Attributes['progress'] -replace '%',''))%"
                }
                # Write-Host "DEBUG Step2: Name with Progress='$currentTaskName'" -ForegroundColor Yellow # コメントアウト
                $currentTags = New-Object System.Collections.Generic.List[string]
                $currentTags.Add("#$($Element.Id.Replace('.','_'))")
                if ($Element.Attributes.ContainsKey("status")) { $currentTags.Add("#$($Element.Attributes['status'])") }
                if ($Element.Attributes.ContainsKey("assignee")) { $currentTags.Add("#assignee-$($Element.Attributes['assignee'] -replace ' ','_')") }
                $tagsString = ""
                if ($currentTags.Count -gt 0) { $tagsString = " $([string]::Join(' ', $currentTags))" }
                # Write-Host "DEBUG Step3: TagsString='$tagsString'" -ForegroundColor Yellow # コメントアウト
                $finalTaskLine = "${currentIndent}${currentSDateStr}/${currentEDateStr}: ${currentTaskName}${tagsString}"
                # Write-Host "DEBUG Step4: FinalTaskLine='$finalTaskLine'" -ForegroundColor Green # コメントアウト
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
    # Write-Host "DEBUG Format-GanttMarkwhen: Final outputLines content (first 20 lines):" -ForegroundColor Green # コメントアウト
    # $outputLines | Select-Object -First 20 | ForEach-Object { Write-Host $_ -ForegroundColor Green } # コメントアウト
    Write-Verbose "Markwhenデータ生成完了 (新仕様)。"
    $result = [string]($outputLines -join [System.Environment]::NewLine)
    return $result
}

# --- メイン処理 ---
try {
    Write-Host "処理を開始します (新MD-WBS仕様)..." -ForegroundColor Green

    $Global:Holidays = Import-HolidayList -InputOfficialHolidayFilePath $OfficialHolidayFilePath -InputCompanyHolidayFilePath $CompanyHolidayFilePath -BaseEncoding $DefaultEncoding
    $parsedResult = Parse-WbsMarkdownAndMetadataAdvanced -FilePath $WbsFilePath -Encoding $DefaultEncoding
    if (-not $parsedResult) { Write-Error "MD-WBSファイル解析失敗。"; exit 1 }
    $projectMetadata = $parsedResult.ProjectMetadata
    $rootWbsElements = $parsedResult.RootElements
    if ($rootWbsElements.Count -eq 0 -and (-not $projectMetadata -or ([string]::IsNullOrEmpty($projectMetadata.Title)) )) { Write-Error "MD-WBS有効データ無。"; exit 1 }

    $Global:AllTaskNodesFlatList = New-Object System.Collections.Generic.List[WbsElementNode]
    function Flatten-WbsTreeNodes {
        param ([System.Collections.Generic.List[WbsElementNode]]$Elements)
        if ($null -eq $Elements) { return }
        foreach ($element in $Elements) {
            if ($element.DurationDays -gt 0 -and $element.Deadline -ne ([datetime]::MinValue)) {
                 $Global:AllTaskNodesFlatList.Add($element)
            }
            if ($element.Children.Count -gt 0) { Flatten-WbsTreeNodes -Elements $element.Children }
        }
    }
    Flatten-WbsTreeNodes -Elements $rootWbsElements

    Write-Verbose "各タスクの日付を算出中..."
    if ($null -ne $Global:AllTaskNodesFlatList) {
        foreach ($taskNode in $Global:AllTaskNodesFlatList) {
            $taskNode.CalculatedStartDate = Calculate-TaskStartDate -DeadlineDate $taskNode.Deadline -DurationBusinessDays $taskNode.DurationDays -Holidays $Global:Holidays
            $taskNode.CalculatedEndDate = Calculate-TaskEndDate -StartDate $taskNode.CalculatedStartDate -DurationBusinessDays $taskNode.DurationDays -Holidays $Global:Holidays
            Write-Verbose "タスク '$($taskNode.Id) $($taskNode.Name)' 開始: $($taskNode.CalculatedStartDate.ToString($Global:DateFormatPattern)), 終了: $($taskNode.CalculatedEndDate.ToString($Global:DateFormatPattern)) (期間: $($taskNode.DurationDays)d)"
        }
    }

    $projectOverallStartDateForHolidayFilter = $projectMetadata.ProjectPlanStartDate
    $projectOverallEndDateForHolidayFilter = $projectMetadata.ProjectPlanOverallDeadline
    if (($null -ne $Global:AllTaskNodesFlatList) -and $Global:AllTaskNodesFlatList.Count -gt 0) {
        $validTasksForDateRange = $Global:AllTaskNodesFlatList | Where-Object {$_.CalculatedStartDate -ne ([datetime]::MinValue) -and $_.CalculatedEndDate -ne ([datetime]::MinValue)}
        if ($validTasksForDateRange.Count -gt 0) {
            if (($projectOverallStartDateForHolidayFilter -eq $null -or $projectOverallStartDateForHolidayFilter -eq ([datetime]::MinValue))) {
                 $projectOverallStartDateForHolidayFilter = ($validTasksForDateRange | Sort-Object CalculatedStartDate | Select-Object -First 1).CalculatedStartDate
            }
            if (($projectOverallEndDateForHolidayFilter -eq $null -or $projectOverallEndDateForHolidayFilter -eq ([datetime]::MinValue))) {
                $projectOverallEndDateForHolidayFilter = ($validTasksForDateRange | Sort-Object CalculatedEndDate -Descending | Select-Object -First 1).CalculatedEndDate
            }
        }
    }
    if($projectOverallStartDateForHolidayFilter -and $projectOverallEndDateForHolidayFilter) {
        Write-Verbose "祝日フィルタ用プロジェクト期間: $($projectOverallStartDateForHolidayFilter.ToString($Global:DateFormatPattern)) - $($projectOverallEndDateForHolidayFilter.ToString($Global:DateFormatPattern))"
    }

    Write-Verbose "依存関係のチェック中..."
    $dependencyWarnings = New-Object System.Collections.Generic.List[string]
    if ($null -ne $Global:AllTaskNodesFlatList) {
        foreach ($taskNode in $Global:AllTaskNodesFlatList) {
            if ($taskNode.Attributes.ContainsKey("depends")) {
                $dependencyIds = ($taskNode.Attributes["depends"] -split ',').Trim()
                foreach ($dependencyId in $dependencyIds) {
                    if ([string]::IsNullOrWhiteSpace($dependencyId)) { continue }
                    $parentTaskNode = $Global:AllTaskNodesFlatList | Where-Object { $_.Id -eq $dependencyId } | Select-Object -First 1
                    if ($parentTaskNode) {
                        if ($parentTaskNode.CalculatedEndDate -ne ([datetime]::MinValue) -and $taskNode.CalculatedStartDate -ne ([datetime]::MinValue) -and $parentTaskNode.CalculatedEndDate -ge $taskNode.CalculatedStartDate) {
                            # 警告メッセージをより分かりやすく
                            $warningMsg = "依存関係警告: タスク `"$($taskNode.Id) $($taskNode.Name)`" (計画開始: $($taskNode.CalculatedStartDate.ToString($Global:DateFormatPattern))) が、" +
                                          "先行タスク `"$($parentTaskNode.Id) $($parentTaskNode.Name)`" (計画終了: $($parentTaskNode.CalculatedEndDate.ToString($Global:DateFormatPattern))) の" +
                                          "終了日以前に開始する計画になっています。"
                            $dependencyWarnings.Add($warningMsg); Write-Warning $warningMsg
                        }
                    } else { Write-Warning "依存関係エラー: タスク '$($taskNode.Id) $($taskNode.Name)' の先行タスクID '$dependencyId' がWBS内タスクIDに見つかりません。" }
                }
            }
        }
    }

    if ($OutputFormat -eq "Markwhen") {
        $ganttChartContent = Format-GanttMarkwhen -Metadata $projectMetadata `
                                                -RootWbsElements $rootWbsElements `
                                                -DatePattern $DateFormatPattern `
                                                -Holidays $Global:Holidays `
                                                -ProjectOverallStartDateForHolidayFilter $projectOverallStartDateForHolidayFilter `
                                                -ProjectOverallEndDateForHolidayFilter $projectOverallEndDateForHolidayFilter `
                                                -HolidayDisplayMode $MarkwhenHolidayDisplayMode
    } else { Write-Error "出力フォーマット '$OutputFormat' はMarkwhenのみサポート"; exit 1 }

    try {
        # ファイル書き出しをUTF8NoBOMに変更
        Set-Content -Path $OutputFilePath -Value $ganttChartContent -Encoding UTF8NoBOM -Force
        Write-Host "ガントチャートデータが正常に出力: $OutputFilePath (Encoding: UTF8NoBOM)" -ForegroundColor Green
    } catch { Write-Error "出力ファイル書込失敗: $OutputFilePath. Error: $($_.Exception.Message)"; exit 1 }

    if ($dependencyWarnings.Count -gt 0) {
        Write-Warning "$($dependencyWarnings.Count) 件の依存関係に関する警告があります。出力ファイルと合わせてMD-WBSの計画を見直してください。"
    }
    Write-Host "処理正常完了。" -ForegroundColor Green

} catch {
    Write-Host "Debug: CRITICAL ERROR - $($_.Exception.GetType().FullName) - $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Debug: StackTrace - $($_.Exception.StackTrace)" -ForegroundColor DarkYellow
    if ($_.InvocationInfo) { Write-Host "Debug: Error at Line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line)" -ForegroundColor Yellow }
    Write-Error "スクリプト実行中致命的エラー: $($_.Exception.Message)"
    exit 1
} finally {
    Remove-Variable -Name Holidays, AllTaskNodesFlatList, projectMetadata, rootWbsElements, parsedResult, DateFormatPattern -ErrorAction SilentlyContinue
}