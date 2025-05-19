param (
    [Parameter(Mandatory = $true)]
    [string]$WbsFilePath,
    [Parameter(Mandatory = $true)]
    [string]$OfficialHolidayFilePath,
    [Parameter(Mandatory = $false)]
    [string]$CompanyHolidayFilePath,
    [Parameter(Mandatory = $true)]
    [string]$OutputFilePath,
    [Parameter(Mandatory = $true)]
    [string]$OutputFormat,
    [Parameter(Mandatory = $false)]
    [string]$DateFormatPattern = "yyyy-MM-dd",
    [Parameter(Mandatory = $false)]
    [string]$DefaultEncoding = "UTF8",
    [Parameter(Mandatory = $false)]
    [string]$MarkwhenHolidayDisplayMode = "All"
)

Write-Host ">>> 実行中のスクリプト: $($MyInvocation.MyCommand.Path)"

class ProjectMetadata {
    [string]$Title
    [string]$Description
    [datetime]$DefinedDate
    [datetime]$ProjectPlanStartDate
    [datetime]$ProjectPlanOverallDeadline
    [string]$View = "month"
}

# Load holidays from CSV files with validation
$officialHolidays = @()
if (Test-Path $OfficialHolidayFilePath) {
    $officialHolidays = Import-Csv -Path $OfficialHolidayFilePath | ForEach-Object {
        if (-not [string]::IsNullOrWhiteSpace($_.Date)) {
            try {
                [datetime]::Parse($_.Date)
            } catch {
                Write-Warning "Invalid date format (official holiday): '$($_.Date)'"
                $null
            }
        }
    } | Where-Object { $_ -ne $null }
}

$companyHolidays = @()
if ($CompanyHolidayFilePath -and (Test-Path $CompanyHolidayFilePath)) {
    $companyHolidays = Import-Csv -Path $CompanyHolidayFilePath | ForEach-Object {
        if (-not [string]::IsNullOrWhiteSpace($_.Date)) {
            try {
                [datetime]::Parse($_.Date)
            } catch {
                Write-Warning "Invalid date format (company holiday): '$($_.Date)'"
                $null
            }
        }
    } | Where-Object { $_ -ne $null }
}

$allHolidays = $officialHolidays + $companyHolidays

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

# Validate and parse the WBS file
$taskNodes = Parse-WbsMarkdownAndMetadataAdvanced -FilePath $WbsFilePath -Encoding $DefaultEncoding

if (-not $taskNodes -or $taskNodes.Count -eq 0) {
    Write-Error "Failed to parse the WBS file or no valid tasks found."
    exit 1
}


# Add project metadata and headers to the Excel sheet
function Export-WbsToExcel {
    param (
        [Parameter(Mandatory)] [System.Collections.Generic.List[PSCustomObject]]$TaskNodes,
        [Parameter(Mandatory)] [string]$ExcelOutputPath,
        [Parameter(Mandatory)] [ProjectMetadata]$ProjectMetadata
    )

    # Excel COM object initialization (Windows only)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)

    # Header rows
    $sheet.Cells.Item(1, 1).Value2 = "案件名"
    $sheet.Cells.Item(1, 2).Value2 = $ProjectMetadata.Title
    $sheet.Cells.Item(1, 3).Value2 = "開始日："
    $sheet.Cells.Item(1, 4).Value2 = $ProjectMetadata.ProjectPlanStartDate.ToString("yyyy-MM-dd")

    $sheet.Cells.Item(2, 1).Value2 = "課題シート：内部用課題の残課題(優先順位 低)の検討 (検討タイミング振り分け、対応方針)"
    $sheet.Cells.Item(2, 3).Value2 = "基準日："
    $sheet.Cells.Item(2, 4).Value2 = $ProjectMetadata.DefinedDate.ToString("yyyy-MM-dd")

    $sheet.Cells.Item(3, 4).Value2 = "週表示："
    $sheet.Cells.Item(3, 5).Value2 = "-1"

    # Column headers
    $sheet.Cells.Item(5, 1).Value2 = "No"
    $sheet.Cells.Item(5, 2).Value2 = "大分類"
    $sheet.Cells.Item(5, 3).Value2 = "中分類"
    $sheet.Cells.Item(5, 4).Value2 = "タスク名称"
    $sheet.Cells.Item(5, 5).Value2 = "先行後続"
    $sheet.Cells.Item(5, 6).Value2 = "関連番号"
    $sheet.Cells.Item(5, 7).Value2 = "先行後続タスク名"
    $sheet.Cells.Item(5, 8).Value2 = "先行有無"
    $sheet.Cells.Item(5, 9).Value2 = "アクションプラン"
    $sheet.Cells.Item(5, 10).Value2 = "進捗日数"
    $sheet.Cells.Item(5, 11).Value2 = "作業遅延"
    $sheet.Cells.Item(5, 12).Value2 = "開始遅延"
    $sheet.Cells.Item(5, 13).Value2 = "遅延日数"
    $sheet.Cells.Item(5, 14).Value2 = "担当組織"
    $sheet.Cells.Item(5, 15).Value2 = "担当者"
    $sheet.Cells.Item(5, 16).Value2 = "表示"
    $sheet.Cells.Item(5, 17).Value2 = "最終更新"
    $sheet.Cells.Item(5, 18).Value2 = "開始入力"
    $sheet.Cells.Item(5, 19).Value2 = "終了入力"
    $sheet.Cells.Item(5, 20).Value2 = "日数入力"
    $sheet.Cells.Item(5, 21).Value2 = "開始計画"
    $sheet.Cells.Item(5, 22).Value2 = "終了計画"
    $sheet.Cells.Item(5, 23).Value2 = "日数計画"
    $sheet.Cells.Item(5, 24).Value2 = "進捗率"
    $sheet.Cells.Item(5, 25).Value2 = "開始実績"
    $sheet.Cells.Item(5, 26).Value2 = "修了実績"

    $row = 6
    foreach ($task in $TaskNodes) {
        if ($task.ElementType -eq "Task") {
            # Hierarchical classification
            $parent = $task.Parent
            $grandparent = $parent?.Parent

            if ($grandparent) { $sheet.Cells.Item($row, 2).Value2 = $grandparent.Name } # B列: 大分類
            if ($parent)      { $sheet.Cells.Item($row, 3).Value2 = $parent.Name }      # C列: 中分類
            $sheet.Cells.Item($row, 4).Value2 = $task.Name                               # D列: タスク名称

            # depends
            if ($task.Attributes["depends"]) {
                $sheet.Cells.Item($row, 5).Value2 = "○"                                  # E列: 先行後続
                $sheet.Cells.Item($row, 6).Value2 = $task.Attributes["depends"]          # F列: 関連番号
            }

            # assignee
            if ($task.Attributes["assignee"]) {
                $sheet.Cells.Item($row, 15).Value2 = $task.Attributes["assignee"]        # O列: 担当者
            }

            # deadline → 終了入力
            if ($task.Deadline -ne $null) {
                $sheet.Cells.Item($row, 19).Value2 = $task.Deadline.ToString("yyyy-MM-dd") # S列
            }

            # duration → 日数入力
            if ($task.DurationDays -gt 0) {
                $sheet.Cells.Item($row, 20).Value2 = $task.DurationDays                    # T列
            }

            # 開始日（逆算）→ 開始入力
            if ($task.CalculatedStartDate -ne $null) {
                $sheet.Cells.Item($row, 18).Value2 = $task.CalculatedStartDate.ToString("yyyy-MM-dd") # R列
            }

            # progress → 進捗率
            if ($task.Attributes["progress"]) {
                $progressValue = $task.Attributes["progress"] -replace "%", ""
                if ($progressValue -match "^[0-9]+$") {
                    $sheet.Cells.Item($row, 24).Value2 = [double]$progressValue / 100.0 # X列
                }
            }

            # status による処理
            if ($task.Attributes["status"]) {
                $status = $task.Attributes["status"].ToLower()
                if ($status -eq "progress" -or $status -eq "completed") {
                    $sheet.Cells.Item($row, 25).Value2 = $task.CalculatedStartDate.ToString("yyyy-MM-dd") # Y列
                }
                if ($status -eq "completed") {
                    $sheet.Cells.Item($row, 24).Value2 = 1.0 # X列: 完了なら100%
                }
            }

            $row++
        }
    }

    # Save and close
    $workbook.SaveAs($ExcelOutputPath)
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

function Export-WbsToExcel {
    param (
        [Parameter(Mandatory)] [System.Collections.Generic.List[PSCustomObject]]$TaskNodes,
        [Parameter(Mandatory)] [string]$ExcelOutputPath
    )

    # Excel COM object initialization (Windows only)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)

    # Header row (adjust as needed)
    $sheet.Cells.Item(1, 2).Value2 = "大分類"
    $sheet.Cells.Item(1, 3).Value2 = "中分類"
    $sheet.Cells.Item(1, 4).Value2 = "タスク名称"
    $sheet.Cells.Item(1, 5).Value2 = "先行後続"
    $sheet.Cells.Item(1, 6).Value2 = "関連番号"
    $sheet.Cells.Item(1, 15).Value2 = "担当者"
    $sheet.Cells.Item(1, 18).Value2 = "開始入力"
    $sheet.Cells.Item(1, 19).Value2 = "終了入力"
    $sheet.Cells.Item(1, 20).Value2 = "日数入力"
    $sheet.Cells.Item(1, 24).Value2 = "進捗率"
    $sheet.Cells.Item(1, 25).Value2 = "開始実績"

    $row = 2
    foreach ($task in $TaskNodes) {
        if ($task.ElementType -eq "Task") {
            # Hierarchical classification
            $parent = $task.Parent
            $grandparent = $parent?.Parent

            if ($grandparent) { $sheet.Cells.Item($row, 2).Value2 = $grandparent.Name } # B列: 大分類
            if ($parent)      { $sheet.Cells.Item($row, 3).Value2 = $parent.Name }      # C列: 中分類
            $sheet.Cells.Item($row, 4).Value2 = $task.Name                               # D列: タスク名称

            # depends
            if ($task.Attributes["depends"]) {
                $sheet.Cells.Item($row, 5).Value2 = "○"                                  # E列: 先行後続
                $sheet.Cells.Item($row, 6).Value2 = $task.Attributes["depends"]          # F列: 関連番号
            }

            # assignee
            if ($task.Attributes["assignee"]) {
                $sheet.Cells.Item($row, 15).Value2 = $task.Attributes["assignee"]        # O列: 担当者
            }

            # deadline → 終了入力
            if ($task.Deadline -ne $null) {
                $sheet.Cells.Item($row, 19).Value2 = $task.Deadline.ToString("yyyy-MM-dd") # S列
            }

            # duration → 日数入力
            if ($task.DurationDays -gt 0) {
                $sheet.Cells.Item($row, 20).Value2 = $task.DurationDays                    # T列
            }

            # 開始日（逆算）→ 開始入力
            if ($task.CalculatedStartDate -ne $null) {
                $sheet.Cells.Item($row, 18).Value2 = $task.CalculatedStartDate.ToString("yyyy-MM-dd") # R列
            }

            # progress → 進捗率
            if ($task.Attributes["progress"]) {
                $progressValue = $task.Attributes["progress"] -replace "%", ""
                if ($progressValue -match "^[0-9]+$") {
                    $sheet.Cells.Item($row, 24).Value2 = [double]$progressValue / 100.0 # X列
                }
            }

            # status による処理
            if ($task.Attributes["status"]) {
                $status = $task.Attributes["status"].ToLower()
                if ($status -eq "progress" -or $status -eq "completed") {
                    $sheet.Cells.Item($row, 25).Value2 = $task.CalculatedStartDate.ToString("yyyy-MM-dd") # Y列
                }
                if ($status -eq "completed") {
                    $sheet.Cells.Item($row, 24).Value2 = 1.0 # X列: 完了なら100%
                }
            }

            $row++
        }
    }

    # Save and close
    $workbook.SaveAs($ExcelOutputPath)
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# WBSファイルの解析
$parsedResult = Parse-WbsMarkdownAndMetadataAdvanced -FilePath $WbsFilePath -Encoding $DefaultEncoding
if (-not $parsedResult -or -not $parsedResult.RootElements) {
    Write-Error "WBSファイルの解析に失敗しました。"
    exit 1
}
$taskNodes = $parsedResult.RootElements

# Main script execution
if ($OutputFormat -eq "Excel") {
    Export-WbsToExcel -TaskNodes $taskNodes -ExcelOutputPath $OutputFilePath
    Write-Host "Excelファイルに出力しました: $OutputFilePath" -ForegroundColor Green
    return
}
