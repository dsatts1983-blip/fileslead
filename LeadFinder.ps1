# WebsitesDoIt Lead Finder — v0.5.0
# Local-first Companies House lead finder for Windows PowerShell 5.1.
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:DataRoot = Join-Path $env:LOCALAPPDATA 'WebsitesDoIt\LeadFinder'
$script:CompanyFile = Join-Path $script:DataRoot 'companies.csv'
$script:LegacyCompanyFile = Join-Path $script:DataRoot 'companies.jsonl'
$script:LeadFile = Join-Path $script:DataRoot 'verified-leads.jsonl'
$script:SearchRoot = Join-Path $script:DataRoot 'searches'
$script:CurrentSearchId = ''
$script:BatchMode = $false
$script:BatchSuccess = 0
$script:BatchFailure = 0
$script:LoadingSearchRows = $false
$script:Importing = $false
$script:CancelImport = $false
$script:CloseAfterImport = $false
$script:CompaniesCount = 0
$script:ResultPage = 0
$script:PageSize = 100
$script:ResultTotal = 0
$script:ResultOffsets = New-Object 'System.Collections.Generic.List[long]'
$script:ResultFile = ''
$script:ResultFormat = 'csv'
$script:SourceMode = 'source'
$script:SourceFile = ''
$script:ResultIndexFile = ''
$script:SourceHeaderMap = @{}
$script:SicCategories = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$script:SicCategoryLookup = @{}
$script:CancelSearch = $false
$script:Searching = $false
$script:CloseAfterSearch = $false
$script:CurrentCompanies = New-Object 'System.Collections.Generic.List[object]'
$script:CurrentLeads = New-Object 'System.Collections.Generic.List[object]'
New-Item -ItemType Directory -Path $script:DataRoot -Force | Out-Null
New-Item -ItemType Directory -Path $script:SearchRoot -Force | Out-Null

function Parse-CsvLine([string]$Line) {
    $items = New-Object 'System.Collections.Generic.List[string]'
    $b = New-Object System.Text.StringBuilder
    $quoted = $false
    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]
        if ($ch -eq '"') {
            if ($quoted -and $i + 1 -lt $Line.Length -and $Line[$i + 1] -eq '"') { [void]$b.Append('"'); $i++ }
            else { $quoted = -not $quoted }
        } elseif ($ch -eq ',' -and -not $quoted) {
            $items.Add($b.ToString()); [void]$b.Clear()
        } else { [void]$b.Append($ch) }
    }
    $items.Add($b.ToString())
    return ,$items.ToArray()
}

function Get-Prop($Object, [string]$Name) {
    $p = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $Name } | Select-Object -First 1
    if ($p) { return [string]$p.Value }
    return ''
}

function Read-JsonLines([string]$Path) {
    $items = New-Object 'System.Collections.Generic.List[object]'
    if (Test-Path -LiteralPath $Path) {
        foreach ($line in [IO.File]::ReadLines($Path)) {
            if ($line) { try { $items.Add(($line | ConvertFrom-Json)) } catch { } }
        }
    }
    return ,$items
}

function Save-Leads {
    $utf8 = [System.Text.UTF8Encoding]::new($false)
    $target = if ($script:CurrentSearchId) { Join-Path $script:SearchRoot ($script:CurrentSearchId + '-leads.jsonl') } else { $script:LeadFile }
    $writer = [IO.StreamWriter]::new($target, $false, $utf8)
    try { foreach ($lead in $script:CurrentLeads) { $writer.WriteLine((ConvertTo-Json -InputObject $lead -Compress -Depth 5)) } }
    finally { $writer.Dispose() }
}

function Normalize-CompanyName([string]$Name) {
    $v = $Name.ToLowerInvariant() -replace '&', ' and '
    $v = $v -replace '\b(limited|ltd\.?)(?=\s|$)', ' '
    $v = $v -replace '[^a-z0-9]+', ' '
    return ($v -replace '\s+', ' ').Trim()
}

function Get-CompanyAddress($Company) {
    $bits = @((Get-Prop $Company 'RegAddress.AddressLine1'), (Get-Prop $Company 'RegAddress.AddressLine2'), (Get-Prop $Company 'RegAddress.PostTown'), (Get-Prop $Company 'RegAddress.County'), (Get-Prop $Company 'RegAddress.PostCode')) | Where-Object { $_ }
    return ($bits -join ', ')
}

function Get-Sic($Company) {
    $codes = @()
    foreach ($p in $Company.PSObject.Properties) { if ($p.Name -match '^SICCode\.SicText_\d+$' -and $p.Value) { $codes += [string]$p.Value } }
    return ($codes -join '; ')
}

function New-CompanyRecord($Fields, $Map) {
    $record = [ordered]@{}
    foreach ($key in $Map.Keys) { $index = $Map[$key]; if ($index -ge 0 -and $index -lt $Fields.Count) { $record[$key] = [string]$Fields[$index] } }
    return [pscustomobject]$record
}

function Show-Status([string]$Text) { $script:statusLabel.Text = $Text; [System.Windows.Forms.Application]::DoEvents() }

function Refresh-LeadGrid {
    $target = if ($script:CurrentSearchId) { Join-Path $script:SearchRoot ($script:CurrentSearchId + '-leads.jsonl') } else { $script:LeadFile }
    $script:CurrentLeads = Read-JsonLines $target
    $script:leadGrid.Rows.Clear()
    foreach ($lead in $script:CurrentLeads) {
        [void]$script:leadGrid.Rows.Add($lead.CompanyName, $lead.CompanyNumber, $lead.Town, $lead.Website, $lead.Email, $lead.Telephone, $lead.Verification)
    }
    $script:leadCountLabel.Text = "$($script:CurrentLeads.Count) verified leads"
}

function Refresh-SearchRuns {
    $script:runPicker.Items.Clear()
    if (Test-Path -LiteralPath $script:SearchRoot) {
        $runs = foreach ($path in (Get-ChildItem -LiteralPath $script:SearchRoot -Filter '*.json' -File)) {
            try { Get-Content -LiteralPath $path.FullName -Raw | ConvertFrom-Json } catch { }
        }
        foreach ($run in ($runs | Sort-Object CreatedAt -Descending)) {
            $item = [pscustomobject]@{ Id=$run.Id; Label="$($run.Name) — $($run.ResultCount) companies — $($run.CreatedAt)" }
            [void]$script:runPicker.Items.Add($item)
            if (-not $script:CurrentSearchId) { $script:CurrentSearchId = $run.Id }
        }
    }
    if ($script:CurrentSearchId) {
        for ($i = 0; $i -lt $script:runPicker.Items.Count; $i++) { if ($script:runPicker.Items[$i].Id -eq $script:CurrentSearchId) { $script:runPicker.SelectedIndex = $i; break } }
    }
}

function Read-SourceHeader([string]$DataFile) {
    $script:SourceHeaderMap = @{}
    if ([IO.Path]::GetExtension($DataFile) -ieq '.csv') {
        $reader = [IO.StreamReader]::new($DataFile, [Text.Encoding]::UTF8, $true)
        try {
            $header = Parse-CsvLine $reader.ReadLine()
            for ($i=0; $i -lt $header.Count; $i++) { $script:SourceHeaderMap[[string]$header[$i].Trim()] = $i }
        } finally { $reader.Dispose() }
    }
}

function Get-SourceLineEndingWidth([string]$DataFile) {
    $stream=[IO.File]::OpenRead($DataFile); $buffer=New-Object byte[] 65536
    try { $read=$stream.Read($buffer,0,$buffer.Length); for($i=0;$i -lt $read;$i++){if($buffer[$i] -eq 10){if($i -gt 0 -and $buffer[$i-1] -eq 13){return 2}else{return 1}}} } finally { $stream.Dispose() }
    return 2
}

function Test-SourceUtf8Bom([string]$DataFile) {
    $stream=[IO.File]::OpenRead($DataFile); $buffer=New-Object byte[] 3
    try { return ($stream.Read($buffer,0,3) -eq 3 -and $buffer[0] -eq 239 -and $buffer[1] -eq 187 -and $buffer[2] -eq 191) } finally { $stream.Dispose() }
}

function Repoint-SavedSearchSources([string]$OldPath, [string]$NewPath) {
    try {
        $oldFull=[IO.Path]::GetFullPath($OldPath)
        foreach($metaPath in (Get-ChildItem -LiteralPath $script:SearchRoot -Filter '*.json' -File)) {
            try {
                $meta=Get-Content -LiteralPath $metaPath.FullName -Raw|ConvertFrom-Json
                if($meta.SourceMode -eq 'source' -and $meta.SourceFile -and [string]::Equals([IO.Path]::GetFullPath([string]$meta.SourceFile),$oldFull,[StringComparison]::OrdinalIgnoreCase)) {
                    $meta.SourceFile=$NewPath
                    [IO.File]::WriteAllText($metaPath.FullName,(ConvertTo-Json -InputObject $meta -Depth 6),[Text.UTF8Encoding]::new($false))
                }
            } catch { }
        }
    } catch { }
}

function Save-SicCategoryCache([string]$DataFile) {
    if ([IO.Path]::GetExtension($DataFile) -ine '.csv') { return }
    $info = Get-Item -LiteralPath $DataFile
    $cache = [pscustomobject]@{ SourceLength=[long]$info.Length; SourceWriteTicks=[long]$info.LastWriteTimeUtc.Ticks; Values=@($script:SicCategories | Sort-Object) }
    $path = Join-Path $script:DataRoot 'sic-categories.json'
    [IO.File]::WriteAllText($path, (ConvertTo-Json -InputObject $cache -Depth 4), [Text.UTF8Encoding]::new($false))
}

function Get-SicDisplayLabel([string]$Value) {
    if ($Value -match '^\s*(\d+)\s*-\s*(.*)$') { return "$($Matches[2].Trim()) ($($Matches[1]))" }
    return $Value
}

function Set-SicCategoryChoices([string]$DataFile, [bool]$UseCache, [bool]$PreserveCurrent=$false) {
    $previousText=[string]$sicBox.Text
    if (-not $PreserveCurrent) { $script:SicCategories.Clear() }
    if ($UseCache -and [IO.Path]::GetExtension($DataFile) -ieq '.csv') {
        $cachePath = Join-Path $script:DataRoot 'sic-categories.json'
        if (Test-Path -LiteralPath $cachePath) {
            try {
                $cache = Get-Content -LiteralPath $cachePath -Raw | ConvertFrom-Json
                $info = Get-Item -LiteralPath $DataFile
                if ([long]$cache.SourceLength -eq [long]$info.Length -and [long]$cache.SourceWriteTicks -eq [long]$info.LastWriteTimeUtc.Ticks) {
                    foreach ($value in $cache.Values) { if ($value) { [void]$script:SicCategories.Add([string]$value) } }
                }
            } catch { $script:SicCategories.Clear() }
        }
    }
    if ($sicBox.Items.Count -gt 0) { $sicBox.Items.Clear() }
    $script:SicCategoryLookup=@{}
    [void]$sicBox.Items.Add('All categories')
    foreach ($value in ($script:SicCategories | Sort-Object)) { $label=Get-SicDisplayLabel ([string]$value); $script:SicCategoryLookup[$label]=[string]$value; [void]$sicBox.Items.Add($label) }
    if ($sicBox.Items.Count -gt 1) {
        if($PreserveCurrent -and $previousText){$sicBox.Text=$previousText}else{$sicBox.SelectedIndex=0}
        $sicBox.Enabled = $true
        $categoryInfoLabel.Text = "$($sicBox.Items.Count - 1) SIC categories loaded — type a category name or code"
    } else {
        if($PreserveCurrent -and $previousText){$sicBox.Text=$previousText}else{$sicBox.SelectedIndex=0}
        $sicBox.Enabled = $true
        $categoryInfoLabel.Text = 'Search once to load categories, then type a name or SIC code here'
    }
}

function Load-SearchCompanies([string]$RunId) {
    $metaPath = Join-Path $script:SearchRoot ($RunId + '.json')
    $meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json
    $script:CurrentSearchId=$RunId; $script:ResultTotal=[int]$meta.ResultCount; $script:ResultOffsets=[System.Collections.Generic.List[long]]::new()
    if ($meta.ResultOffsets) { foreach($offset in $meta.ResultOffsets){$script:ResultOffsets.Add([long]$offset)} }
    if ($meta.HeaderMap) { $script:SourceHeaderMap=@{}; foreach($property in $meta.HeaderMap.PSObject.Properties){$script:SourceHeaderMap[$property.Name]=[int]$property.Value} }
    $script:ResultFormat = if ($meta.Format) { [string]$meta.Format } else { 'jsonl' }
    $script:SourceMode = if ($meta.SourceMode) { [string]$meta.SourceMode } else { 'spool' }
    $script:SourceFile = [string]$meta.SourceFile
    $script:ResultIndexFile = [string]$meta.ResultIndexFile
    $extension = if ($script:ResultFormat -eq 'csv') { '-results.csv' } else { '-results.jsonl' }
    $script:ResultFile=Join-Path $script:SearchRoot ($RunId+$extension)
    if ($script:SourceMode -eq 'source') { $script:ResultIndexFile=Join-Path $script:SearchRoot ($RunId+'-offsets.bin') }

    # Upgrade older saved searches (which kept at most 1,000 JSON rows) into the page-based format.
    if (-not (Test-Path -LiteralPath $script:ResultFile)) {
        $oldPath = Join-Path $script:SearchRoot ($RunId+'-companies.jsonl')
        if (Test-Path -LiteralPath $oldPath) {
            $script:ResultFormat='jsonl'; $script:SourceMode='spool'; $script:ResultFile=Join-Path $script:SearchRoot ($RunId+'-results.jsonl')
            $writer=[IO.StreamWriter]::new($script:ResultFile,$false,[Text.UTF8Encoding]::new($false)); $stream=$writer.BaseStream; $stateWriter=$null
            $script:ResultOffsets=[System.Collections.Generic.List[long]]::new(); $count=0; $page=-1
            try {
                foreach($savedLine in [IO.File]::ReadLines($oldPath)) {
                    if(-not $savedLine){continue}; $saved=$savedLine|ConvertFrom-Json
                    if($count % $script:PageSize -eq 0){if($stateWriter){$stateWriter.Dispose()};$writer.Flush();$script:ResultOffsets.Add([long]$stream.Position);$page++;$statePath=Join-Path $script:SearchRoot ($RunId+'-state-page'+$page+'.jsonl');$stateWriter=[IO.StreamWriter]::new($statePath,$false,[Text.UTF8Encoding]::new($false))}
                    $writer.WriteLine((ConvertTo-Json -InputObject $saved.Company -Compress -Depth 5)); $stateWriter.WriteLine((ConvertTo-Json -InputObject ([pscustomobject]@{Pick=[bool]$saved.Pick;CandidateWebsite=[string]$saved.CandidateWebsite}) -Compress -Depth 3)); $count++
                }
            } finally { $writer.Dispose(); if($stateWriter){$stateWriter.Dispose()} }
            $script:ResultTotal=$count; $meta.ResultCount=$count; $meta.Format='jsonl'; $meta.SourceMode='spool'; $meta.SourceFile=$script:ResultFile; $meta.ResultOffsets=@($script:ResultOffsets.ToArray())
            [IO.File]::WriteAllText($metaPath,(ConvertTo-Json -InputObject $meta -Depth 6),[Text.UTF8Encoding]::new($false))
        }
    }
    if ($script:ResultOffsets.Count -eq 0 -and $script:ResultTotal -gt 0) { $script:ResultOffsets.Add(0) }
    $script:ResultPage=0; Show-CompanyPage
}

function Show-CompanyPage {
    $script:LoadingSearchRows=$true
    try {
        $script:CurrentCompanies.Clear(); $companyGrid.Rows.Clear()
        $start=$script:ResultPage*$script:PageSize; $end=[Math]::Min($start+$script:PageSize,$script:ResultTotal)
        if($end -gt $start -and $script:SourceMode -eq 'source' -and (Test-Path -LiteralPath $script:ResultIndexFile) -and (Test-Path -LiteralPath $script:SourceFile)) {
            $indexStream=[IO.File]::OpenRead($script:ResultIndexFile); $indexReader=[IO.BinaryReader]::new($indexStream)
            $sourceStream=[IO.File]::OpenRead($script:SourceFile); $sourceReader=[IO.StreamReader]::new($sourceStream,[Text.Encoding]::UTF8,$true,65536,$true)
            try {
                [void]$indexStream.Seek([long]($start*8),[IO.SeekOrigin]::Begin)
                for($i=$start;$i -lt $end;$i++) {
                    $sourceOffset=$indexReader.ReadInt64(); $sourceReader.DiscardBufferedData(); [void]$sourceStream.Seek($sourceOffset,[IO.SeekOrigin]::Begin); $line=$sourceReader.ReadLine()
                    if($null -eq $line){break}
                    if($script:ResultFormat -eq 'csv'){$co=New-CompanyRecord (Parse-CsvLine $line) $script:SourceHeaderMap}else{$co=$line|ConvertFrom-Json}
                    $script:CurrentCompanies.Add($co)
                    [void]$companyGrid.Rows.Add($false,(Get-Prop $co 'CompanyName'),(Get-Prop $co 'CompanyNumber'),(Get-Prop $co 'RegAddress.PostTown'),(Get-Prop $co 'RegAddress.PostCode'),(Get-Prop $co 'CompanyStatus'),(Get-Sic $co),'')
                }
            } finally { $sourceReader.Dispose(); $sourceStream.Dispose(); $indexReader.Dispose(); $indexStream.Dispose() }
        } elseif($end -gt $start -and (Test-Path -LiteralPath $script:ResultFile)) {
            $fs=[IO.File]::OpenRead($script:ResultFile)
            try {
                [void]$fs.Seek($script:ResultOffsets[$script:ResultPage],[IO.SeekOrigin]::Begin)
                $reader=[IO.StreamReader]::new($fs,[Text.Encoding]::UTF8,$false,65536,$true)
                try {
                    for($i=$start;$i -lt $end;$i++) {
                        $line=$reader.ReadLine(); if($null -eq $line){break}
                        if($script:ResultFormat -eq 'csv') { $co=New-CompanyRecord (Parse-CsvLine $line) $script:SourceHeaderMap }
                        else { $co=$line | ConvertFrom-Json }
                        $script:CurrentCompanies.Add($co)
                        [void]$companyGrid.Rows.Add($false,(Get-Prop $co 'CompanyName'),(Get-Prop $co 'CompanyNumber'),(Get-Prop $co 'RegAddress.PostTown'),(Get-Prop $co 'RegAddress.PostCode'),(Get-Prop $co 'CompanyStatus'),(Get-Sic $co),'')
                    }
                } finally { $reader.Dispose() }
            } finally { $fs.Dispose() }
        }
        $statePath=Join-Path $script:SearchRoot ($script:CurrentSearchId+'-state-page'+$script:ResultPage+'.jsonl')
        if(Test-Path -LiteralPath $statePath) {
            $index=0
            foreach($line in [IO.File]::ReadLines($statePath)) {
                if($index -ge $companyGrid.Rows.Count){break}; if(-not $line){$index++;continue}
                try { $saved=$line|ConvertFrom-Json; $companyGrid.Rows[$index].Cells['Pick'].Value=[bool]$saved.Pick; $companyGrid.Rows[$index].Cells['Website'].Value=[string]$saved.CandidateWebsite } catch { }
                $index++
            }
        }
        $pages=[Math]::Max(1,[Math]::Ceiling($script:ResultTotal/[double]$script:PageSize));$first=if($script:ResultTotal -eq 0){0}else{$start+1};$last=[Math]::Min($end,$script:ResultTotal)
        $companyCountLabel.Text="Matching companies: $($script:ResultTotal.ToString('N0'))   |   Showing $first–$last   |   Page $($script:ResultPage+1) of $pages"
        $prevPageButton.Enabled=$script:ResultPage -gt 0;$nextPageButton.Enabled=($script:ResultPage+1) -lt $pages
    } finally { $script:LoadingSearchRows=$false }
}

function Save-SearchCompanies {
    if ($script:LoadingSearchRows -or -not $script:CurrentSearchId) { return }
    $path = Join-Path $script:SearchRoot ($script:CurrentSearchId + '-state-page' + $script:ResultPage + '.jsonl')
    $writer = [IO.StreamWriter]::new($path, $false, [Text.UTF8Encoding]::new($false))
    try {
        for ($i = 0; $i -lt $script:CurrentCompanies.Count -and $i -lt $companyGrid.Rows.Count; $i++) {
            $row = $companyGrid.Rows[$i]
            $saved = [pscustomobject]@{ Pick=[bool]$row.Cells['Pick'].Value; CandidateWebsite=[string]$row.Cells['Website'].Value }
            $writer.WriteLine((ConvertTo-Json -InputObject $saved -Compress -Depth 3))
        }
    } finally { $writer.Dispose() }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'WebsitesDoIt Lead Finder — v0.5.0'
$form.Size = [System.Drawing.Size]::new(1180, 760)
$form.MinimumSize = [System.Drawing.Size]::new(1000, 650)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.Font = [System.Drawing.Font]::new('Segoe UI', 9)

$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'; $header.Height = 84; $header.BackColor = [System.Drawing.Color]::FromArgb(24, 42, 72)
$form.Controls.Add($header)
$title = New-Object System.Windows.Forms.Label
$title.Text = 'WebsitesDoIt  |  Lead Finder'; $title.ForeColor = [System.Drawing.Color]::White
$title.Font = [System.Drawing.Font]::new('Segoe UI Semibold', 18); $title.Location = [System.Drawing.Point]::new(24, 13); $title.AutoSize = $true
$header.Controls.Add($title)
$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = 'Step 1: import your Companies House ZIP or CSV, then filter and search companies.'; $subtitle.ForeColor = [System.Drawing.Color]::FromArgb(201, 214, 232)
$subtitle.Location = [System.Drawing.Point]::new(27, 51); $subtitle.AutoSize = $true; $header.Controls.Add($subtitle)
$importButton = New-Object System.Windows.Forms.Button; $importButton.Text = 'Import Companies House ZIP / CSV'; $importButton.Location = [System.Drawing.Point]::new(740, 21); $importButton.Size = [System.Drawing.Size]::new(390, 42)
$importButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$importButton.BackColor = [System.Drawing.Color]::FromArgb(33, 105, 180); $importButton.ForeColor = [System.Drawing.Color]::White; $importButton.FlatStyle = 'Flat'; $header.Controls.Add($importButton)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'; $tabs.Padding = [System.Drawing.Point]::new(14, 7); $form.Controls.Add($tabs)
$searchTab = [System.Windows.Forms.TabPage]::new('Find companies'); $searchTab.BackColor = [System.Drawing.Color]::White
$leadsTab = [System.Windows.Forms.TabPage]::new('Verified leads'); $leadsTab.BackColor = [System.Drawing.Color]::White
[void]$tabs.TabPages.Add($searchTab); [void]$tabs.TabPages.Add($leadsTab)

$top = New-Object System.Windows.Forms.Panel; $top.Dock = 'Top'; $top.Height = 204; $top.Padding = [System.Windows.Forms.Padding]::new(18, 14, 18, 8); $searchTab.Controls.Add($top)
$importInfo = New-Object System.Windows.Forms.Label; $importInfo.Text = 'No company data imported yet. Use Import above to select your ZIP or CSV.'; $importInfo.Location = [System.Drawing.Point]::new(20, 23); $importInfo.AutoSize = $true; $top.Controls.Add($importInfo)

$labels = @(@('Company name contains', 20, 70), @('Town / city', 330, 70), @('Postcode', 550, 70), @('Category / SIC', 770, 70))
foreach ($it in $labels) { $l = New-Object System.Windows.Forms.Label; $l.Text = $it[0]; $l.Location = [System.Drawing.Point]::new($it[1], $it[2]); $l.AutoSize = $true; $top.Controls.Add($l) }
$nameBox = New-Object System.Windows.Forms.TextBox; $nameBox.Location = [System.Drawing.Point]::new(20, 93); $nameBox.Size = [System.Drawing.Size]::new(290, 26); $top.Controls.Add($nameBox)
$townBox = New-Object System.Windows.Forms.TextBox; $townBox.Location = [System.Drawing.Point]::new(330, 93); $townBox.Size = [System.Drawing.Size]::new(200, 26); $top.Controls.Add($townBox)
$postcodeBox = New-Object System.Windows.Forms.TextBox; $postcodeBox.Location = [System.Drawing.Point]::new(550, 93); $postcodeBox.Size = [System.Drawing.Size]::new(200, 26); $top.Controls.Add($postcodeBox)
$sicBox = New-Object System.Windows.Forms.ComboBox; $sicBox.DropDownStyle='DropDown'; $sicBox.AutoCompleteSource='ListItems'; $sicBox.AutoCompleteMode='SuggestAppend'; $sicBox.Location = [System.Drawing.Point]::new(770, 93); $sicBox.Size = [System.Drawing.Size]::new(330, 26); $sicBox.MaxDropDownItems=18; $sicBox.IntegralHeight=$false; [void]$sicBox.Items.Add('All categories'); $sicBox.SelectedIndex=0; $top.Controls.Add($sicBox)
$activeBox = New-Object System.Windows.Forms.CheckBox; $activeBox.Text = 'Active companies only'; $activeBox.Checked = $true; $activeBox.Location = [System.Drawing.Point]::new(20, 132); $activeBox.AutoSize = $true; $top.Controls.Add($activeBox)
$searchButton = New-Object System.Windows.Forms.Button; $searchButton.Text = 'Search imported companies'; $searchButton.Location = [System.Drawing.Point]::new(250, 125); $searchButton.Size = [System.Drawing.Size]::new(210, 34); $searchButton.BackColor = [System.Drawing.Color]::FromArgb(24, 42, 72); $searchButton.ForeColor = [System.Drawing.Color]::White; $searchButton.FlatStyle = 'Flat'; $top.Controls.Add($searchButton)
$cancelSearchButton = New-Object System.Windows.Forms.Button; $cancelSearchButton.Text='Stop search'; $cancelSearchButton.Location=[System.Drawing.Point]::new(470,125); $cancelSearchButton.Size=[System.Drawing.Size]::new(90,34); $cancelSearchButton.Enabled=$false; $top.Controls.Add($cancelSearchButton)
$companyCountLabel = New-Object System.Windows.Forms.Label; $companyCountLabel.Text='Run a search to see the matching total'; $companyCountLabel.Location = [System.Drawing.Point]::new(575, 134); $companyCountLabel.AutoSize = $true; $companyCountLabel.Font=[System.Drawing.Font]::new('Segoe UI Semibold',10); $top.Controls.Add($companyCountLabel)
$categoryInfoLabel = New-Object System.Windows.Forms.Label; $categoryInfoLabel.Location=[System.Drawing.Point]::new(20,169); $categoryInfoLabel.AutoSize=$true; $categoryInfoLabel.ForeColor=[System.Drawing.Color]::FromArgb(85,97,112); $top.Controls.Add($categoryInfoLabel)

$companyGrid = New-Object System.Windows.Forms.DataGridView; $companyGrid.Dock = 'Fill'; $companyGrid.ReadOnly = $false; $companyGrid.AllowUserToAddRows = $false; $companyGrid.SelectionMode = 'FullRowSelect'; $companyGrid.MultiSelect = $false; $companyGrid.AutoSizeColumnsMode = 'Fill'; $companyGrid.BackgroundColor = [System.Drawing.Color]::White; $companyGrid.ColumnHeadersVisible=$true; $companyGrid.ColumnHeadersHeight=34; $companyGrid.RowTemplate.Height=26; $companyGrid.EnableHeadersVisualStyles=$false; $companyGrid.ColumnHeadersDefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(232,238,246); $companyGrid.ColumnHeadersDefaultCellStyle.Font=[System.Drawing.Font]::new('Segoe UI Semibold',9)
$pickCol = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $pickCol.Name = 'Pick'; $pickCol.HeaderText = 'Check'; $pickCol.Width = 55; [void]$companyGrid.Columns.Add($pickCol)
foreach ($c in @(@('CompanyName','Company'),@('CompanyNumber','Company number'),@('Town','Town'),@('Postcode','Postcode'),@('Status','Status'),@('SIC','SIC codes'))) { $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col.Name = $c[0]; $col.HeaderText = $c[1]; $col.ReadOnly = $true; [void]$companyGrid.Columns.Add($col) }
$webCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $webCol.Name = 'Website'; $webCol.HeaderText = 'Candidate website (paste here)'; $webCol.ReadOnly = $false; [void]$companyGrid.Columns.Add($webCol)
$companyGrid.Add_CurrentCellDirtyStateChanged({ if ($companyGrid.IsCurrentCellDirty) { $companyGrid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } })
$companyGrid.Add_CellValueChanged({ Save-SearchCompanies })
$searchTab.Controls.Add($companyGrid)
$actions = New-Object System.Windows.Forms.Panel; $actions.Dock = 'Bottom'; $actions.Height = 154; $actions.Padding = [System.Windows.Forms.Padding]::new(18, 10, 18, 8); $searchTab.Controls.Add($actions)
$selectAllButton = New-Object System.Windows.Forms.Button; $selectAllButton.Text = 'Select all shown'; $selectAllButton.Location = [System.Drawing.Point]::new(20, 12); $selectAllButton.Size = [System.Drawing.Size]::new(150, 34); $actions.Controls.Add($selectAllButton)
$googleButton = New-Object System.Windows.Forms.Button; $googleButton.Text = 'Search Google for selected row'; $googleButton.Location = [System.Drawing.Point]::new(190, 12); $googleButton.Size = [System.Drawing.Size]::new(250, 34); $googleButton.BackColor = [System.Drawing.Color]::FromArgb(238, 243, 249); $googleButton.FlatStyle = 'Flat'; $actions.Controls.Add($googleButton)
$websiteLabel = New-Object System.Windows.Forms.Label; $websiteLabel.Text = 'Website for selected row'; $websiteLabel.Location = [System.Drawing.Point]::new(450, 21); $websiteLabel.AutoSize = $true; $actions.Controls.Add($websiteLabel)
$websiteBox = New-Object System.Windows.Forms.TextBox; $websiteBox.Location = [System.Drawing.Point]::new(590, 16); $websiteBox.Size = [System.Drawing.Size]::new(155, 26); $actions.Controls.Add($websiteBox)
$verifyButton = New-Object System.Windows.Forms.Button; $verifyButton.Text = 'Verify & collect contacts'; $verifyButton.Location = [System.Drawing.Point]::new(765, 12); $verifyButton.Size = [System.Drawing.Size]::new(220, 34); $verifyButton.BackColor = [System.Drawing.Color]::FromArgb(24, 130, 95); $verifyButton.ForeColor = [System.Drawing.Color]::White; $verifyButton.FlatStyle = 'Flat'; $actions.Controls.Add($verifyButton)
$batchButton = New-Object System.Windows.Forms.Button; $batchButton.Text = 'Run checks on checked rows'; $batchButton.Location = [System.Drawing.Point]::new(765, 55); $batchButton.Size = [System.Drawing.Size]::new(220, 34); $batchButton.BackColor = [System.Drawing.Color]::FromArgb(24, 130, 95); $batchButton.ForeColor = [System.Drawing.Color]::White; $batchButton.FlatStyle = 'Flat'; $actions.Controls.Add($batchButton)
$prevPageButton = New-Object System.Windows.Forms.Button; $prevPageButton.Text='‹ Previous'; $prevPageButton.Location=[System.Drawing.Point]::new(465,55); $prevPageButton.Size=[System.Drawing.Size]::new(120,34); $prevPageButton.Enabled=$false; $actions.Controls.Add($prevPageButton)
$nextPageButton = New-Object System.Windows.Forms.Button; $nextPageButton.Text='Next ›'; $nextPageButton.Location=[System.Drawing.Point]::new(595,55); $nextPageButton.Size=[System.Drawing.Size]::new(120,34); $nextPageButton.Enabled=$false; $actions.Controls.Add($nextPageButton)
$hint = New-Object System.Windows.Forms.Label; $hint.Text = 'Select all shown checks the 100 rows on this page. Add candidate websites, then run checks; Google searches open for manual review.'; $hint.Location = [System.Drawing.Point]::new(20, 102); $hint.AutoSize = $true; $hint.ForeColor = [System.Drawing.Color]::FromArgb(85, 97, 112); $actions.Controls.Add($hint)

$leadTop = New-Object System.Windows.Forms.Panel; $leadTop.Dock = 'Top'; $leadTop.Height = 62; $leadTop.Padding = [System.Windows.Forms.Padding]::new(18, 12, 18, 8); $leadsTab.Controls.Add($leadTop)
$script:leadCountLabel = New-Object System.Windows.Forms.Label; $script:leadCountLabel.Location = [System.Drawing.Point]::new(20, 18); $script:leadCountLabel.AutoSize = $true; $leadTop.Controls.Add($script:leadCountLabel)
$script:runPicker = New-Object System.Windows.Forms.ComboBox; $script:runPicker.DropDownStyle = 'DropDownList'; $script:runPicker.Location = [System.Drawing.Point]::new(250, 12); $script:runPicker.Size = [System.Drawing.Size]::new(400, 28); $script:runPicker.DisplayMember = 'Label'; $leadTop.Controls.Add($script:runPicker)
$exportButton = New-Object System.Windows.Forms.Button; $exportButton.Text = 'Export this search (4 fields)'; $exportButton.Location = [System.Drawing.Point]::new(850, 10); $exportButton.Size = [System.Drawing.Size]::new(230, 34); $exportButton.BackColor = [System.Drawing.Color]::FromArgb(33, 105, 180); $exportButton.ForeColor = [System.Drawing.Color]::White; $exportButton.FlatStyle = 'Flat'; $leadTop.Controls.Add($exportButton)
$script:leadGrid = New-Object System.Windows.Forms.DataGridView; $script:leadGrid.Dock = 'Fill'; $script:leadGrid.ReadOnly = $true; $script:leadGrid.AllowUserToAddRows = $false; $script:leadGrid.AutoSizeColumnsMode = 'Fill'; $script:leadGrid.BackgroundColor = [System.Drawing.Color]::White
foreach ($c in @(@('CompanyName','Company'),@('CompanyNumber','Company number'),@('Town','Town'),@('Website','Website'),@('Email','Generic email'),@('Telephone','Telephone'),@('Verification','Status'))) { $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col.Name = $c[0]; $col.HeaderText = $c[1]; [void]$script:leadGrid.Columns.Add($col) }
$leadsTab.Controls.Add($script:leadGrid)
$statusStrip = New-Object System.Windows.Forms.StatusStrip; $statusStrip.Dock = 'Bottom'; $script:statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $script:statusLabel.Text = 'Ready'; [void]$statusStrip.Items.Add($script:statusLabel); $form.Controls.Add($statusStrip)

if (Test-Path -LiteralPath $script:CompanyFile) { $importInfo.Text = 'Companies House data is stored locally and ready to search' }
Set-SicCategoryChoices $script:CompanyFile $true
Refresh-SearchRuns
Refresh-LeadGrid

$importButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog; $dialog.Title = 'Select Companies House ZIP or CSV'; $dialog.Filter = 'Companies House ZIP or CSV (*.zip;*.csv)|*.zip;*.csv|ZIP files (*.zip)|*.zip|CSV files (*.csv)|*.csv'; $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -ne 'OK') { return }
    $path = $dialog.FileName
    $hasCurrent = (Test-Path -LiteralPath $script:CompanyFile) -or (Test-Path -LiteralPath $script:LegacyCompanyFile)
    if ($hasCurrent) {
        $choice = [System.Windows.Forms.MessageBox]::Show("Companies House data is already imported. Replace it with this file?`n`nA timestamped backup will be kept. Existing data is not replaced unless the new file copies and validates successfully.", 'Replace existing company data?', 'YesNo', 'Warning', 'Button2')
        if ($choice -ne 'Yes') { return }
    }
    $tempFile = Join-Path $script:DataRoot ('companies-' + [Guid]::NewGuid().ToString('N') + '.csv.importing')
    $archive = $null; $sourceStream = $null; $targetStream = $null; $backupFile = $null
    $script:Importing = $true; $script:CancelImport = $false; $script:CloseAfterImport = $false
    $importButton.Enabled=$false; $searchButton.Enabled=$false; $nameBox.Enabled=$false; $townBox.Enabled=$false; $postcodeBox.Enabled=$false; $sicBox.Enabled=$false; $activeBox.Enabled=$false
    try {
        Show-Status 'Copying the source file into local storage…'
        if ([IO.Path]::GetExtension($path) -ieq '.zip') {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $archive = [System.IO.Compression.ZipFile]::OpenRead($path)
            $entry = $archive.Entries | Where-Object { $_.FullName -match '(?i)\.csv$' } | Sort-Object Length -Descending | Select-Object -First 1
            if (-not $entry) { throw 'The ZIP does not contain a CSV file.' }
            $sourceStream = $entry.Open(); $totalLength = [long]$entry.Length
        } else {
            $sourceStream = [IO.File]::OpenRead($path); $totalLength = (Get-Item -LiteralPath $path).Length
        }
        $targetStream = [IO.FileStream]::new($tempFile, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None, 8388608)
        $buffer = New-Object byte[] 8388608; $copied = [long]0; $lastUpdate = [long]0
        while (($read = $sourceStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            if ($script:CancelImport) { throw [OperationCanceledException]::new('Import cancelled. Existing data was kept.') }
            $targetStream.Write($buffer, 0, $read); $copied += $read
            if (($copied - $lastUpdate) -ge 67108864) { $lastUpdate = $copied; $percent = if ($totalLength -gt 0) { [math]::Min(100,[math]::Floor(100*$copied/$totalLength)) } else { 0 }; Show-Status "Copying company file: $percent% — click the window close button to cancel safely" }
        }
        $targetStream.Dispose(); $targetStream = $null; $sourceStream.Dispose(); $sourceStream = $null; if ($archive) { $archive.Dispose(); $archive = $null }
        $checkReader = [IO.StreamReader]::new($tempFile, [Text.Encoding]::UTF8, $true)
        try { $headerLine = $checkReader.ReadLine(); if (-not $headerLine) { throw 'The selected file is empty.' }; if ($headerLine -notmatch '(?i)(^|,)\s*"?CompanyName"?\s*(,|$)' -or $headerLine -notmatch '(?i)(^|,)\s*"?CompanyNumber"?\s*(,|$)') { throw 'This is not the expected Companies House basic company CSV.' } } finally { $checkReader.Dispose() }
        if ($script:CancelImport) { throw [OperationCanceledException]::new('Import cancelled. Existing data was kept.') }
        if (Test-Path -LiteralPath $script:CompanyFile) { $backupFile = $script:CompanyFile + '.' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff') + '.bak'; [IO.File]::Replace($tempFile, $script:CompanyFile, $backupFile, $true); Repoint-SavedSearchSources $script:CompanyFile $backupFile }
        else { [IO.File]::Move($tempFile, $script:CompanyFile) }
        $script:Importing = $false; $importButton.Enabled=$true; $searchButton.Enabled=$true; $nameBox.Enabled=$true; $townBox.Enabled=$true; $postcodeBox.Enabled=$true; $sicBox.Enabled=$true; $activeBox.Enabled=$true; $importInfo.Text = 'Companies House data stored locally. Search once to load its SIC categories.'; Set-SicCategoryChoices $script:CompanyFile $false; Show-Status 'Import complete — source CSV stored locally'
        if ($backupFile) { [System.Windows.Forms.MessageBox]::Show("Import complete. The previous data is backed up at:`n$backupFile", 'Import complete', 'OK', 'Information') | Out-Null }
        else { [System.Windows.Forms.MessageBox]::Show('Import complete. The Companies House CSV is ready to search.', 'Import complete', 'OK', 'Information') | Out-Null }
        if ($script:CloseAfterImport) { $form.Close() }
    } catch {
        if ($targetStream) { $targetStream.Dispose() }; if ($sourceStream) { $sourceStream.Dispose() }; if ($archive) { $archive.Dispose() }
        if (Test-Path -LiteralPath $tempFile) { Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue }
        $script:Importing = $false; $importButton.Enabled=$true; $searchButton.Enabled=$true; $nameBox.Enabled=$true; $townBox.Enabled=$true; $postcodeBox.Enabled=$true; $sicBox.Enabled=$true; $activeBox.Enabled=$true
        if ($_.Exception -is [OperationCanceledException]) { Show-Status 'Import cancelled — existing data is unchanged' }
        else { Show-Status 'Import failed — existing data is unchanged'; [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Import could not be completed', 'OK', 'Error') | Out-Null }
        if ($script:CloseAfterImport) { $form.Close() }
    }
})

$form.Add_FormClosing({
    param($sender, $eventArgs)
    if ($script:Importing) { $eventArgs.Cancel = $true; $script:CancelImport = $true; $script:CloseAfterImport = $true; $script:statusLabel.Text = 'Stopping import safely before closing…' }
    elseif ($script:Searching) { $eventArgs.Cancel=$true; $script:CancelSearch=$true; $script:CloseAfterSearch=$true; $script:statusLabel.Text='Stopping the search safely before closing…' }
})

$selectAllButton.Add_Click({
    $select = $false
    foreach ($row in $companyGrid.Rows) { if (-not [bool]$row.Cells['Pick'].Value) { $select = $true; break } }
    $script:LoadingSearchRows = $true
    foreach ($row in $companyGrid.Rows) { $row.Cells['Pick'].Value = $select }
    $script:LoadingSearchRows = $false; Save-SearchCompanies
})

$searchButton.Add_Click({
    $dataFile = if (Test-Path -LiteralPath $script:CompanyFile) { $script:CompanyFile } else { $script:LegacyCompanyFile }
    if (-not (Test-Path -LiteralPath $dataFile)) { [System.Windows.Forms.MessageBox]::Show('Import a Companies House ZIP or CSV first.', 'No company data', 'OK', 'Information') | Out-Null; return }
    $csvMode = ([IO.Path]::GetExtension($dataFile) -ieq '.csv')
    $script:LoadingSearchRows = $true; $script:CancelSearch=$false; $script:Searching=$true; $script:CloseAfterSearch=$false; $searchButton.Enabled=$false; $importButton.Enabled=$false; $nameBox.Enabled=$false; $townBox.Enabled=$false; $postcodeBox.Enabled=$false; $sicBox.Enabled=$false; $activeBox.Enabled=$false; $cancelSearchButton.Enabled=$true
    $companyGrid.Rows.Clear(); $script:CurrentCompanies.Clear(); $found = 0; $scanned = 0; $indexWriter=$null; $indexStream=$null; $resultPath=''; $runSaved=$false
    $n = $nameBox.Text.Trim(); $town = $townBox.Text.Trim(); $pc = $postcodeBox.Text.Trim(); $sicInput = $sicBox.Text.Trim(); $sic = if($script:SicCategoryLookup.ContainsKey($sicInput)){$script:SicCategoryLookup[$sicInput]}else{$sicInput}; if($sic -eq 'All categories'){$sic=''}
    $runId = [Guid]::NewGuid().ToString('N'); $script:ResultOffsets=[System.Collections.Generic.List[long]]::new(); $buildSicCategories=($script:SicCategories.Count -eq 0); if($buildSicCategories){$script:SicCategories.Clear()}; $script:SourceMode='source'; $script:SourceFile=$dataFile
    try {
        Read-SourceHeader $dataFile
        $script:ResultFormat=if($csvMode){'csv'}else{'jsonl'}
        $resultPath=Join-Path $script:SearchRoot ($runId + '-offsets.bin'); $script:ResultIndexFile=$resultPath
        $indexStream=[IO.FileStream]::new($resultPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read,1048576)
        $indexWriter=[IO.BinaryWriter]::new($indexStream)
        $lineEndingWidth=Get-SourceLineEndingWidth $dataFile; $hasBom=Test-SourceUtf8Bom $dataFile; $lineByteOffset=[long]0; $fileLength=[long](Get-Item -LiteralPath $dataFile).Length
        if($buildSicCategories){Show-Status 'Searching imported data and loading SIC categories for the first time…'}else{Show-Status 'Searching all imported company records…'}
        foreach ($line in [IO.File]::ReadLines($dataFile)) {
            if ($script:CancelSearch) { throw [OperationCanceledException]::new('Search stopped by user.') }
            $recordOffset=$lineByteOffset; $lineByteOffset += [Text.Encoding]::UTF8.GetByteCount([string]$line) + $lineEndingWidth; if($scanned -eq 0 -and $hasBom){$lineByteOffset+=3}; $scanned++
            if (-not $line) { continue }
            if ($csvMode) { if($scanned -eq 1){continue}; $co=New-CompanyRecord (Parse-CsvLine $line) $script:SourceHeaderMap }
            else { try { $co=$line | ConvertFrom-Json } catch { continue } }
            $sicVal=Get-Sic $co
            if($buildSicCategories){foreach($part in ($sicVal -split ';')) { $categoryValue=$part.Trim(); if($categoryValue -and $categoryValue -ne 'None Supplied'){[void]$script:SicCategories.Add($categoryValue)} }}
            if ($activeBox.Checked -and (Get-Prop $co 'CompanyStatus') -notmatch '^active$') { continue }
            if ($n -and (Get-Prop $co 'CompanyName') -notlike "*$n*") { continue }
            if ($town -and (Get-Prop $co 'RegAddress.PostTown') -notlike "*$town*") { continue }
            if ($pc -and (Get-Prop $co 'RegAddress.PostCode') -notlike "*$pc*") { continue }
            if ($sic -and $sicVal -notlike "*$sic*") { continue }
            $indexWriter.Write([long]$recordOffset)
            $found++
            if (($scanned % 25000) -eq 0) { $percent=if($fileLength -gt 0){[Math]::Min(100,[int](100*$lineByteOffset/$fileLength))}else{0}; $companyCountLabel.Text="Searching… $($scanned.ToString('N0')) scanned, $($found.ToString('N0')) matches"; Show-Status "Reading imported data: $percent% — $($scanned.ToString('N0')) records scanned. Stop search is available." }
        }
        $indexWriter.Dispose(); $indexWriter=$null; $indexStream=$null
        if($script:CancelSearch){throw [OperationCanceledException]::new('Search stopped by user.')}
        if($buildSicCategories){Save-SicCategoryCache $dataFile}
        Set-SicCategoryChoices $dataFile $false $true
        $created = Get-Date; $area = if ($town) { $town } elseif ($pc) { $pc } else { 'UK-wide' }
        $category = if ($sic) { $sic } else { 'All categories' }
        $nameParts=New-Object 'System.Collections.Generic.List[string]'; $nameParts.Add($area); if($pc -and $town){$nameParts.Add("Postcode $pc")}; if($n){$nameParts.Add("Name contains $n")}; $nameParts.Add($category); $nameParts.Add($(if($activeBox.Checked){'Active only'}else{'All statuses'})); $runName=$nameParts -join ' — '
        $run = [pscustomobject]@{ Id=$runId; Name=$runName; Area=$area; Category=$category; NameFilter=$n; Town=$town; Postcode=$pc; SIC=$sic; ActiveOnly=$activeBox.Checked; ResultCount=$found; Format=$script:ResultFormat; SourceMode='source'; SourceFile=$dataFile; ResultIndexFile=$resultPath; HeaderMap=$script:SourceHeaderMap; CreatedAt=$created.ToString('yyyy-MM-dd HH:mm:ss') }
        $runPath = Join-Path $script:SearchRoot ($runId + '.json')
        [IO.File]::WriteAllText($runPath, (ConvertTo-Json -InputObject $run -Depth 6), [Text.UTF8Encoding]::new($false))
        $runSaved=$true
        $script:CurrentSearchId=$runId; $script:ResultFile=''; $script:ResultIndexFile=$resultPath; $script:SourceFile=$dataFile; $script:SourceMode='source'; $script:ResultTotal=$found; $script:ResultPage=0; Refresh-SearchRuns; Refresh-LeadGrid; Show-CompanyPage
        Show-Status "Search complete — $($found.ToString('N0')) matching companies across $([Math]::Max(1,[Math]::Ceiling($found/[double]$script:PageSize))) page(s)"
    } catch {
        if($indexWriter){$indexWriter.Dispose()} elseif($indexStream){$indexStream.Dispose()}
        if(-not $runSaved -and $resultPath -and (Test-Path -LiteralPath $resultPath)){Remove-Item -LiteralPath $resultPath -Force -ErrorAction SilentlyContinue}
        if($script:CurrentSearchId){try{Load-SearchCompanies $script:CurrentSearchId}catch{}}else{$companyCountLabel.Text='Search stopped — no partial results were saved'}
        if($_.Exception -is [OperationCanceledException]) { Show-Status 'Search stopped — no partial search was saved' }
        else { Show-Status 'Search failed'; [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Search could not be completed','OK','Error')|Out-Null }
    } finally { $script:LoadingSearchRows=$false; $script:Searching=$false; $searchButton.Enabled=$true; $importButton.Enabled=$true; $nameBox.Enabled=$true; $townBox.Enabled=$true; $postcodeBox.Enabled=$true; $sicBox.Enabled=$true; $activeBox.Enabled=$true; $cancelSearchButton.Enabled=$false; if($script:CloseAfterSearch){$form.Close()} }
})

$cancelSearchButton.Add_Click({ $script:CancelSearch=$true; Show-Status 'Stopping search safely…' })
$prevPageButton.Add_Click({ if($script:ResultPage -gt 0){$companyGrid.EndEdit();Save-SearchCompanies;$script:ResultPage--;Show-CompanyPage} })
$nextPageButton.Add_Click({ if((($script:ResultPage+1)*$script:PageSize) -lt $script:ResultTotal){$companyGrid.EndEdit();Save-SearchCompanies;$script:ResultPage++;Show-CompanyPage} })

$googleButton.Add_Click({
    if ($companyGrid.SelectedRows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Select a company first.', 'Select a company', 'OK', 'Information') | Out-Null; return }
    $r = $companyGrid.SelectedRows[0].Index; $co = $script:CurrentCompanies[$r]
    $query = '"' + (Get-Prop $co 'CompanyName') + '" ' + (Get-Prop $co 'RegAddress.PostTown')
    Start-Process ('https://www.google.com/search?q=' + [Uri]::EscapeDataString($query))
    Show-Status 'Google search opened in your browser; check the result before using the website.'
})

$verifyButton.Add_Click({
    if ($companyGrid.SelectedRows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Select a company first.', 'Select a company', 'OK', 'Information') | Out-Null; return }
    $rowIndex = $companyGrid.SelectedRows[0].Index; $co = $script:CurrentCompanies[$rowIndex]
    $url = $websiteBox.Text.Trim()
    if (-not $url) { [System.Windows.Forms.MessageBox]::Show('Paste the candidate website address.', 'Website required', 'OK', 'Information') | Out-Null; return }
    if ($url -notmatch '^https?://') { $url = 'https://' + $url; $websiteBox.Text = $url }
    try {
        $uri = [Uri]$url; if ($uri.Scheme -notin @('http','https')) { throw 'Use a public http or https website.' }
        Show-Status "Checking $($uri.Host) for the registered company name…"
        $pages = New-Object 'System.Collections.Generic.List[string]'; $seen = @{}
        $queue = New-Object 'System.Collections.Generic.Queue[string]'; $queue.Enqueue($uri.AbsoluteUri)
        $root = $uri.GetLeftPart([UriPartial]::Authority).TrimEnd('/')
        $pageLimit = 8
        while ($queue.Count -gt 0 -and $pages.Count -lt $pageLimit) {
            $candidate = $queue.Dequeue(); if ($seen.ContainsKey($candidate)) { continue }; $seen[$candidate] = $true
            try {
                $response = Invoke-WebRequest -Uri $candidate -UseBasicParsing -TimeoutSec 12 -MaximumRedirection 4 -Headers @{ 'User-Agent'='WebsitesDoItLeadFinder/0.5 (+public business contact discovery)' }
                $html = [string]$response.Content
                if ($html) { $pages.Add($html) }
                if ($pages.Count -eq 1) {
                    foreach ($m in [regex]::Matches($html, '(?i)href\s*=\s*["'']([^"'']+)["'']')) {
                        $href = [Net.WebUtility]::HtmlDecode($m.Groups[1].Value)
                        if ($href -match '(?i)(contact|about|privacy|terms|location)' -and $href -notmatch '^(mailto:|tel:|#|javascript:)') {
                            try { $link = [Uri]::new([Uri]$response.BaseResponse.ResponseUri, $href); if ($link.Host -eq $uri.Host -and $link.Scheme -in @('http','https')) { $queue.Enqueue($link.AbsoluteUri) } } catch { }
                        }
                    }
                }
            } catch { }
        }
        if ($pages.Count -eq 0) { throw 'The website could not be read. Check the address and try again.' }
        $companyName = Get-Prop $co 'CompanyName'; $needle = Normalize-CompanyName $companyName
        $combined = [Net.WebUtility]::HtmlDecode(($pages -join "`n")) -replace '<[^>]+>', ' ' -replace '\s+', ' '
        $normalText = Normalize-CompanyName $combined
        if (-not $needle -or $normalText.IndexOf($needle, [StringComparison]::Ordinal) -lt 0) {
            throw "The registered company name was not found on the website pages checked. No lead was saved.`n`nCompany: $companyName`nWebsite: $($uri.Host)"
        }
        $emails = New-Object 'System.Collections.Generic.List[string]'; $generic = 'info|sales|support|accounts|help|contact|hello|office|enquiries|enquiry|admin|reception|bookings|team'
        $emailPattern = '(?i)\b(?:' + $generic + ')[a-z0-9._%+\-]*@(?:[a-z0-9\-]+\.)+[a-z]{2,}\b'
        foreach ($m in [regex]::Matches($combined, $emailPattern)) { $e = $m.Value.ToLowerInvariant(); if (-not $emails.Contains($e)) { $emails.Add($e) } }
        $phones = New-Object 'System.Collections.Generic.List[string]'
        foreach ($p in [regex]::Matches($combined, '(?<!\w)(?:\+44\s?\(?0?\)?|0)(?:\s?\(?\d{2,5}\)?)[\s.\-/]*\d{3,4}[\s.\-/]*\d{3,4}(?!\w)')) { $v = ($p.Value -replace '\s+', ' ').Trim(); if (-not $phones.Contains($v)) { $phones.Add($v) } }
        $existing = @($script:CurrentLeads | Where-Object { $_.CompanyNumber -eq (Get-Prop $co 'CompanyNumber') })
        if ($existing.Count) { foreach ($old in $existing) { [void]$script:CurrentLeads.Remove($old) } }
        $email = if ($emails.Count) { $emails[0] } else { '' }; $telephone = if ($phones.Count) { $phones[0] } else { '' }
        $lead = [pscustomobject]@{ CompanyName=$companyName; CompanyNumber=(Get-Prop $co 'CompanyNumber'); Town=(Get-Prop $co 'RegAddress.PostTown'); Postcode=(Get-Prop $co 'RegAddress.PostCode'); SIC=(Get-Sic $co); Website=$uri.GetLeftPart([UriPartial]::Authority); Email=$email; AllGenericEmails=($emails -join '; '); Telephone=$telephone; AllTelephones=($phones -join '; '); Verification='Legal name matched'; VerifiedAt=(Get-Date).ToString('s'); SourcePages=($pages.Count); Evidence='Registered company name found on public website pages' }
        $script:CurrentLeads.Add($lead); Save-Leads; Refresh-LeadGrid
        if ($script:BatchMode) { $script:BatchSuccess++ }
        else { [System.Windows.Forms.MessageBox]::Show("Website identity matched and saved.`n`nGeneric emails: $($emails.Count)`nTelephone numbers: $($phones.Count)`nPages checked: $($pages.Count)`n`nA blank contact field means no suitable public contact was found on the pages checked.", 'Verified lead saved', 'OK', 'Information') | Out-Null }
        Show-Status "Verified $companyName — lead saved locally"
    } catch {
        if ($script:BatchMode) { $script:BatchFailure++ }
        else { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Website check', 'OK', 'Warning') | Out-Null }
        Show-Status 'Website check ended'
    }
})

$batchButton.Add_Click({
    $selected = New-Object 'System.Collections.Generic.List[object]'; $missing = 0
    foreach ($row in $companyGrid.Rows) {
        if ([bool]$row.Cells['Pick'].Value) {
            $url = [string]$row.Cells['Website'].Value
            if ($url.Trim()) { $selected.Add([pscustomobject]@{ Index=$row.Index; Url=$url.Trim() }) } else { $missing++ }
        }
    }
    if ($selected.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Tick companies and paste a candidate website into each row before running checks.', 'Websites needed', 'OK', 'Information') | Out-Null; return }
    $script:BatchMode = $true; $script:BatchSuccess = 0; $script:BatchFailure = 0
    for ($i = 0; $i -lt $selected.Count; $i++) {
        $item = $selected[$i]; $companyGrid.ClearSelection(); $companyGrid.Rows[$item.Index].Selected = $true; $companyGrid.CurrentCell = $companyGrid.Rows[$item.Index].Cells['CompanyName']; $websiteBox.Text = $item.Url
        Show-Status "Checking selected company $($i + 1) of $($selected.Count)…"
        $verifyButton.PerformClick()
    }
    $script:BatchMode = $false; Refresh-LeadGrid
    [System.Windows.Forms.MessageBox]::Show("Checks complete for this saved search.`n`nVerified and saved: $($script:BatchSuccess)`nNot matched or unavailable: $($script:BatchFailure)`nSkipped because no website was entered: $missing`n`nOnly verified matches are included in this search's export.", 'Search checks complete', 'OK', 'Information') | Out-Null
    Show-Status 'Selected website checks complete'
})

$script:runPicker.Add_SelectedIndexChanged({
    if ($script:runPicker.SelectedItem) { $companyGrid.EndEdit(); Save-SearchCompanies; $script:CurrentSearchId = $script:runPicker.SelectedItem.Id; Refresh-LeadGrid; Load-SearchCompanies $script:CurrentSearchId }
})

$exportButton.Add_Click({
    if ($script:CurrentLeads.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('There are no verified leads to export yet.', 'No leads', 'OK', 'Information') | Out-Null; return }
    $defaultName = 'WebsitesDoIt-Verified-Leads.csv'
    if ($script:CurrentSearchId) {
        $metaPath = Join-Path $script:SearchRoot ($script:CurrentSearchId + '.json')
        if (Test-Path -LiteralPath $metaPath) { $meta = Get-Content -LiteralPath $metaPath -Raw | ConvertFrom-Json; $safe = $meta.Name -replace '[^A-Za-z0-9_-]+','-'; $defaultName = "$safe-Verified-Leads.csv" }
    }
    $dialog = New-Object System.Windows.Forms.SaveFileDialog; $dialog.Title = 'Export this saved search'; $dialog.Filter = 'CSV files (*.csv)|*.csv'; $dialog.FileName = $defaultName
    if ($dialog.ShowDialog() -eq 'OK') {
        $rows = foreach ($lead in $script:CurrentLeads) { [pscustomobject]@{ Company=$lead.CompanyName; Email=$lead.Email; Website=$lead.Website; Telephone=$lead.Telephone } }
        $csv = $rows | ConvertTo-Csv -NoTypeInformation
        [IO.File]::WriteAllLines($dialog.FileName, $csv, ([Text.UTF8Encoding]::new($true)))
        Show-Status "Exported $($script:CurrentLeads.Count) leads to $($dialog.FileName)"
    }
})

$tabs.Add_SelectedIndexChanged({ if ($tabs.SelectedTab -eq $leadsTab) { $companyGrid.EndEdit(); Save-SearchCompanies; Refresh-LeadGrid } })
[void]$form.ShowDialog()
