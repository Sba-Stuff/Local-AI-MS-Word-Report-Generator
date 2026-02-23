#Requires -Version 5.1
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# ---------- CONFIGURATION ----------
$apiBase = "http://192.168.10.4:1234/v1"
$model   = "liquid/lfm2-1.2b"
$chatEndpoint = "$apiBase/chat/completions"

# Timeout settings (seconds) – increase for slow systems
$script:apiTimeout = 120
$script:delayBetweenCalls = 500

# ---------- GUI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "AI Report Generator (Word)"
$form.Size = New-Object System.Drawing.Size(600,500)
$form.StartPosition = "CenterScreen"

# Title
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Report Title:"
$lblTitle.Location = New-Object System.Drawing.Point(20,20)
$lblTitle.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($lblTitle)

$txtTitle = New-Object System.Windows.Forms.TextBox
$txtTitle.Location = New-Object System.Drawing.Point(130,18)
$txtTitle.Size = New-Object System.Drawing.Size(430,20)
$form.Controls.Add($txtTitle)

# Author
$lblAuthor = New-Object System.Windows.Forms.Label
$lblAuthor.Text = "Author:"
$lblAuthor.Location = New-Object System.Drawing.Point(20,50)
$lblAuthor.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($lblAuthor)

$txtAuthor = New-Object System.Windows.Forms.TextBox
$txtAuthor.Location = New-Object System.Drawing.Point(130,48)
$txtAuthor.Size = New-Object System.Drawing.Size(430,20)
$txtAuthor.Text = $env:USERNAME
$form.Controls.Add($txtAuthor)

# Document Type (simplified)
$lblType = New-Object System.Windows.Forms.Label
$lblType.Text = "Document Type:"
$lblType.Location = New-Object System.Drawing.Point(20,80)
$lblType.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($lblType)

$cmbType = New-Object System.Windows.Forms.ComboBox
$cmbType.Location = New-Object System.Drawing.Point(130,78)
$cmbType.Size = New-Object System.Drawing.Size(200,20)
$cmbType.DropDownStyle = "DropDownList"
$cmbType.Items.AddRange(@("Technical Report","Research Paper","Proposal","Case Study","Analysis"))
$cmbType.SelectedIndex = 0
$form.Controls.Add($cmbType)

# Number of Sections
$lblSections = New-Object System.Windows.Forms.Label
$lblSections.Text = "Number of Sections:"
$lblSections.Location = New-Object System.Drawing.Point(20,110)
$lblSections.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($lblSections)

$numSections = New-Object System.Windows.Forms.NumericUpDown
$numSections.Location = New-Object System.Drawing.Point(130,108)
$numSections.Size = New-Object System.Drawing.Size(60,20)
$numSections.Minimum = 3
$numSections.Maximum = 15
$numSections.Value = 6
$form.Controls.Add($numSections)

# Detailed Instructions (multiline textbox)
$lblInstructions = New-Object System.Windows.Forms.Label
$lblInstructions.Text = "Detailed Instructions:"
$lblInstructions.Location = New-Object System.Drawing.Point(20,140)
$lblInstructions.Size = New-Object System.Drawing.Size(120,20)
$form.Controls.Add($lblInstructions)

$txtInstructions = New-Object System.Windows.Forms.TextBox
$txtInstructions.Location = New-Object System.Drawing.Point(20,165)
$txtInstructions.Size = New-Object System.Drawing.Size(540,150)
$txtInstructions.Multiline = $true
$txtInstructions.ScrollBars = "Vertical"
$txtInstructions.Text = "Write a comprehensive report on the given topic. Include background, current trends, challenges, and future outlook. Use technical language suitable for professionals."
$form.Controls.Add($txtInstructions)

# Generate Button
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Generate Report"
$btnGenerate.Location = New-Object System.Drawing.Point(130,330)
$btnGenerate.Size = New-Object System.Drawing.Size(180,30)
$btnGenerate.Add_Click({
    $btnGenerate.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Generate-Report -Title $txtTitle.Text -Author $txtAuthor.Text `
                        -DocType $cmbType.SelectedItem -SectionCount $numSections.Value `
                        -Instructions $txtInstructions.Text
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    } finally {
        $btnGenerate.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
$form.Controls.Add($btnGenerate)

# Status label (simple)
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready"
$lblStatus.Location = New-Object System.Drawing.Point(20,370)
$lblStatus.Size = New-Object System.Drawing.Size(540,20)
$form.Controls.Add($lblStatus)

# ---------- AI Functions ----------
function Update-Status {
    param([string]$message)
    $lblStatus.Text = $message
    $lblStatus.Refresh()
    Start-Sleep -Milliseconds 50
}

function Invoke-AI {
    param([string]$prompt, [int]$maxTokens = 1000)
    
    $body = @{
        model = $model
        messages = @(
            @{ role = "system"; content = "You are a professional technical writer. Create well-structured, detailed paragraphs with proper technical language." }
            @{ role = "user";   content = $prompt }
        )
        temperature = 0.7
        max_tokens = $maxTokens
        stream = $false
    } | ConvertTo-Json -Depth 4

    try {
        Write-Host "  [WAIT] Waiting for AI response (timeout: $script:apiTimeout seconds)..." -ForegroundColor DarkYellow
        $response = Invoke-RestMethod -Uri $chatEndpoint -Method Post -Body $body -ContentType "application/json" -TimeoutSec $script:apiTimeout
        return $response.choices[0].message.content.Trim()
    } catch {
        Write-Warning "AI call failed: $($_.Exception.Message)"
        return $null
    }
}

function Generate-Outline {
    param([string]$title, [string]$docType, [int]$sectionCount, [string]$instructions)
    
    Update-Status "Generating outline..."
    Write-Host "[OUTLINE] Generating outline for '$title'..." -ForegroundColor Yellow
    
    $prompt = @"
Create a numbered outline for a $docType titled "$title".
Follow these instructions: $instructions
Generate exactly $sectionCount main sections.
Format: 1. Section Title, 2. Section Title, etc.
Only output the numbered list, no extra text.
"@
    $result = Invoke-AI -prompt $prompt -maxTokens 400
    if (-not $result) {
        Write-Host "  [WARN] Using fallback outline" -ForegroundColor DarkYellow
        $result = "1. Introduction`n2. Background`n3. Current State`n4. Analysis`n5. Recommendations`n6. Conclusion"
    }
    
    $sections = @()
    $result -split "`n" | ForEach-Object {
        if ($_ -match '^\d+\.\s*(.+)$') {
            $sections += $matches[1].Trim()
        }
    }
    return $sections
}

function Generate-SectionContent {
    param([string]$title, [string]$docType, [string]$section, [string]$instructions, [int]$sectionNumber, [int]$totalSections)
    
    Update-Status "Writing section $sectionNumber of $totalSections : $section..."
    Write-Host "  [PARAGRAPH] Generating content for: $section" -ForegroundColor Gray
    
    $prompt = @"
Write detailed paragraph content for the section "$section" of a $docType titled "$title".
Overall instructions: $instructions
Requirements:
- Write 3 to 5 substantial paragraphs
- Use professional, technical language
- Be specific, include relevant details
- Do not include the section title in the response
- Separate paragraphs with blank lines
"@
    $result = Invoke-AI -prompt $prompt -maxTokens 1500
    if (-not $result) {
        $result = "This section discusses $section in the context of $title.`n`nKey aspects include fundamental principles and effective application.`n`nFurther analysis reveals important considerations for practitioners.`n`nIn conclusion, this area requires continued attention."
    }
    return $result
}

# ---------- Word Document Creation ----------
function New-ReportDocument {
    param([string]$title, [string]$author, [string]$docType, [array]$sections, [hashtable]$content)
    
    Update-Status "Creating Word document..."
    Write-Host "[WORD] Creating Word document..." -ForegroundColor Cyan
    
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $true
    $document = $wordApp.Documents.Add()
    $selection = $wordApp.Selection

    # Title page
    $selection.Style = "Title"
    $selection.TypeText($title)
    $selection.TypeParagraph()
    
    $selection.Style = "Subtitle"
    $selection.TypeText($docType)
    $selection.TypeParagraph()
    
    $selection.Style = "Normal"
    $selection.Font.Bold = $true
    $selection.TypeText("Author: ")
    $selection.Font.Bold = $false
    $selection.TypeText($author)
    $selection.TypeParagraph()
    
    $selection.Font.Bold = $true
    $selection.TypeText("Date: ")
    $selection.Font.Bold = $false
    $selection.TypeText((Get-Date -Format "MMMM dd, yyyy"))
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    # Sections
    $sectionNumber = 1
    foreach ($section in $sections) {
        # Heading
        $selection.Style = "Heading 1"
        $selection.TypeText("$sectionNumber. $section")
        $selection.TypeParagraph()
        
        # Content paragraphs
        $selection.Style = "Normal"
        $paragraphs = $content[$section] -split "`n`n|`r`n`r`n" | Where-Object { $_.Trim() -ne "" }
        foreach ($para in $paragraphs) {
            $selection.TypeText($para.Trim())
            $selection.TypeParagraph()
            $selection.TypeParagraph()
        }
        $sectionNumber++
    }

    # Save to Desktop
    $desktop = [Environment]::GetFolderPath("Desktop")
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $safeTitle = $title -replace '[^\w\s-]','' -replace '\s+','_'
    $filePath = "$desktop\$safeTitle-$timestamp.docx"
    
    $document.SaveAs([ref]$filePath)
    Write-Host "[SAVED] Document saved: $filePath" -ForegroundColor Green

    # Cleanup
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selection) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    return $filePath
}

# ---------- Main Generation Function ----------
function Generate-Report {
    param(
        [string]$Title,
        [string]$Author,
        [string]$DocType,
        [int]$SectionCount,
        [string]$Instructions
    )
    
    if ([string]::IsNullOrWhiteSpace($Title)) {
        throw "Please enter a report title."
    }
    
    $sections = Generate-Outline -title $Title -docType $DocType -sectionCount $SectionCount -instructions $Instructions
    if ($sections.Count -eq 0) { throw "No sections generated." }
    
    Write-Host "[SECTIONS] $($sections -join ', ')" -ForegroundColor Cyan
    
    $content = @{}
    $i = 1
    foreach ($section in $sections) {
        $content[$section] = Generate-SectionContent -title $Title -docType $DocType -section $section `
                                                      -instructions $Instructions -sectionNumber $i -totalSections $sections.Count
        if ($i -lt $sections.Count) { Start-Sleep -Milliseconds $script:delayBetweenCalls }
        $i++
    }
    
    $output = New-ReportDocument -title $Title -author $Author -docType $DocType `
                                  -sections $sections -content $content
    
    [System.Windows.Forms.MessageBox]::Show("Report created!`n$output", "Success", "OK", "Information")
    Update-Status "Done. Report saved to desktop."
}

# ---------- Launch ----------
$form.ShowDialog() | Out-Null