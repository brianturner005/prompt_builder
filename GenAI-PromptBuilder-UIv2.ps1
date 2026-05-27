#requires -Version 5.1
<#
  GenAI Integrated Suite - Mission UI (PowerShell 5.1)
  - Combines the Prompt Builder and the Prompt Launcher into a single application.
#>

# --- Environment & Setup ---
try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13 } catch { try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {} }

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing       | Out-Null
try { Add-Type -AssemblyName System.Net.Http | Out-Null } catch {}

$appDataDir = Join-Path $env:LOCALAPPDATA "GenAI-PromptBuilder"
if (-not (Test-Path $appDataDir)) { New-Item -ItemType Directory -Force -Path $appDataDir | Out-Null }
$configFile = Join-Path $appDataDir "config.json"
$templateFile = Join-Path $appDataDir "templates.json"

# ==========================================================================================
# 1. INTEGRATED PROMPT LAUNCHER FUNCTION (The Pop-up)
# ==========================================================================================
function Show-CopilotPromptLauncher {
    $launcherForm = New-Object System.Windows.Forms.Form
    $launcherForm.Text = "Copilot Prompt Launcher"
    $launcherForm.Size = New-Object System.Drawing.Size(700, 600)
    $launcherForm.StartPosition = "CenterParent"
    $launcherForm.FormBorderStyle = "FixedDialog"
    $launcherForm.MaximizeBox = $false

    $promptLibrary = @(
        # --- DoW GenAI: Summarization ---
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Summarization"; Name="BLUF Executive Summary"; Prompt="Review the attached document and provide a Bottom Line Up Front (BLUF) summary. Follow the BLUF with three concise bullet points outlining the strategic impact, resource requirements, and any immediate deadlines." }
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Summarization"; Name="Intelligence Report Synthesis"; Prompt="Synthesize the key findings from this intelligence report. Focus on identifying specific threats, geographic areas of concern, and recommended posture adjustments. Present the information in a format suitable for a morning briefing." }
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Summarization"; Name="Technical Manual Condensation"; Prompt="Condense the following technical manual section into a 'Quick Start Guide' for operators. Include only the essential safety warnings and the step-by-step procedure for system initialization." }
        # --- DoW GenAI: Drafting ---
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Drafting"; Name="Official Correspondence (Standard Memo)"; Prompt="Draft an official memorandum for record (MFR) regarding [Subject]. Ensure the tone is formal and follows standard military correspondence formatting. Use active voice and keep the body to three paragraphs: Situation, Action, and Conclusion." }
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Drafting"; Name="Award Citation Narrative"; Prompt="Draft a narrative for a [Award Name] for [Rank/Name]. Key achievements include: [Achievement 1], [Achievement 2], and [Achievement 3]. Focus on 'impact to the mission' and 'leadership' using strong action verbs." }
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Drafting"; Name="Information Paper Outline"; Prompt="Create a detailed outline for an Information Paper regarding the deployment of [Technology/System]. Include sections for Background, Discussion, and Recommendation." }
        # --- DoW GenAI: Analysis ---
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Analysis"; Name="Strategic Risk Assessment"; Prompt="Analyze the provided project plan for potential strategic risks. Identify gaps in logistics, staffing, and timeline. For each risk, suggest a potential mitigation strategy consistent with IL5 security protocols." }
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Analysis"; Name="COA Comparison"; Prompt="Compare the two Courses of Action (COAs) described in the text. Create a table evaluating each against the following criteria: Cost-effectiveness, Time-to-Deployment, and Operational Risk." }
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Analysis"; Name="Policy Gap Analysis"; Prompt="Read the provided policy and compare it against the latest [Regulation Name/Number]. Identify any sections that are outdated or in conflict with the new regulation." }
        # --- DoW GenAI: Explanation ---
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Explanation"; Name="Simplify Doctrine Jargon"; Prompt="Explain the concept of [Joint Publication/Doctrine Term] in plain English. Avoid acronyms and provide a real-world analogy that would be understandable to a junior service member." }
        [PSCustomObject]@{ Platform="DoW GenAI"; Category="Explanation"; Name="Zero Trust for Leadership"; Prompt="Explain the core principles of Zero Trust Architecture and why it is critical for DoW mission security. Focus on the 'Never Trust, Always Verify' mindset without using overly technical jargon." }
        # --- Microsoft Copilot for M365: Summarization ---
        [PSCustomObject]@{ Platform="Microsoft Copilot for M365"; Category="Summarization"; Name="Teams Meeting Recap"; Prompt="Summarize the transcript of this meeting. List the main discussion topics, any decisions that were finalized, and specific action items assigned to individuals." }
        [PSCustomObject]@{ Platform="Microsoft Copilot for M365"; Category="Summarization"; Name="Email Thread Summary"; Prompt="Summarize this email thread. Who is waiting on information from whom? Provide a timeline of the conversation and the current status of the request." }
        # --- Microsoft Copilot for M365: Drafting ---
        [PSCustomObject]@{ Platform="Microsoft Copilot for M365"; Category="Drafting"; Name="Meeting Agenda Draft"; Prompt="Draft an agenda for a 30-minute sync meeting regarding [Project Name]. Include sections for status updates, blocker resolution, and next steps. Add a 'Pre-read' section at the top." }
        [PSCustomObject]@{ Platform="Microsoft Copilot for M365"; Category="Drafting"; Name="Follow-up Email"; Prompt="Write a professional follow-up email to [Name] regarding our meeting earlier today. Thank them for their time and restate that I will provide the updated project tracker by COB tomorrow." }
        # --- Microsoft Copilot for M365: Analysis ---
        [PSCustomObject]@{ Platform="Microsoft Copilot for M365"; Category="Analysis"; Name="Excel Data Trends"; Prompt="Look at this data and identify the three most significant trends over the last quarter. Highlight any outliers that may require further investigation." }
        # --- Microsoft Copilot for M365: Explanation ---
        [PSCustomObject]@{ Platform="Microsoft Copilot for M365"; Category="Explanation"; Name="Software How-To Guide"; Prompt="Provide a step-by-step guide on how to [Action, e.g., create a Pivot Table] in Microsoft Excel. Use numbered steps and mention specific tab names." }
    )

    $lblPlatform = New-Object System.Windows.Forms.Label; $lblPlatform.Text = "1. Select Platform:"; $lblPlatform.Location = New-Object System.Drawing.Point(20, 20); $lblPlatform.AutoSize = $true; $lblPlatform.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold); $launcherForm.Controls.Add($lblPlatform)
    $cmbPlatform = New-Object System.Windows.Forms.ComboBox; $cmbPlatform.Location = New-Object System.Drawing.Point(20, 45); $cmbPlatform.Size = New-Object System.Drawing.Size(300, 25); $cmbPlatform.DropDownStyle = "DropDownList"; @("DoW GenAI", "Microsoft Copilot for M365") | ForEach-Object { [void]$cmbPlatform.Items.Add($_) }; $cmbPlatform.SelectedIndex = 0; $launcherForm.Controls.Add($cmbPlatform)

    $lblCategory = New-Object System.Windows.Forms.Label; $lblCategory.Text = "2. Select Category:"; $lblCategory.Location = New-Object System.Drawing.Point(340, 20); $lblCategory.AutoSize = $true; $lblCategory.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold); $launcherForm.Controls.Add($lblCategory)
    $cmbCategory = New-Object System.Windows.Forms.ComboBox; $cmbCategory.Location = New-Object System.Drawing.Point(340, 45); $cmbCategory.Size = New-Object System.Drawing.Size(320, 25); $cmbCategory.DropDownStyle = "DropDownList"; @("Summarization", "Drafting", "Analysis", "Explanation") | ForEach-Object { [void]$cmbCategory.Items.Add($_) }; $cmbCategory.SelectedIndex = 0; $launcherForm.Controls.Add($cmbCategory)

    $lblPrompts = New-Object System.Windows.Forms.Label; $lblPrompts.Text = "3. Select a Prompt to Discover:"; $lblPrompts.Location = New-Object System.Drawing.Point(20, 90); $lblPrompts.AutoSize = $true; $lblPrompts.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold); $launcherForm.Controls.Add($lblPrompts)
    $lstPrompts = New-Object System.Windows.Forms.ListBox; $lstPrompts.Location = New-Object System.Drawing.Point(20, 115); $lstPrompts.Size = New-Object System.Drawing.Size(640, 180); $lstPrompts.Font = New-Object System.Drawing.Font("Segoe UI", 9); $launcherForm.Controls.Add($lstPrompts)

    $txtPromptDetail = New-Object System.Windows.Forms.TextBox; $txtPromptDetail.Location = New-Object System.Drawing.Point(20, 310); $txtPromptDetail.Size = New-Object System.Drawing.Size(640, 160); $txtPromptDetail.Multiline = $true; $txtPromptDetail.ScrollBars = "Vertical"; $txtPromptDetail.ReadOnly = $true; $txtPromptDetail.BackColor = "White"; $txtPromptDetail.Font = New-Object System.Drawing.Font("Consolas", 10); $launcherForm.Controls.Add($txtPromptDetail)
    
    $btnCopy = New-Object System.Windows.Forms.Button; $btnCopy.Text = "📋 Copy Prompt to Clipboard"; $btnCopy.Location = New-Object System.Drawing.Point(20, 490); $btnCopy.Size = New-Object System.Drawing.Size(250, 45); $btnCopy.BackColor = [System.Drawing.Color]::LightBlue; $btnCopy.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold); $launcherForm.Controls.Add($btnCopy)

    $UpdateListAction = {
        $lstPrompts.Items.Clear()
        $txtPromptDetail.Text = ""
        $filteredPrompts = $promptLibrary | Where-Object { $_.Platform -eq $cmbPlatform.SelectedItem -and $_.Category -eq $cmbCategory.SelectedItem }
        foreach ($p in $filteredPrompts) { [void]$lstPrompts.Items.Add($p.Name) }
    }
    $cmbPlatform.Add_SelectedIndexChanged($UpdateListAction)
    $cmbCategory.Add_SelectedIndexChanged($UpdateListAction)

    $lstPrompts.Add_SelectedIndexChanged({
        if ($lstPrompts.SelectedItem) {
            $p = $promptLibrary | Where-Object { $_.Name -eq $lstPrompts.SelectedItem -and $_.Platform -eq $cmbPlatform.SelectedItem }
            if ($p) { $txtPromptDetail.Text = $p.Prompt }
        }
    })

    $btnCopy.Add_Click({
        if (-not [string]::IsNullOrWhiteSpace($txtPromptDetail.Text)) {
            [System.Windows.Forms.Clipboard]::SetText($txtPromptDetail.Text)
            [System.Windows.Forms.MessageBox]::Show("Prompt copied successfully!", "System Notification", "OK", "Information")
        }
    })
    
    . $UpdateListAction
    [void]$launcherForm.ShowDialog()
    $launcherForm.Dispose()
}

# ==========================================================================================
# 2. ORIGINAL APP HELPER FUNCTIONS
# ==========================================================================================
function Get-NormalizedBaseUrl {
    param([string]$BaseUrl)
    if ([string]::IsNullOrWhiteSpace($BaseUrl)) { return "" }
    $u = $BaseUrl.Trim() -replace '/v1/.*$','' -replace '/server/v1/.*$','' -replace '/openai/.*$',''
    if ($u.EndsWith('/')) { $u = $u.Substring(0, $u.Length - 1) }
    return $u
}

function New-HttpClient {
    try {
        $handler = New-Object System.Net.Http.HttpClientHandler
        $handler.UseDefaultCredentials = $true
        $handler.Proxy = [System.Net.WebRequest]::GetSystemWebProxy()
        if ($handler.Proxy) { $handler.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials }
        $client = New-Object System.Net.Http.HttpClient($handler)
        $client.Timeout = [System.TimeSpan]::FromSeconds(120)
        return $client
    } catch { return $null }
}

function Normalize-ModelIds {
    param($Obj)
    $ids = New-Object System.Collections.Generic.List[string]
    if (-not $Obj) { return $ids.ToArray() }
    $data = if ($Obj.data) { $Obj.data } elseif ($Obj.models) { $Obj.models } elseif ($Obj -is [System.Collections.IEnumerable]) { $Obj } else { $null }
    if ($data) {
        foreach ($m in $data) {
            $id = if ($m.id) { $m.id } elseif ($m.model) { $m.model } elseif ($m -is [string]) { $m } else { $null }
            if ($id) { $ids.Add([string]$id) }
        }
    }
    return ($ids.ToArray() | Sort-Object -Unique)
}

function Extract-AssistantText {
    param($Obj)
    try {
        if ($Obj.choices[0].message.content) { return [string]$Obj.choices[0].message.content }
        if ($Obj.choices[0].text) { return [string]$Obj.choices[0].text }
    } catch {}
    return $null
}

function Update-Status {
    param([string]$Message)
    $statusLabel.Text = $Message
    $statusBar.Refresh()
}

# ==========================================================================================
# 3. MAIN UI DEFINITION
# ==========================================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "GenAI Prompt Builder - Mission UI"
$form.Size = New-Object System.Drawing.Size(1050, 950)
$form.MinimumSize = New-Object System.Drawing.Size(850, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White

$fontLabel = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$fontText  = New-Object System.Drawing.Font("Segoe UI", 9)
$fontMono  = New-Object System.Drawing.Font("Consolas", 9)
$tooltip = New-Object System.Windows.Forms.ToolTip

$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusBar.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusBar)

$inputPanel = New-Object System.Windows.Forms.Panel
$inputPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$inputPanel.Height = 700
$inputPanel.AutoScroll = $true

$splitter = New-Object System.Windows.Forms.Splitter
$splitter.Dock = [System.Windows.Forms.DockStyle]::Top
$splitter.Height = 6
$splitter.BackColor = [System.Drawing.Color]::LightGray

$outputTabs = New-Object System.Windows.Forms.TabControl
$outputTabs.Dock = [System.Windows.Forms.DockStyle]::Fill
$outputTabs.Padding = New-Object System.Drawing.Point(10, 3)

$form.Controls.Add($outputTabs)
$form.Controls.Add($splitter)
$form.Controls.Add($inputPanel)

[int]$curY = 15; [int]$padX = 20; [int]$fullWidth = 960
$anchorAll = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$anchorTopRight = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

# --- ROW 1: Actions ---
[int]$btnGenX = $padX + 130
[int]$btnClrX = $padX + 260
[int]$btnLauncherX = $padX + 390

$btnListModels = New-Object System.Windows.Forms.Button; $btnListModels.Text = "List Models"; $btnListModels.Location = New-Object System.Drawing.Point($padX, $curY); $btnListModels.Size = New-Object System.Drawing.Size(120, 30); $inputPanel.Controls.Add($btnListModels)
$btnGenerate = New-Object System.Windows.Forms.Button; $btnGenerate.Text = "Generate"; $btnGenerate.Location = New-Object System.Drawing.Point($btnGenX, $curY); $btnGenerate.Size = New-Object System.Drawing.Size(120, 30); $btnGenerate.BackColor = [System.Drawing.Color]::LightBlue; $inputPanel.Controls.Add($btnGenerate)
$btnClear = New-Object System.Windows.Forms.Button; $btnClear.Text = "Clear Form"; $btnClear.Location = New-Object System.Drawing.Point($btnClrX, $curY); $btnClear.Size = New-Object System.Drawing.Size(120, 30); $inputPanel.Controls.Add($btnClear)

# *** THE NEW PROMPT LAUNCHER BUTTON ***
$btnOpenLauncher = New-Object System.Windows.Forms.Button; $btnOpenLauncher.Text = "🚀 Open Prompt Launcher"; $btnOpenLauncher.Location = New-Object System.Drawing.Point($btnLauncherX, $curY); $btnOpenLauncher.Size = New-Object System.Drawing.Size(200, 30); $btnOpenLauncher.BackColor = [System.Drawing.Color]::LightGreen; $btnOpenLauncher.Font = $fontLabel; $inputPanel.Controls.Add($btnOpenLauncher)
$curY += 45

# --- ROW 2: Templates ---
[int]$saveBtnX = $padX + $fullWidth - 190
[int]$saveBtnY = $curY - 5
$lblTemplate = New-Object System.Windows.Forms.Label; $lblTemplate.Text = "Load Preset Template"; $lblTemplate.Font = $fontLabel; $lblTemplate.Location = New-Object System.Drawing.Point($padX, $curY); $lblTemplate.AutoSize = $true; $inputPanel.Controls.Add($lblTemplate)
$btnSaveTemp = New-Object System.Windows.Forms.Button; $btnSaveTemp.Text = "Save Current as Template"; $btnSaveTemp.Location = New-Object System.Drawing.Point($saveBtnX, $saveBtnY); $btnSaveTemp.Size = New-Object System.Drawing.Size(190, 25); $btnSaveTemp.Anchor = $anchorTopRight; $inputPanel.Controls.Add($btnSaveTemp)
$curY += 20
$comboTemplate = New-Object System.Windows.Forms.ComboBox; $comboTemplate.Location = New-Object System.Drawing.Point($padX, $curY); $comboTemplate.Size = New-Object System.Drawing.Size($fullWidth, 25); $comboTemplate.Anchor = $anchorAll; $comboTemplate.DropDownStyle = "DropDownList"
$comboTemplate.Items.AddRange(@("--- Custom Prompt ---", "1. Document Summarization (BLUF)", "2. Draft Official Memo", "3. Draft Professional Email", "4. Meeting Minutes", "5. Code or Script Review", "6. Data Extraction & Formatting"))
$comboTemplate.SelectedIndex = 0
$inputPanel.Controls.Add($comboTemplate)
$curY += 40

# --- Model, Temp, Tokens ---
[int]$tempX = $padX + 450; [int]$tokenX = $padX + 600
$lblModel = New-Object System.Windows.Forms.Label; $lblModel.Text = "Model ID"; $lblModel.Font = $fontLabel; $lblModel.Location = New-Object System.Drawing.Point($padX, $curY); $lblModel.AutoSize = $true; $inputPanel.Controls.Add($lblModel)
$lblTemp = New-Object System.Windows.Forms.Label; $lblTemp.Text = "Temperature"; $lblTemp.Font = $fontLabel; $lblTemp.Location = New-Object System.Drawing.Point($tempX, $curY); $lblTemp.AutoSize = $true; $inputPanel.Controls.Add($lblTemp); $tooltip.SetToolTip($lblTemp, "0 = Strict/Factual`n1 = Creative/Varied")
$lblMaxTokens = New-Object System.Windows.Forms.Label; $lblMaxTokens.Text = "Max Tokens"; $lblMaxTokens.Font = $fontLabel; $lblMaxTokens.Location = New-Object System.Drawing.Point($tokenX, $curY); $lblMaxTokens.AutoSize = $true; $inputPanel.Controls.Add($lblMaxTokens)
$curY += 20
$comboModels = New-Object System.Windows.Forms.ComboBox; $comboModels.Location = New-Object System.Drawing.Point($padX, $curY); $comboModels.Size = New-Object System.Drawing.Size(430, 25); $inputPanel.Controls.Add($comboModels)
$nudTemp = New-Object System.Windows.Forms.NumericUpDown; $nudTemp.Location = New-Object System.Drawing.Point($tempX, $curY); $nudTemp.Size = New-Object System.Drawing.Size(130, 25); $nudTemp.DecimalPlaces=2; $nudTemp.Increment=0.05; $nudTemp.Maximum=1; $nudTemp.Value=0.3; $inputPanel.Controls.Add($nudTemp)
$nudMaxTokens = New-Object System.Windows.Forms.NumericUpDown; $nudMaxTokens.Location = New-Object System.Drawing.Point($tokenX, $curY); $nudMaxTokens.Size = New-Object System.Drawing.Size(130, 25); $nudMaxTokens.Maximum=32000; $nudMaxTokens.Increment=100; $nudMaxTokens.Value=2000; $inputPanel.Controls.Add($nudMaxTokens)
$curY += 40

# --- Prompt Fields ---
$lblPersona = New-Object System.Windows.Forms.Label; $lblPersona.Text = "Persona or Role"; $lblPersona.Font = $fontLabel; $lblPersona.Location = New-Object System.Drawing.Point($padX, $curY); $lblPersona.AutoSize = $true; $inputPanel.Controls.Add($lblPersona)
$curY += 20; $txtPersona = New-Object System.Windows.Forms.ComboBox; $txtPersona.Location = New-Object System.Drawing.Point($padX, $curY); $txtPersona.Size = New-Object System.Drawing.Size($fullWidth, 25); $txtPersona.Anchor = $anchorAll; $txtPersona.Items.AddRange(@("Professional Technical Writer", "Executive Officer (XO)", "Subject Matter Expert", "Data Analyst", "Legal Counsel", "Project Manager")); $txtPersona.Text = "Professional Technical Writer"; $inputPanel.Controls.Add($txtPersona); $curY += 40

[int]$toneX = $padX + 490; [int]$halfW = 470
$lblAudience = New-Object System.Windows.Forms.Label; $lblAudience.Text = "Audience"; $lblAudience.Font = $fontLabel; $lblAudience.Location = New-Object System.Drawing.Point($padX, $curY); $lblAudience.AutoSize = $true; $inputPanel.Controls.Add($lblAudience)
$lblTone = New-Object System.Windows.Forms.Label; $lblTone.Text = "Tone"; $lblTone.Font = $fontLabel; $lblTone.Location = New-Object System.Drawing.Point($toneX, $curY); $lblTone.AutoSize = $true; $inputPanel.Controls.Add($lblTone)
$curY += 20; $txtAudience = New-Object System.Windows.Forms.ComboBox; $txtAudience.Location = New-Object System.Drawing.Point($padX, $curY); $txtAudience.Size = New-Object System.Drawing.Size($halfW, 25); $txtAudience.Items.AddRange(@("Non-technical stakeholders", "Senior Leadership (GO/SES)", "Technical Staff", "Front-line Operators")); $txtAudience.Text = "Non-technical stakeholders"; $inputPanel.Controls.Add($txtAudience)
$txtTone = New-Object System.Windows.Forms.ComboBox; $txtTone.Location = New-Object System.Drawing.Point($toneX, $curY); $txtTone.Size = New-Object System.Drawing.Size($halfW, 25); $txtTone.Anchor = $anchorAll; $txtTone.Items.AddRange(@("Formal, concise, objective", "Instructional & Detailed", "Urgent & Direct", "Analytical")); $txtTone.Text = "Formal, concise, objective"; $inputPanel.Controls.Add($txtTone); $curY += 40

$lblFormat = New-Object System.Windows.Forms.Label; $lblFormat.Text = "Output Format"; $lblFormat.Font = $fontLabel; $lblFormat.Location = New-Object System.Drawing.Point($padX, $curY); $lblFormat.AutoSize = $true; $inputPanel.Controls.Add($lblFormat)
$lblScope = New-Object System.Windows.Forms.Label; $lblScope.Text = "Scope and Constraints"; $lblScope.Font = $fontLabel; $lblScope.Location = New-Object System.Drawing.Point($toneX, $curY); $lblScope.AutoSize = $true; $inputPanel.Controls.Add($lblScope)
$curY += 20; $txtOutputFormat = New-Object System.Windows.Forms.ComboBox; $txtOutputFormat.Location = New-Object System.Drawing.Point($padX, $curY); $txtOutputFormat.Size = New-Object System.Drawing.Size($halfW, 25); $txtOutputFormat.Items.AddRange(@("Markdown", "Executive Summary (BLUF)", "Official Email Draft", "Situation Report (SITREP)", "JSON Data", "Meeting Minutes")); $txtOutputFormat.Text = "Markdown"; $inputPanel.Controls.Add($txtOutputFormat)
$txtScope = New-Object System.Windows.Forms.TextBox; $txtScope.Location = New-Object System.Drawing.Point($toneX, $curY); $txtScope.Size = New-Object System.Drawing.Size($halfW, 25); $txtScope.Anchor = $anchorAll; $txtScope.Text = "Unclassified only. No PII."; $inputPanel.Controls.Add($txtScope); $curY += 40

$lblGoal = New-Object System.Windows.Forms.Label; $lblGoal.Text = "Goal or Task"; $lblGoal.Font = $fontLabel; $lblGoal.Location = New-Object System.Drawing.Point($padX, $curY); $lblGoal.AutoSize = $true; $inputPanel.Controls.Add($lblGoal)
$curY += 20; $txtGoal = New-Object System.Windows.Forms.TextBox; $txtGoal.Location = New-Object System.Drawing.Point($padX, $curY); $txtGoal.Size = New-Object System.Drawing.Size($fullWidth, 60); $txtGoal.Anchor = $anchorAll; $txtGoal.Multiline = $true; $txtGoal.ScrollBars = "Vertical"; $inputPanel.Controls.Add($txtGoal); $curY += 75

$lblContext = New-Object System.Windows.Forms.Label; $lblContext.Text = "Context Text"; $lblContext.Font = $fontLabel; $lblContext.Location = New-Object System.Drawing.Point($padX, $curY); $lblContext.AutoSize = $true; $inputPanel.Controls.Add($lblContext)
$curY += 20; $txtContext = New-Object System.Windows.Forms.TextBox; $txtContext.Location = New-Object System.Drawing.Point($padX, $curY); $txtContext.Size = New-Object System.Drawing.Size($fullWidth, 80); $txtContext.Anchor = $anchorAll; $txtContext.Multiline = $true; $txtContext.ScrollBars = "Vertical"; $inputPanel.Controls.Add($txtContext); $curY += 95

# --- File Attachments ---
[int]$browseX = $padX + $fullWidth - 100
[int]$browseY = $curY - 5
$lblFiles = New-Object System.Windows.Forms.Label; $lblFiles.Text = "File Attachments (Browse or Drag & Drop)"; $lblFiles.Font = $fontLabel; $lblFiles.Location = New-Object System.Drawing.Point($padX, $curY); $lblFiles.AutoSize = $true; $inputPanel.Controls.Add($lblFiles)
$btnBrowse = New-Object System.Windows.Forms.Button; $btnBrowse.Text = "Browse..."; $btnBrowse.Location = New-Object System.Drawing.Point($browseX, $browseY); $btnBrowse.Size = New-Object System.Drawing.Size(100, 25); $btnBrowse.Anchor = $anchorTopRight; $inputPanel.Controls.Add($btnBrowse)
$curY += 25
$txtFiles = New-Object System.Windows.Forms.TextBox; $txtFiles.Location = New-Object System.Drawing.Point($padX, $curY); $txtFiles.Size = New-Object System.Drawing.Size($fullWidth, 50); $txtFiles.Anchor = $anchorAll; $txtFiles.Multiline = $true; $txtFiles.ScrollBars = "Vertical"; $txtFiles.AllowDrop = $true; $inputPanel.Controls.Add($txtFiles)

# Drag and Drop Logic
$txtFiles.Add_DragEnter({ if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy } })
$txtFiles.Add_DragDrop({
    $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    foreach ($f in $files) {
        if ([string]::IsNullOrWhiteSpace($txtFiles.Text)) { $txtFiles.Text = $f } else { $txtFiles.Text += "`r`n$f" }
    }
})

$btnBrowse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Multiselect = $true
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($f in $ofd.FileNames) {
            if ([string]::IsNullOrWhiteSpace($txtFiles.Text)) { $txtFiles.Text = $f } else { $txtFiles.Text += "`r`n$f" }
        }
    }
})

# --- Output Tabs & Actions ---
$tabResponse = New-Object System.Windows.Forms.TabPage; $tabResponse.Text = "Assistant Response"; $outputTabs.TabPages.Add($tabResponse)
$tabRequest = New-Object System.Windows.Forms.TabPage; $tabRequest.Text = "Request Payload"; $outputTabs.TabPages.Add($tabRequest)
$tabRaw = New-Object System.Windows.Forms.TabPage; $tabRaw.Text = "Raw JSON Output"; $outputTabs.TabPages.Add($tabRaw)

$pnlResponseActions = New-Object System.Windows.Forms.Panel; $pnlResponseActions.Dock = "Top"; $pnlResponseActions.Height = 35; $pnlResponseActions.BackColor = [System.Drawing.Color]::WhiteSmoke
$btnCopy = New-Object System.Windows.Forms.Button; $btnCopy.Text = "Copy to Clipboard"; $btnCopy.Location = New-Object System.Drawing.Point(10, 5); $btnCopy.Size = New-Object System.Drawing.Size(150, 25); $pnlResponseActions.Controls.Add($btnCopy)
$btnExport = New-Object System.Windows.Forms.Button; $btnExport.Text = "Export to File"; $btnExport.Location = New-Object System.Drawing.Point(170, 5); $btnExport.Size = New-Object System.Drawing.Size(130, 25); $pnlResponseActions.Controls.Add($btnExport)
$tabResponse.Controls.Add($pnlResponseActions)

$rtOutput = New-Object System.Windows.Forms.RichTextBox; $rtOutput.Dock = "Fill"; $rtOutput.ReadOnly = $true; $rtOutput.Font = $fontText; $rtOutput.BackColor = [System.Drawing.Color]::White; $tabResponse.Controls.Add($rtOutput)
$rtRequest = New-Object System.Windows.Forms.RichTextBox; $rtRequest.Dock = "Fill"; $rtRequest.ReadOnly = $true; $rtRequest.Font = $fontMono; $rtRequest.BackColor = [System.Drawing.Color]::WhiteSmoke; $tabRequest.Controls.Add($rtRequest)
$rtRaw = New-Object System.Windows.Forms.RichTextBox; $rtRaw.Dock = "Fill"; $rtRaw.ReadOnly = $true; $rtRaw.Font = $fontMono; $rtRaw.BackColor = [System.Drawing.Color]::WhiteSmoke; $tabRaw.Controls.Add($rtRaw)

$rtOutput.BringToFront()

# ==========================================================================================
# 4. MAIN APP LOGIC 
# ==========================================================================================

# *** CLICK EVENT FOR THE NEW LAUNCHER BUTTON ***
$btnOpenLauncher.Add_Click({
    Show-CopilotPromptLauncher
})

# Copy & Export Logic
$btnCopy.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($rtOutput.Text)) {
        [System.Windows.Forms.Clipboard]::SetText($rtOutput.Text)
        Update-Status "Response copied to clipboard!"
    }
})

$btnExport.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($rtOutput.Text)) {
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "Text Files (*.txt)|*.txt|Markdown Files (*.md)|*.md|All Files (*.*)|*.*"
        $sfd.FileName = "GenAI_Response.txt"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Set-Content -Path $sfd.FileName -Value $rtOutput.Text -Encoding UTF8
            Update-Status "Saved to $($sfd.FileName)"
        }
    }
})

# Template Logic
function Load-CustomTemplates {
    if (Test-Path $templateFile) {
        try {
            $templates = Get-Content $templateFile -Raw | ConvertFrom-Json
            foreach ($t in $templates) {
                if (-not $comboTemplate.Items.Contains($t.Name)) { [void]$comboTemplate.Items.Add($t.Name) }
            }
        } catch { Update-Status "Failed to load custom templates." }
    }
}

$btnSaveTemp.Add_Click({
    $name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a name for this Custom Template:", "Save Template", "My Custom Template")
    if ([string]::IsNullOrWhiteSpace($name)) { return }
    
    $newTemplate = @{
        Name = $name; Persona = $txtPersona.Text; Audience = $txtAudience.Text
        Tone = $txtTone.Text; Format = $txtOutputFormat.Text; Scope = $txtScope.Text; Goal = $txtGoal.Text
    }
    
    $existing = @()
    if (Test-Path $templateFile) { $existing = @(Get-Content $templateFile -Raw | ConvertFrom-Json) }
    $existing = $existing | Where-Object { $_.Name -ne $name }
    $existing += $newTemplate
    $existing | ConvertTo-Json -Depth 5 | Set-Content $templateFile -Encoding UTF8
    
    Update-Status "Template '$name' saved!"
    Load-CustomTemplates
    $comboTemplate.Text = $name
})

$comboTemplate.Add_SelectedIndexChanged({
    $selected = $comboTemplate.Text
    if ($selected -match "1. Document Summarization") {
        $txtPersona.Text = "Intelligence Analyst"; $txtAudience.Text = "Senior Leadership (GO/SES)"; $txtTone.Text = "Formal, concise, objective"; $txtOutputFormat.Text = "Executive Summary (BLUF)"; $txtScope.Text = "Unclassified only. Highlight key decisions and timelines."; $txtGoal.Text = "Read the provided text and write a summary. Place the Bottom Line Up Front (BLUF) at the top, followed by 3 to 5 key bullet points."
    } elseif ($selected -match "2. Draft Official Memo") {
        $txtPersona.Text = "Executive Officer (XO)"; $txtAudience.Text = "General Unit Personnel"; $txtTone.Text = "Instructional, professional, authoritative"; $txtOutputFormat.Text = "Official Memorandum format"; $txtScope.Text = "Unclassified only. Be direct. Avoid overly complex jargon."; $txtGoal.Text = "Draft an official memorandum addressing the topic provided in the context. Ensure clear action items and deadlines are stated."
    } elseif ($selected -match "3. Draft Professional Email") {
        $txtPersona.Text = "Action Officer"; $txtAudience.Text = "Designated Recipients"; $txtTone.Text = "Professional, clear, action-oriented"; $txtOutputFormat.Text = "Official Email Draft"; $txtScope.Text = "Unclassified only. Include BLUF. Clearly state any required actions."; $txtGoal.Text = "Draft a professional email based on the provided context. Include a clear, descriptive subject line at the top."
    } elseif ($selected -match "4. Meeting Minutes") {
        $txtPersona.Text = "Executive Assistant"; $txtAudience.Text = "Meeting Attendees and Leadership"; $txtTone.Text = "Objective, clear, concise"; $txtOutputFormat.Text = "Meeting Minutes (Markdown)"; $txtScope.Text = "Unclassified. Capture attendees, key discussion points, decisions made, and action items (with POC and deadline)."; $txtGoal.Text = "Review the provided meeting transcript/notes and generate formal meeting minutes."
    } elseif ($selected -match "5. Code or Script Review") {
        $txtPersona.Text = "Senior Software Engineer"; $txtAudience.Text = "Technical Staff"; $txtTone.Text = "Analytical, constructive, detailed"; $txtOutputFormat.Text = "Markdown with code blocks"; $txtScope.Text = "Identify bugs, security vulnerabilities, and inefficiencies."; $txtGoal.Text = "Review the provided script or code. Explain any errors, suggest performance optimizations, and provide a refactored version of the code."
    } elseif ($selected -match "6. Data Extraction") {
        $txtPersona.Text = "Data Analyst"; $txtAudience.Text = "Technical Staff"; $txtTone.Text = "Objective, analytical"; $txtOutputFormat.Text = "JSON Data or Markdown Table"; $txtScope.Text = "Extract ONLY the requested data points. Do not hallucinate data."; $txtGoal.Text = "Analyze the provided unstructured text or file. Extract key metrics (Dates, Names, Locations, Values) and format them cleanly as requested."
    } elseif ($selected -eq "--- Custom Prompt ---") {
        $txtGoal.Text = ""
    } else {
        if (Test-Path $templateFile) {
            $templates = Get-Content $templateFile -Raw | ConvertFrom-Json
            $match = $templates | Where-Object { $_.Name -eq $selected }
            if ($match) {
                $txtPersona.Text = $match.Persona; $txtAudience.Text = $match.Audience; $txtTone.Text = $match.Tone; $txtOutputFormat.Text = $match.Format; $txtScope.Text = $match.Scope; $txtGoal.Text = $match.Goal
            }
        }
    }
})

$btnListModels.Add_Click({
    $baseUrl = [System.Environment]::GetEnvironmentVariable('GENAI_BASE_URL','User')
    $apiKey  = [System.Environment]::GetEnvironmentVariable('GENAI_API_KEY','User')
    if (!$baseUrl -or !$apiKey) { [System.Windows.Forms.MessageBox]::Show("Set GENAI_BASE_URL and GENAI_API_KEY."); return }
    
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; Update-Status "Fetching models..."
    $client = New-HttpClient
    $normBase = Get-NormalizedBaseUrl $baseUrl
    $allIds = @()
    foreach ($path in @('/v1/models', '/models', '/openai/deployments')) {
        try {
            $client.DefaultRequestHeaders.Clear()
            $client.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue('Bearer', $apiKey)
            $resp = $client.GetAsync("$normBase$path").GetAwaiter().GetResult()
            if ($resp.IsSuccessStatusCode) {
                $json = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                $allIds += Normalize-ModelIds (ConvertFrom-Json -InputObject $json)
            }
        } catch { }
    }
    $comboModels.Items.Clear()
    if ($allIds.Count -gt 0) { 
        $allIds | Sort-Object -Unique | ForEach-Object { [void]$comboModels.Items.Add($_) }
        
        if (Test-Path $configFile) {
            try {
                $cfg = Get-Content $configFile -Raw | ConvertFrom-Json
                if ($allIds -contains $cfg.Model) { $comboModels.Text = $cfg.Model } else { $comboModels.SelectedIndex = 0 }
                $nudTemp.Value = $cfg.Temp; $nudMaxTokens.Value = $cfg.MaxTokens
            } catch { $comboModels.SelectedIndex = 0 }
        } else { $comboModels.SelectedIndex = 0 }
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default; Update-Status "Ready"
    if ($client) { $client.Dispose() }
})

$btnGenerate.Add_Click({
    $baseUrl = [System.Environment]::GetEnvironmentVariable('GENAI_BASE_URL','User')
    $apiKey  = [System.Environment]::GetEnvironmentVariable('GENAI_API_KEY','User')
    if (!$baseUrl -or !$apiKey) { [System.Windows.Forms.MessageBox]::Show("Check Environment Variables."); return }
    
    $cfgObj = @{ Model = $comboModels.Text; Temp = $nudTemp.Value; MaxTokens = $nudMaxTokens.Value }
    $cfgObj | ConvertTo-Json | Set-Content $configFile -Encoding UTF8

    $sys = "Persona: $($txtPersona.Text)`nAudience: $($txtAudience.Text)`nTone: $($txtTone.Text)`nFormat: $($txtOutputFormat.Text)`nConstraints: $($txtScope.Text)"
    $usr = "GOAL:`n$($txtGoal.Text)`n`nCONTEXT:`n$($txtContext.Text)"

    $filePaths = $txtFiles.Lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if ($filePaths.Length -gt 0) {
        $usr += "`n`n--- ATTACHED FILES ---"
        foreach ($path in $filePaths) {
            $cleanPath = $path.Trim() -replace '^"|"$', ''
            if (Test-Path -LiteralPath $cleanPath -PathType Leaf) {
                $fileContent = Get-Content -LiteralPath $cleanPath -Raw
                $usr += "`nFile: $cleanPath`nContent:`n$fileContent`n"
            } else {
                $usr += "`nFile: $cleanPath`n[ERROR: File not found]`n"
            }
        }
    }

    $payload = @{
        model = $comboModels.Text
        messages = @( @{role="system"; content=$sys}, @{role="user"; content=$usr} )
        temperature = [double]$nudTemp.Value
        max_tokens = [int]$nudMaxTokens.Value
    }
    $jsonPay = $payload | ConvertTo-Json -Depth 10
    $rtRequest.Text = $jsonPay
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; Update-Status "Generating response..."
    $client = New-HttpClient
    try {
        $client.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue('Bearer', $apiKey)
        $content = New-Object System.Net.Http.StringContent($jsonPay, [System.Text.Encoding]::UTF8, "application/json")
        $url = "$(Get-NormalizedBaseUrl $baseUrl)/v1/chat/completions"
        $resp = $client.PostAsync($url, $content).GetAwaiter().GetResult()
        $body = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $rtRaw.Text = $body
        if ($resp.IsSuccessStatusCode) {
            $rtOutput.Text = Extract-AssistantText (ConvertFrom-Json -InputObject $body)
            $outputTabs.SelectedTab = $tabResponse
            Update-Status "Success!"
        } else {
            $rtOutput.Text = "ERROR: $($resp.StatusCode)`n$body"
            $outputTabs.SelectedTab = $tabRaw
            Update-Status "Failed with error code $($resp.StatusCode)"
        }
    } catch { $rtOutput.Text = "Exception: $($_.Exception.Message)"; Update-Status "Error executing request." }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    if ($client) { $client.Dispose() }
})

$btnClear.Add_Click({
    $comboTemplate.SelectedIndex = 0
    $txtGoal.Text = ""; $txtContext.Text = ""; $txtFiles.Text = ""; $rtOutput.Text = ""; $rtRaw.Text = ""; $rtRequest.Text = ""
    Update-Status "Form cleared."
})

# Boot Logic
$form.Add_Shown({
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    Load-CustomTemplates
    $btnListModels.PerformClick()
})

[void]$form.ShowDialog()
$form.Dispose()
