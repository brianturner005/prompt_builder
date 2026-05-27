Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Copilot Prompt Launcher"
$form.Size = New-Object System.Drawing.Size(700, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# --- THE PROMPT LIBRARY ---
$global:promptLibrary = @(
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

# --- UI LOGIC ---

# Platform Selection
$lblPlatform = New-Object System.Windows.Forms.Label
$lblPlatform.Text = "1. Select Platform:"
$lblPlatform.Location = New-Object System.Drawing.Point(20, 20)
$lblPlatform.AutoSize = $true
$lblPlatform.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblPlatform)

$cmbPlatform = New-Object System.Windows.Forms.ComboBox
$cmbPlatform.Location = New-Object System.Drawing.Point(20, 45)
$cmbPlatform.Size = New-Object System.Drawing.Size(300, 25)
$cmbPlatform.DropDownStyle = "DropDownList"
@("DoW GenAI", "Microsoft Copilot for M365") | ForEach-Object { [void]$cmbPlatform.Items.Add($_) }
$cmbPlatform.SelectedIndex = 0
$form.Controls.Add($cmbPlatform)

# Category Selection
$lblCategory = New-Object System.Windows.Forms.Label
$lblCategory.Text = "2. Select Category:"
$lblCategory.Location = New-Object System.Drawing.Point(340, 20)
$lblCategory.AutoSize = $true
$lblCategory.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblCategory)

$cmbCategory = New-Object System.Windows.Forms.ComboBox
$cmbCategory.Location = New-Object System.Drawing.Point(340, 45)
$cmbCategory.Size = New-Object System.Drawing.Size(320, 25)
$cmbCategory.DropDownStyle = "DropDownList"
@("Summarization", "Drafting", "Analysis", "Explanation") | ForEach-Object { [void]$cmbCategory.Items.Add($_) }
$cmbCategory.SelectedIndex = 0
$form.Controls.Add($cmbCategory)

# Prompt List
$lblPrompts = New-Object System.Windows.Forms.Label
$lblPrompts.Text = "3. Select a Prompt to Discover:"
$lblPrompts.Location = New-Object System.Drawing.Point(20, 90)
$lblPrompts.AutoSize = $true
$lblPrompts.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblPrompts)

$lstPrompts = New-Object System.Windows.Forms.ListBox
$lstPrompts.Location = New-Object System.Drawing.Point(20, 115)
$lstPrompts.Size = New-Object System.Drawing.Size(640, 180)
$lstPrompts.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lstPrompts)

# Prompt Details
$txtPromptDetail = New-Object System.Windows.Forms.TextBox
$txtPromptDetail.Location = New-Object System.Drawing.Point(20, 310)
$txtPromptDetail.Size = New-Object System.Drawing.Size(640, 160)
$txtPromptDetail.Multiline = $true
$txtPromptDetail.ScrollBars = "Vertical"
$txtPromptDetail.ReadOnly = $true
$txtPromptDetail.BackColor = "White"
$txtPromptDetail.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($txtPromptDetail)

# Copy Button
$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "📋 Copy Prompt to Clipboard"
$btnCopy.Location = New-Object System.Drawing.Point(20, 490)
$btnCopy.Size = New-Object System.Drawing.Size(250, 45)
$btnCopy.BackColor = [System.Drawing.Color]::LightBlue
$btnCopy.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnCopy)

# --- Event Handlers ---

function Update-PromptList {
    $lstPrompts.Items.Clear()
    $txtPromptDetail.Text = ""
    $selectedPlatform = $cmbPlatform.SelectedItem
    $selectedCategory = $cmbCategory.SelectedItem
    $filteredPrompts = $global:promptLibrary | Where-Object { $_.Platform -eq $selectedPlatform -and $_.Category -eq $selectedCategory }
    foreach ($p in $filteredPrompts) { [void]$lstPrompts.Items.Add($p.Name) }
}

$cmbPlatform.Add_SelectedIndexChanged({ Update-PromptList })
$cmbCategory.Add_SelectedIndexChanged({ Update-PromptList })

$lstPrompts.Add_SelectedIndexChanged({
    if ($lstPrompts.SelectedItem) {
        $p = $global:promptLibrary | Where-Object { $_.Name -eq $lstPrompts.SelectedItem -and $_.Platform -eq $cmbPlatform.SelectedItem }
        if ($p) { $txtPromptDetail.Text = $p.Prompt }
    }
})

$btnCopy.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($txtPromptDetail.Text)) {
        [System.Windows.Forms.Clipboard]::SetText($txtPromptDetail.Text)
        [System.Windows.Forms.MessageBox]::Show("Prompt copied successfully!", "System Notification", "OK", "Information")
    }
})

Update-PromptList
[void]$form.ShowDialog()
