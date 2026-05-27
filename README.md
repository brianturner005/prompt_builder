🛡️ GenAI Integrated Suite - Mission UI
A secure, IL5-optimized PowerShell graphical interface that combines two powerful tools: a Prompt Builder for creating detailed, structured requests and a Prompt Launcher for discovering high-quality, mission-ready prompts.

🚀 Quick Start
Set Environment Variables: Ensure your workstation has your API credentials set at the User level:

GENAI_BASE_URL: The endpoint for your authorized GenAI gateway.

GENAI_API_KEY: Your mission-authorized API key.

Run the Script:

Save the final script as GenAI_Integrated_Suite.ps1.

Right-click the file and select Run with PowerShell.

⚙️ Operation
Click List Models to discover available mission models in the Prompt Builder.

Click OPEN PROMPT LAUNCHER to find and copy pre-built prompts for common tasks.

Paste a copied prompt into the "Goal or Task" field, or build your own from scratch.

Use the dropdowns and text boxes to refine your request.

Click Generate to submit the final prompt to the API.

🛠️ Key Features
Feature	Description
Integrated Prompt Launcher	Access a pop-up library of curated prompts for both the DoW GenAI and M365 Copilot platforms to kick-start your workflow.
Guided Prompting	Dropdown menus for Persona, Audience, Tone, and Output Format help non-engineers build expert-level prompts.
File Injection	Use the "Browse" or Drag-and-Drop functionality to attach .txt, .csv, or other text-based files. Content is automatically injected into the context.
Custom Templates	Save your most-used prompt configurations as a custom template for one-click loading.
Mission Ready	Proxy-aware and TLS 1.2/1.3 hardened for operation behind enterprise gateways.
 
✏️ Customization
You can easily change the default presets to better fit your specific mission or unit by editing the GenAI_Integrated_Suite.ps1 file in a text editor.

1. Customizing the Prompt Builder Dropdowns:

Locate the MAIN UI DEFINITION section and find the Items.AddRange commands.

Example - Modifying the Persona Dropdown:

powershell
# Find this line:
$txtPersona.Items.AddRange(@("Professional Technical Writer", "Executive Officer (XO)", "Data Analyst"))

# Change it to add your own roles:
$txtPersona.Items.AddRange(@("J2 Intelligence Analyst", "Cyber Security Engineer", "Logistics Officer"))
2. Customizing the Prompt Launcher Library:

Locate the INTEGRATED PROMPT LAUNCHER FUNCTION section and find the $promptLibrary array. You can add, edit, or remove [PSCustomObject] blocks to manage the prompt library.

Example - Adding a New Prompt:

powershell
$promptLibrary = @(
    # ... existing prompts ...

    # Add your new prompt here
    [PSCustomObject]@{ Platform="DoW GenAI"; Category="Analysis"; Name="New Custom Analysis"; Prompt="Analyze the data for anomalies and report findings." }
)
🔒 Security & Compliance
Classification: This tool is for Unclassified/CUI operations only. Do not process classified data through this interface.

PII/PHI: Users are prohibited from injecting real PII (Social Security Numbers, etc.) or PHI (Medical Records) into the prompt fields.

Proxy Support: Uses System.Net.WebRequest::GetSystemWebProxy() to ensure traffic is correctly routed through authorized network egress points.

📂 Requirements
PowerShell: Version 5.1 (Standard on Windows 10/11).

Network: Access to the GENAI_BASE_URL endpoint.

Permissions: Ability to execute local PowerShell scripts.
