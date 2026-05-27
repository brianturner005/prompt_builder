# GenAI Prompt Builder (CLI) - Windows PowerShell 5.1
# - OpenAI-style payload to /v1/chat/completions
# - Proxy-aware using default domain credentials (fixes HTTP 407)
# - Base URL normalization (strips accidental endpoint suffixes)
# - Optional model discovery function

# ---------------------------------------
# POLICY BANNER (single-quoted here-string)
# ---------------------------------------
$PolicyBanner = @'
[IL5/CUI Reminder]
- Approved for Controlled Unclassified Information (CUI) on IL5 platforms.
- Do NOT include PII/PHI or classified data.
- Validate all AI output; non-deterministic, may be inaccurate or biased.
'@

# ---------------------------------------
# Session-level proxy credentials (PS 5.1)
# ---------------------------------------
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

# ---------------------------------------
# Utilities
# ---------------------------------------
function Normalize-BaseUrl {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $Url }
    $u = $Url.TrimEnd('/')

    # Strip common resource suffixes if a full endpoint was pasted by mistake
    $badSuffixes = @(
        '/v1/chat/completions',
        '/chat/completions',
        '/v1/models',
        '/models',
        '/server/v1/models',
        '/server/models',
        '/get-models',
        '/v1/generate',
        '/server/generate'
    )
    foreach ($s in $badSuffixes) {
        if ($u.EndsWith($s)) { $u = $u.Substring(0, $u.Length - $s.Length); break }
    }
    return $u
}

function Get-GenAIConfig {
    param(
        [string]$BaseUrl = $env:GENAI_BASE_URL,
        [string]$ApiKey  = $env:GENAI_API_KEY
    )
    if ([string]::IsNullOrWhiteSpace($BaseUrl)) { throw "GENAI_BASE_URL is not set." }
    if ([string]::IsNullOrWhiteSpace($ApiKey))  { throw "GENAI_API_KEY is not set." }
    $BaseUrl = Normalize-BaseUrl $BaseUrl
    [pscustomobject]@{ BaseUrl = $BaseUrl.TrimEnd('/'); ApiKey = $ApiKey }
}

function New-PromptProfile {
    param(
        [string]$Model        = "gemini-2.5-pro",  # Use an ID from your tenant's model list
        [string]$OutputFormat = "markdown",
        [string]$Goal,
        [string]$Audience,
        [string]$Persona,
        [string]$Tone         = "formal, concise, evidence-backed",
        [string]$Constraints  = "Avoid assumptions; keep content unclassified.",
        [double]$Temperature  = 0.3,
        [int]$MaxTokens       = 300,
        [string]$ContextText  = "",
        [string[]]$Files      = @(),
        [string]$Examples     = ""
    )

    $gen = @{
        temperature       = [Math]::Round([double]$Temperature, 2)
        max_output_tokens = [int]$MaxTokens
    }

    $profile = [ordered]@{
        model             = $Model
        output_format     = $OutputFormat
        goal              = $Goal
        audience          = $Audience
        persona           = $Persona
        tone              = $Tone
        scope_constraints = $Constraints
        generation        = $gen
        context           = @(
            @{ type='text'; title='Context'; content=$ContextText }
        )
        files             = $Files
        examples          = $Examples
        policy_banner     = $PolicyBanner
    }

    return ($profile | ConvertTo-Json -Depth 6)
}

function Compose-SystemAndUserText {
    param([string]$ProfileJson)

    $p  = $ProfileJson | ConvertFrom-Json
    $nl = [Environment]::NewLine

    $system =
        "You are a $($p.persona) writing for $($p.audience)." + $nl +
        "Tone: $($p.tone)" + $nl +
        "Scope and Constraints: $($p.scope_constraints)" + $nl +
        "Output format: $($p.output_format)" + $nl +
        "Follow CUI/IL5 guardrails and provide citations when summarizing source docs."

    $ctxText   = ($p.context | Where-Object { $_.type -eq 'text' } | ForEach-Object { $_.content }) -join $nl
    $filesList = ($p.files   | ForEach-Object { "- $($_)" }) -join $nl

    $user =
        "Goal or Task:" + $nl +
        "$($p.goal)" + $nl + $nl +
        "Context:" + $nl +
        "$ctxText" + $nl + $nl +
        "Files referenced:" + $nl +
        "$filesList" + $nl + $nl +
        "Examples (few-shot):" + $nl +
        "$($p.examples)"

    [pscustomobject]@{ system = $system; user = $user }
}

function Invoke-GenAI {
    param(
        [Parameter(Mandatory)][string]$BaseUrl,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter(Mandatory)][string]$ProfileJson,
        [string]$GeneratePath = "/v1/chat/completions"   # OpenAI-style route confirmed by your probe
    )

    $uri = "$($BaseUrl.TrimEnd('/'))$GeneratePath"

    # Defensive typing
    $p = $ProfileJson | ConvertFrom-Json
    $p.generation.temperature       = [double]$p.generation.temperature
    $p.generation.max_output_tokens = [int]   $p.generation.max_output_tokens

    # System + user strings
    $sysUser = Compose-SystemAndUserText -ProfileJson $ProfileJson

    # ----- OpenAI-style payload -----
    $payloadObj = [ordered]@{
        model    = $p.model
        messages = @(
            @{ role = "system"; content = $sysUser.system },
            @{ role = "user";   content = $sysUser.user   }
        )
        temperature = $p.generation.temperature
        max_tokens  = $p.generation.max_output_tokens
        # Optional: remove or adjust if your gateway rejects it
        # response_format = @{ type = "json_object" }
    }
    $json = $payloadObj | ConvertTo-Json -Depth 6

    # Headers
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type"  = "application/json"
    }

    # Determine proxy for THIS URL (PS 5.1)
    $proxyUriObj = [System.Net.WebRequest]::DefaultWebProxy.GetProxy([Uri]$uri)
    $proxyUrl    = $null
    if ($proxyUriObj -and $proxyUriObj.AbsoluteUri -and $proxyUriObj.AbsoluteUri -ne $uri) {
        $proxyUrl = $proxyUriObj.AbsoluteUri
    }

    # Debug
    Write-Host "DEBUG → POST $uri" -ForegroundColor Yellow
    if ($proxyUrl) { Write-Host "DEBUG → Proxy: $proxyUrl (default credentials)" -ForegroundColor Yellow }
    else           { Write-Host "DEBUG → No proxy (direct)" -ForegroundColor Yellow }

    try {
        if ($proxyUrl) {
            $resp = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $json `
                    -Proxy $proxyUrl -ProxyUseDefaultCredentials `
                    -UseDefaultCredentials `
                    -ErrorAction Stop
        } else {
            $resp = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $json `
                    -UseDefaultCredentials `
                    -ErrorAction Stop
        }
        return $resp
    } catch {
        throw "GENAI call failed: $($_.Exception.Message)"
    }
}

function Save-ProfileJson {
    param([string]$ProfileJson)
    $path = Join-Path $env:USERPROFILE ("PromptProfile_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".json")
    $ProfileJson | Set-Content -Path $path -Encoding UTF8
    return $path
}

# ---------------------------------------
# Optional: Model discovery (probe paths)
# ---------------------------------------
function Get-GenAIModels {
    param(
        [Parameter(Mandatory)][string]$BaseUrl,
        [Parameter(Mandatory)][string]$ApiKey,
        [switch]$VerboseUrls
    )

    $base    = $BaseUrl.TrimEnd('/')
    $headers = @{ Authorization = "Bearer $ApiKey" }

    $candidates = @(
        "$base/v1/models",
        "$base/models",
        "$base/server/v1/models",
        "$base/server/models",
        "$base/get-models"
    )

    $resp = $null; $hit = $null
    foreach ($u in $candidates) {
        if ($VerboseUrls) { Write-Host "Trying: $u" -ForegroundColor Yellow }
        try {
            $resp = Invoke-RestMethod -Uri $u -Headers $headers -Method GET -ErrorAction Stop
            $hit  = $u
            break
        } catch { if ($VerboseUrls) { Write-Host "  -> $($_.Exception.Message)" -ForegroundColor DarkGray } }
    }

    if (-not $resp) { throw "No model listing endpoint found. Tried: $($candidates -join ', ')" }

    $models =
        if ($resp.models) { $resp.models }
        elseif ($resp.data) { $resp.data }
        elseif ($resp -is [System.Collections.IEnumerable]) { $resp }
        else { @($resp) }

    $list = foreach ($m in $models) {
        $id   = if ($m.PSObject.Properties['id']) { $m.id } elseif ($m.PSObject.Properties['model']) { $m.model } elseif ($m.PSObject.Properties['name']) { $m.name } else { "<unknown>" }
        $name = if ($m.PSObject.Properties['displayName']) { $m.displayName } elseif ($m.PSObject.Properties['title']) { $m.title } elseif ($m.PSObject.Properties['name']) { $m.name } else { $null }
        $desc = if ($m.PSObject.Properties['description']) { $m.description } else { $null }
        [pscustomobject]@{ Id = $id; DisplayName = $name; Description = $desc }
    }

    Write-Host "Discovered via: $hit" -ForegroundColor Cyan
    $list | Format-Table -AutoSize
    return $list
}

# ---------------------------------------
# MAIN (CLI)
# ---------------------------------------
try {
    Write-Host ("="*70)
    Write-Host $PolicyBanner -ForegroundColor Green
    Write-Host ("="*70)

    $cfg = Get-GenAIConfig

    # Optional: show models first (press Enter to skip)
    $seeModels = Read-Host "Press Enter to skip listing models, or type 'list' to show them"
    if ($seeModels -eq 'list') {
        $models = Get-GenAIModels -BaseUrl $cfg.BaseUrl -ApiKey $cfg.ApiKey -VerboseUrls
        Write-Host "`nEnter a model id from the list above." -ForegroundColor Cyan
    }

    $model = Read-Host "Model id (e.g., gemini-2.5-pro)"
    $fmt   = Read-Host "Output format (markdown or json)"
    $goal  = Read-Host "Goal or Task"
    $aud   = Read-Host "Audience"
    $per   = Read-Host "Persona or Role"
    $tone  = Read-Host "Tone"
    $cons  = Read-Host "Scope and constraints"
    $temp  = Read-Host "Temperature (0 to 1, e.g., 0.3)"
    $tok   = Read-Host "Max output tokens (e.g., 300)"
    $ctx   = Read-Host "Context text (paste, keep under 10000 chars)"
    $ex    = Read-Host "Examples (few-shot; optional)"

    $files = @()
    while ($true) {
        $f = Read-Host "Add file path (press Enter to finish)"
        if ([string]::IsNullOrWhiteSpace($f)) { break }
        if (Test-Path $f) { $files += $f } else { Write-Warning "Not found: $f" }
    }

    $profile = New-PromptProfile `
        -Model $model `
        -OutputFormat $fmt `
        -Goal $goal `
        -Audience $aud `
        -Persona $per `
        -Tone $tone `
        -Constraints $cons `
        -Temperature ([double]$temp) `
        -MaxTokens ([int]$tok) `
        -ContextText $ctx `
        -Files $files `
        -Examples $ex

    $path = Save-ProfileJson -ProfileJson $profile
    Write-Host "Saved profile to: $path" -ForegroundColor Cyan

    # Call using OpenAI-style route (confirmed by your probe)
    $resp = Invoke-GenAI -BaseUrl $cfg.BaseUrl -ApiKey $cfg.ApiKey -ProfileJson $profile -GeneratePath "/v1/chat/completions"
    $resp | ConvertTo-Json -Depth 8
}
catch {
    Write-Error $_.Exception.Message
}