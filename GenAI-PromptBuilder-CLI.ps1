<#
  GenAI-PromptBuilder-CLI.ps1
  Windows PowerShell 5.1 CLI Prompt Builder for GENAI tenants

  - OpenAI-style payload to /v1/chat/completions
  - Works behind NIPR enterprise proxies (uses DefaultNetworkCredentials)
  - Automatically uses HttpClient if available; falls back to Invoke-RestMethod if not
  - Base URL normalization (prevents accidental /v1/... suffixes)
  - Optional model discovery (probes common listing paths)
  - Progress + elapsed time + assistant-text extraction
#>

# ---------------------------------------
# POLICY BANNER (single-quoted here-string; no interpolation)
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

# Try to load System.Net.Http; if unavailable, we will fallback later
try { Add-Type -AssemblyName System.Net.Http } catch { }

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
        [string]$Model        = "gemini-2.5-pro",  # Use an ID from your tenant's list
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
        context           = @(@{ type='text'; title='Context'; content=$ContextText })
        files             = $Files
        examples          = $Examples
        policy_banner     = $PolicyBanner
    }
    return ($profile | ConvertTo-Json -Depth 6)
}

function Compose-SystemAndUserText {
    param([string]$ProfileJson)

    $p  = $ProfileJson | ConvertFrom-Json

    # Build system content (explicit line breaks)
    $system  = "You are a $($p.persona) writing for $($p.audience)."
    $system += "`nTone: $($p.tone)"
    $system += "`nScope and Constraints: $($p.scope_constraints)"
    $system += "`nOutput format: $($p.output_format)"
    $system += "`nFollow CUI/IL5 guardrails and provide citations when summarizing source docs."

    # Context blocks (text only for now)
    $ctxText   = ($p.context | Where-Object { $_.type -eq 'text' } | ForEach-Object { $_.content }) -join "`n"
    $filesList = ($p.files   | ForEach-Object { "- $($_)" }) -join "`n"

    $user  = "Goal or Task:`n$($p.goal)`n"
    $user += "`nContext:`n$ctxText`n"
    $user += "`nFiles referenced:`n$filesList`n"
    $user += "`nExamples (few-shot):`n$($p.examples)"

    [pscustomobject]@{ system = $system; user = $user }
}

function Get-AssistantText {
    param($Resp)
    if ($Resp -and $Resp.choices -and $Resp.choices.Count -gt 0) {
        return $Resp.choices[0].message.content
    }
    # Fallback: raw JSON
    try { return ($Resp | ConvertTo-Json -Depth 8) } catch { return "$Resp" }
}

function Ensure-HttpClientAvailable {
    try {
        # Try creating and disposing a client instance
        $tmp = [System.Net.Http.HttpClient]::new()
        $tmp.Dispose()
        return $true
    } catch {
        try {
            Add-Type -AssemblyName System.Net.Http -ErrorAction Stop
            $tmp = [System.Net.Http.HttpClient]::new()
            $tmp.Dispose()
            return $true
        } catch {
            return $false
        }
    }
}

function Invoke-GenAI {
    param(
        [Parameter(Mandatory)][string]$BaseUrl,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter(Mandatory)][string]$ProfileJson,
        [string]$GeneratePath = "/v1/chat/completions"   # OpenAI-style route (confirmed)
    )

    $uri = "$($BaseUrl.TrimEnd('/'))$GeneratePath"

    # Defensive typing
    $p = $ProfileJson | ConvertFrom-Json
    $p.generation.temperature       = [double]$p.generation.temperature
    $p.generation.max_output_tokens = [int]   $p.generation.max_output_tokens

    $sysUser = Compose-SystemAndUserText -ProfileJson $ProfileJson

    # OpenAI-style payload
    $payloadObj = [ordered]@{
        model      = $p.model
        messages   = @(
            @{ role = "system"; content = $sysUser.system },
            @{ role = "user";   content = $sysUser.user   }
        )
        temperature = $p.generation.temperature
        max_tokens  = $p.generation.max_output_tokens
        stream      = $false
    }
    $json = $payloadObj | ConvertTo-Json -Depth 6

    Write-Host "DEBUG: POST $uri"

    # Headers
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type"  = "application/json"
        "Accept"        = "application/json"
    }

    # Determine proxy for THIS URL
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    $proxyUriObj = [System.Net.WebRequest]::DefaultWebProxy.GetProxy([Uri]$uri)
    $proxyUrl    = $null
    if ($proxyUriObj -and $proxyUriObj.AbsoluteUri -and $proxyUriObj.AbsoluteUri -ne $uri) {
        $proxyUrl = $proxyUriObj.AbsoluteUri
        Write-Host "DEBUG: Proxy $proxyUrl (default credentials)"
    } else {
        Write-Host "DEBUG: No proxy (direct)"
    }

    if (Ensure-HttpClientAvailable) {
        # --- HttpClient path (preferred when available) ---
        $handler = [System.Net.Http.HttpClientHandler]::new()
        $handler.UseProxy = $true
        $handler.Proxy = [System.Net.WebRequest]::DefaultWebProxy
        $handler.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

        $client = [System.Net.Http.HttpClient]::new($handler)
        $client.Timeout = [TimeSpan]::FromSeconds(120)

        $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $uri)
        $req.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $ApiKey)
        $req.Headers.Accept.Clear()
        $req.Headers.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new('application/json'))
        $req.Content = New-Object System.Net.Http.StringContent($json, [System.Text.Encoding]::UTF8, 'application/json')

        $resp   = $client.SendAsync($req).Result
        $status = [int]$resp.StatusCode
        $text   = $resp.Content.ReadAsStringAsync().Result
        Write-Host ("DEBUG: Status {0}" -f $status)

        if ($status -ge 200 -and $status -lt 300) {
            try { return ($text | ConvertFrom-Json) } catch { return $text }
        } else {
            Write-Warning "GENAI returned HTTP $status"
            Write-Host "Response body:"
            Write-Host $text
            throw "GENAI call failed ($status)."
        }
    } else {
        # --- Fallback: Invoke-RestMethod (PS 5.1) ---
        try {
            $respObj = $null
            if ($proxyUrl) {
                $respObj = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $json `
                           -Proxy $proxyUrl -ProxyUseDefaultCredentials `
                           -UseDefaultCredentials `
                           -ErrorAction Stop
            } else {
                $respObj = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $json `
                           -UseDefaultCredentials `
                           -ErrorAction Stop
            }

            # IRM may already return an object; if string, parse it
            if ($respObj -is [string]) {
                try { return ($respObj | ConvertFrom-Json) } catch { return $respObj }
            } else {
                return $respObj
            }
        } catch {
            throw "GENAI call failed (Invoke-RestMethod): $($_.Exception.Message)"
        }
    }
}

function Save-ProfileJson {
    param([string]$ProfileJson)
    $path = Join-Path $env:USERPROFILE ("PromptProfile_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".json")
    $ProfileJson | Set-Content -Path $path -Encoding UTF8
    return $path
}

# Optional: Model discovery (probe common paths)
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
        if ($VerboseUrls) { Write-Host "Trying: $u" }
        try {
            $resp = Invoke-RestMethod -Uri $u -Headers $headers -Method GET -ErrorAction Stop
            $hit  = $u
            break
        } catch {
            if ($VerboseUrls) { Write-Host ("  -> " + $_.Exception.Message) }
        }
    }
    if (-not $resp) { throw "No model listing endpoint found. Tried: $($candidates -join ', ')" }

    $models =
        if ($resp.models) { $resp.models }
        elseif ($resp.data) { $resp.data }
        elseif ($resp -is [System.Collections.IEnumerable]) { $resp }
        else { @($resp) }

    $list = foreach ($m in $models) {
        $id   = $null
        $name = $null
        $desc = $null

        if ($m.PSObject.Properties['id'])         { $id = $m.id }
        elseif ($m.PSObject.Properties['model'])  { $id = $m.model }
        elseif ($m.PSObject.Properties['name'])   { $id = $m.name }
        else { $id = "<unknown>" }

        if ($m.PSObject.Properties['displayName']) { $name = $m.displayName }
        elseif ($m.PSObject.Properties['title'])   { $name = $m.title }
        elseif ($m.PSObject.Properties['name'])    { $name = $m.name }
        else { $name = $null }

        if ($m.PSObject.Properties['description']) { $desc = $m.description }

        [pscustomobject]@{ Id = $id; DisplayName = $name; Description = $desc }
    }

    Write-Host "Discovered via: $hit"
    $list | Format-Table -AutoSize
    return $list
}

# ---------------------------------------
# MAIN (CLI)
# ---------------------------------------
try {
    Write-Host ("="*70)
    Write-Host $PolicyBanner
    Write-Host ("="*70)

    $cfg = Get-GenAIConfig

    # Quick Modes
    $presets = @{
        "summary" = @{ Temperature = 0.2; MaxTokens = 300 }
        "brief"   = @{ Temperature = 0.3; MaxTokens = 600 }
        "detail"  = @{ Temperature = 0.4; MaxTokens = 1200 }
    }

    $seeModels = Read-Host "Press Enter to skip listing models, or type 'list' to show them"
    if ($seeModels -eq 'list') {
        $models = Get-GenAIModels -BaseUrl $cfg.BaseUrl -ApiKey $cfg.ApiKey -VerboseUrls
        Write-Host "`nEnter a model id from the list above."
    }

    $model = Read-Host "Model id (e.g., gemini-2.5-pro)"
    $fmt   = Read-Host "Output format (markdown or json)"
    $goal  = Read-Host "Goal or Task"
    $aud   = Read-Host "Audience"
    $per   = Read-Host "Persona or Role"
    $tone  = Read-Host "Tone"
    $cons  = Read-Host "Scope and constraints"
    $mode  = Read-Host "Choose quick mode (summary | brief | detail) or press Enter to skip"

    if (-not [string]::IsNullOrWhiteSpace($mode) -and $presets.ContainsKey($mode)) {
        $temp = $presets[$mode].Temperature
        $tok  = $presets[$mode].MaxTokens
        Write-Host ("Using preset: {0} (temp={1}, max_tokens={2})" -f $mode, $temp, $tok)
    } else {
        $temp = Read-Host "Temperature (0 to 1, e.g., 0.3)"
        $tok  = Read-Host "Max output tokens (e.g., 300)"
    }

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
    Write-Host "Saved profile to: $path"

    # Progress + elapsed time
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity "Contacting GENAI" -Status "Sending request..." -PercentComplete 10

    $resp = Invoke-GenAI -BaseUrl $cfg.BaseUrl -ApiKey $cfg.ApiKey -ProfileJson $profile -GeneratePath "/v1/chat/completions"

    Write-Progress -Activity "Contacting GENAI" -Completed
    $sw.Stop()
    Write-Host ("Completed in {0:N1} seconds." -f $sw.Elapsed.TotalSeconds)

    $assistantText = Get-AssistantText -Resp $resp
    Write-Host "`n--- Assistant Output ---`n"
    Write-Host $assistantText
}
catch {
    Write-Error $_.Exception.Message
}