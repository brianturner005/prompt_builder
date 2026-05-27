<#
  GenAI-ListModels.ps1
  Discover available model IDs from a GENAI.mil tenant.

  - Reads GENAI_BASE_URL and GENAI_API_KEY from env vars by default
  - Normalizes base URL (removes accidental /v1/... suffixes)
  - Probes several likely listing endpoints
  - Prints a table and optional JSON export

  Author: Brian Turner & M365 Copilot
#>

param(
  [string]$BaseUrl,
  [string]$ApiKey,
  [string]$OutJson,            # Optional: path to write JSON output
  [switch]$VerboseUrls         # Optional: echo attempted URLs
)

function Get-EnvOrParam {
  param(
    [string]$Value,
    [string]$EnvName,
    [string]$Description
  )
  if ([string]::IsNullOrWhiteSpace($Value)) {
    $envVal = (Get-ChildItem "Env:$EnvName" -ErrorAction SilentlyContinue).Value
    if ([string]::IsNullOrWhiteSpace($envVal)) {
      throw "$Description is not set. Provide -$($Description.Replace(' ', '')) or set $EnvName."
    }
    return $envVal
  }
  return $Value
}

function Normalize-BaseUrl {
  param([string]$Url)
  if ([string]::IsNullOrWhiteSpace($Url)) { return $Url }
  $u = $Url.TrimEnd('/')

  # Strip common resource suffixes if the user pasted a full endpoint
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
    if ($u.EndsWith($s)) {
      $u = $u.Substring(0, $u.Length - $s.Length)
      break
    }
  }
  return $u
}

function Invoke-ModelList {
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

  $resp = $null
  $hit  = $null

  foreach ($u in $candidates) {
    if ($VerboseUrls) { Write-Host "Trying: $u" -ForegroundColor Yellow }
    try {
      $resp = Invoke-RestMethod -Uri $u -Headers $headers -Method GET -ErrorAction Stop
      $hit  = $u
      break
    } catch {
      if ($VerboseUrls) { Write-Host "  -> $($_.Exception.Message)" -ForegroundColor DarkGray }
      continue
    }
  }

  if (-not $resp) {
    throw "No model listing endpoint found. Tried: $($candidates -join ', ')"
  }

  return @{ Response = $resp; Url = $hit }
}

function Parse-Models {
  param([Parameter(Mandatory)]$Response)

  # Normalize common shapes:
  # 1) { models: [...] }
  # 2) { data: [...] }
  # 3) [ ... ]
  $models =
    if ($Response.models) { $Response.models }
    elseif ($Response.data) { $Response.data }
    elseif ($Response -is [System.Collections.IEnumerable]) { $Response }
    else { @($Response) }

  # Project to stable shape
  $list = foreach ($m in $models) {
    $id  = $null
    $name= $null
    $desc= $null

    if ($m.PSObject.Properties['id'])          { $id   = $m.id }
    elseif ($m.PSObject.Properties['model'])   { $id   = $m.model }
    elseif ($m.PSObject.Properties['name'])    { $id   = $m.name }
    else                                       { $id   = "<unknown>" }

    if ($m.PSObject.Properties['displayName']) { $name = $m.displayName }
    elseif ($m.PSObject.Properties['title'])   { $name = $m.title }
    elseif ($m.PSObject.Properties['name'])    { $name = $m.name }
    else                                       { $name = $null }

    if ($m.PSObject.Properties['description']) { $desc = $m.description }

    [pscustomobject]@{
      Id          = $id
      DisplayName = $name
      Description = $desc
    }
  }

  return $list
}

# -----------------------------
# Main
# -----------------------------
try {
  $resolvedBase = Get-EnvOrParam -Value $BaseUrl -EnvName "GENAI_BASE_URL" -Description "Base Url"
  $resolvedKey  = Get-EnvOrParam -Value $ApiKey  -EnvName "GENAI_API_KEY"  -Description "Api Key"

  $resolvedBase = Normalize-BaseUrl $resolvedBase
  Write-Host "Base URL: $resolvedBase" -ForegroundColor Cyan

  $result = Invoke-ModelList -BaseUrl $resolvedBase -ApiKey $resolvedKey -VerboseUrls:$VerboseUrls
  Write-Host "Discovered via: $($result.Url)" -ForegroundColor Cyan

  $models = Parse-Models -Response $result.Response

  if (-not $models -or $models.Count -eq 0) {
    Write-Warning "The endpoint returned no models."
    return
  }

  # Display results
  Write-Host "`nAvailable models:" -ForegroundColor Green
  $models | Format-Table -AutoSize

  # Optional export
  if (-not [string]::IsNullOrWhiteSpace($OutJson)) {
    $models | ConvertTo-Json -Depth 5 | Set-Content -Path $OutJson -Encoding UTF8
    Write-Host "`nSaved JSON to: $OutJson" -ForegroundColor Green
  }
}
catch {
  Write-Error $_.Exception.Message
  exit 1
}