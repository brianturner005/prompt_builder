param(
    [Parameter(Mandatory=$true)]
    [string]$Path,

    # Choose how to replace non-ASCII characters:
    # "conservative" -> map dashes to "-", quotes to '"/'', spaces to " ", else remove
    # "strict"       -> replace everything non-ASCII with "?"
    [ValidateSet("conservative","strict")]
    [string]$Mode = "conservative"
)

# Read entire file
$text = Get-Content -LiteralPath $Path -Raw

# Build normalized output
$sb = New-Object System.Text.StringBuilder

foreach ($ch in $text.ToCharArray()) {
    $code = [int][char]$ch
    if ($code -le 127) {
        [void]$sb.Append($ch)
        continue
    }

    if ($Mode -eq "strict") {
        [void]$sb.Append("?")
        continue
    }

    # Conservative mapping by Unicode category
    $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
    switch ($cat) {
        # Any dash punctuation -> ASCII hyphen-minus
        "DashPunctuation"      { [void]$sb.Append("-") }

        # Space-like separators -> ASCII space
        "SpaceSeparator"       { [void]$sb.Append(" ") }

        # Quotes
        "InitialQuotePunctuation" { [void]$sb.Append('"') }
        "FinalQuotePunctuation"   { [void]$sb.Append('"') }
        "OtherPunctuation" {
            # Common curly quotes that sometimes show as punctuation
            if ($ch -eq [char]0x2018 -or $ch -eq [char]0x2019) { [void]$sb.Append("'") }
            elseif ($ch -eq [char]0x201C -or $ch -eq [char]0x201D) { [void]$sb.Append('"') }
            else { [void]$sb.Append("-") }
        }

        # Controls, surrogates, non-spacing marks -> drop
        "Control"              { }
        "Surrogate"            { }
        "NonSpacingMark"       { }
        "EnclosingMark"        { }

        default {
            # For anything else (letters/symbols), safest is to replace with a space
            [void]$sb.Append(" ")
        }
    }
}

# Save as UTF-8 (ASCII-safe)
Set-Content -LiteralPath $Path -Value $sb.ToString() -Encoding UTF8
Write-Host "Normalized non-ASCII to safe ASCII in:" $Path