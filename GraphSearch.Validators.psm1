# Copyright 2026 Fentron Security Solutions
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

<#
    Global identifier validators for GraphSearch Stage 2.

    Design: a small set of reusable checksum PRIMITIVES (Luhn, IBAN mod-97,
    11-proof, EAN-13, German RVNR cross-sum) plus a data-driven REGISTRY of
    identifier definitions. Each definition carries a country, a category, a
    find-pattern, and a validator. Callers select by name, country, or
    category - so an engagement can run "all payment data" or "all EU national
    IDs" without code changes.

    Every checksum is verified against a known-valid reference number in the
    project test notes; formats without a checksum (US SSN, UK NINo) use
    structural + range rules and return ChecksumValid according to those rules.

    Coverage is intentionally global-but-partial: this is a framework, not a
    claim to validate every country. Adding a jurisdiction is one entry in
    $script:Definitions plus (if needed) one primitive. See
    docs/CONTEXT-VALIDATION.md.
#>

Set-StrictMode -Version Latest

# ===========================================================================
#  Checksum primitives
# ===========================================================================

function Test-Luhn {
    param([Parameter(Mandatory)][string]$Value)
    $s = $Value -replace '\D',''
    if ($s.Length -lt 2) { return $false }
    $sum = 0; $alt = $false
    for ($i = $s.Length - 1; $i -ge 0; $i--) {
        $d = [int]([string]$s[$i])
        if ($alt) { $d *= 2; if ($d -gt 9) { $d -= 9 } }
        $sum += $d; $alt = -not $alt
    }
    return ($sum % 10 -eq 0)
}

function Test-IbanMod97 {
    param([Parameter(Mandatory)][string]$Value)
    $s = ($Value -replace '\s','').ToUpper()
    if ($s -notmatch '^[A-Z]{2}\d{2}[A-Z0-9]+$') { return $false }
    if ($s.Length -lt 15 -or $s.Length -gt 34) { return $false }
    $rot = $s.Substring(4) + $s.Substring(0,4)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $rot.ToCharArray()) {
        if ([char]::IsLetter($c)) { [void]$sb.Append(([int][char]$c - 55)) }
        else                      { [void]$sb.Append($c) }
    }
    # mod-97 over a long numeric string, processed in chunks
    $rem = 0
    foreach ($ch in $sb.ToString().ToCharArray()) {
        $rem = (($rem * 10) + [int][string]$ch) % 97
    }
    return ($rem -eq 1)
}

function Test-Elfproef {
    # Dutch 11-proof (BSN). Weights 9..2 then -1; sum divisible by 11.
    param([Parameter(Mandatory)][string]$Value)
    $s = $Value -replace '\D',''
    if ($s.Length -eq 8) { $s = '0' + $s }
    if ($s.Length -ne 9) { return $false }
    $w = 9,8,7,6,5,4,3,2,-1
    $sum = 0
    for ($i = 0; $i -lt 9; $i++) { $sum += [int]([string]$s[$i]) * $w[$i] }
    return (($sum % 11) -eq 0)
}

function Test-EAN13 {
    param([Parameter(Mandatory)][string]$Value)
    $s = $Value -replace '\D',''
    if ($s.Length -ne 13) { return $false }
    $sum = 0
    for ($i = 0; $i -lt 12; $i++) {
        $w = if ($i % 2 -eq 0) { 1 } else { 3 }
        $sum += [int]([string]$s[$i]) * $w
    }
    $calc = (10 - ($sum % 10)) % 10
    return ($calc -eq [int]([string]$s[12]))
}

function Test-GermanRVNRChecksum {
    param([Parameter(Mandatory)][string]$Value)
    $s = ($Value -replace '\s','').ToUpper()
    if ($s -notmatch '^\d{8}[A-Z]\d{3}$') { return $false }
    $day = [int]$s.Substring(2,2)
    if     ($day -ge 65) { $day -= 64 } elseif ($day -ge 33) { $day -= 32 }
    $month = [int]$s.Substring(4,2)
    if ($day -lt 1 -or $day -gt 31 -or $month -lt 1 -or $month -gt 12) { return $false }
    $letterNum = '{0:D2}' -f ([int][char]$s[8] - [int][char]'A' + 1)
    $digits = $s.Substring(0,8) + $letterNum + $s.Substring(9,2)
    $factors = 2,1,2,5,7,1,2,1,2,1,2,1
    $total = 0
    for ($i = 0; $i -lt 12; $i++) {
        $p = [int]([string]$digits[$i]) * $factors[$i]
        foreach ($c in ([string]$p).ToCharArray()) { $total += [int][string]$c }
    }
    return (($total % 10) -eq [int]([string]$s[11]))
}

# ===========================================================================
#  Structural / range validators (no checksum exists)
# ===========================================================================

function Test-USSSN {
    param([Parameter(Mandatory)][string]$Value)
    $m = [regex]::Match($Value, '^(\d{3})-?(\d{2})-?(\d{4})$')
    if (-not $m.Success) { return $false }
    $a = $m.Groups[1].Value; $g = $m.Groups[2].Value; $ser = $m.Groups[3].Value
    if ($a -eq '000' -or $a -eq '666' -or [int]$a -ge 900) { return $false }
    if ($g -eq '00')   { return $false }
    if ($ser -eq '0000') { return $false }
    return $true
}

function Test-UKNINo {
    param([Parameter(Mandatory)][string]$Value)
    $s = ($Value -replace '\s','').ToUpper()
    if ($s -notmatch '^[A-Z]{2}\d{6}[A-D]?$') { return $false }
    $p = $s.Substring(0,2)
    # first letter not D F I Q U V; second not D F I O Q U V
    if ('DFIQUV'.Contains([string]$p[0])) { return $false }
    if ('DFIOQUV'.Contains([string]$p[1])) { return $false }
    if (@('BG','GB','NK','KN','TN','NT','ZZ') -contains $p) { return $false }
    return $true
}

# ===========================================================================
#  Definition registry
#  Each entry: Name, Country, Category, Description, Pattern (find), and
#  Validate (scriptblock -> bool) OR $null when unverifiable (heuristic).
# ===========================================================================

$script:Definitions = @(
    [pscustomobject]@{
        Name='CreditCard'; Country='Global'; Category='payment'
        Description='Payment card number (Luhn, 13-19 digits)'
        Pattern='\b(?:\d[ -]?){12,18}\d\b'
        Validate={ param($v) (Test-Luhn $v) -and ((($v -replace '\D','').Length) -ge 13) -and ((($v -replace '\D','').Length) -le 19) }
    }
    [pscustomobject]@{
        Name='IBAN'; Country='Global'; Category='bank'
        Description='International Bank Account Number (mod-97)'
        Pattern='\b[A-Z]{2}\d{2}[ ]?(?:[A-Z0-9][ ]?){11,30}\b'
        Validate={ param($v) Test-IbanMod97 $v }
    }
    [pscustomobject]@{
        Name='US-SSN'; Country='US'; Category='national-id'
        Description='US Social Security Number (format + range)'
        Pattern='\b\d{3}-\d{2}-\d{4}\b'
        Validate={ param($v) Test-USSSN $v }
    }
    [pscustomobject]@{
        Name='UK-NINo'; Country='UK'; Category='national-id'
        Description='UK National Insurance Number (format + prefix rules)'
        Pattern='\b[A-Za-z]{2}\s?\d{2}\s?\d{2}\s?\d{2}\s?[A-Da-d]?\b'
        Validate={ param($v) Test-UKNINo $v }
    }
    [pscustomobject]@{
        Name='DE-RVNR'; Country='DE'; Category='national-id'
        Description='German Rentenversicherungsnummer / Sozialversicherungsnummer (Mod-10)'
        Pattern='\b\d{2}\s?\d{6}\s?[A-Za-z]\s?\d{2}\s?\d\b'
        Validate={ param($v) Test-GermanRVNRChecksum $v }
    }
    [pscustomobject]@{
        Name='CH-AHV'; Country='CH'; Category='national-id'
        Description='Swiss AHV/AVS Sozialversicherungsnummer (EAN-13)'
        Pattern='\b756[.\s]?\d{4}[.\s]?\d{4}[.\s]?\d{2}\b'
        Validate={ param($v) Test-EAN13 $v }
    }
    [pscustomobject]@{
        Name='NL-BSN'; Country='NL'; Category='national-id'
        Description='Dutch Burgerservicenummer (11-proof)'
        Pattern='\b\d{9}\b'
        Validate={ param($v) Test-Elfproef $v }
    }
    [pscustomobject]@{
        Name='ES-DNI'; Country='ES'; Category='national-id'
        Description='Spanish DNI/NIE (mod-23 letter)'
        Pattern='\b[XYZxyz]?\d{7,8}[A-Za-z]\b'
        Validate={
            param($v)
            $s = ($v -replace '[-\s]','').ToUpper()
            $m = [regex]::Match($s,'^([XYZ]?)(\d{7,8})([TRWAGMYFPDXBNJZSQVHLCKE])$')
            if (-not $m.Success) { return $false }
            $num = $m.Groups[2].Value
            switch ($m.Groups[1].Value) { 'X' {$num="0$num"} 'Y' {$num="1$num"} 'Z' {$num="2$num"} }
            return ('TRWAGMYFPDXBNJZSQVHLCKE'[[int64]$num % 23] -eq $m.Groups[3].Value[0])
        }
    }
    [pscustomobject]@{
        Name='CA-SIN'; Country='CA'; Category='national-id'
        Description='Canadian Social Insurance Number (Luhn)'
        Pattern='\b\d{3}[ -]?\d{3}[ -]?\d{3}\b'
        Validate={ param($v) Test-Luhn $v }
    }
    [pscustomobject]@{
        Name='DE-Personalnummer'; Country='DE'; Category='employee-id'
        Description='German employee number (HEURISTIC label-adjacent; no checksum exists)'
        Pattern=$null   # special-cased finder below
        Validate=$null
    }
)

# ===========================================================================
#  Finders
# ===========================================================================

function Get-ContextSnippet {
    param([string]$Text,[int]$Index,[int]$Length,[int]$Context)
    $start = [Math]::Max(0, $Index - $Context)
    $len   = [Math]::Min($Text.Length - $start, $Length + (2 * $Context))
    return (($Text.Substring($start, $len)) -replace '\s+',' ').Trim()
}

function Find-ByDefinition {
    param($Def,[string]$Text,[int]$Context = 60)
    $out = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrEmpty($Text)) { return $out }

    if ($Def.Name -eq 'DE-Personalnummer') {
        $pattern = '(?i)(personalnummer|pers\.?\-?\s?nr\.?|personalnr\.?)\s*[:#]?\s*([A-Z0-9\-]{3,20})'
        foreach ($m in [regex]::Matches($Text, $pattern)) {
            $out.Add([pscustomobject]@{
                Type=$Def.Name; Country=$Def.Country; Category=$Def.Category
                Match=$m.Groups[2].Value; ChecksumValid=$null
                Offset=$m.Index; Context=(Get-ContextSnippet $Text $m.Index $m.Length $Context)
            })
        }
        return $out
    }

    foreach ($m in [regex]::Matches($Text, $Def.Pattern)) {
        $valid = & $Def.Validate $m.Value
        $out.Add([pscustomobject]@{
            Type=$Def.Name; Country=$Def.Country; Category=$Def.Category
            Match=$m.Value; ChecksumValid=$valid
            Offset=$m.Index; Context=(Get-ContextSnippet $Text $m.Index $m.Length $Context)
        })
    }
    return $out
}

# ===========================================================================
#  Selection + public API
# ===========================================================================

function Get-ValidatorDefinitions {
    param([string[]]$Name,[string[]]$Country,[string[]]$Category)
    $defs = $script:Definitions
    if ($Name)     { $defs = $defs | Where-Object { $Name     -contains $_.Name } }
    if ($Country)  { $defs = $defs | Where-Object { $Country  -contains $_.Country } }
    if ($Category) { $defs = $defs | Where-Object { $Category -contains $_.Category } }
    return $defs
}

function Get-AvailableValidators {
    $script:Definitions | Select-Object Name,Country,Category,Description
}

function Invoke-Validators {
    <#
        Run selected validators over text. Select by any combination of
        -Name, -Country, -Category. With no selector, runs ALL definitions.
        Note: broad no-checksum patterns (NL-BSN = any 9 digits) are noisy
        unless scoped - prefer selecting by country/category for a run.
    #>
    param(
        [string]$Text,
        [string[]]$Name,
        [string[]]$Country,
        [string[]]$Category,
        [int]$Context = 60
    )
    $defs = Get-ValidatorDefinitions -Name $Name -Country $Country -Category $Category
    $results = New-Object System.Collections.Generic.List[object]
    foreach ($def in $defs) {
        foreach ($hit in (Find-ByDefinition -Def $def -Text $Text -Context $Context)) {
            $results.Add($hit)
        }
    }
    return $results
}

Export-ModuleMember -Function `
    Test-Luhn, Test-IbanMod97, Test-Elfproef, Test-EAN13, Test-GermanRVNRChecksum,
    Test-USSSN, Test-UKNINo,
    Get-ValidatorDefinitions, Get-AvailableValidators, Invoke-Validators
