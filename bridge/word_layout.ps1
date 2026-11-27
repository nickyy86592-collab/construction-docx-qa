param([Parameter(Mandatory=$true)][ValidateSet('offer','measure')][string]$Mode)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName System.Security
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$work = Join-Path $env:RUNNER_TEMP ("word-layout-" + $env:GITHUB_RUN_ID)
$seal = [Security.Cryptography.DataProtectionScope]::CurrentUser
$rsa = [Security.Cryptography.RSACryptoServiceProvider]::new(3072)
$rsa.PersistKeyInCsp = $false
$keyfile = Join-Path $work 'receiver.dat'
if ($Mode -eq 'offer') {
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $p = $rsa.ExportParameters($false)
    $key = @{
        run_id = $env:GITHUB_RUN_ID
        modulus = [Convert]::ToBase64String($p.Modulus)
        exponent = [Convert]::ToBase64String($p.Exponent)
    }
    $key | ConvertTo-Json -Compress | Set-Content -Encoding ASCII (Join-Path $work 'receiver.json')
    $private = [Security.Cryptography.ProtectedData]::Protect($rsa.ExportCspBlob($true), $null, $seal)
    [IO.File]::WriteAllBytes($keyfile, $private)
    $rsa.Dispose()
    exit 0
}

function Join-Bytes([byte[]]$A, [byte[]]$B) {
    $result = New-Object byte[] ($A.Length + $B.Length)
    [Array]::Copy($A, 0, $result, 0, $A.Length)
    [Array]::Copy($B, 0, $result, $A.Length, $B.Length)
    return ,$result
}

function Test-Tag([byte[]]$A, [byte[]]$B) {
    if ($A.Length -ne $B.Length) { return $false }
    $diff = 0
    for ($i = 0; $i -lt $A.Length; $i++) { $diff = $diff -bor ($A[$i] -bxor $B[$i]) }
    return ($diff -eq 0)
}

function Protect-Packet([byte[]]$Data, [string]$Direction, [bool]$Encrypt) {
    $mac = [Security.Cryptography.HMACSHA256]::new([byte[]]$material[32..63])
    $aes = [Security.Cryptography.Aes]::Create()
    $aes.Key = [byte[]]$material[0..31]
    $aes.Mode = [Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [Security.Cryptography.PaddingMode]::PKCS7
    $aad = [Text.Encoding]::ASCII.GetBytes("word-layout-v1/$env:GITHUB_RUN_ID/$Direction/")
    try {
        if ($Encrypt) {
            $aes.GenerateIV()
            $transform = $aes.CreateEncryptor()
            try { $cipher = $transform.TransformFinalBlock($Data, 0, $Data.Length) }
            finally { $transform.Dispose() }
            $message = Join-Bytes $aes.IV $cipher
            $tag = $mac.ComputeHash((Join-Bytes $aad $message))
            return ,(Join-Bytes $message $tag)
        }
        if ($Data.Length -lt 64) { throw 'Invalid encrypted request length' }
        $offset = $Data.Length - 32
        $message = [byte[]]$Data[0..($offset - 1)]
        $tag = [byte[]]$Data[$offset..($Data.Length - 1)]
        if (-not (Test-Tag $tag ($mac.ComputeHash((Join-Bytes $aad $message))))) { throw 'Request integrity check failed' }
        $aes.IV = [byte[]]$message[0..15]
        $transform = $aes.CreateDecryptor()
        try { return , ($transform.TransformFinalBlock($message, 16, $message.Length - 16)) }
        finally { $transform.Dispose() }
    } finally {
        $mac.Dispose()
        $aes.Dispose()
    }
}

$word = $null
$out = Join-Path $work 'proof'
$docs = Join-Path $work 'docs'
try {
    $rsa.ImportCspBlob([Security.Cryptography.ProtectedData]::Unprotect([IO.File]::ReadAllBytes($keyfile), $null, $seal))
    Remove-Item -LiteralPath $keyfile -Force
    $headers = @{Authorization="Bearer $env:GH_TOKEN";Accept='application/vnd.github+json';'X-GitHub-Api-Version'='2022-11-28'}
    $base = "https://api.github.com/repos/$env:GITHUB_REPOSITORY"
    $ref = [Uri]::EscapeDataString($env:GITHUB_REF_NAME)
    $descriptor = "$base/contents/bridge/delivery/$env:GITHUB_RUN_ID.json?ref=$ref"
    $payload = $null
    for ($attempt = 0; $attempt -lt 48 -and -not $payload; $attempt++) {
        try { $response = Invoke-RestMethod -Uri $descriptor -Headers $headers -TimeoutSec 12 }
        catch {
            if ([int]$_.Exception.Response.StatusCode -ne 404) { throw }
            Start-Sleep -Seconds 5
            continue
        }
        $payload = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($response.content)) | ConvertFrom-Json
    }
    if (-not $payload -or $payload.run_id -ne $env:GITHUB_RUN_ID -or $payload.blobs.Count -lt 1 -or $payload.blobs.Count -gt 32) { throw 'No valid request for this run' }
    $material = $rsa.Decrypt([Convert]::FromBase64String($payload.wrapped), $true)
    if ($material.Length -ne 64) { throw 'Invalid session key' }
    $buffer = [IO.MemoryStream]::new()
    try {
        foreach ($sha in $payload.blobs) {
            if ($sha -cnotmatch '^[0-9a-f]{40}$') { throw 'Invalid request reference' }
            $blob = Invoke-RestMethod -Uri "$base/git/blobs/$sha" -Headers $headers -TimeoutSec 25
            $bytes = [Convert]::FromBase64String($blob.content)
            $buffer.Write($bytes, 0, $bytes.Length)
        }
        $packet = Protect-Packet ($buffer.ToArray()) 'in' $false
    } finally { $buffer.Dispose() }
    $zipPath = Join-Path $work 'request.zip'
    [IO.File]::WriteAllBytes($zipPath, $packet)
    New-Item -ItemType Directory -Path $out,$docs -Force | Out-Null
    $zip = [IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $entries = @($zip.Entries)
        if ($entries.Count -lt 1 -or $entries.Count -gt 14) { throw 'Invalid Word input count' }
        foreach ($entry in $entries) {
            if ($entry.FullName -cnotmatch '^(0[1-9]|10|stage40|stage75|compile)[.]docx$' -or $entry.Length -gt 10000000) { throw 'Invalid Word input' }
            [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, (Join-Path $docs $entry.FullName))
        }
    } finally { $zip.Dispose() }
    Write-Output "REQUEST_READY"
    & $env:AUDIT_PYTHON (Join-Path $PSScriptRoot 'render_word.py') $docs $out
    if ($LASTEXITCODE -ne 0) { throw 'Microsoft Word pagination failed' }
    $outputZip = Join-Path $work 'proof.zip'
    [IO.Compression.ZipFile]::CreateFromDirectory($out, $outputZip)
    [IO.File]::WriteAllBytes((Join-Path $work 'proof.bin'), (Protect-Packet ([IO.File]::ReadAllBytes($outputZip)) 'out' $true))
    Write-Output "Encrypted output ready"
} finally {
    if ($null -ne $word) { $word.Quit() }
    $rsa.Dispose()
    foreach ($name in @('docs','proof','request.zip','proof.zip','receiver.dat')) {
        $item = Join-Path $work $name
        if (Test-Path -LiteralPath $item) { Remove-Item -LiteralPath $item -Recurse -Force }
    }
}
