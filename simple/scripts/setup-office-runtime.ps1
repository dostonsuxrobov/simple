param([string]$Destination = (Join-Path $env:LOCALAPPDATA 'simple\office-runtime'))

$ErrorActionPreference = 'Stop'
$version = '26.2.6'
$expectedHash = 'f9877032fd908beb9c0ddf06df4af5c2e85f419c42e14876c4cce5aae5fb2660'
$downloadUrl = 'https://mirror.fcix.net/tdf/libreoffice/stable/26.2.6/win/x86_64/LibreOffice_26.2.6_Win_x86-64.msi'
$checksumSource = 'https://download.documentfoundation.org/libreoffice/stable/26.2.6/win/x86_64/LibreOffice_26.2.6_Win_x86-64.msi.sha256'
$targetDirectory = [IO.Path]::GetFullPath($Destination)
$runtimeExe = Join-Path $targetDirectory 'program\soffice.exe'
$manifestPath = Join-Path $targetDirectory 'simple-runtime.json'
if ((Test-Path -LiteralPath $runtimeExe) -and (Test-Path -LiteralPath $manifestPath)) {
  $installed = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
  if ($installed.version -eq $version -and $installed.installerSha256 -eq $expectedHash) {
    Write-Output "The document engine is ready: $runtimeExe"
    exit 0
  }
}

$jobDirectory = Join-Path ([IO.Path]::GetTempPath()) ('simple-office-setup-' + [Guid]::NewGuid())
New-Item -ItemType Directory -Path $jobDirectory | Out-Null
$installer = Join-Path $jobDirectory 'LibreOffice.msi'
try {
  Write-Output 'Downloading the document engine (356 MB)…'
  Invoke-WebRequest -Uri $downloadUrl -OutFile $installer
  if ((Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant() -ne $expectedHash) {
    throw 'The download does not match the verified release. Setup stopped before extraction.'
  }
  $signature = Get-AuthenticodeSignature -LiteralPath $installer
  if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'O=The Document Foundation') {
    throw 'The document engine publisher signature could not be verified.'
  }
  Write-Output 'Preparing the local document engine…'
  $arguments = '/a "' + $installer + '" /qn TARGETDIR="' + $targetDirectory + '" /L*v "' + (Join-Path $jobDirectory 'setup.log') + '"'
  $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -PassThru -Wait -WindowStyle Hidden
  if ($process.ExitCode -ne 0 -or !(Test-Path -LiteralPath $runtimeExe)) { throw "Extraction failed (code $($process.ExitCode)). Log: $jobDirectory" }
  @{
    version = $version
    installerSha256 = $expectedHash
    downloadUrl = $downloadUrl
    checksumSource = $checksumSource
    publisher = 'The Document Foundation'
    preparedAt = [DateTime]::UtcNow.ToString('o')
  } | ConvertTo-Json | Set-Content -LiteralPath $manifestPath -Encoding utf8
  Write-Output "The document engine is ready: $runtimeExe"
} finally {
  # Only our two files are removed. A failed job's log remains for diagnosis.
  if (Test-Path -LiteralPath $installer) { Remove-Item -LiteralPath $installer -Force }
}
