$ErrorActionPreference = 'Stop'

$workspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$backupRoot = [IO.Path]::GetFullPath((Join-Path $workspace ('.cleanup-backups\simple_open_with_' + (Get-Date -Format 'yyyyMMdd-HHmmss'))))
if (-not $backupRoot.StartsWith($workspace, [StringComparison]::OrdinalIgnoreCase)) {
  throw "Backup path escaped workspace: $backupRoot"
}
New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null

$exported = 0
function Export-RegKey([string]$key, [string]$name) {
  $providerPath = 'Registry::' + $key.Replace('HKCU', 'HKEY_CURRENT_USER')
  if (Test-Path -LiteralPath $providerPath) {
    $target = Join-Path $backupRoot ($name + '.reg')
    & reg.exe EXPORT $key $target /y | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Could not back up $key" }
    $script:exported++
  }
}

$corruptSlots = [ordered]@{
  '.blend' = 'a'
  '.casc' = 'a'
  '.docx' = 'a'
  '.epub' = 'ba'
  '.fbx' = 'a'
  '.html' = 'a'
  '.ics' = 'a'
  '.ipynb' = 'a'
  '.jpg' = 'abdc'
  '.m4a' = 'ba'
  '.md' = 'a'
  '.mkv' = 'ba'
  '.mp3' = 'ab'
  '.mp4' = 'badec'
  '.pdf' = 'cadb'
  '.tres' = 'a'
  '.txt' = 'a'
  '.unity' = 'a'
  '.webp' = 'cab'
  '.xlsx' = 'a'
}

$supported = @(
  '.docx','.pdf','.txt','.md','.doc','.png','.jpg','.jpeg','.webp','.gif','.bmp','.svg','.avif',
  '.mp4','.m4v','.webm','.ogv','.mov','.mkv','.xlsx','.xlsm','.xlsb','.xls','.xltx','.xltm',
  '.xlt','.xlam','.xla','.xml','.ods','.fods','.csv','.tsv','.tab','.numbers','.slk','.sylk',
  '.dif','.dbf','.prn','.wk1','.wk2','.wk3','.wk4','.wks','.wq1','.wq2','.wb1','.wb2','.wb3',
  '.123','.qpw','.html','.htm'
)

$staleOpenWithData = @(
  'simple_calc.exe',
  'simple_pdf.exe',
  'electron.exe',
  'WINWORD.EXE',
  'Acrobat.exe',
  'iTunes.exe',
  'AudioBookConverter.exe',
  'Telegram.exe',
  'vlc.exe',
  'cascadeur.exe',
  'msedgewebview2.exe',
  'Microsoft.ZuneMusic_8wekyb3d8bbwe!Microsoft.ZuneMusic'
)

$touchedExts = @($supported + @($corruptSlots.Keys) + @('.ico')) | Sort-Object -Unique
foreach ($ext in $touchedExts) {
  Export-RegKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext" ('FileExts_' + $ext.TrimStart('.'))
}
Export-RegKey 'HKCU\Software\Classes\Applications\simple.exe' 'Applications_simple_exe'
Export-RegKey 'HKCU\Software\Classes\Applications\simple_calc.exe' 'Applications_simple_calc_exe'
Export-RegKey 'HKCU\Software\Classes\Applications\simple_pdf.exe' 'Applications_simple_pdf_exe'
Export-RegKey 'HKCU\Software\Classes\simple.Document' 'ProgId_simple_Document'
Export-RegKey 'HKCU\Software\Classes\ods_auto_file' 'ProgId_ods_auto_file'
Export-RegKey 'HKCU\Software\Classes\.pdf' 'Classes_pdf'
Export-RegKey 'HKCU\Software\Classes\.ods' 'Classes_ods'
Export-RegKey 'HKCU\Software\Classes\Simple.Unified.File' 'ProgId_Simple_Unified_File'
Export-RegKey 'HKCU\Software\Microsoft\Windows\CurrentVersion\App Paths\simple.exe' 'AppPaths_simple_exe'
Export-RegKey 'HKCU\Software\simple\Capabilities' 'Legacy_simple_Capabilities'
Export-RegKey 'HKCU\Software\RegisteredApplications' 'RegisteredApplications'

$removedOpenWith = [System.Collections.Generic.List[string]]::new()
function Repair-OpenWithList([string]$extension, [string[]]$removeData) {
  $key = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$extension\OpenWithList"
  if (-not (Test-Path -LiteralPath $key)) { return }

  $property = Get-ItemProperty -LiteralPath $key
  $entryProps = @($property.PSObject.Properties | Where-Object { $_.Name -match '^[a-z]$' })
  foreach ($entry in $entryProps) {
    if ($removeData -contains [string]$entry.Value) {
      Remove-ItemProperty -LiteralPath $key -Name $entry.Name -Force
      $script:removedOpenWith.Add(('{0}:{1}={2}' -f $extension, $entry.Name, $entry.Value))
    }
  }

  $updated = Get-ItemProperty -LiteralPath $key
  $remaining = @($updated.PSObject.Properties | Where-Object { $_.Name -match '^[a-z]$' } | ForEach-Object { $_.Name })
  $originalMru = [string]$updated.MRUList
  $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  $newOrder = [System.Collections.Generic.List[string]]::new()
  foreach ($character in $originalMru.ToCharArray()) {
    $name = [string]$character
    if ($remaining -contains $name -and $seen.Add($name)) { $newOrder.Add($name) }
  }
  foreach ($name in ($remaining | Sort-Object)) {
    if ($seen.Add($name)) { $newOrder.Add($name) }
  }

  if ($newOrder.Count -gt 0) {
    Set-ItemProperty -LiteralPath $key -Name 'MRUList' -Value ($newOrder -join '')
  } elseif ($updated.PSObject.Properties.Name -contains 'MRUList') {
    Remove-ItemProperty -LiteralPath $key -Name 'MRUList' -Force
  }
}

foreach ($ext in $supported) {
  $targets = [System.Collections.Generic.List[string]]::new()
  foreach ($data in $staleOpenWithData) { $targets.Add($data) }
  if ($corruptSlots.Contains($ext)) { $targets.Add($corruptSlots[$ext]) }
  Repair-OpenWithList $ext $targets.ToArray()
}
foreach ($ext in $corruptSlots.Keys) {
  if ($supported -notcontains $ext) { Repair-OpenWithList $ext @($corruptSlots[$ext]) }
}
Repair-OpenWithList '.ico' @('simple.exe')

$explorerProgIdRemovals = [ordered]@{
  '.docx' = @('Word.Document.12','docxfile')
  '.pdf' = @('Acrobat.Document.DC','pdf_auto_file','simple.Document')
  '.doc' = @('Word.Document.8')
  '.bmp' = @('Paint.Picture')
  '.mp4' = @('VLC.mp4')
  '.m4v' = @('VLC.m4v')
  '.webm' = @('VLC.webm')
  '.ogv' = @('VLC.ogv')
  '.mov' = @('VLC.mov')
  '.mkv' = @('VLC.mkv')
  '.xlsx' = @('Excel.Sheet.12')
  '.xlsm' = @('Excel.SheetMacroEnabled.12')
  '.xlsb' = @('Excel.SheetBinaryMacroEnabled.12')
  '.xls' = @('Excel.Sheet.8')
  '.xltx' = @('Excel.Template')
  '.xltm' = @('Excel.TemplateMacroEnabled')
  '.xlt' = @('Excel.Template.8')
  '.xlam' = @('Excel.AddInMacroEnabled')
  '.ods' = @('Excel.OpenDocumentSpreadsheet.12','ods_auto_file')
  '.csv' = @('Excel.CSV')
}

$removedProgIds = [System.Collections.Generic.List[string]]::new()
foreach ($ext in $explorerProgIdRemovals.Keys) {
  $key = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext\OpenWithProgids"
  if (-not (Test-Path -LiteralPath $key)) { continue }
  $names = (Get-ItemProperty -LiteralPath $key).PSObject.Properties.Name
  foreach ($progId in $explorerProgIdRemovals[$ext]) {
    if ($names -contains $progId) {
      Remove-ItemProperty -LiteralPath $key -Name $progId -Force
      $removedProgIds.Add(('{0}:Explorer:{1}' -f $ext, $progId))
    }
  }
}

$classPdfOpenWith = 'Registry::HKEY_CURRENT_USER\Software\Classes\.pdf\OpenWithProgids'
if (Test-Path -LiteralPath $classPdfOpenWith) {
  $names = (Get-ItemProperty -LiteralPath $classPdfOpenWith).PSObject.Properties.Name
  if ($names -contains 'SimplePDF.Document') {
    Remove-ItemProperty -LiteralPath $classPdfOpenWith -Name 'SimplePDF.Document' -Force
    $removedProgIds.Add('.pdf:Classes:SimplePDF.Document')
  }
}

$removedLegacyDefaults = [System.Collections.Generic.List[string]]::new()
$legacyDefaults = [ordered]@{
  '.pdf' = 'simple.Document'
  '.ods' = 'ods_auto_file'
}
foreach ($ext in $legacyDefaults.Keys) {
  $key = "Registry::HKEY_CURRENT_USER\Software\Classes\$ext"
  if ((Test-Path -LiteralPath $key) -and ((Get-Item -LiteralPath $key).GetValue('') -eq $legacyDefaults[$ext])) {
    & reg.exe DELETE "HKCU\Software\Classes\$ext" /ve /f | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Could not remove the stale default value for $ext" }
    $removedLegacyDefaults.Add(('{0}={1}' -f $ext, $legacyDefaults[$ext]))
  }
}

$currentApplication = 'Registry::HKEY_CURRENT_USER\Software\Classes\Applications\simple.exe'
if (Test-Path -LiteralPath $currentApplication) {
  $names = (Get-ItemProperty -LiteralPath $currentApplication).PSObject.Properties.Name
  if ($names -contains 'ApplicationIcon') {
    Remove-ItemProperty -LiteralPath $currentApplication -Name 'ApplicationIcon' -Force
  }
}

$exactKeysToRemove = @(
  'Registry::HKEY_CURRENT_USER\Software\Classes\Applications\simple_calc.exe',
  'Registry::HKEY_CURRENT_USER\Software\Classes\Applications\simple_pdf.exe',
  'Registry::HKEY_CURRENT_USER\Software\Classes\simple.Document',
  'Registry::HKEY_CURRENT_USER\Software\Classes\ods_auto_file',
  'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths\simple.exe',
  'Registry::HKEY_CURRENT_USER\Software\simple\Capabilities'
)
$removedKeys = [System.Collections.Generic.List[string]]::new()
foreach ($key in $exactKeysToRemove) {
  if (Test-Path -LiteralPath $key) {
    Remove-Item -LiteralPath $key -Recurse -Force
    $removedKeys.Add($key)
  }
}

if (-not ('AssociationShellRefresh' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class AssociationShellRefresh {
  [DllImport("shell32.dll")]
  public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
'@
}
[AssociationShellRefresh]::SHChangeNotify(0x08000000, 0x0000, [IntPtr]::Zero, [IntPtr]::Zero)

[pscustomobject]@{
  BackupPath = $backupRoot
  ExportedKeys = $exported
  RemovedOpenWithEntries = $removedOpenWith.Count
  RemovedOpenWithDetail = ($removedOpenWith -join '; ')
  RemovedProgIdReferences = $removedProgIds.Count
  RemovedProgIdDetail = ($removedProgIds -join '; ')
  RemovedLegacyDefaults = $removedLegacyDefaults.Count
  RemovedLegacyDefaultDetail = ($removedLegacyDefaults -join '; ')
  RemovedRegistryKeys = $removedKeys.Count
  RemovedKeyDetail = ($removedKeys -join '; ')
} | ConvertTo-Json -Depth 3
