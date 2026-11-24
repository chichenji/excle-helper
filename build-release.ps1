# Build GitHub Release runnable artifacts into _work\release\
# Outputs:
#   1. ExcleHelper-Setup-v<ver>.exe       (single-file GUI installer/uninstaller with embedded DLL)
#   2. ExcleHelper-v<ver>-win-anycpu.zip  (portable ZIP: bin\ExcleHelper.dll + install/uninstall .bat/.ps1 + README + LICENSE)
#   3. SHA256SUMS.txt
param(
  [string]$Version = '1.1.0'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

# 1. Rebuild main DLL cleanly
& (Join-Path $root 'build.ps1')
$dll = Join-Path $root 'bin\ExcleHelper.dll'
if (-not (Test-Path $dll)) { throw "missing built dll: $dll" }

$outDir = Join-Path $root '_work\release'
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

# 2. Build single-file Setup EXE embedding bin\ExcleHelper.dll
$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { $csc = 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (-not (Test-Path $csc)) { throw 'csc.exe not found' }

$setupExe = Join-Path $outDir ("ExcleHelper-Setup-v{0}.exe" -f $Version)
$setupSrc = Join-Path $root 'tools\Setup.cs'
$args = @(
  '/nologo',
  '/target:winexe',
  '/platform:anycpu',
  '/optimize+',
  '/warn:4',
  '/utf8output',
  "/out:$setupExe",
  "/resource:$dll,ExcleHelper.dll",
  '/r:System.dll',
  '/r:System.Drawing.dll',
  '/r:System.Windows.Forms.dll',
  $setupSrc
)
Write-Host "csc Setup.cs -> $setupExe"
& $csc @args
if ($LASTEXITCODE -ne 0) { throw "Setup.cs compile failed ($LASTEXITCODE)" }

# 3. Build portable runnable ZIP
$stage = Join-Path $outDir ("stage\ExcleHelper-v{0}" -f $Version)
New-Item -ItemType Directory -Force -Path (Join-Path $stage 'bin') | Out-Null
Copy-Item $dll (Join-Path $stage 'bin\ExcleHelper.dll') -Force
foreach ($f in @('install.bat', 'uninstall.bat', 'install.ps1', 'uninstall.ps1', 'README.md', 'LICENSE')) {
  Copy-Item (Join-Path $root $f) (Join-Path $stage $f) -Force
}
$zipPath = Join-Path $outDir ("ExcleHelper-v{0}-win-anycpu.zip" -f $Version)
Compress-Archive -Path $stage -DestinationPath $zipPath -CompressionLevel Optimal -Force
Remove-Item (Join-Path $outDir 'stage') -Recurse -Force

# 4. Generate SHA256SUMS.txt
$sumFile = Join-Path $outDir 'SHA256SUMS.txt'
$lines = @()
foreach ($asset in @($setupExe, $zipPath)) {
  $h = (Get-FileHash -Path $asset -Algorithm SHA256).Hash.ToLowerInvariant()
  $lines += ("{0}  {1}" -f $h, (Split-Path -Leaf $asset))
}
Set-Content -Path $sumFile -Value $lines -Encoding ASCII

Write-Host ''
Write-Host 'Release artifacts ready:'
Get-ChildItem $outDir -File | ForEach-Object {
  Write-Host ("  {0,-38} {1,8:N1} KB" -f $_.Name, ($_.Length / 1KB))
}
