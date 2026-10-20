# Build ExcleHelper.dll (EXCLE助手) with the in-box .NET Framework 4.8 compiler.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$src  = Join-Path $here 'src'
$bin  = Join-Path $here 'bin'
New-Item -ItemType Directory -Force -Path $bin | Out-Null

$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { throw "csc not found: $csc" }

$outDll = Join-Path $bin 'ExcleHelper.dll'
if (Test-Path $outDll) {
  $free = $true
  try { $fs = [System.IO.File]::Open($outDll, 'Open', 'ReadWrite', 'None'); $fs.Close() } catch { $free = $false }
  if (-not $free) {
    $parked = Join-Path $bin ('ExcleHelper.locked-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.dll')
    Move-Item $outDll $parked -Force
    Write-Host "note: previous dll in use -> parked as $(Split-Path -Leaf $parked)"
  }
}

$files = @(Get-ChildItem $src -Filter *.cs -Recurse | Sort-Object Name | ForEach-Object { $_.FullName })
if ($files.Count -eq 0) { throw "no sources in $src" }

$refs = @(
  '/r:System.dll', '/r:System.Core.dll', '/r:System.Windows.Forms.dll',
  '/r:System.Drawing.dll', '/r:Microsoft.VisualBasic.dll', '/r:Microsoft.CSharp.dll',
  '/r:System.Xml.dll', '/r:System.Security.dll'
)

$cscArgs = @(
  '/target:library', '/platform:anycpu', '/optimize+', '/nologo', '/warn:2',
  "/out:$outDll",
  "/res:$src\ribbon.xml,ExcleHelper.ribbon.xml"
) + $refs + $files

Write-Host "csc $($files.Count) sources -> $outDll"
& $csc $cscArgs
if ($LASTEXITCODE -ne 0) { throw "compile failed (exit $LASTEXITCODE)" }

$dll = Get-Item $outDll
Write-Host ("OK  {0}  {1:N1} KB" -f $dll.FullName, ($dll.Length / 1KB))
