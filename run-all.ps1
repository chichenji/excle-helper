# 统一回归入口：逐套运行，套间清理 Excel 与"文档恢复"状态，避免互相污染
$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$suites = @(
  'uia-verify2.ps1',
  'uia-verify3.ps1',
  'uia-verify4.ps1',
  'uia-verify5.ps1',
  'uia-verify6-ai.ps1',
  'uia-verify7-join.ps1',
  'uia-verify8-ai2.ps1',
  'uia-verify9-quality.ps1',
  'uia-verify10-adv.ps1'
)

# 强制杀 Excel 会累积"文档恢复"条目；下次启动时恢复面板抢焦点，导致自动化点击全部失效。
# 这里把它们备份到 _work\excel-recovery-backup（移动文件 + 导出注册表）后清理。
function Clear-ExcelRecovery {
  $backup = Join-Path $root '_work\excel-recovery-backup'
  if (-not (Test-Path $backup)) { New-Item -ItemType Directory -Force -Path $backup | Out-Null }
  $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'

  $excelDir = Join-Path $env:APPDATA 'Microsoft\Excel'
  if (Test-Path $excelDir) {
    Get-ChildItem $excelDir -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -match 'Autorecovered|\.xar$' } |
      ForEach-Object {
        try { Move-Item $_.FullName (Join-Path $backup ($stamp + '-' + $_.Name)) -Force } catch { }
      }
  }

  $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DocumentRecovery'
  if (Test-Path $key) {
    $reg = Join-Path $backup ($stamp + '-DocumentRecovery.reg')
    try { reg.exe export 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Resiliency\DocumentRecovery' $reg /y 2>&1 | Out-Null } catch { }
    try { Remove-Item $key -Recurse -Force } catch { }
  }
}

function Reset-Excel {
  Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
    if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } }
  }
  $deadline = (Get-Date).AddSeconds(30)
  while ((Get-Date) -lt $deadline) {
    $still = Get-Process EXCEL -ErrorAction SilentlyContinue
    if (-not $still) { break }
    Start-Sleep -Milliseconds 500
  }
  Clear-ExcelRecovery
  Start-Sleep -Seconds 3
}

$total = 0
$failedSuites = @()

Reset-Excel
foreach ($s in $suites) {
  $path = Join-Path $root $s
  if (-not (Test-Path $path)) { Write-Host ("  跳过（不存在）：" + $s); continue }

  $out = pwsh -NoProfile -File $path 2>&1
  $line = ($out | Select-String -Pattern 'RESULT' | ForEach-Object { $_.Line.Trim() }) -join ' '

  $bad = (-not $line) -or ($line -match 'fail=(\d+)' -and [int]$Matches[1] -gt 0)
  if ($bad) {
    # Excel 会话状态偶发不稳（页签懒加载 / 恢复面板），失败重试一次
    Write-Host ("        （首次未通过，重置后重试一次）")
    Reset-Excel
    $out = pwsh -NoProfile -File $path 2>&1
    $line2 = ($out | Select-String -Pattern 'RESULT' | ForEach-Object { $_.Line.Trim() }) -join ' '
    if ($line2) { $line = $line2 }
    $bad = (-not $line) -or ($line -match 'fail=(\d+)' -and [int]$Matches[1] -gt 0)
  }

  if ($line -match 'pass=(\d+)') { $total += [int]$Matches[1] }
  if ($bad) { $failedSuites += $s; if (-not $line) { $line = '(无输出，可能异常退出)' } }

  Write-Host ("  {0,-22} {1}" -f $s, $line)
  if ($failedSuites -contains $s) {
    $out | Select-String -Pattern 'FAIL|EXCEPTION|not found|not actionable' | Select-Object -First 3 |
      ForEach-Object { "        " + $_.Line.Trim() }
  }
  Reset-Excel
}

Write-Host ''
Write-Host ("合计通过：{0}" -f $total)
if ($failedSuites.Count -gt 0) {
  Write-Host ("有失败的套件：" + ($failedSuites -join ', '))
  exit 1
} else {
  Write-Host '全部通过 ✓'
}
