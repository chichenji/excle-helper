$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$env:EXCLE_SILENT = '1'
$env:EXCLE_TRACE = '1'
$env:EXCLE_SHOT = '1'

$pass = 0; $fail = 0; $failed = @()

$script:xlWin = $null
function FindAny([string]$name) {
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $name)
  if ($null -ne $script:xlWin) {
    $e = $script:xlWin.FindFirst($TS::Descendants, $c)
    if ($null -ne $e) { return $e }
  }
  return $AE::RootElement.FindFirst($TS::Descendants, $c)
}
function WaitFor([string]$name, [int]$sec) {
  $end = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $end) {
    $e = FindAny $name
    if ($null -ne $e) { return $e }
    Start-Sleep -Milliseconds 350
  }
  return $null
}
function Activate([object]$e) {
  try { $p = $e.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern); if ($p) { $p.Invoke(); return $true } } catch { }
  try {
    $p = $e.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern)
    if ($p) {
      if ($p.Current.ExpandCollapseState -eq [System.Windows.Automation.ExpandCollapseState]::Expanded) { return $true }
      $p.Expand(); return $true
    }
  } catch { }
  try { $p = $e.GetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern); if ($p) { $p.DoDefaultAction(); return $true } } catch { }
  return $false
}
function Click-Ribbon([string]$item, [string]$menu) {
  for ($attempt = 1; $attempt -le 4; $attempt++) {
    Start-Sleep -Milliseconds 400
    $menuOk = $true
    if ($menu) {
      $m = WaitFor $menu 8
      if ($null -eq $m) { $menuOk = $false }
      elseif (-not (Activate $m)) { $menuOk = $false }
      else { Start-Sleep -Milliseconds 1200 }
    }
    if ($menuOk) {
      $b = WaitFor $item 8
      if ($null -ne $b -and (Activate $b)) { Start-Sleep -Milliseconds 1500; return }
    }
    Start-Sleep -Milliseconds 700
  }
  throw "item '$item' not found after 3 attempts"
}
$script:curTab = $null
function Select-Tab([string]$name) {
  $script:curTab = $name
  $t = WaitFor $name 15
  if ($null -eq $t) { throw "tab '$name' not found" }
  try { $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() } catch { Activate $t | Out-Null }
  Start-Sleep -Milliseconds 900
}
function Check([string]$name, $actual, $expected) {
  $a = if ($null -eq $actual) { '' } else { ([string]$actual).Trim() }
  $e = if ($null -eq $expected) { '' } else { ([string]$expected).Trim() }
  if ($a -eq $e) { $script:pass++; Write-Host ("  PASS  {0,-24} = {1}" -f $name, $a) }
  else { $script:fail++; $script:failed += $name; Write-Host ("  FAIL  {0,-24} got [{1}] want [{2}]" -f $name, $a, $e) }
}

Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Start-Sleep -Milliseconds 700
Remove-Item (Join-Path $env:TEMP 'exclehelper-panel.png') -ErrorAction SilentlyContinue
Remove-Item (Join-Path $env:TEMP 'exclehelper-about.png') -ErrorAction SilentlyContinue

$xl = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $true
  $xl.DisplayAlerts = $false
  Start-Sleep -Seconds 2
  $wb = $xl.Workbooks.Add()
  $xl.WindowState = -4137
  $ws = $xl.ActiveSheet
  $ws.Activate()

  Write-Host ''
  Write-Host '=== 选项卡结构 ==='
  foreach ($t in @('EXCLE助手','格式与数据','工具与输出')) {
    $el = WaitFor $t 8
    Check ("tab " + $t) ($null -ne $el) $true
  }

  Write-Host ''
  Write-Host '=== 功能面板 / 关于（离屏渲染）==='
  Select-Tab 'EXCLE助手'
  Click-Ribbon '功能面板'
  Start-Sleep -Milliseconds 1200
  $panelPng = Join-Path $env:TEMP 'exclehelper-panel.png'
  Check 'panel-shot' (Test-Path $panelPng) $true
  if (Test-Path $panelPng) { Write-Host ("  panel png: " + (Get-Item $panelPng).Length + " bytes") }

  Click-Ribbon '关于'
  Start-Sleep -Milliseconds 1200
  $aboutPng = Join-Path $env:TEMP 'exclehelper-about.png'
  Check 'about-shot' (Test-Path $aboutPng) $true
  if (Test-Path $aboutPng) { Write-Host ("  about png: " + (Get-Item $aboutPng).Length + " bytes") }

  Write-Host ''
  Write-Host '=== 编辑 / 格式 / 工作表 新功能 ==='
  Select-Tab 'EXCLE助手'
  $ws.Range('A1').Value2 = 1
  $ws.Range('A2').Value2 = 2
  $ws.Range('A1:A2').Select() | Out-Null
  Click-Ribbon '行顺序上下倒置' '行列重排'
  Check 'e.flip.rows' $ws.Range('A1').Value2 2

  $ws.Range('B1').Value2 = 5
  $ws.Range('B2').Value2 = $null
  $ws.Range('B3').Value2 = 7
  $ws.Range('B1:B3').Select() | Out-Null
  Click-Ribbon '删除空单元格(上移)' '清空与删除'
  Check 'e.del.blankcells' $ws.Range('B2').Value2 7

  $ws.Range('C1').Value2 = 'x'
  $ws.Range('C2').Value2 = $null
  $ws.Range('C1:C2').Select() | Out-Null
  Click-Ribbon '空白并入上方' '合并单元格'
  Check 'e.merge.blank' $ws.Range('C1').MergeCells $true

  Select-Tab 'EXCLE助手'
  $ws.Range('D1').Value2 = 'abc@'
  $ws.Range('D1').Select() | Out-Null
  Select-Tab '格式与数据'; Click-Ribbon '邮箱校验' '校验标记'
  Check 'v.email.check' $ws.Range('E1').Value2 '错误'

  $ws.Range('D3').Value2 = $null
  $ws.Range('D3').Select() | Out-Null
  Select-Tab '格式与数据'; Click-Ribbon '标记空值' '校验标记'
  $c = $ws.Range('D3').Interior.Color
  Check 'v.null.highlight' ($c -eq 65535) $true

  $ws.Range('F1').Value2 = 9
  $ws.Range('F1:F1').Select() | Out-Null
  Select-Tab '格式与数据'; Click-Ribbon '自动调整'

  Select-Tab '格式与数据'
  $ws.Activate()
  Click-Ribbon '智能索引' '表名与目录'
  $hit = $null
  foreach ($s in $wb.Worksheets) { if ($s.Name -eq '目录') { $hit = $s.Name } }
  Check 'w.smart.index' $hit '目录'
}
catch {
  Write-Host ('EXCEPTION: ' + $_.Exception.Message)
  Write-Host $_.ScriptStackTrace
  $fail++
}
finally {
  Start-Sleep -Milliseconds 400
  Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
}

Write-Host ''
Write-Host ("RESULT: pass={0} fail={1}" -f $pass, $fail)
if ($failed.Count -gt 0) { Write-Host ("failed: " + ($failed -join ', ')) }
if ($fail -gt 0) { exit 1 }
exit 0
