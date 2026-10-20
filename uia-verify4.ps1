$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$env:EXCLE_SILENT = '1'
$env:EXCLE_TRACE = '1'

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
  try { $p = $e.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern); if ($p) { $p.Expand(); return $true } } catch { }
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
  if ($a -eq $e) { $script:pass++; Write-Host ("  PASS  {0,-26} = {1}" -f $name, $a) }
  else { $script:fail++; $script:failed += $name; Write-Host ("  FAIL  {0,-26} got [{1}] want [{2}]" -f $name, $a, $e) }
}

Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Start-Sleep -Milliseconds 700

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
Start-Sleep -Seconds 2
$wb = $xl.Workbooks.Add()
$xl.WindowState = -4137
Start-Sleep -Seconds 2
$script:xlWin = $AE::RootElement.FindFirst($TS::Children,
  (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'XLMAIN')))
$ws = $wb.Worksheets[1]

Write-Host ''
Write-Host '=== 回退系统 ==='
# 数据：A 列数字 + B 列公式
$ws.Range('A1').Value2 = 1
$ws.Range('A2').Value2 = 2
$ws.Range('A3').Value2 = 3
$ws.Range('B1').Formula = '=A1*10'
$ws.Range('A1:A3').Select() | Out-Null
Start-Sleep -Milliseconds 400

Select-Tab 'EXCLE助手'
Click-Ribbon '行顺序上下倒置' '行列重排'
Check '倒置后 A1' $ws.Range('A1').Value2 3

Select-Tab 'EXCLE助手'
Click-Ribbon '撤销上次操作' '回退'
Check '回退后 A1' $ws.Range('A1').Value2 1
Check '回退后 A3' $ws.Range('A3').Value2 3
Check '公式仍在' $ws.Range('B1').HasFormula $true

Click-Ribbon '重做上次操作' '回退'
Check '重做后 A1' $ws.Range('A1').Value2 3

Write-Host ''
Write-Host '=== 回退覆盖「写右侧」的操作 ==='
Click-Ribbon '撤销上次操作' '回退'
$ws.Range('E1').NumberFormat = '@'
$ws.Range('E1').Value2 = '2024年3月5日'
$ws.Range('E2').Value2 = '2024/6/1'
$ws.Range('E1:E2').Select() | Out-Null
Start-Sleep -Milliseconds 400
Click-Ribbon '日期拆分(6 列)' '日期时间'
Check '拆分写入了 F1' $ws.Range('F1').Value2 2024
Click-Ribbon '撤销上次操作' '回退'
$cleared = ($null -eq $ws.Range('F1').Value2) -or ([string]$ws.Range('F1').Value2 -eq '')
Check '回退清掉了右侧写入' $cleared $true
Check '源数据仍在' $ws.Range('E1').Value2 '2024年3月5日'

Write-Host ''
Write-Host '=== 回退新增工作表 ==='
Select-Tab '格式与数据'
Click-Ribbon '智能索引' '表名与目录'
$hasDir = $false
foreach ($s in $wb.Worksheets) { if ($s.Name -eq '目录') { $hasDir = $true } }
Check '目录表已生成' $hasDir $true
Select-Tab 'EXCLE助手'
Click-Ribbon '撤销上次操作' '回退'
$stillDir = $false
foreach ($s in $wb.Worksheets) { if ($s.Name -eq '目录') { $stillDir = $true } }
Check '回退删除了新增表' $stillDir $false

$xl.Quit()
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }

Write-Host ''
Write-Host ("RESULT: pass={0} fail={1}" -f $pass, $fail)
if ($fail -gt 0) { Write-Host ("failed: " + ($failed -join ', ')); exit 1 }
