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
  if ($a -eq $e) { $script:pass++; Write-Host ("  PASS  {0,-24} = {1}" -f $name, $a) }
  else { $script:fail++; $script:failed += $name; Write-Host ("  FAIL  {0,-24} got [{1}] want [{2}]" -f $name, $a, $e) }
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

Write-Host ''
Write-Host '=== 选项卡 ==='
foreach ($t in @('EXCLE助手','格式与数据','工具与输出')) {
  $found = $null -ne (WaitFor $t 6)
  Check ("tab " + $t) $found $true
}

Write-Host ''
Write-Host '=== 批次 1：日期 ==='
$ws = $wb.Worksheets[1]
$ws.Cells.Item(1,1).NumberFormat = '@'
$ws.Cells.Item(1,1).Value2 = '2024年3月5日'
$ws.Cells.Item(2,1).NumberFormat = '@'
$ws.Cells.Item(2,1).Value2 = '2023-12-31'
$ws.Cells.Item(3,1).NumberFormat = '@'
$ws.Cells.Item(3,1).Value2 = '2024/6/1'
$rng = $ws.Range('A1:A3'); $rng.Select() | Out-Null
Start-Sleep -Milliseconds 400
Select-Tab 'EXCLE助手'
Click-Ribbon '日期拆分(6 列)' '日期时间'
Check 'dt.split 年' $ws.Cells.Item(1,2).Value2 2024
Check 'dt.split 月' $ws.Cells.Item(1,3).Value2 3
Check 'dt.split 日' $ws.Cells.Item(1,4).Value2 5
Check 'dt.split 星期' $ws.Cells.Item(1,5).Value2 '星期二'
Check 'dt.split 季度' $ws.Cells.Item(1,6).Value2 '第1季度'

Write-Host ''
Write-Host '=== 批次 4：统计 ==='
$ws.Cells.Item(1,9).Value2 = 10
$ws.Cells.Item(2,9).Value2 = 30
$ws.Cells.Item(3,9).Value2 = 60
$rng2 = $ws.Range('I1:I3'); $rng2.Select() | Out-Null
Start-Sleep -Milliseconds 400
Select-Tab '格式与数据'
Click-Ribbon '占比' '统计分析'
Check 'st.share' ([math]::Round([double]$ws.Cells.Item(1,10).Value2, 4)) 0.1

Write-Host ''
Write-Host '=== 批次 4：条件格式 ==='
$rng2.Select() | Out-Null
Start-Sleep -Milliseconds 400
Select-Tab '格式与数据'
Click-Ribbon '数据条' '条件格式'
$fc = 0
try { $fc = $ws.Range('I1:I3').FormatConditions.Count } catch { }
Check 'cf.bars 条件格式数' $fc 1
$rng2.Select() | Out-Null
Start-Sleep -Milliseconds 300
Click-Ribbon '三色阶' '条件格式'
$fc2 = 0
try { $fc2 = $ws.Range('I1:I3').FormatConditions.Count } catch { }
Check 'cf.scale 条件格式数' $fc2 1

Write-Host ''
Write-Host '=== 批次 2：公式视图 ==='
Select-Tab '工具与输出'
Click-Ribbon '显示全部公式' '公式工具'
Check 'fm.show.all 开' $xl.ActiveWindow.DisplayFormulas $true
Click-Ribbon '显示全部公式' '公式工具'
Check 'fm.show.all 关' $xl.ActiveWindow.DisplayFormulas $false

Write-Host ''
Write-Host '=== 注册表 ==='
$ro = New-Object -ComObject 'ExcleHelper.Connect'
Check 'ops 总数 >= 240' ([int](($ro.Ping() -replace '.*ops=(\d+).*','$1')) -ge 240) $true

Write-Host ''
Write-Host '=== 回归：中文日期标准化 + 公式不被覆盖 ==='
$ws.Range('M1').NumberFormat = '@'
$ws.Range('M1').Value2 = '2024年3月5日'
$ws.Range('M2').Formula = '=1+1'
$ws.Range('M1:M2').Select() | Out-Null
Start-Sleep -Milliseconds 400
Select-Tab 'EXCLE助手'
Click-Ribbon '中文日期标准化' '日期时间'
Check 'dt.normalize 序列号' ([math]::Round([double]$ws.Range('M1').Value2, 0)) 45356
Check '公式未被覆盖' $ws.Range('M2').HasFormula $true

$xl.Quit()
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }

Write-Host ''
Write-Host ("RESULT: pass={0} fail={1}" -f $pass, $fail)
if ($fail -gt 0) { Write-Host ("failed: " + ($failed -join ', ')); exit 1 }
