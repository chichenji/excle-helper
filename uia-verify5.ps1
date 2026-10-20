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
Write-Host '=== 文本高级 ==='
$ws.Range('A1').Value2 = 'apple'
$ws.Range('B1').Value2 = 'appel'
$ws.Range('A1:B1').Select() | Out-Null
Select-Tab 'EXCLE助手'
Click-Ribbon '两列相似度' '高级文本'
Check 't2.similar' $ws.Range('C1').Value2 0.6

$ws.Range('A3').Value2 = 'aabbcc'
$ws.Range('A3').Select() | Out-Null
Click-Ribbon '删除重复字符' '高级文本'
Check 't2.uniqchar' $ws.Range('A3').Value2 'abc'

Write-Host ''
Write-Host '=== 隐私与地址 ==='
Select-Tab '格式与数据'
$ws.Range('E1').Value2 = '广东省深圳市南山区科技园1号'
$ws.Range('E1').Select() | Out-Null
Click-Ribbon '地址拆分（省市区）' '隐私与安全'
Check 'p2.addr 省' $ws.Range('F1').Value2 '广东省'
Check 'p2.addr 市' $ws.Range('G1').Value2 '深圳市'
Check 'p2.addr 区' $ws.Range('H1').Value2 '南山区'

$phone = '139' + '12345678'
$ws.Range('E3').NumberFormat = '@'
$ws.Range('E3').Value2 = $phone
$ws.Range('E3').Select() | Out-Null
Click-Ribbon '一键脱敏' '隐私与安全'
Check 'p2.mask 手机号' $ws.Range('E3').Value2 '139****5678'

Write-Host ''
Write-Host '=== 数据体检与结构 ==='
Select-Tab '格式与数据'
$ws.Range('A5').Value2 = 'x'
$ws.Range('A6').Value2 = 'x'
$ws.Range('B5').Value2 = $null
$ws.Range('A5:B6').Select() | Out-Null
Click-Ribbon '一键数据体检' '数据体检'
$hasAudit = $false
foreach ($s in $wb.Worksheets) { if ($s.Name -eq '数据体检') { $hasAudit = $true } }
Check 'q2.audit 生成报告表' $hasAudit $true

$ws.Activate()
$ws.Range('M1').Value2 = 'a'
$ws.Range('M2').Value2 = 'a'
$ws.Range('M3').Value2 = 'b'
$ws.Range('M4').Value2 = 'b'
$ws.Range('M1:M4').Select() | Out-Null
Click-Ribbon '分组序号' '数据体检'
Check 'q2.seq 组内序号1' $ws.Range('N1').Value2 1
Check 'q2.seq 组内序号2' $ws.Range('N2').Value2 2
Check 'q2.seq 新组重新计数' $ws.Range('N3').Value2 1

Write-Host ''
Write-Host '=== 注册总数 ==='
$ro = New-Object -ComObject 'ExcleHelper.Connect'
$ops = [int](($ro.Ping() -replace '.*ops=(\d+).*', '$1'))
Check 'ops >= 275' ($ops -ge 275) $true

Write-Host ''
Write-Host '=== 二维码 ==='
Select-Tab '工具与输出'
$ws.Range('O1').NumberFormat = '@'
$ws.Range('O1').Value2 = 'excle-helper-qr-test'
$ws.Range('O1').Select() | Out-Null
Start-Sleep -Milliseconds 400
$before = $ws.Shapes.Count
Click-Ribbon '生成二维码（取单元格文本）' '二维码与截图'
$after = $ws.Shapes.Count
Check 'pic.qrcode 插入图片' ($after - $before) 1

$xl.Quit()
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }

Write-Host ''
Write-Host ("RESULT: pass={0} fail={1}" -f $pass, $fail)
if ($fail -gt 0) { Write-Host ("failed: " + ($failed -join ', ')); exit 1 }
