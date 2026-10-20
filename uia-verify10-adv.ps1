$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$env:EXCLE_SILENT = '1'
$env:EXCLE_AI_KEY = 'test-key-123456'
$env:EXCLE_AI_AUTOAPPLY = '1'

$aiInput = Join-Path $env:TEMP 'excle-ai-inputs.txt'
$pfInput = Join-Path $env:TEMP 'excle-pf-inputs.txt'
function Set-PfInput([string]$k, [string]$v) { Set-Content -Path $pfInput -Value ("{0}={1}" -f $k, $v) -Encoding UTF8 }
$mockDir = Join-Path $env:TEMP 'excle-ai3-mock'
New-Item -ItemType Directory -Force -Path $mockDir | Out-Null
$modeFile = Join-Path $mockDir 'mode.txt'
$port = 8900 + (Get-Random -Maximum 90)
$mockBase = "http://127.0.0.1:$port/"
Set-Content -Path $modeFile -Value 'synth' -Encoding ASCII

$pass = 0; $fail = 0; $failed = @()
$script:xlWin = $null
$script:xl = $null

function Set-AiInput([hashtable]$kv) {
  $lines = @(); foreach ($k in $kv.Keys) { $lines += ("{0}={1}" -f $k, $kv[$k]) }
  Set-Content -Path $aiInput -Value $lines -Encoding UTF8
}
function FindAny([string]$name) {
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $name)
  if ($null -ne $script:xlWin) { $e = $script:xlWin.FindFirst($TS::Descendants, $c); if ($null -ne $e) { return $e } }
  return $AE::RootElement.FindFirst($TS::Descendants, $c)
}
function WaitFor([string]$name, [int]$sec) {
  $end = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $end) { $e = FindAny $name; if ($null -ne $e) { return $e }; Start-Sleep -Milliseconds 300 }
  return $null
}
function Activate([object]$e) {
  try { $p = $e.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern); if ($p) { $p.Invoke(); return $true } } catch { }
  try { $p = $e.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern); if ($p) { if ($p.Current.ExpandCollapseState -eq 'Expanded') { return $true }; $p.Expand(); return $true } } catch { }
  try { $p = $e.GetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern); if ($p) { $p.DoDefaultAction(); return $true } } catch { }
  return $false
}
$script:curTab = $null
function Select-Tab([string]$name) {
  $script:curTab = $name
  $t = WaitFor $name 15
  if ($null -eq $t) { throw "tab '$name' not found" }
  try { $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() } catch { Activate $t | Out-Null }
  Start-Sleep -Milliseconds 900
}
function Focus-Tab([string]$name) {
  Select-Tab 'EXCLE助手'
  Start-Sleep -Milliseconds 400
  Select-Tab $name
  Start-Sleep -Milliseconds 400
}
function Wait-ExcelIdle([int]$sec) {
  $end = (Get-Date).AddSeconds($sec)
  while ((Get-Date) -lt $end) {
    try { $n = $script:xl.ActiveSheet.Name; if ($null -ne $n) { Start-Sleep -Milliseconds 250; return $true } } catch { Start-Sleep -Milliseconds 400 }
  }
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
      if ($null -ne $b -and (Activate $b)) { Start-Sleep -Milliseconds 1500; try { Wait-ExcelIdle 25 | Out-Null } catch { }; return }
    }
    Start-Sleep -Milliseconds 700
  }
  throw "item '$item' not found after 3 attempts"
}
function Check([string]$name, $actual, $expected) {
  $a = if ($null -eq $actual) { '' } else { ([string]$actual).Trim() }
  $e = if ($null -eq $expected) { '' } else { ([string]$expected).Trim() }
  if ($a -eq $e) { $script:pass++; Write-Host ("  PASS  {0,-26} = {1}" -f $name, $a) }
  else { $script:fail++; $script:failed += $name; Write-Host ("  FAIL  {0,-26} got [{1}] want [{2}]" -f $name, $a, $e) }
}
function SheetExists($wb, [string]$n) { foreach ($s in $wb.Worksheets) { if ($s.Name -eq $n) { return $true } }; return $false }

# mock：按 mode 返回 synth / translate
$mockScript = {
  param($modeFile, $port)
  $payloads = @{
    'synth'     = '{"rows":[{"地区":"华东","金额":"120"},{"地区":"华南","金额":"80"},{"地区":"华北","金额":"150"}]}'
    'translate' = '{"values":["Hello","World"]}'
  }
  $listener = New-Object System.Net.HttpListener
  $listener.Prefixes.Add("http://127.0.0.1:$port/")
  $listener.Start()
  while ($true) {
    try {
      $ctx = $listener.GetContext()
      $mode = 'synth'
      if (Test-Path $modeFile) { $mode = (Get-Content $modeFile -Raw).Trim() }
      $reader = New-Object System.IO.StreamReader($ctx.Request.InputStream); $null = $reader.ReadToEnd()
      $content = '{}'; if ($payloads.ContainsKey($mode)) { $content = $payloads[$mode] }
      $body = '{"choices":[{"message":{"content":' + (ConvertTo-Json $content -Compress) + '}}]}'
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
      $ctx.Response.ContentType = 'application/json'
      $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length); $ctx.Response.Close()
    } catch { }
  }
}
$job = Start-Job -ScriptBlock $mockScript -ArgumentList $modeFile, $port
Start-Sleep -Seconds 2

$cfgDir = Join-Path $env:APPDATA 'ExcleHelper'
$cfgLines = @(
  'provider=custom',
  ('baseurl=' + $mockBase.TrimEnd('/') + '/v1'),
  'model=mock-model',
  'temperature=0', 'maxtokens=1024', 'timeout=20',
  'streaming=0', 'jsonmode=0', 'redact=1', 'contextrows=3'
)
Set-Content -Path (Join-Path $cfgDir 'ai.ini') -Value $cfgLines -Encoding UTF8

Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Start-Sleep -Milliseconds 800
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
Start-Sleep -Seconds 2
$wb = $xl.Workbooks.Add(); $xl.WindowState = -4137
Start-Sleep -Seconds 2
$script:xl = $xl
$script:xlWin = $AE::RootElement.FindFirst($TS::Children,
  (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'XLMAIN')))

$ws = $wb.Worksheets[1]

Write-Host ''
Write-Host '=== A3-2 单元格样式库 ==='
$ws.Activate(); $ws.Range('C1:C2').Select() | Out-Null
Set-PfInput 'style' '2'
Focus-Tab '格式与数据'
Click-Ribbon '单元格样式库…' '样式与图表'
$filled = $ws.Range('C1').Interior.Color
Check 'A3-2 表头样式已套用' ($filled -ne 16777215) $true

Write-Host ''
Write-Host '=== A3-3 生成条码 ==='
$ws.Range('E1').Value2 = 'ABC123'
$ws.Range('E1').Select() | Out-Null
Select-Tab '格式与数据'
$before = $ws.Shapes.Count
Click-Ribbon '生成条码（Code39）' '工具与效率'
Check 'A3-3 插入条码图片' ($ws.Shapes.Count - $before) 1

Write-Host ''
Write-Host '=== A4-1 工作簿体检 ==='
$ws.Activate()
Focus-Tab '格式与数据'
Click-Ribbon '工作簿体检…' '工作簿治理'
Check 'A4-1 生成体检表' (SheetExists $wb '工作簿体检') $true
Check 'A4-1 体检行数=9' $wb.Worksheets['工作簿体检'].UsedRange.Rows.Count 9

Write-Host ''
Write-Host '=== C-1 AI 合成测试数据 ==='
$ws.Activate()
$ws.Range('A10').Value2 = '地区'; $ws.Range('B10').Value2 = '金额'
$ws.Range('A10:B10').Select() | Out-Null
Set-Content -Path $modeFile -Value 'synth' -Encoding ASCII
Set-AiInput @{ ask='3' }
Select-Tab '工具与输出'
Focus-Tab '工具与输出'
Click-Ribbon '合成测试数据…' 'AI 进阶'
Check 'C-1 生成合成数据表' (SheetExists $wb '合成数据') $true
Check 'C-1 数据行数=表头+3' $wb.Worksheets['合成数据'].UsedRange.Rows.Count 4

Write-Host ''
Write-Host '=== C-2 AI 批量翻译 ==='
$ws.Activate()
$ws.Range('G1').Value2 = '你好'
$ws.Range('G2').Value2 = '世界'
$ws.Range('G1:G2').Select() | Out-Null
Set-Content -Path $modeFile -Value 'translate' -Encoding ASCII
Set-AiInput @{ ask='英文' }
Focus-Tab '工具与输出'
Click-Ribbon '批量翻译…' 'AI 进阶'
Check 'C-2 译文写入 H1' $ws.Range('H1').Value2 'Hello'
Check 'C-2 译文写入 H2' $ws.Range('H2').Value2 'World'

$xl.Quit()
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Stop-Job $job -ErrorAction SilentlyContinue | Out-Null
Remove-Job $job -Force -ErrorAction SilentlyContinue | Out-Null
Remove-Item $aiInput -Force -ErrorAction SilentlyContinue
Remove-Item $pfInput -Force -ErrorAction SilentlyContinue

Write-Host ''
Write-Host ("RESULT: pass={0} fail={1}" -f $pass, $fail)
if ($fail -gt 0) { Write-Host ("failed: " + ($failed -join ', ')); exit 1 }
