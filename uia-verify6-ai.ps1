$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$env:EXCLE_SILENT = '1'
$env:EXCLE_AI_ASK = '按地区汇总销售额'
$env:EXCLE_AI_AUTOAPPLY = '1'
$env:EXCLE_AI_KEY = 'test-key-123456'

$mockDir = Join-Path $env:TEMP 'excle-ai-mock'
New-Item -ItemType Directory -Force -Path $mockDir | Out-Null
$modeFile = Join-Path $mockDir 'mode.txt'
Set-Content -Path $modeFile -Value 'ok' -Encoding ASCII

# ---------- 启动 mock 供应商 ----------
$mockScript = {
  param($dir, $modeFile)
  $ok = '{"formula":"=SUMIF(B2:B5,\"华东\",D2:D5)","target":"F2","explanation":"按地区汇总销售额","assumptions":["B 列是地区","D 列是销售额"],"confidence":0.92}'
  $risky = '{"formula":"=WEBSERVICE(\"http://x\")","target":"F2","explanation":"拉取外部数据","assumptions":[],"confidence":0.5}'
  $listener = New-Object System.Net.HttpListener
  $listener.Prefixes.Add('http://127.0.0.1:8765/')
  $listener.Start()
  while ($true) {
    try {
      $ctx = $listener.GetContext()
      $mode = 'ok'
      if (Test-Path $modeFile) { $mode = (Get-Content $modeFile -Raw).Trim() }
      $reader = New-Object System.IO.StreamReader($ctx.Request.InputStream)
      $null = $reader.ReadToEnd()
      if ($mode -eq '401') {
        $ctx.Response.StatusCode = 401
        $body = '{"error":{"message":"invalid api key"}}'
      } elseif ($mode -eq 'badjson') {
        $body = '{"choices":[{"message":{"content":"抱歉，我无法给出公式。"}}]}'
      } elseif ($mode -eq 'risky') {
        $body = '{"choices":[{"message":{"content":' + (ConvertTo-Json $risky -Compress) + '}}]}'
      } else {
        $body = '{"choices":[{"message":{"content":' + (ConvertTo-Json $ok -Compress) + '}}]}'
      }
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
      $ctx.Response.ContentType = 'application/json'
      $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
      $ctx.Response.Close()
    } catch { }
  }
}
$job = Start-Job -ScriptBlock $mockScript -ArgumentList $mockDir, $modeFile
Start-Sleep -Seconds 2
"mock 供应商已启动（job $($job.Id)）"

# ---------- 写 AI 配置（指向 mock） ----------
$cfgDir = Join-Path $env:APPDATA 'ExcleHelper'
New-Item -ItemType Directory -Force -Path $cfgDir | Out-Null
@'
provider=custom
baseurl=http://127.0.0.1:8765/v1
model=mock-model
temperature=0
maxtokens=512
timeout=20
streaming=0
jsonmode=0
redact=1
contextrows=3
'@ | Set-Content -Path (Join-Path $cfgDir 'ai.ini') -Encoding UTF8

# ---------- 工具函数 ----------
$pass = 0; $fail = 0; $failed = @()
$script:xlWin = $null
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
function Check([string]$name, $actual, $expected) {
  $a = if ($null -eq $actual) { '' } else { ([string]$actual).Trim() }
  $e = if ($null -eq $expected) { '' } else { ([string]$expected).Trim() }
  if ($a -eq $e) { $script:pass++; Write-Host ("  PASS  {0,-26} = {1}" -f $name, $a) }
  else { $script:fail++; $script:failed += $name; Write-Host ("  FAIL  {0,-26} got [{1}] want [{2}]" -f $name, $a, $e) }
}

# ---------- 启动 Excel ----------
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Start-Sleep -Milliseconds 700
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
Start-Sleep -Seconds 2
$wb = $xl.Workbooks.Add()
$xl.WindowState = -4137
Start-Sleep -Seconds 2
$script:xlWin = $AE::RootElement.FindFirst($TS::Children,
  (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'XLMAIN')))
$ws = $wb.Worksheets[1]

# 数据表
$ws.Range('A1').Value2 = '日期'; $ws.Range('B1').Value2 = '地区'; $ws.Range('C1').Value2 = '产品'; $ws.Range('D1').Value2 = '销售额'
$ws.Range('A2').Value2 = '2024-01-05'; $ws.Range('B2').Value2 = '华东'; $ws.Range('C2').Value2 = 'A'; $ws.Range('D2').Value2 = 1200
$ws.Range('A3').Value2 = '2024-02-11'; $ws.Range('B3').Value2 = '华南'; $ws.Range('C3').Value2 = 'B'; $ws.Range('D3').Value2 = 900
$ws.Range('A4').Value2 = '2024-03-02'; $ws.Range('B4').Value2 = '华东'; $ws.Range('C4').Value2 = 'A'; $ws.Range('D4').Value2 = 1500
$ws.Range('A5').Value2 = '2024-04-18'; $ws.Range('B5').Value2 = '华北'; $ws.Range('C5').Value2 = 'C'; $ws.Range('D5').Value2 = 700
$ws.Range('A1:D5').Select() | Out-Null
Start-Sleep -Milliseconds 500

Write-Host ''
Write-Host '=== AI 生成公式（正常返回）==='
Select-Tab '工具与输出'
Click-Ribbon '生成公式'
Check 'ai.formula 写入 F2' $ws.Range('F2').Formula '=SUMIF(B2:B5,"华东",D2:D5)'

Write-Host ''
Write-Host '=== 危险函数必须被拦下 ==='
$ws.Range('F2').ClearContents()
Set-Content -Path $modeFile -Value 'risky' -Encoding ASCII
Start-Sleep -Milliseconds 300
$ws.Range('A1:D5').Select() | Out-Null
Click-Ribbon '生成公式'
Check 'WEBSERVICE 未被写入' $ws.Range('F2').Formula $null

Write-Host ''
Write-Host '=== 模型返回非 JSON ==='
Set-Content -Path $modeFile -Value 'badjson' -Encoding ASCII
Start-Sleep -Milliseconds 300
$ws.Range('A1:D5').Select() | Out-Null
Click-Ribbon '生成公式'
Check '非 JSON 时不写入' $ws.Range('F2').Formula $null

Write-Host ''
Write-Host '=== 接口 401 ==='
Set-Content -Path $modeFile -Value '401' -Encoding ASCII
Start-Sleep -Milliseconds 300
$ws.Range('A1:D5').Select() | Out-Null
Click-Ribbon '生成公式'
Check '401 时不写入' $ws.Range('F2').Formula $null

Write-Host ''
Write-Host '=== AI 设置窗 ==='
Click-Ribbon 'AI 设置…' '设置与测试'
Start-Sleep -Seconds 1
$winCond = New-Object System.Windows.Automation.AndCondition(
  (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, 'EXCLE助手 · AI 设置')),
  (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)))
$dlg = $null
$end = (Get-Date).AddSeconds(8)
while ((Get-Date) -lt $end -and $null -eq $dlg) {
  $dlg = $AE::RootElement.FindFirst($TS::Descendants, $winCond)
  if ($null -eq $dlg) { Start-Sleep -Milliseconds 400 }
}
Check '设置窗已打开' ($null -ne $dlg) $true
if ($null -ne $dlg) {
  $all = $dlg.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
  $joined = ''
  foreach ($e in $all) {
    $joined += $e.Current.Name + '|'
    try { $joined += $e.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value + '|' } catch { }
  }
  Check '设置窗显示 mock 地址' ($joined -match '8765') $true
  Check '设置窗显示模型名' ($joined -match 'mock-model') $true
  try { $dlg.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch { }
}

$xl.Quit()
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { if (-not $_.HasExited) { try { Stop-Process -Id $_.Id -Force } catch { } } }
Stop-Job $job -ErrorAction SilentlyContinue | Out-Null
Remove-Job $job -Force -ErrorAction SilentlyContinue | Out-Null

Write-Host ''
Write-Host ("RESULT: pass={0} fail={1}" -f $pass, $fail)
if ($fail -gt 0) { Write-Host ("failed: " + ($failed -join ', ')); exit 1 }
