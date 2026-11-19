# Two Word instances on one local file, without the pane: what does the
# second one get? Word asks "file in use" in a dialog that automation cannot
# answer, so the second instance runs in a process of its own with a time
# limit and is stopped by its process id if it hangs. Only Word processes that
# this script started are touched. Prints one finding per line.
param(
  [Parameter(Mandatory)][string]$File,
  [int]$Wait = 20
)
$ErrorActionPreference = 'Stop'
Add-Type -Namespace SiglumRig -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$dir = Split-Path $File
$pidFile = Join-Path $dir "second.pid"
$outFile = Join-Path $dir "second.txt"
$second = @'
param($File, $PidFile, $OutFile)
$ErrorActionPreference = 'Stop'
Add-Type -Namespace SiglumRig -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$b = $null
$blank = $null
$d = $null
$mine = $null
try {
  $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
  $b = New-Object -ComObject Word.Application
  $new = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
  if ($new.Count -eq 0) {
    "second instance: Word did not start an instance of its own" | Add-Content $OutFile
    return
  }
  if ($b.Windows.Count -gt 0) {
    $wpid = 0
    [void][SiglumRig.Win]::GetWindowThreadProcessId([IntPtr]$b.Windows.Item(1).Hwnd, [ref]$wpid)
    if ($new -notcontains [int]$wpid) {
      "second instance: Word reused an existing process ($wpid)" | Add-Content $OutFile
      return
    }
  }
  $blank = $b.Documents.Add()
  $wpid = 0
  $hwnd = [IntPtr]::Zero
  try { $hwnd = [IntPtr]$blank.ActiveWindow.Hwnd } catch { }
  if ($hwnd -eq [IntPtr]::Zero) {
    $b.Visible = $true
    $hwnd = [IntPtr]$blank.ActiveWindow.Hwnd
  }
  [void][SiglumRig.Win]::GetWindowThreadProcessId($hwnd, [ref]$wpid)
  if ($wpid -eq 0 -or $new -notcontains [int]$wpid) {
    "second instance: window belongs to Word process $wpid, not started by this script" | Add-Content $OutFile
    return
  }
  $mine = [int]$wpid
  Set-Content $PidFile $mine
  $blank.Close([ref]0)
  $blank = $null
  $b.Visible = $false
  $b.DisplayAlerts = 0
  try {
    $d = $b.Documents.Open($File, $false, $false, $false)
    "second instance opened it, read-only: " + $d.ReadOnly | Add-Content $OutFile
    try { $d.Save(); "second instance Save(): no error, read-only " + $d.ReadOnly | Add-Content $OutFile } catch { "second instance Save(): " + $_.Exception.Message | Add-Content $OutFile }
    $d.Close([ref]0)
    $d = $null
  } catch { "second instance Open(): " + $_.Exception.Message | Add-Content $OutFile }
} finally {
  if ($b -and $mine) {
    if ($d) { try { $d.Close([ref]0) } catch { } }
    if ($blank) { try { $blank.Close([ref]0) } catch { } }
    try { $b.Quit([ref]0) } catch { }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($b) } catch { }
    $p = Get-Process -Id $mine -ErrorAction SilentlyContinue
    if ($p -and -not $p.WaitForExit(30000)) { Stop-Process -Id $mine -Force -ErrorAction SilentlyContinue }
  } elseif ($b) {
    if ($d) { try { $d.Close([ref]0) } catch { } }
    if ($blank) { try { $blank.Close([ref]0) } catch { } }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($b) } catch { }
  }
}
'@
$secondPs1 = Join-Path $dir "second.ps1"
$a = $null
$blank = $null
$da = $null
$mine = $null
$proc = $null
try {
  Set-Content $secondPs1 $second
  $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
  $a = New-Object -ComObject Word.Application
  $new = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
  if ($new.Count -eq 0) {
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)
    $a = $null
    throw "Word did not start an instance of its own."
  }
  if ($a.Windows.Count -gt 0) {
    $wpid = 0
    [void][SiglumRig.Win]::GetWindowThreadProcessId([IntPtr]$a.Windows.Item(1).Hwnd, [ref]$wpid)
    if ($new -notcontains [int]$wpid) {
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)
      $a = $null
      throw "Word reused an existing process ($wpid)."
    }
  }
  $blank = $a.Documents.Add()
  $wpid = 0
  $hwnd = [IntPtr]::Zero
  try { $hwnd = [IntPtr]$blank.ActiveWindow.Hwnd } catch { }
  if ($hwnd -eq [IntPtr]::Zero) {
    $a.Visible = $true
    $hwnd = [IntPtr]$blank.ActiveWindow.Hwnd
  }
  [void][SiglumRig.Win]::GetWindowThreadProcessId($hwnd, [ref]$wpid)
  if ($wpid -eq 0 -or $new -notcontains [int]$wpid) {
    throw "The application's window is in Word process $wpid, which this script did not start."
  }
  $mine = [int]$wpid
  $blank.Close([ref]0)
  $blank = $null
  $a.Visible = $false
  $a.DisplayAlerts = 0
  $da = $a.Documents.Open($File, $false, $false, $false)
  $dhwnd = [IntPtr]::Zero
  try { $dhwnd = [IntPtr]$da.ActiveWindow.Hwnd } catch { }
  if ($dhwnd -ne [IntPtr]::Zero) {
    $dpid = 0
    [void][SiglumRig.Win]::GetWindowThreadProcessId($dhwnd, [ref]$dpid)
    if ([int]$dpid -ne $mine) { throw "The document opened in Word process $dpid, not in $mine." }
  }
  "first instance read-only: " + $da.ReadOnly
  $proc = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $secondPs1, $File, $pidFile, $outFile) -PassThru -WindowStyle Hidden
  if (-not $proc.WaitForExit($Wait * 1000)) {
    "second instance: no answer after $Wait s (Open() waits on Word's 'file in use' prompt)"
    if (Test-Path $pidFile) { $wp = [int](Get-Content $pidFile); Stop-Process -Id $wp -Force -ErrorAction SilentlyContinue; "  stopped the second instance (pid $wp)" }
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
  }
  if (Test-Path $outFile) { Get-Content $outFile }
  $da.Close([ref]0)
  $da = $null
} finally {
  try {
    if ($proc -and -not $proc.HasExited) {
      if (Test-Path $pidFile) { $wp = [int](Get-Content $pidFile); Stop-Process -Id $wp -Force -ErrorAction SilentlyContinue }
      Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
  } catch { }
  try {
    if ($a -and $mine) {
      if ($da) { try { $da.Close([ref]0) } catch { } }
      if ($blank) { try { $blank.Close([ref]0) } catch { } }
      try { $a.Quit([ref]0) } catch { }
      try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($a) } catch { }
      $p = Get-Process -Id $mine -ErrorAction SilentlyContinue
      if ($p -and -not $p.WaitForExit(30000)) { Stop-Process -Id $mine -Force -ErrorAction SilentlyContinue }
    } elseif ($a) {
      if ($da) { try { $da.Close([ref]0) } catch { } }
      if ($blank) { try { $blank.Close([ref]0) } catch { } }
      try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($a) } catch { }
    }
  } catch { }
  try {
    Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\Word\Reading Locations' -ErrorAction SilentlyContinue |
      Where-Object { "$($_.GetValue('File Path'))".StartsWith($dir) } | Remove-Item -ErrorAction SilentlyContinue
  } catch { }
  try { Remove-Item -Force -ErrorAction SilentlyContinue $secondPs1, $pidFile, $outFile } catch { }
}
