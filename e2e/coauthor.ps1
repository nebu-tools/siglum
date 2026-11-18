# Two Word instances on one local file, without the pane: what does the second
# one get? Word asks "file in use" in a dialog that automation cannot answer,
# so the second instance runs in its own process with a time limit and is
# stopped by its process id if it hangs. Only instances started here are
# touched. Prints one finding per line.
param(
  [Parameter(Mandatory)][string]$File,
  [int]$Wait = 20
)
$ErrorActionPreference = 'Stop'
Add-Type -Namespace Siglum -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$dir = Split-Path $File
$pidFile = Join-Path $dir "second.pid"
$outFile = Join-Path $dir "second.txt"
$second = @'
param($File, $PidFile, $OutFile)
Add-Type -Namespace Siglum -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$b = New-Object -ComObject Word.Application
$b.DisplayAlerts = 0
$blank = $b.Documents.Add()
$p = 0; [void][Siglum.Win]::GetWindowThreadProcessId([IntPtr]$blank.ActiveWindow.Hwnd, [ref]$p)
Set-Content $PidFile $p
$blank.Close([ref]0)
try {
  $d = $b.Documents.Open($File, $false, $false, $false)
  "second instance opened it, read-only: " + $d.ReadOnly | Add-Content $OutFile
  try { $d.Save(); "second instance Save(): no error, read-only " + $d.ReadOnly | Add-Content $OutFile } catch { "second instance Save(): " + $_.Exception.Message | Add-Content $OutFile }
  $d.Close([ref]0)
} catch { "second instance Open(): " + $_.Exception.Message | Add-Content $OutFile }
$b.Quit([ref]0)
'@
$secondPs1 = Join-Path $dir "second.ps1"
Set-Content $secondPs1 $second
$a = $null
try {
  $a = New-Object -ComObject Word.Application
  $a.DisplayAlerts = 0
  $da = $a.Documents.Open($File, $false, $false, $false)
  "first instance read-only: " + $da.ReadOnly
  $proc = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $secondPs1, $File, $pidFile, $outFile) -PassThru -WindowStyle Hidden
  if (-not $proc.WaitForExit($Wait * 1000)) {
    "second instance: no answer after $Wait s (Open() blocks, presumably on Word's 'file in use' prompt)"
    if (Test-Path $pidFile) { $wp = [int](Get-Content $pidFile); Stop-Process -Id $wp -Force -ErrorAction SilentlyContinue; "  stopped the second instance (pid $wp)" }
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
  }
  if (Test-Path $outFile) { Get-Content $outFile }
  $da.Close([ref]0)
} finally {
  if ($a) { $a.Quit([ref]0) }
}
