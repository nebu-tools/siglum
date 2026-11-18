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
$dir = Split-Path $File
$pidFile = Join-Path $dir "second.pid"
$outFile = Join-Path $dir "second.txt"
$second = @'
param($File, $PidFile, $OutFile)
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$b = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
if ($new.Count -ne 1) { "second instance: Word did not start one process of its own" | Add-Content $OutFile; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($b); exit }
Set-Content $PidFile $new[0]
$b.DisplayAlerts = 0
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
$mine = $null
try {
  $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
  $a = New-Object -ComObject Word.Application
  $new = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
  if ($new.Count -ne 1) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($a); $a = $null; throw "Word did not start one process of its own." }
  $mine = $new[0]
  $a.DisplayAlerts = 0
  $da = $a.Documents.Open($File, $false, $false, $false)
  "first instance read-only: " + $da.ReadOnly
  $proc = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $secondPs1, $File, $pidFile, $outFile) -PassThru -WindowStyle Hidden
  if (-not $proc.WaitForExit($Wait * 1000)) {
    "second instance: no answer after $Wait s (Open() waits on Word's 'file in use' prompt)"
    if (Test-Path $pidFile) { $wp = [int](Get-Content $pidFile); Stop-Process -Id $wp -Force -ErrorAction SilentlyContinue; "  stopped the second instance (pid $wp)" }
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
  }
  if (Test-Path $outFile) { Get-Content $outFile }
  $da.Close([ref]0)
} finally {
  if ($a) {
    try { $a.Quit([ref]0) } catch { }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)
    $p = Get-Process -Id $mine -ErrorAction SilentlyContinue
    if ($p -and -not $p.WaitForExit(30000)) { Stop-Process -Id $mine -Force -ErrorAction SilentlyContinue }
  }
}
