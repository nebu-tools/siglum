# Opens a docx in a Word instance of its own with the add-in sideloaded, runs
# the commands the harness drops into $Dir, then saves to out.docx and quits.
#
# Commands are files named cmd-<n>.ps1. Each is a script that may use $word,
# $doc and $Dir; its output goes to ack-<n>.json. A file named "done" ends the
# session with a save, "abort" ends it without one.
#
# Writes info.json: Word's version and build, the instance's process id, how
# long opening and saving took, the peak memory, the times Word's UI was
# frozen for a second or more ("stalls": [seconds after open, ms]) and, with
# -Pages, the page count.
#
# It touches only the Word process it starts, found by its process id, and
# the one registry value it adds, which it removes again, together with the
# entry Word makes under "Reading Locations" for the document. If Word hands back
# an instance that was already running, it stops without touching it.
param(
  [Parameter(Mandatory)][string]$Dir,
  [Parameter(Mandatory)][string]$AddinId,
  [int]$Timeout = 300,
  [switch]$Pages
)
$ErrorActionPreference = 'Stop'
Add-Type -Namespace SiglumRig -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$madeKey = -not (Test-Path $key)
if ($madeKey) { New-Item -Path $key -Force | Out-Null }
New-ItemProperty -Path $key -Name $AddinId -Value "$Dir\manifest.xml" -PropertyType String -Force | Out-Null
$word = $null
$doc = $null
$mine = $null
$info = [ordered]@{}
$proc = $null
function Sample { if ($proc) { $proc.Refresh(); $info.peakWorkingSetMB = [math]::Round($proc.PeakWorkingSet64 / 1MB); $info.privateMB = [math]::Round($proc.PrivateMemorySize64 / 1MB) } }
try {
  $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $word = New-Object -ComObject Word.Application
  $info.startSeconds = $sw.Elapsed.TotalSeconds
  $new = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
  if ($new.Count -eq 0) {
    # Word reused an instance that someone else started: leave it alone.
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    $word = $null
    throw "Word did not start an instance of its own."
  }
  $info.version = $word.Version
  $info.build = $word.Build
  $word.Visible = $true
  $word.DisplayAlerts = 0
  $sw.Restart()
  $doc = $word.Documents.Open("$Dir\in.docx", $false, $false, $false)
  $info.openSeconds = $sw.Elapsed.TotalSeconds
  $wpid = 0
  [void][SiglumRig.Win]::GetWindowThreadProcessId([IntPtr]$doc.ActiveWindow.Hwnd, [ref]$wpid)
  if ($new -notcontains [int]$wpid) { throw "The document opened in Word process $wpid, which this script did not start." }
  $mine = [int]$wpid
  $info.pid = $mine
  Set-Content -Path "$Dir\pid" -Value $mine
  $proc = Get-Process -Id $mine
  $sw.Restart()
  # A trivial COM call every half second shows whether Word's UI thread is
  # free: a slow answer means Word was frozen that long.
  $stalls = New-Object System.Collections.ArrayList
  $probe = [Diagnostics.Stopwatch]::new()
  $end = $null
  while ($sw.Elapsed.TotalSeconds -lt $Timeout) {
    if (Test-Path "$Dir\done") { $end = 'done'; break }
    if (Test-Path "$Dir\abort") { $end = 'abort'; break }
    foreach ($cmd in @(Get-ChildItem -Path $Dir -Filter 'cmd-*.ps1' | Sort-Object Name)) {
      $n = $cmd.BaseName.Substring(4)
      $ack = [ordered]@{}
      try { $ack.value = (& ([scriptblock]::Create((Get-Content -Raw -Encoding UTF8 $cmd.FullName)))) } catch { $ack.error = $_.Exception.Message }
      Remove-Item $cmd.FullName
      ($ack | ConvertTo-Json -Depth 6 -Compress) | Set-Content -Encoding UTF8 "$Dir\ack-$n.tmp"
      Move-Item "$Dir\ack-$n.tmp" "$Dir\ack-$n.json"
    }
    Start-Sleep -Milliseconds 250
    $probe.Restart()
    try { $null = $doc.Saved } catch { [void]$stalls.Add(@([math]::Round($sw.Elapsed.TotalSeconds, 1), "rejected: " + $_.Exception.Message)) }
    if ($probe.ElapsedMilliseconds -ge 1000) { [void]$stalls.Add(@([math]::Round($sw.Elapsed.TotalSeconds - $probe.Elapsed.TotalSeconds, 1), $probe.ElapsedMilliseconds)) }
  }
  $info.stalls = $stalls
  $info.paneSeconds = $sw.Elapsed.TotalSeconds
  Sample
  if (-not $end) { Write-Error "Timeout: the add-in did not report back." }
  if ($end -eq 'done') {
    if ($Pages) { $sw.Restart(); $info.pages = $doc.ComputeStatistics(2); $info.pagesSeconds = $sw.Elapsed.TotalSeconds }
    $sw.Restart()
    # Not in the list of recent files.
    $doc.SaveAs2([ref]"$Dir\out.docx", [ref]16, [ref]$false, [ref]"", [ref]$false)
    $info.saveSeconds = $sw.Elapsed.TotalSeconds
    Sample
  }
  $doc.Close([ref]0)
  $doc = $null
} finally {
  ($info | ConvertTo-Json -Depth 4) | Set-Content -Encoding UTF8 "$Dir\info.json"
  if ($word -and $mine) {
    try { $word.Quit([ref]0) } catch { }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    $p = Get-Process -Id $mine -ErrorAction SilentlyContinue
    if ($p -and -not $p.WaitForExit(30000)) { Stop-Process -Id $mine -Force -ErrorAction SilentlyContinue }
  } elseif ($word) {
    # The instance is not known to be ours: close only the document this script opened.
    if ($doc) { try { $doc.Close([ref]0) } catch { } }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  }
  Remove-ItemProperty -Path $key -Name $AddinId -ErrorAction SilentlyContinue
  # Word remembers where a reader was in a closed document; forget this one.
  Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\Word\Reading Locations' -ErrorAction SilentlyContinue |
    Where-Object { "$($_.GetValue('File Path'))".StartsWith($Dir) } | Remove-Item -ErrorAction SilentlyContinue
  if ($madeKey -and -not (Get-ItemProperty $key).PSObject.Properties.Where({ $_.Name -notlike 'PS*' })) { Remove-Item $key -ErrorAction SilentlyContinue }
}
