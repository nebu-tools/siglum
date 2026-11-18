# Opens a docx in a new Word instance with the add-in sideloaded, waits until the
# harness page reports done (a file named "done" appears), saves and quits.
# Writes info.json next to the document: Word's version and build, how long
# opening and saving took, the instance's peak memory, the times Word's UI was
# frozen for a second or more ("stalls": [seconds after open, ms]) and, with
# -Pages, the page count.
#
# It only touches the Word instance it starts and the registry value it adds,
# which it removes again. Other Word windows on the machine are left alone.
param(
  [Parameter(Mandatory)][string]$Dir,
  [Parameter(Mandatory)][string]$AddinId,
  [int]$Timeout = 300,
  [switch]$Pages
)
$ErrorActionPreference = 'Stop'
Add-Type -Namespace Siglum -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
New-ItemProperty -Path $key -Name $AddinId -Value "$Dir\manifest.xml" -PropertyType String -Force | Out-Null
$word = $null
$info = [ordered]@{}
$proc = $null
function Sample { if ($proc) { $proc.Refresh(); $info.peakWorkingSetMB = [math]::Round($proc.PeakWorkingSet64 / 1MB); $info.privateMB = [math]::Round($proc.PrivateMemorySize64 / 1MB) } }
try {
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $word = New-Object -ComObject Word.Application
  $info.startSeconds = $sw.Elapsed.TotalSeconds
  $info.version = $word.Version
  $info.build = $word.Build
  $word.Visible = $true
  $word.DisplayAlerts = 0
  $sw.Restart()
  $doc = $word.Documents.Open("$Dir\in.docx", $false, $false, $false)
  $info.openSeconds = $sw.Elapsed.TotalSeconds
  $wpid = 0
  [void][Siglum.Win]::GetWindowThreadProcessId([IntPtr]$doc.ActiveWindow.Hwnd, [ref]$wpid)
  $info.pid = $wpid
  $proc = Get-Process -Id $wpid
  $sw.Restart()
  # While the pane works, a trivial COM call every half second shows whether
  # Word's UI thread is free: a slow answer means Word was frozen that long.
  $stalls = New-Object System.Collections.ArrayList
  $probe = [Diagnostics.Stopwatch]::new()
  while (-not (Test-Path "$Dir\done") -and $sw.Elapsed.TotalSeconds -lt $Timeout) {
    Start-Sleep -Milliseconds 500
    $probe.Restart()
    try { $null = $doc.Saved } catch { [void]$stalls.Add(@([math]::Round($sw.Elapsed.TotalSeconds, 1), "rejected: " + $_.Exception.Message)) }
    if ($probe.ElapsedMilliseconds -ge 1000) { [void]$stalls.Add(@([math]::Round($sw.Elapsed.TotalSeconds - $probe.Elapsed.TotalSeconds, 1), $probe.ElapsedMilliseconds)) }
  }
  $info.stalls = $stalls
  $info.paneSeconds = $sw.Elapsed.TotalSeconds
  Sample
  if (-not (Test-Path "$Dir\done")) { Write-Error "Timeout: the add-in did not report back." }
  if ($Pages) { $sw.Restart(); $info.pages = $doc.ComputeStatistics(2); $info.pagesSeconds = $sw.Elapsed.TotalSeconds }
  $sw.Restart()
  $doc.SaveAs2([ref]"$Dir\out.docx", [ref]16)
  $info.saveSeconds = $sw.Elapsed.TotalSeconds
  Sample
  $doc.Close([ref]0)
} finally {
  ($info | ConvertTo-Json) | Set-Content -Encoding UTF8 "$Dir\info.json"
  if ($word) {
    $word.Quit([ref]0)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  }
  Remove-ItemProperty -Path $key -Name $AddinId -ErrorAction SilentlyContinue
}
