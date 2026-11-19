# Opens a docx in a Word instance of its own with the add-in sideloaded, runs
# the commands the harness drops into $Dir, then saves to out.docx and quits.
#
# Commands are files named cmd-<n>.ps1. Each is a script that may use $word,
# $doc and $Dir; its output goes to ack-<n>.json. A file named "done" ends the
# session with a save, "abort" ends it without one. The harness touches a file
# named "alive"; when it is a minute old, the session ends without a save.
#
# Writes info.json: Word's version and build, the instance's process id, how
# long opening and saving took, the peak memory, the times Word's UI was
# frozen for a second or more ("stalls": [seconds after open, ms]) and, with
# -Pages, the page count.
#
# It touches only the Word process it starts, found by its process id from a
# window of the application it created, and the one registry value it sets.
# Any pre-existing registry value of the same name is restored with its
# original type and content at the end; otherwise the value is removed, along
# with the entry Word makes under "Reading Locations" for the document. If Word
# hands back an instance that was already running, it stops without touching it.
param(
  [Parameter(Mandatory)][string]$Dir,
  [Parameter(Mandatory)][string]$AddinId,
  [int]$Timeout = 300,
  [switch]$Pages
)
$ErrorActionPreference = 'Stop'
Add-Type -Namespace SiglumRig -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$keyPath = 'Software\Microsoft\Office\16.0\WEF\Developer'
$key = "HKCU:\$keyPath"
$madeKey = $false
$hadPrior = $false
$prior = $null
$priorKind = $null
$wroteReg = $false
$word = $null
$doc = $null
$blank = $null
$mine = $null
$end = $null
$info = [ordered]@{}
$proc = $null
function Sample { if ($proc) { $proc.Refresh(); $info.peakWorkingSetMB = [math]::Round($proc.PeakWorkingSet64 / 1MB); $info.privateMB = [math]::Round($proc.PrivateMemorySize64 / 1MB) } }
try {
  $madeKey = -not (Test-Path $key)
  if ($madeKey) {
    New-Item -Path $key -Force | Out-Null
  } else {
    $rk = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($keyPath, $false)
    if ($rk) {
      try {
        if ($rk.GetValueNames() -contains $AddinId) {
          $hadPrior = $true
          $prior = $rk.GetValue($AddinId, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
          $priorKind = $rk.GetValueKind($AddinId)
        }
      } finally {
        $rk.Close()
      }
    }
  }
  New-ItemProperty -Path $key -Name $AddinId -Value "$Dir\manifest.xml" -PropertyType String -Force | Out-Null
  $wroteReg = $true

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
  # If this COM object already has a window, verify its process before touching it.
  if ($word.Windows.Count -gt 0) {
    $wpid = 0
    [void][SiglumRig.Win]::GetWindowThreadProcessId([IntPtr]$word.Windows.Item(1).Hwnd, [ref]$wpid)
    if ($new -notcontains [int]$wpid) {
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
      $word = $null
      throw "Word reused an existing process ($wpid)."
    }
  }
  # Identify the Word process from a blank document's window before opening
  # in.docx, so a failed Open() still knows which Word process to close.
  $blank = $word.Documents.Add()
  $wpid = 0
  $hwnd = [IntPtr]::Zero
  try { $hwnd = [IntPtr]$blank.ActiveWindow.Hwnd } catch { }
  if ($hwnd -eq [IntPtr]::Zero) {
    $word.Visible = $true
    $hwnd = [IntPtr]$blank.ActiveWindow.Hwnd
  }
  [void][SiglumRig.Win]::GetWindowThreadProcessId($hwnd, [ref]$wpid)
  if ($wpid -eq 0 -or $new -notcontains [int]$wpid) {
    throw "The application's window is in Word process $wpid, which this script did not start."
  }
  $mine = [int]$wpid
  $info.pid = $mine
  Set-Content -Path "$Dir\pid" -Value $mine
  $proc = Get-Process -Id $mine
  # Close before Open(): Word drops a pristine blank document when another opens.
  $blank.Close([ref]0)
  $blank = $null

  $info.version = $word.Version
  $info.build = $word.Build
  $word.Visible = $true
  $word.DisplayAlerts = 0
  $sw.Restart()
  $doc = $word.Documents.Open("$Dir\in.docx", $false, $false, $false)
  $info.openSeconds = $sw.Elapsed.TotalSeconds
  $dpid = 0
  [void][SiglumRig.Win]::GetWindowThreadProcessId([IntPtr]$doc.ActiveWindow.Hwnd, [ref]$dpid)
  if ([int]$dpid -ne $mine) { throw "The document opened in Word process $dpid, not in $mine." }
  $sw.Restart()
  # A trivial COM call every half second shows whether Word's UI thread is
  # free: a slow answer means Word was frozen that long.
  $stalls = New-Object System.Collections.ArrayList
  $probe = [Diagnostics.Stopwatch]::new()
  while ($sw.Elapsed.TotalSeconds -lt $Timeout) {
    if (Test-Path "$Dir\done") { $end = 'done'; break }
    if (Test-Path "$Dir\abort") { $end = 'abort'; break }
    # The harness touches "alive" every few seconds; without it, nobody will end the session.
    if ((Test-Path "$Dir\alive") -and ((Get-Date) - (Get-Item "$Dir\alive").LastWriteTime).TotalSeconds -gt 60) { $end = 'orphaned'; break }
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
  if ($end -eq 'orphaned') { Write-Error "The harness stopped; closing without saving." }
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
  try { ($info | ConvertTo-Json -Depth 4) | Set-Content -Encoding UTF8 "$Dir\info.json" } catch { }
  try {
    if ($word -and $mine) {
      if ($doc) { try { $doc.Close([ref]0) } catch { } }
      if ($blank) { try { $blank.Close([ref]0) } catch { } }
      try { $word.Quit([ref]0) } catch { }
      try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
      $p = Get-Process -Id $mine -ErrorAction SilentlyContinue
      if ($p -and -not $p.WaitForExit(30000)) { Stop-Process -Id $mine -Force -ErrorAction SilentlyContinue }
    } elseif ($word) {
      # The instance is not known to be ours: close only documents this script opened.
      if ($doc) { try { $doc.Close([ref]0) } catch { } }
      if ($blank) { try { $blank.Close([ref]0) } catch { } }
      try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    }
  } catch { }
  try {
    if ($wroteReg) {
      if ($hadPrior) {
        $rk = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($keyPath, $true)
        if ($rk) {
          try { $rk.SetValue($AddinId, $prior, $priorKind) } finally { $rk.Close() }
        }
      } else {
        Remove-ItemProperty -Path $key -Name $AddinId -ErrorAction SilentlyContinue
      }
    }
  } catch { }
  try {
    if ($madeKey) {
      $rk = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($keyPath, $false)
      if ($rk) {
        $empty = ($rk.ValueCount -eq 0) -and ($rk.SubKeyCount -eq 0)
        $rk.Close()
        if ($empty) { Remove-Item $key -ErrorAction SilentlyContinue }
      }
    }
  } catch { }
  # Word remembers where a reader was in a closed document; forget this one.
  try {
    Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\Word\Reading Locations' -ErrorAction SilentlyContinue |
      Where-Object { "$($_.GetValue('File Path'))".StartsWith($Dir) } | Remove-Item -ErrorAction SilentlyContinue
  } catch { }
  # Nobody is left to clean up after an orphaned session, so it does it itself.
  try {
    if ($end -eq 'orphaned') { Set-Location $env:TEMP; Remove-Item -Recurse -Force $Dir -ErrorAction SilentlyContinue }
  } catch { }
}
