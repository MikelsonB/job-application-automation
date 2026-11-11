# Used by tools/joblist.mjs (don't run by hand).
#   -Mode read  -Xlsx <file>                        prints {"J0001":{"n":my_notes,"f":fit,"a":application},...}
#   -Mode write -Csv <file> -Xlsx <file> -Today D   builds the formatted Excel file from the CSV
param([string]$Mode, [string]$Csv, [string]$Xlsx, [string]$Today, [int]$SoonDays = 7)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
function ColOf($ws, $name) {
  $n = $ws.UsedRange.Columns.Count
  for ($c = 1; $c -le $n; $c++) { if ([string]$ws.Cells.Item(1, $c).Value2 -eq $name) { return $c } }
  return 0
}
try {
  if ($Mode -eq 'read') {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $kc = ColOf $ws 'key'; $nc = ColOf $ws 'my_notes'; $fc = ColOf $ws 'fit'; $ac = ColOf $ws 'application'
    $cells = @{}
    if ($kc) {
      $last = $ws.UsedRange.Rows.Count
      for ($r = 2; $r -le $last; $r++) {
        $k = [string]$ws.Cells.Item($r, $kc).Value2
        if (-not $k) { continue }
        $cells[$k] = @{
          n = $(if ($nc) { [string]$ws.Cells.Item($r, $nc).Value2 } else { '' })
          f = $(if ($fc) { $ws.Cells.Item($r, $fc).Value2 } else { $null })
          a = $(if ($ac) { [string]$ws.Cells.Item($r, $ac).Value2 } else { $null })
        }
      }
    }
    $wb.Close($false)
    if ($cells.Count -eq 0) { '{}' } else { $cells | ConvertTo-Json -Compress -Depth 3 }
    return
  }

  $wb = $xl.Workbooks.Open($Csv)
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'Jobs'
  $last = $ws.UsedRange.Rows.Count
  $ncol = $ws.UsedRange.Columns.Count
  $c = @{}
  foreach ($h in 'rank','fit','application','posting','deadline','date_posted','role','notes','my_notes','last_checked','link') { $c[$h] = ColOf $ws $h }

  foreach ($h in 'deadline','date_posted','last_checked') { if ($c[$h]) { $ws.Range($ws.Cells.Item(2, $c[$h]), $ws.Cells.Item([Math]::Max($last, 2), $c[$h])).NumberFormat = 'yyyy-mm-dd' } }
  $hdr = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $ncol))
  $hdr.Font.Bold = $true
  $hdr.Interior.Color = 0xE6D8C9
  $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($last, $ncol)).AutoFilter() | Out-Null
  $ws.Activate()
  $xl.ActiveWindow.SplitRow = 1
  $xl.ActiveWindow.FreezePanes = $true
  $ws.UsedRange.Columns.AutoFit() | Out-Null
  foreach ($h in 'deadline','date_posted','last_checked') { if ($c[$h]) { $ws.Columns.Item($c[$h]).ColumnWidth = 12 } }
  if ($c['role']) { $ws.Columns.Item($c['role']).ColumnWidth = 55 }
  if ($c['notes']) { $ws.Columns.Item($c['notes']).ColumnWidth = 35 }
  if ($c['my_notes']) { $ws.Columns.Item($c['my_notes']).ColumnWidth = 30; $ws.Columns.Item($c['my_notes']).Interior.Color = 0xF2F2F2 }

  $todayD = [datetime]::ParseExact($Today, 'yyyy-MM-dd', $null)
  $green = 0xCEEFC6; $orange = 0x9CEBFF; $grey = 0x808080; $lightgrey = 0xD9D9D9
  for ($r = 2; $r -le $last; $r++) {
    $app = [string]$ws.Cells.Item($r, $c['application']).Value2
    $post = [string]$ws.Cells.Item($r, $c['posting']).Value2
    if ($app -eq 'skipped') { $ws.Cells.Item($r, $c['application']).Interior.Color = $lightgrey }
    elseif ($app -and $app -ne 'not applied') { $ws.Cells.Item($r, $c['application']).Interior.Color = $green }
    if ($post -eq 'closed' -or $post -eq 'deadline passed') { $ws.Range($ws.Cells.Item($r, 1), $ws.Cells.Item($r, $ncol)).Font.Color = $grey }
    elseif ($post -eq 'open') { $ws.Cells.Item($r, $c['posting']).Font.Color = 0x006100 }
    $dl = $ws.Cells.Item($r, $c['deadline']).Value2
    if ($dl -is [double] -and $post -ne 'closed' -and $post -ne 'deadline passed') {
      $days = ([datetime]::FromOADate($dl) - $todayD).TotalDays
      if ($days -le $SoonDays) { $ws.Cells.Item($r, $c['deadline']).Interior.Color = $orange }
    }
    $lc = $ws.Cells.Item($r, $c['link'])
    $u = [string]$lc.Value2
    if ($u -like 'http*') { $ws.Hyperlinks.Add($lc, $u, '', '', $u) | Out-Null }
  }
  # editable columns: yellow header + dropdown on "application" (free text still allowed)
  foreach ($h in 'fit','application','my_notes') { if ($c[$h]) { $ws.Cells.Item(1, $c[$h]).Interior.Color = 0x99FFFF } }
  if ($c['application'] -and $last -ge 2) {
    $v = $ws.Range($ws.Cells.Item(2, $c['application']), $ws.Cells.Item($last, $c['application'])).Validation
    $v.Delete()
    $v.Add(3, 1, 1, 'not applied,applied,skipped,hide')
    $v.IgnoreBlank = $true
    $v.ShowError = $false
    $v.InputTitle = 'Application'
    $v.InputMessage = 'Pick applied, skipped (not interested), not applied, or hide (remove from the list). Saved at the next update.'
  }
  if ($c['fit'] -and $last -ge 2) {
    $v = $ws.Range($ws.Cells.Item(2, $c['fit']), $ws.Cells.Item($last, $c['fit'])).Validation
    $v.Delete()
    $v.Add(1, 1, 1, '0', '100')
    $v.ErrorMessage = 'Fit must be a number from 0 to 100.'
  }
  $ws.Columns.Item($c['link']).ColumnWidth = 14
  $ws.Range($ws.Cells.Item(2, $c['link']), $ws.Cells.Item([Math]::Max($last, 2), $c['link'])).HorizontalAlignment = 5
  $ws.Cells.Item(1, 1).Select() | Out-Null
  $wb.SaveAs($Xlsx, 51)
  $wb.Close($false)
  "Saved job_listing.xlsx ($($last - 1) jobs)"
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
