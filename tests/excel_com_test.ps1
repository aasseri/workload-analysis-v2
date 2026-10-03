# excel_com_test.ps1 - opens the generated workbook in real Microsoft Excel,
# scans for formula errors, verifies named results, and runs the mandatory recalculation test.
# Usage: powershell -ExecutionPolicy Bypass -File excel_com_test.ps1 <xlsx path> <json output path>
param([string]$Path, [string]$Out)

$ErrorActionPreference = 'Stop'
$r = [ordered]@{ file = $Path }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($Path)
  $r.sheetNames = @($wb.Worksheets | ForEach-Object { $_.Name })
  $r.names = @($wb.Names | ForEach-Object { $_.Name + ' => ' + $_.RefersTo })
  $r.tables = @($wb.Worksheets | ForEach-Object { $ws = $_; $ws.ListObjects | ForEach-Object { $ws.Index.ToString() + ':' + $_.Name + ' ' + $_.Range.Address() } })
  $xl.CalculateFull()

  function Scan-Errors($wb) {
    $found = @()
    foreach ($ws in $wb.Worksheets) {
      try {
        $cells = $ws.UsedRange.SpecialCells(-4123, 16)   # xlCellTypeFormulas, xlErrors
        foreach ($c in $cells) { $found += ($ws.Index.ToString() + '!' + $c.Address() + ' ' + $c.Text) }
      } catch { }
    }
    return $found
  }
  function Snap($wb) {
    $s = [ordered]@{}
    foreach ($n in 'TotalHours','TaskCount','AnnualHours','ExactNeed','CalcNeed','ActualCount','Gap','Status') {
      $s[$n] = $wb.Names.Item($n).RefersToRange.Value2
    }
    return $s
  }

  # Positions sheet (need by actual job title): non-empty rows as "title|count|hours|exact|need", plus totals/checks text
  function PosSnap($wb) {
    $ps = $wb.Worksheets.Item(6)
    $out = @()
    for ($i = 5; $i -le 45; $i++) {
      $a = $ps.Range("A$i").Text; $b = $ps.Range("B$i").Text
      if ($a -or $b) { $out += ($a + ' | ' + $b + ' | ' + $ps.Range("C$i").Text + ' | ' + $ps.Range("D$i").Text + ' | ' + $ps.Range("E$i").Text + ' | ' + $ps.Range("F$i").Text + ' | ' + $ps.Range("G$i").Text) }
    }
    return $out
  }
  $r.sheet6Name = $wb.Worksheets.Item(6).Name
  $r.positionsInitial = @(PosSnap $wb)
  $r.errorsInitial = @(Scan-Errors $wb)
  $r.initial = Snap $wb
  $r.roundCheck_1_5 = $xl.Evaluate('ROUND(1.5,0)')
  $r.roundCheck_1_3 = $xl.Evaluate('ROUND(1.3,0)')
  $r.roundCheck_float = $xl.Evaluate('ROUND(3.3-1.8,0)')   # computed 1.4999999999999998 -> 2 (same as JS engine)

  $tasks = $wb.Worksheets.Item(3)
  $f = $wb.Names.Item('TaskTitles').RefersToRange.Row
  $r.firstTaskRow = $f
  $r.protected = @($wb.Worksheets | ForEach-Object { $_.ProtectContents })
  $r.formulaHours = $tasks.Range("I$f").Formula
  $r.formulaPerYear = $tasks.Range("F$f").Formula
  $r.formulaShare = $tasks.Range("J$f").Formula
  $r.posFormulas = @($wb.Worksheets.Item(6).Range('D5').Formula, $wb.Worksheets.Item(6).Range('E5').Formula, $wb.Worksheets.Item(6).Range('G5').Formula)
  $r.summaryInfo = @(4..6 | ForEach-Object { $wb.Worksheets.Item(1).Range("C$_").Text })
  $r.sectionNames = @($wb.Names.Item('SectionNames').RefersToRange | ForEach-Object { $_.Value2 })
  $r.needFormulas = @(5..11 | ForEach-Object { $wb.Worksheets.Item(5).Range("C$_").Formula })
  $r.summaryNarrative = $wb.Worksheets.Item(1).Range('A15').Text

  # Locked formula cell must reject edits
  try { $tasks.Range("I$f").Value2 = 1; $r.lockedCellWritable = $true } catch { $r.lockedCellWritable = $false }

  # Recalculation test 1: change repetitions of task 1 (E5)
  $r.E5_before = $tasks.Range("G$f").Value2
  $tasks.Range("G$f").Value2 = 50
  $xl.CalculateFull()
  $r.afterE5_50 = Snap $wb
  $tasks.Range("G$f").Value2 = 70
  $xl.CalculateFull()
  $r.afterE5_70 = Snap $wb

  # Recalculation test 2: add a new task into the first spare row
  $spare = $f + [int]$r.initial.TaskCount
  $firstFreq = $wb.Names.Item('FreqNames').RefersToRange.Cells.Item(1,1).Value2
  $tasks.Range("B$spare").Value2 = 'new task'
  $tasks.Range("C$spare").Value2 = $wb.Names.Item('SectionNames').RefersToRange.Cells.Item(1,1).Value2
  $tasks.Range("D$spare").Value2 = 'new position'
  $tasks.Range("E$spare").Value2 = $firstFreq
  $tasks.Range("G$spare").Value2 = 1
  $tasks.Range("H$spare").Value2 = 60
  $xl.CalculateFull()
  $r.afterNewTask = Snap $wb
  $r.positionsAfterNewTask = @(PosSnap $wb)
  $r.newTaskRow = @($tasks.Range("A$spare").Value2, $tasks.Range("F$spare").Value2, $tasks.Range("I$spare").Value2, $tasks.Range("J$spare").Value2)
  # fill a spare row in the section x title sheet with the new pair -> unlisted hours back to 0
  $ps6 = $wb.Worksheets.Item(6)
  $pr = 5; while ($ps6.Range("B$pr").Text -ne '' -or $ps6.Range("A$pr").Text -ne '') { $pr++ }
  $ps6.Range("B$pr").Value2 = $tasks.Range("C$spare").Value2
  $ps6.Range("C$pr").Value2 = 'new position'
  $xl.CalculateFull()
  $r.positionsAfterSpareRow = @(PosSnap $wb)

  # Change annual hours in settings
  $wb.Names.Item('AnnualHours').RefersToRange.Value2 = 2000
  $xl.CalculateFull()
  $r.afterAnnual2000 = Snap $wb

  $r.errorsAfter = @(Scan-Errors $wb)
  $r.printAreaName = @($wb.Names | Where-Object { $_.Name -like '*Print_Area' } | ForEach-Object { $_.Name + ' => ' + $_.RefersTo })

  $ps = $tasks.PageSetup
  $r.print = [ordered]@{
    tasksOrientation = $ps.Orientation; tasksPaper = $ps.PaperSize; tasksTitleRows = $ps.PrintTitleRows
    tasksArea = $ps.PrintArea; tasksFitWide = $ps.FitToPagesWide
    summaryOrientation = $wb.Worksheets.Item(1).PageSetup.Orientation
    summaryPaper = $wb.Worksheets.Item(1).PageSetup.PaperSize
  }
  $r.rtl = @($wb.Worksheets | ForEach-Object { $_.Activate(); $xl.ActiveWindow.DisplayRightToLeft })
  $r.freeze = @($wb.Worksheets | ForEach-Object { $_.Activate(); $xl.ActiveWindow.FreezePanes })
  $r.validationFreq = $tasks.Range("E$f").Validation.Formula1
  $r.validationSection = $tasks.Range("C$f").Validation.Formula1
  $r.cfCountStatus = $wb.Worksheets.Item(5).Range('C11').FormatConditions.Count
  $wb.Close($false)
  $r.ok = $true
} catch {
  $r.ok = $false
  $r.error = $_.Exception.Message
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
$r | ConvertTo-Json -Depth 5 | Out-File -FilePath $Out -Encoding utf8
