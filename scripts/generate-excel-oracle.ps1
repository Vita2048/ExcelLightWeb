<#
.SYNOPSIS
  Captures Excel ground-truth formula results via Windows Excel COM automation.
.PARAMETER SpecPath
  Path to a JSON spec file with { id, cells, other?, expressions }.
.PARAMETER OutPath
  Path to write the captured JSON oracle fixture.
#>
param(
  [Parameter(Mandatory = $true)][string]$SpecPath,
  [Parameter(Mandatory = $true)][string]$OutPath
)

$ErrorActionPreference = 'Stop'

$resolvedSpec = (Resolve-Path -LiteralPath $SpecPath).Path
$spec = Get-Content -LiteralPath $resolvedSpec -Raw -Encoding UTF8 | ConvertFrom-Json

$excelErrorCodes = @{
  (-2146826288) = '#NULL!'
  (-2146826281) = '#DIV/0!'
  (-2146826273) = '#VALUE!'
  (-2146826265) = '#REF!'
  (-2146826259) = '#NAME?'
  (-2146826252) = '#NUM!'
  (-2146826246) = '#N/A'
  (-2146826243) = '#SPILL!'
  (-2146826239) = '#CALC!'
}

function Set-InputCells($worksheet, $entries) {
  if (-not $entries) { return }
  if ($entries -is [System.Collections.IList] -and $entries.Count -gt 0 -and -not ($entries[0] -is [System.Collections.IList])) {
    $entries = ,$entries
  }
  foreach ($entry in $entries) {
    $row = [int]$entry[0]
    $col = [int]$entry[1]
    $val = $entry[2]
    $cell = $worksheet.Cells.Item($row, $col)
    if ($null -eq $val) {
      $cell.ClearContents() | Out-Null
    } elseif ($val -is [string]) {
      if ($val -eq '') {
        $cell.Formula = '=""'
      } elseif ($val.StartsWith('=')) {
        try {
          $cell.Formula2 = $val
        } catch {
          $cell.Formula = $val
        }
      } else {
        $cell.NumberFormat = '@'
        $cell.Value2 = $val
      }
    } elseif ($val -is [bool]) {
      $cell.Value2 = [bool]$val
    } else {
      $cell.Value2 = [double]$val
    }
  }
}

function Read-ScalarResult($cell) {
  $val = $cell.Value2
  if ($val -is [int]) {
    if ($excelErrorCodes.ContainsKey($val)) {
      return @{ kind = 'error'; value = $excelErrorCodes[$val] }
    }
    $text = [string]$cell.Text
    return @{ kind = 'error'; value = $text.Trim().ToUpperInvariant() }
  }
  if ($null -eq $val) {
    return @{ kind = 'blank'; value = '' }
  }
  if ($val -is [bool]) {
    return @{ kind = 'boolean'; value = $(if ($val) { 'true' } else { 'false' }) }
  }
  if ($val -is [string]) {
    return @{ kind = 'text'; value = $val }
  }
  $num = [double]$val
  return @{ kind = 'number'; value = $num.ToString('G17', [System.Globalization.CultureInfo]::InvariantCulture) }
}

$excel = $null
$workbook = $null

$stage = 'init'
try {
  $stage = 'create Excel.Application'
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false

  $stage = 'create workbook'
  $workbook = $excel.Workbooks.Add()
  $workbook.Date1904 = $false

  $stage = 'setup Sheet1'
  $sheet1 = $workbook.Worksheets.Item(1)
  $sheet1.Name = 'Sheet1'
  $sheet1.Range('T:AZ').ColumnWidth = 30

  if ($spec.other -and $spec.other.Count -gt 0) {
    $stage = 'setup other sheet'
    $sheet2 = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $sheet1)
    $sheet2.Name = "O'Brien data"
    Set-InputCells $sheet2 $spec.other
  }

  $stage = 'set input cells'
  Set-InputCells $sheet1 $spec.cells

  if ($spec.pivots) {
    $pivotFunctions = @{
      'sum'          = -4157
      'count'        = -4112
      'average'      = -4106
      'max'          = -4136
      'min'          = -4139
      'product'      = -4149
      'countNumbers' = -4113
    }
    foreach ($p in $spec.pivots) {
      $stage = "create PivotTable $($p.name)"
      $cache = $workbook.PivotCaches().Create(1, $sheet1.Range([string]$p.source))
      $pt = $cache.CreatePivotTable($sheet1.Range([string]$p.target), [string]$p.name)
      $pt.ManualUpdate = $true
      $pos = 1
      if ($p.rows) {
        foreach ($f in $p.rows) {
          $pf = $pt.PivotFields([string]$f)
          $pf.Orientation = 1
          $pf.Position = $pos++
        }
      }
      $pos = 1
      if ($p.columns) {
        foreach ($f in $p.columns) {
          $pf = $pt.PivotFields([string]$f)
          $pf.Orientation = 2
          $pf.Position = $pos++
        }
      }
      $pos = 1
      if ($p.filters) {
        foreach ($f in $p.filters) {
          $pf = $pt.PivotFields([string]$f)
          $pf.Orientation = 3
          $pf.Position = $pos++
        }
      }
      if ($p.values) {
        foreach ($v in $p.values) {
          $df = $pt.AddDataField($pt.PivotFields([string]$v.field))
          if ($v.aggregation -and $pivotFunctions.ContainsKey([string]$v.aggregation)) {
            $df.Function = $pivotFunctions[[string]$v.aggregation]
          }
          if ($v.customName -and ([string]$df.Caption -ne [string]$v.customName)) {
            $df.Caption = [string]$v.customName
          }
        }
      }
      if ($p.grandTotals) {
        $pt.ColumnGrand = [bool]$p.grandTotals.rows
        $pt.RowGrand = [bool]$p.grandTotals.columns
      }
      $pt.ManualUpdate = $false
      $pt.RefreshTable() | Out-Null

      if ($p.hiddenItems) {
        foreach ($rule in $p.hiddenItems) {
          $pf = $pt.PivotFields([string]$rule.field)
          $hiddenSet = @{}
          foreach ($itemName in $rule.items) { $hiddenSet[[string]$itemName] = $true }
          foreach ($pi in $pf.PivotItems()) {
            if ($hiddenSet.ContainsKey([string]$pi.Name)) {
              $pi.Visible = $false
            }
          }
        }
      }
      if ($p.collapsed) {
        foreach ($colRule in $p.collapsed) {
          $pf = $pt.PivotFields([string]$colRule.field)
          $pf.PivotItems([string]$colRule.item).ShowDetail = $false
        }
      }
      $pt.RefreshTable() | Out-Null
    }
  }

  $cases = @()
  foreach ($entry in $spec.expressions) {
    $expr = if ($entry -is [string]) { $entry } else { [string]$entry.expression }
    $setup = if ($entry -isnot [string] -and $entry.setup) { $entry.setup } else { $null }
    $stage = "clear range before $expr"
    $sheet1.Range('T1:AZ50').Clear() | Out-Null
    if ($setup) {
      $stage = "setup cells for $expr"
      Set-InputCells $sheet1 $setup
    }
    $stage = "assign $expr"
    $target = $sheet1.Cells.Item(1, 20)
    try {
      $target.Formula2 = $expr
    } catch {
      try {
        $target.Formula = $expr
      } catch {
        throw "Failed to assign formula '$expr': $($_.Exception.Message)"
      }
    }
    $stage = "recalc $expr"
    $excel.CalculateFullRebuild()

    $stage = "read spill/result $expr"
    $spill = $null
    try { $spill = $target.SpillingToRange } catch { $spill = $null }

    if ($null -ne $spill -and ($spill.Rows.Count -gt 1 -or $spill.Columns.Count -gt 1)) {
      $rowCount = [int]$spill.Rows.Count
      $colCount = [int]$spill.Columns.Count
      $matrix = @()
      for ($r = 1; $r -le $rowCount; $r++) {
        $rowItems = @()
        for ($c = 1; $c -le $colCount; $c++) {
          $rowItems += ,(Read-ScalarResult $spill.Cells.Item($r, $c))
        }
        $matrix += ,$rowItems
      }
      $cases += [ordered]@{
        expression = [string]$expr
        kind       = 'array'
        rows       = $rowCount
        columns    = $colCount
        matrix     = $matrix
      }
    } else {
      $scalar = Read-ScalarResult $target
      $cases += [ordered]@{
        expression = [string]$expr
        kind       = $scalar.kind
        value      = $scalar.value
      }
    }
  }

  $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
  $relativeSpec = $resolvedSpec
  if ($resolvedSpec.StartsWith($repoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    $relativeSpec = $resolvedSpec.Substring($repoRoot.Length).TrimStart('\', '/').Replace('\', '/')
  }

function To-JsonRows($entries) {
  $list = [System.Collections.ArrayList]::new()
  if (-not $entries) { return ,$list }
  if ($entries -is [System.Collections.IList] -and $entries.Count -gt 0 -and -not ($entries[0] -is [System.Collections.IList])) {
    [void]$list.Add($entries)
    return ,$list
  }
  foreach ($row in $entries) {
    [void]$list.Add($row)
  }
  return ,$list
}

  $cellsOut = To-JsonRows $spec.cells
  $otherOut = To-JsonRows $spec.other
  $output = [ordered]@{
    id           = [string]$spec.id
    sourceSpec   = $relativeSpec
    excelVersion = [string]$excel.Version
    excelBuild   = [string]$excel.Build
    capturedAt   = (Get-Date).ToString('o')
    cells        = $cellsOut
    other        = $otherOut
    cases        = $cases
  }
  if ($spec.pivots) {
    $output['pivots'] = $spec.pivots
  }

  $outDir = Split-Path -Parent $OutPath
  if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $json = $output | ConvertTo-Json -Depth 10
  [System.IO.File]::WriteAllText($OutPath, $json + "`n", [System.Text.UTF8Encoding]::new($false))
  Write-Host "Captured $($cases.Count) Excel oracle cases (Excel $($excel.Version) build $($excel.Build)) -> $OutPath"
} catch {
  throw "Oracle failed at stage [$stage]: $($_.Exception.Message)"
} finally {
  if ($null -ne $workbook) {
    try { $workbook.Close($false) } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
  }
  if ($null -ne $excel) {
    try { $excel.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
