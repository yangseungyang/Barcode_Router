param([string]$html,[string]$out)
$wd = New-Object -ComObject Word.Application; $wd.Visible=$false; $wd.DisplayAlerts=0
try {
  $d = $wd.Documents.Open($html, $false, $false)
  $ps = $d.PageSetup; $ps.PaperSize = 7; $ps.TopMargin = 56.7; $ps.BottomMargin = 56.7; $ps.LeftMargin = 56.7; $ps.RightMargin = 56.7
  $d.ActiveWindow.View.Type = 3
  foreach($s in @($d.InlineShapes)){ try { $s.LinkFormat.SavePictureWithDocument = $true; $s.LinkFormat.BreakLink() } catch {} }
  foreach($f in @($d.Fields)){ try { if($f.Type -eq 67 -or $f.Type -eq 68){ $f.Unlink() } } catch {} }
  $ft = $d.Sections.Item(1).Footers.Item(1)
  $ft.PageNumbers.Add(1, $true) | Out-Null
  $hd = $d.Sections.Item(1).Headers.Item(1).Range
  $hd.Text = "Barcode Router 사용설명서  ·  Ver 0.00  ·  Rev 2"; $hd.Font.Size = 8; $hd.Font.Name = "Malgun Gothic"; $hd.ParagraphFormat.Alignment = 2
  foreach($t in @($d.Tables)){
    $t.Rows.AllowBreakAcrossPages = $true; $t.Rows.Item(1).HeadingFormat = $true
    $t.AllowAutoFit = $true; $t.AutoFitBehavior(1); $t.AutoFitBehavior(2)
    $t.PreferredWidthType = 2; $t.PreferredWidth = 100
  }
  # 목차 (클릭하면 해당 절로 이동)
  $f = $d.Content; $f.Find.ClearFormatting()
  if($f.Find.Execute('[[TOC]]')){ $f.Text = ''; $toc = $d.TablesOfContents.Add($f, $true, 2, 3, $false, [Type]::Missing, $true, $true, [Type]::Missing, $true) }
  $d.Repaginate(); foreach($t in @($d.TablesOfContents)){ $t.Update() }
  $d.SaveAs2($out, 16)
  "pages=" + $d.ComputeStatistics(2) + " shapes=" + $d.InlineShapes.Count
  $d.ExportAsFixedFormat(($out -replace '\.docx$','.pdf'), 17, $false, 0, 0, 1, 1, 0, $true, $true, 1, $true, $true, $false)
  $d.Close($false)
} finally { $wd.Quit() }
