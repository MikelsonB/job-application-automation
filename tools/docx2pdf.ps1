# Converts .docx files to PDF with Microsoft Word, embeds fonts in the .docx,
# and reports the page count of each document.
# Usage: powershell -ExecutionPolicy Bypass -File tools\docx2pdf.ps1 file1.docx [file2.docx ...]
param([Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)][string[]]$Docx)

$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  foreach ($p in $Docx) {
    $full = (Resolve-Path $p).Path
    $pdf = [IO.Path]::ChangeExtension($full, '.pdf')
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
    $doc = $word.Documents.Open($full, $false, $false, $false)
    try {
      # Embed Raleway so the .docx looks right on a recruiter's computer too.
      $doc.EmbedTrueTypeFonts = $true
      $doc.SaveSubsetFonts = $true
      $doc.DoNotEmbedSystemFonts = $true
      $doc.Save()
      $pages = $doc.ComputeStatistics(2)  # wdStatisticPages
      # ExportAsFixedFormat(Out, wdExportFormatPDF, OpenAfter, OptimizeForPrint, wdExportAllDocument,
      #   From, To, wdExportDocumentContent, IncludeDocProps, KeepIRM, CreateBookmarks=none,
      #   DocStructureTags (tagged PDF = cleaner text for ATS parsers), BitmapMissingFonts, PDF/A)
      $doc.ExportAsFixedFormat($pdf, 17, $false, 0, 0, 1, 1, 0, $true, $true, 0, $true, $true, $false)
      Write-Output "OK | pages=$pages | $pdf"
    } finally {
      $doc.Close(0)
    }
  }
} finally {
  $word.Quit()
  [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
