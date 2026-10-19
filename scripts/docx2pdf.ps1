# Render a .docx to PDF using the installed Word COM automation.
#
# Why: the editor SDK's doc_to_image service is remote and frequently
# unavailable (50xxx/437xxx errors). Word is already installed on any machine
# that edits these reports, and its PDF export is the ground truth for how the
# document will actually paginate.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File docx2pdf.ps1 -Src in.docx -Pdf out.pdf
#
# Then rasterise with:  python pdf2png.py out.pdf pages/
param(
  [Parameter(Mandatory = $true)][string]$Src,
  [Parameter(Mandatory = $true)][string]$Pdf
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Src)) {
  Write-Output "SRC NOT FOUND: $Src"; exit 2
}
if (Test-Path -LiteralPath $Pdf) { Remove-Item -LiteralPath $Pdf -Force }

# Open read-only (3rd arg $true) so Word never prompts to save the source.
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($Src, $false, $true)
  # 17 = wdFormatPDF
  $doc.SaveAs([ref]$Pdf, [ref]17)
  $doc.Close($false)
  Write-Output ("PDF OK: " + (Get-Item -LiteralPath $Pdf).Length)
}
catch {
  # Word rejects files whose namespace prefixes were dropped by a naive
  # ElementTree re-serialization with "the file may be corrupted".
  Write-Output ("PDF FAILED: " + $_.Exception.Message)
  exit 1
}
finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
