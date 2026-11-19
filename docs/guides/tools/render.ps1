param([string]$Docx, [string]$Pdf)
# Opens the .docx in Word (hidden), updates the table of contents, saves a PDF copy. The .docx is not changed.
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($Docx, $false, $true)
    foreach ($t in $doc.TablesOfContents) { $t.Update() }
    $doc.SaveAs([ref]$Pdf, [ref]17)
    $doc.Close([ref]0)
    "pdf ok"
} finally {
    $word.Quit()
}
