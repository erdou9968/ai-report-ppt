param([string]$pptx, [string]$pdf)
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
try {
    $pres = $app.Presentations.Open($pptx, $true, $false, $false)
    $pres.SaveAs($pdf, 32)
    Write-Output ("slides: " + $pres.Slides.Count)
    $pres.Close()
} finally {
    $app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}
