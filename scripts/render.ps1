param([string]$pptx, [string]$outdir, [int]$width = 1600)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $outdir | Out-Null
Get-ChildItem $outdir -Filter *.png | Remove-Item -Force
$app = New-Object -ComObject PowerPoint.Application
try {
    $pres = $app.Presentations.Open($pptx, $true, $false, $false)
    $h = [int]($width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
    $i = 1
    foreach ($s in $pres.Slides) {
        $s.Export((Join-Path $outdir ("s{0:D2}.png" -f $i)), 'PNG', $width, $h)
        $i++
    }
    $pres.Close()
    Write-Output "exported $($i-1)"
} finally {
    $app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}
