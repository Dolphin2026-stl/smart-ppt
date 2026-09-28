param([Parameter(Mandatory=$true)][string]$InputPath,[Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $InputPath).Path
$target = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $target -Force | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$deck = $null
try {
    $deck = $app.Presentations.Open($source, -1, 0, 0)
    $deck.Export($target, 'PNG', 1600, 900)
    Write-Output "Rendered $($deck.Slides.Count) slides to $target"
} finally {
    if ($null -ne $deck) { $deck.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($deck) }
    # Do not Quit: PowerPoint may share an existing user application instance.
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
