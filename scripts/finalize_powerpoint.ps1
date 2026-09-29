param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [Parameter(Mandatory = $true)][string]$Plan,
    [Parameter(Mandatory = $true)][string]$Audit
)

$ErrorActionPreference = 'Stop'
$pptxPath = (Resolve-Path -LiteralPath $Pptx).Path
$planData = Get-Content -LiteralPath $Plan -Raw -Encoding UTF8 | ConvertFrom-Json
$auditData = Get-Content -LiteralPath $Audit -Raw -Encoding UTF8 | ConvertFrom-Json
$slides = @($auditData.slides)

function Get-NoteForSlide($slideInfo, [int]$outputIndex, $planObject) {
    if ($planObject.sequence -and $outputIndex -lt $planObject.sequence.Count -and $planObject.sequence[$outputIndex].notes) {
        return [string]$planObject.sequence[$outputIndex].notes
    }
    $sourcePage = [string]$slideInfo.source_page
    if ($planObject.speaker_notes -and $planObject.speaker_notes.PSObject.Properties[$sourcePage]) {
        return [string]$planObject.speaker_notes.$sourcePage
    }
    foreach ($edit in @($planObject.edits)) {
        if ($edit.page -eq $slideInfo.source_page -and $edit.notes) { return [string]$edit.notes }
    }
    return $null
}

$notes = @()
for ($i = 0; $i -lt $slides.Count; $i++) {
    $note = Get-NoteForSlide $slides[$i] $i $planData
    if ([string]::IsNullOrWhiteSpace($note)) { throw "Missing speaker notes for output slide $($i + 1). Provide sequence[].notes, speaker_notes, or edits[].notes in the plan JSON." }
    $notes += $note
}

$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($pptxPath, $false, $false, $false)
    if ($presentation.Slides.Count -ne $notes.Count) { throw "The audit slide count does not match the presentation." }
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $notesPage = $presentation.Slides.Item($i).NotesPage
        $body = $null
        foreach ($shape in $notesPage.Shapes) {
            try {
                if ($shape.PlaceholderFormat.Type -eq 2) { $body = $shape; break }
            } catch { }
        }
        if ($null -eq $body) { throw "Slide $i has no speaker notes body placeholder." }
        $body.TextFrame.TextRange.Text = $notes[$i - 1]
    }
    $presentation.Save()
}
finally {
    if ($null -ne $presentation) { $presentation.Close() }
    if ($null -ne $powerPoint -and $powerPoint.Presentations.Count -eq 0) { $powerPoint.Quit() }
    if ($null -ne $presentation) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) }
    if ($null -ne $powerPoint) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
python (Join-Path $scriptDir 'write_source_report.py') $pptxPath --plan $Plan --audit $Audit
if ($LASTEXITCODE -ne 0) { throw 'Could not create the source report.' }
Write-Output "Saved notes for $($notes.Count) slides and created the source report."
