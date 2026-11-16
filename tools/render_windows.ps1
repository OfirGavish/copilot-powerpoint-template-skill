param(
    [string]$InputDirectory = (Join-Path (Split-Path $PSScriptRoot -Parent) 'outputs\checks'),
    [string]$OutputDirectory = (Join-Path (Split-Path $PSScriptRoot -Parent) 'outputs\rendered')
)
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$overflow = [Collections.Generic.List[object]]::new()
try {
    foreach ($file in Get-ChildItem -LiteralPath $InputDirectory -Filter '*.pptx' | Sort-Object Name) {
        $dest = Join-Path $OutputDirectory $file.BaseName
        [void](New-Item -ItemType Directory -Force -Path $dest)
        $staging = Join-Path ([IO.Path]::GetTempPath()) ('template-render-' + [guid]::NewGuid().ToString() + '.pptx')
        $presentation = $null
        try {
            # Unique, read-only copies avoid Office's cached-content behavior.
            Copy-Item -LiteralPath $file.FullName -Destination $staging
            $presentation = $app.Presentations.Open($staging, -1, 0, 0)
            for ($index = 1; $index -le $presentation.Slides.Count; $index++) {
                $slide = $presentation.Slides.Item($index)
                try {
                    $png = Join-Path $dest ('slide-{0:D2}.png' -f $index)
                    if (Test-Path -LiteralPath $png) { Remove-Item -LiteralPath $png }
                    $slide.Export($png, 'PNG', 1600, 900)
                    foreach ($shape in $slide.Shapes) {
                        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                            $frame = $shape.TextFrame2
                            $available = $shape.Height - $frame.MarginTop - $frame.MarginBottom
                            if ($frame.TextRange.BoundHeight -gt $available + 1.5) {
                                $overflow.Add(@{
                                    file = $file.Name
                                    slide = $index
                                    text = $frame.TextRange.Text
                                    boundHeight = $frame.TextRange.BoundHeight
                                    availableHeight = $available
                                })
                            }
                        }
                    }
                } finally {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($slide)
                }
            }
            Write-Output ("Rendered {0}: {1} slides" -f $file.Name, $presentation.Slides.Count)
        } finally {
            if ($null -ne $presentation) {
                $presentation.Close()
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
            }
            if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging }
        }
    }
} finally {
    # Never quit a shared Office instance or close existing user presentations.
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
ConvertTo-Json -InputObject @($overflow.ToArray()) -Depth 4 |
    Set-Content -LiteralPath (Join-Path $OutputDirectory 'text-overflow.json') -Encoding utf8
if ($overflow.Count -gt 0) { throw "Detected $($overflow.Count) text-height issues." }
