# Swaps the logo in the three .xlsx templates for src/webparts/timesheet/assets/peak-logo.jpg.
# Edits the zip directly (no Excel, no ExcelJS). Safe to rerun: keeps the image's top-left and
# width, drops any crop, and sets the height from the logo's aspect ratio.
param([string[]]$Types = @('Admin', 'Instructor', 'Medic'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing
$root = Split-Path -Parent $PSScriptRoot
$logoPath = Join-Path $root 'src/webparts/timesheet/assets/peak-logo.jpg'
$logo = [IO.File]::ReadAllBytes($logoPath)
$img = [Drawing.Image]::FromFile($logoPath)
$ratio = $img.Height / $img.Width
$img.Dispose()
$rowEmu = 54.95 * 12700 # header row height; the logo must stay inside it

foreach ($type in $Types) {
    $file = Join-Path $root "Templates/xlsx/$type Timesheet Template.xlsx"
    $zip = [IO.Compression.ZipFile]::Open($file, 'Update')
    try {
        $media = @($zip.Entries | Where-Object { $_.FullName -like 'xl/media/*' })
        if ($media.Count -ne 1) { throw "$type : expected one image, found $($media.Count)" }
        $name = $media[0].FullName
        $media[0].Delete()
        $s = $zip.CreateEntry($name).Open(); $s.Write($logo, 0, $logo.Length); $s.Close()

        $entry = $zip.GetEntry('xl/drawings/drawing1.xml')
        $r = New-Object IO.StreamReader($entry.Open()); $xml = $r.ReadToEnd(); $r.Close()
        $entry.Delete()
        $xml = $xml -replace '<a:srcRect[^>]*/>', ''
        $null = $xml -match '<xdr:from>.*?<xdr:rowOff>(\d+)</xdr:rowOff>'
        $top = [long]$Matches[1]
        $null = $xml -match '<a:ext cx="(\d+)" cy="\d+"/>'
        $cx = [long]$Matches[1]
        $cy = [long][math]::Round($cx * $ratio)
        if ($top + $cy -gt $rowEmu) { # too tall: shrink both sides proportionally
            $cy = [long]($rowEmu - $top); $cx = [long][math]::Round($cy / $ratio)
        }
        $xml = $xml -replace '<a:ext cx="\d+" cy="\d+"/>', "<a:ext cx=`"$cx`" cy=`"$cy`"/>"
        $xml = $xml -replace '(<xdr:to>.*?<xdr:rowOff>)\d+(</xdr:rowOff>)', "`${1}$($top + $cy)`${2}"
        $w = New-Object IO.StreamWriter($zip.CreateEntry('xl/drawings/drawing1.xml').Open(), (New-Object Text.UTF8Encoding($false)))
        $w.Write($xml); $w.Close()
    } finally {
        $zip.Dispose()
    }
    Write-Host "$type : cx=$cx cy=$cy to.rowOff=$($top + $cy)"
}
