# Converts the macro-enabled templates to macro-free .xlsx copies.
# Edits the OOXML package directly (no Excel): drops vbaProject.bin and its relationship,
# and switches the workbook content type from macroEnabled to the plain sheet type.
# Excel COM was dropped because it hangs opening the Medic template.
param([string[]]$Names = @('Admin Timesheet Template', 'Instructor Timesheet Template', 'Medic Timesheet Template', 'Expense Report Template'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $root 'Templates'
$out = Join-Path $src 'xlsx'
New-Item -ItemType Directory -Force $out | Out-Null

function Edit-Entry($zip, $name, [scriptblock]$change) {
    $entry = $zip.GetEntry($name)
    $reader = New-Object IO.StreamReader($entry.Open())
    $text = $reader.ReadToEnd()
    $reader.Close()
    $entry.Delete()
    $writer = New-Object IO.StreamWriter($zip.CreateEntry($name).Open(), (New-Object Text.UTF8Encoding($false)))
    $writer.Write((& $change $text))
    $writer.Close()
}

foreach ($name in $Names) {
    $in = Join-Path $src "$name (current).xlsm"
    $dest = Join-Path $out "$name.xlsx"
    Copy-Item -LiteralPath $in -Destination $dest -Force
    $zip = [IO.Compression.ZipFile]::Open($dest, 'Update')
    try {
        $vba = $zip.GetEntry('xl/vbaProject.bin')
        if ($vba) { $vba.Delete() }
        Edit-Entry $zip '[Content_Types].xml' {
            param($t)
            $t = $t -replace 'application/vnd\.ms-excel\.sheet\.macroEnabled\.main\+xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
            $t -replace '<(Default|Override) [^>]*ContentType="application/vnd\.ms-office\.vbaProject"/>', ''
        }
        Edit-Entry $zip 'xl/_rels/workbook.xml.rels' {
            param($t)
            $t -replace '<Relationship [^>]*relationships/vbaProject"[^>]*/>', ''
        }
    } finally {
        $zip.Dispose()
    }
    Write-Host "Wrote $dest"
}
