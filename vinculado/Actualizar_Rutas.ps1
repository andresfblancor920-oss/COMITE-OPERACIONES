# Apunta todos los vínculos de las presentaciones de esta carpeta al archivo
# Base_Comite_Ventas.xlsx de esta misma carpeta. Cierre PowerPoint antes de ejecutar.
$ErrorActionPreference = "Stop"
$carpeta = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsx = Join-Path $carpeta "Base_Comite_Ventas.xlsx"
if (-not (Test-Path $xlsx)) { Write-Host "No se encontró $xlsx"; Read-Host "Enter para salir"; exit 1 }

function Encode-Ruta([string]$ruta) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($b in [System.Text.Encoding]::UTF8.GetBytes($ruta)) {
        $ch = [char]$b
        if (($b -lt 128) -and ("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.~:\/()',;=+$!".IndexOf($ch) -ge 0)) {
            [void]$sb.Append($ch)
        } else {
            [void]$sb.Append('%' + $b.ToString('X2'))
        }
    }
    return $sb.ToString()
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$nuevo = 'Target="file:///' + (Encode-Ruta $xlsx)
$patron = 'Target="file:///[^"!]*?Base_Comite_Ventas\.xlsx'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$total = 0
foreach ($pptx in Get-ChildItem -Path $carpeta -Filter *.pptx) {
    try {
        $zip = [System.IO.Compression.ZipFile]::Open($pptx.FullName, 'Update')
    } catch {
        Write-Host ("{0}: no se pudo abrir (¿está abierto en PowerPoint? ciérrelo y repita)" -f $pptx.Name)
        continue
    }
    $n = 0
    try {
        foreach ($e in @($zip.Entries)) {
            if (-not $e.FullName.EndsWith('.rels')) { continue }
            $s = $e.Open()
            $sr = New-Object System.IO.StreamReader($s, $utf8, $true, 4096, $true)
            $txt = $sr.ReadToEnd()
            $sr.Dispose()
            $cuenta = ([regex]::Matches($txt, $patron)).Count
            if ($cuenta -gt 0) {
                $txt2 = [regex]::Replace($txt, $patron, $nuevo.Replace('$', '$$'))
                $s.SetLength(0)
                $s.Position = 0
                $sw = New-Object System.IO.StreamWriter($s, $utf8)
                $sw.Write($txt2)
                $sw.Flush()
                $sw.Dispose()
                $n += $cuenta
            } else {
                $s.Dispose()
            }
        }
    } finally { $zip.Dispose() }
    Write-Host ("{0}: {1} vínculos actualizados" -f $pptx.Name, $n)
    $total += $n
}
Write-Host ""
Write-Host "Listo. Vínculos apuntando a: $xlsx"
Write-Host "Abra la presentación y elija 'Actualizar vínculos'."
Read-Host "Enter para salir"
