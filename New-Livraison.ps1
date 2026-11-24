#Requires -Version 5.1
<#
    Construit Livraison.zip a partir d'une liste blanche.

    Ce qui part sur un poste : les quatre elements livres (l'EXE, son dossier Maquette, le
    moteur dont le worker est extrait a chaque deploiement, setup.exe), le README, le
    verificateur que l'enseignant lance sans git, et un SHA256SUMS calcule au passage.
    Rien d'autre n'entre : ni config.json, ni journal, ni sources Office\ (plusieurs Go),
    ni la suite de tests, qui a besoin du depot.

    Usage :
      powershell -NoProfile -ExecutionPolicy Bypass -File .\New-Livraison.ps1
      ... -Destination C:\tmp\Livraison.zip     (la suite l'utilise dans un dossier de travail)

    Code de sortie : 0 si le ZIP est ecrit, 1 si un element de la liste manque.
#>
param(
    [string]$Destination = "",
    [string]$Racine = ""
)

$ErrorActionPreference = "Stop"

if ($Racine -eq "") { $Racine = $PSScriptRoot }
if ($Destination -eq "") { $Destination = Join-Path $Racine "Livraison.zip" }

# Liste blanche, dans l'ordre ou le poste les lit. Un dossier est pris recursivement.
$listeBlanche = @(
    "Install-Office2024-KMS.exe"
    "Maquette"
    "Install-Office2024-KMS.ps1"
    "setup.exe"
    "README.md"
    "Verifier-Livraison.ps1"
)

$fichiers = @()
$manquants = @()
foreach ($element in $listeBlanche) {
    $chemin = Join-Path $Racine $element
    if (Test-Path -LiteralPath $chemin -PathType Container) {
        foreach ($f in @(Get-ChildItem -LiteralPath $chemin -Recurse -File | Sort-Object FullName)) { $fichiers += $f }
    } elseif (Test-Path -LiteralPath $chemin -PathType Leaf) {
        $fichiers += (Get-Item -LiteralPath $chemin)
    } else {
        $manquants += $element
    }
}
if ($manquants.Count -gt 0) {
    Write-Host ("ECHEC : elements de la liste blanche absents : " + ($manquants -join ", ")) -ForegroundColor Red
    exit 1
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# SHA256SUMS : une ligne par fichier, chemin relatif en barres obliques, format sha256sum.
$sommes = New-Object System.Collections.Generic.List[string]
$sha = [System.Security.Cryptography.SHA256]::Create()
$relatifs = @()
foreach ($f in $fichiers) {
    $rel = $f.FullName.Substring($Racine.Length).TrimStart([char]'\').Replace('\', '/')
    $relatifs += $rel
    $octets = [System.IO.File]::ReadAllBytes($f.FullName)
    $sommes.Add(([System.BitConverter]::ToString($sha.ComputeHash($octets)).Replace("-", "").ToLowerInvariant() + "  " + $rel))
}

if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Force }
$dossierDest = Split-Path -Parent $Destination
if ($dossierDest -ne "" -and -not (Test-Path -LiteralPath $dossierDest)) {
    New-Item -ItemType Directory -Path $dossierDest -Force | Out-Null
}

$zip = [System.IO.Compression.ZipFile]::Open($Destination, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    for ($i = 0; $i -lt $fichiers.Count; $i++) {
        [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip, $fichiers[$i].FullName, $relatifs[$i], [System.IO.Compression.CompressionLevel]::Optimal)
    }
    $entree = $zip.CreateEntry("SHA256SUMS", [System.IO.Compression.CompressionLevel]::Optimal)
    $flux = $entree.Open()
    $ecrivain = New-Object System.IO.StreamWriter($flux, (New-Object System.Text.UTF8Encoding($false)))
    try { $ecrivain.Write(($sommes -join "`n") + "`n") } finally { $ecrivain.Dispose() }
} finally {
    $zip.Dispose()
}

Write-Host ("Livraison.zip ecrit : " + $Destination)
Write-Host ("  " + $fichiers.Count + " fichiers + SHA256SUMS, " + (Get-Item -LiteralPath $Destination).Length + " octets")
foreach ($l in $sommes) { Write-Host ("  " + $l) }
exit 0
