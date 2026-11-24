#Requires -Version 5.1
<#
    Verifie une livraison extraite d'un ZIP, sur un poste sans git, sans module et sans
    droits particuliers. Rien n'est installe ni modifie : ce script ne fait que lire.

      - SHA256SUMS : chaque fichier annonce est present et a l'empreinte annoncee ;
      - les quatre elements livres sont la (EXE, les cinq XAML, le moteur, setup.exe) ;
      - setup.exe porte une signature Authenticode valide emise par Microsoft ;
      - l'EXE porte le nom et le produit attendus ;
      - le dossier ne contient ni config.json ni journal (fichiers du poste, pas de la
        livraison).

    Usage, apres extraction du ZIP :
      powershell -NoProfile -ExecutionPolicy Bypass -File .\Verifier-Livraison.ps1
      ... -Dossier C:\chemin\de\la\livraison

    Code de sortie : 0 si la livraison est conforme, 1 sinon.
#>
param([string]$Dossier = "")

$ErrorActionPreference = "Stop"
if ($Dossier -eq "") { $Dossier = $PSScriptRoot }

$script:Pass = 0
$script:Fail = 0

function Check {
    param([string]$Nom, [bool]$Ok, [string]$Detail = "")
    if ($Ok) { $script:Pass++ ; Write-Host ("  PASS   " + $Nom) }
    else { $script:Fail++ ; Write-Host ("  ECHEC  " + $Nom + "   -> " + $Detail) -ForegroundColor Red }
}

function Get-Hash {
    param([string]$Chemin)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $flux = [System.IO.File]::OpenRead($Chemin)
    try { return ([System.BitConverter]::ToString($sha.ComputeHash($flux)).Replace("-", "").ToLowerInvariant()) }
    finally { $flux.Dispose() }
}

Write-Host ("Livraison verifiee : " + $Dossier)

# 1. Sommes annoncees.
$fichierSommes = Join-Path $Dossier "SHA256SUMS"
$sommes = @()
if (Test-Path -LiteralPath $fichierSommes) {
    foreach ($l in [System.IO.File]::ReadAllLines($fichierSommes)) {
        if ($l -match '^([0-9a-fA-F]{64})\s+\*?(.+)$') {
            $sommes += @{ Hash = $matches[1].ToLowerInvariant() ; Chemin = $matches[2].Trim() }
        }
    }
}
Check "SHA256SUMS lisible" ($sommes.Count -gt 0) ("fichier absent ou vide : " + $fichierSommes + " - la livraison a-t-elle ete extraite en entier ?")

# 2. Chaque fichier annonce est la, et son empreinte correspond.
$ecarts = @()
foreach ($s in $sommes) {
    $chemin = Join-Path $Dossier ($s.Chemin.Replace('/', '\'))
    if (-not (Test-Path -LiteralPath $chemin)) {
        $ecarts += ($s.Chemin + " : absent")
    } else {
        $h = Get-Hash $chemin
        if ($h -ne $s.Hash) { $ecarts += ($s.Chemin + " : annonce=" + $s.Hash.Substring(0, 12) + " trouve=" + $h.Substring(0, 12)) }
    }
}
Check ("SHA256SUMS : les " + $sommes.Count + " fichiers annonces sont conformes") ($ecarts.Count -eq 0) ($ecarts -join " ; ")

# 3. Les quatre elements livres, et les cinq XAML un par un.
$attendus = @(
    "Install-Office2024-KMS.exe"
    "Install-Office2024-KMS.ps1"
    "setup.exe"
    "Maquette\Principal.xaml"
    "Maquette\Styles.xaml"
    "Maquette\Applications.xaml"
    "Maquette\Etats.xaml"
    "Maquette\ReliquatProxy.xaml"
)
$absents = @($attendus | Where-Object { -not (Test-Path -LiteralPath (Join-Path $Dossier $_)) })
Check "les quatre elements livres sont la (EXE, cinq XAML, moteur, setup.exe)" ($absents.Count -eq 0) ("absents : " + ($absents -join ", "))

# 4. setup.exe signe par Microsoft.
$cheminSetup = Join-Path $Dossier "setup.exe"
$sig = $null
if (Test-Path -LiteralPath $cheminSetup) { $sig = Get-AuthenticodeSignature -LiteralPath $cheminSetup }
$sujet = if ($null -ne $sig -and $null -ne $sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { "" }
$etat = if ($null -ne $sig) { [string]$sig.Status } else { "(absent)" }
Check "setup.exe : signature Authenticode valide emise par Microsoft" (($etat -eq "Valid") -and ($sujet -match 'O=Microsoft Corporation')) ("Status=" + $etat + " sujet=[" + $sujet + "]")

# 5. Proprietes du binaire livre.
$cheminExe = Join-Path $Dossier "Install-Office2024-KMS.exe"
$vi = $null
if (Test-Path -LiteralPath $cheminExe) { $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($cheminExe) }
$nomExe = if ($null -ne $vi) { [string]$vi.OriginalFilename } else { "(absent)" }
$produit = if ($null -ne $vi) { [string]$vi.ProductName } else { "(absent)" }
Check "EXE : nom livre et produit attendus" (($nomExe -eq "Install-Office2024-KMS.exe") -and ($produit -eq "InstallOffice")) ("OriginalFilename=[" + $nomExe + "] ProductName=[" + $produit + "]")

# 6. Aucun fichier du poste ne s'est glisse dans la livraison.
$locaux = @(Get-ChildItem -LiteralPath $Dossier -File |
    Where-Object { $_.Name -match '^config\.json$' -or $_.Name -match '\.log$' } |
    ForEach-Object { $_.Name })
Check "aucun config.json ni journal dans la livraison" ($locaux.Count -eq 0) ("trouves : " + ($locaux -join ", ") + " - ce sont des fichiers du poste, pas de la livraison")

Write-Host ""
if ($script:Fail -eq 0) {
    Write-Host ("RESULTAT : livraison conforme (" + $script:Pass + " controles).") -ForegroundColor Green
    exit 0
}
Write-Host ("RESULTAT : " + $script:Fail + " ecart(s) sur " + ($script:Pass + $script:Fail) + " controles.") -ForegroundColor Red
exit 1
