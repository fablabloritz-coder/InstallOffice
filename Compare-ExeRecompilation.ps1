#Requires -Version 5.1
<#
    Recompile Apercu-Maquette.ps1 avec la commande documentee (README, section
    "Compilation de l'EXE") et compare le resultat au binaire commite :

      - le script embarque (SHA-256 des octets de la ressource .ps1) doit etre identique :
        c'est ce qui prouve que l'EXE livre contient bien le script du depot ;
      - les proprietes du binaire (FileVersionInfo) et le manifeste requireAdmin doivent
        etre identiques : c'est ce qui prouve qu'il a ete compile avec la commande
        documentee, pas avec d'autres options.

    Les octets des deux EXE ne sont PAS compares : PS2EXE n'est pas reproductible
    (horodatage et identifiants de ressources varient d'une compilation a l'autre). La
    taille et les empreintes des binaires sont affichees pour information seulement.

    Usage (c'est ce que lance le job "recompilation" du workflow) :
      powershell -NoProfile -ExecutionPolicy Bypass -File .\Compare-ExeRecompilation.ps1

    Code de sortie : 0 si le binaire commite embarque le script actuel avec les proprietes
    de la commande documentee, 1 sinon.
#>
param(
    [string]$ExeLivre = "",
    [string]$ScriptSource = ""
)

$ErrorActionPreference = "Stop"

if ($ExeLivre -eq "")   { $ExeLivre   = Join-Path $PSScriptRoot "Install-Office2024-KMS.exe" }
if ($ScriptSource -eq "") { $ScriptSource = Join-Path $PSScriptRoot "Apercu-Maquette.ps1" }

# Options de la commande documentee. Si elles changent dans le README, elles changent ici :
# un ecart avec le binaire commite fait echouer ce controle, ce qui force a recompiler.
$options = @{
    NoConsole   = $true
    RequireAdmin = $true
    x64         = $true
    Title       = "Deploiement Office 2024"
    Product     = "InstallOffice"
    Description = "Installation Office LTSC 2024 et activation KMS"
    Company     = "Lycee Loritz"
    Version     = "8.4.0.0"
}

function Get-ScriptEmbarque {
    param([string]$Chemin)
    $octets = [System.IO.File]::ReadAllBytes($Chemin)
    $asm = [System.Reflection.Assembly]::Load($octets)
    $nom = @($asm.GetManifestResourceNames() | Where-Object { $_ -like "*.ps1" }) | Select-Object -First 1
    if (-not $nom) { return $null }
    $flux = $asm.GetManifestResourceStream($nom)
    $mem = New-Object System.IO.MemoryStream
    $flux.CopyTo($mem)
    return $mem.ToArray()
}

function Get-Empreinte {
    param([byte[]]$Octets)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    return ([System.BitConverter]::ToString($sha.ComputeHash($Octets)).Replace("-", ""))
}

function Get-Proprietes {
    param([string]$Chemin)
    $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Chemin)
    $octets = [System.IO.File]::ReadAllBytes($Chemin)
    $ascii = [System.Text.Encoding]::ASCII.GetString($octets)
    $niveau = if ($ascii -match 'requestedExecutionLevel[^>]*level\s*=\s*"(?<lvl>[A-Za-z]+)"') { $matches["lvl"] } else { "(absent)" }
    return @{
        FileDescription  = [string]$vi.FileDescription
        FileVersion      = [string]$vi.FileVersion
        ProductName      = [string]$vi.ProductName
        ProductVersion   = [string]$vi.ProductVersion
        CompanyName      = [string]$vi.CompanyName
        LegalCopyright   = [string]$vi.LegalCopyright
        InternalName     = [string]$vi.InternalName
        OriginalFilename = [string]$vi.OriginalFilename
        Comments         = [string]$vi.Comments
        Manifeste        = $niveau
    }
}

$ecarts = @()
Write-Host ("EXE livre  : " + $ExeLivre)
if (-not (Test-Path -LiteralPath $ExeLivre)) {
    Write-Host ("ECHEC : binaire absent : " + $ExeLivre + " - le livrer ou corriger -ExeLivre.") -ForegroundColor Red
    exit 1
}
if (-not (Test-Path -LiteralPath $ScriptSource)) {
    Write-Host ("ECHEC : script source absent : " + $ScriptSource) -ForegroundColor Red
    exit 1
}

$module = Get-Module -ListAvailable ps2exe | Select-Object -First 1
if (-not $module) {
    Write-Host "ECHEC : module ps2exe absent - Install-Module ps2exe -Scope CurrentUser (README, Compilation de l'EXE)." -ForegroundColor Red
    exit 1
}
Import-Module ps2exe -ErrorAction Stop

$dossier = Join-Path $env:TEMP ("recompile_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $dossier -Force | Out-Null
$recompile = Join-Path $dossier "Install-Office2024-KMS.exe"
try {
    Invoke-PS2EXE -InputFile $ScriptSource -OutputFile $recompile `
        -NoConsole -RequireAdmin -x64 `
        -Title $options.Title -Product $options.Product `
        -Description $options.Description -Company $options.Company `
        -Version $options.Version | Out-Null
    if (-not (Test-Path -LiteralPath $recompile)) { throw "Invoke-PS2EXE n'a rien ecrit : " + $recompile }

    $octetsLivre = [System.IO.File]::ReadAllBytes($ExeLivre)
    $octetsRecomp = [System.IO.File]::ReadAllBytes($recompile)
    Write-Host ("recompile  : " + $recompile + " (ps2exe " + $module.Version + ")")
    # Information, jamais critere : deux compilations du meme script n'ont pas les memes octets.
    Write-Host ("octets     : livre=" + $octetsLivre.Length + " (sha256=" + (Get-Empreinte $octetsLivre).Substring(0, 12) + ") recompile=" + $octetsRecomp.Length + " (sha256=" + (Get-Empreinte $octetsRecomp).Substring(0, 12) + ") - non compares : PS2EXE n'est pas reproductible")

    $scriptLivre = Get-ScriptEmbarque $ExeLivre
    $scriptRecomp = Get-ScriptEmbarque $recompile
    if ($null -eq $scriptLivre -or $null -eq $scriptRecomp) {
        $ecarts += ("script embarque illisible : livre=" + $(if ($null -eq $scriptLivre) { "aucune ressource .ps1" } else { "ok" }) + " recompile=" + $(if ($null -eq $scriptRecomp) { "aucune ressource .ps1" } else { "ok" }) + " - verifier le format d'embarquement de ps2exe")
    } else {
        $hLivre = Get-Empreinte $scriptLivre
        $hRecomp = Get-Empreinte $scriptRecomp
        Write-Host ("script     : livre=" + $hLivre.Substring(0, 12) + " (" + $scriptLivre.Length + " octets) recompile=" + $hRecomp.Substring(0, 12) + " (" + $scriptRecomp.Length + " octets)")
        if ($hLivre -ne $hRecomp) {
            $ecarts += ("script embarque different : livre=" + $hLivre.Substring(0, 12) + " recompile=" + $hRecomp.Substring(0, 12) + " - recompiler avec la commande du README et committer le binaire")
        }
    }

    $propLivre = Get-Proprietes $ExeLivre
    $propRecomp = Get-Proprietes $recompile
    $comparees = 0
    foreach ($cle in @("FileDescription", "FileVersion", "ProductName", "ProductVersion", "CompanyName", "LegalCopyright", "InternalName", "OriginalFilename", "Comments", "Manifeste")) {
        $comparees++
        if ([string]$propLivre[$cle] -ne [string]$propRecomp[$cle]) {
            $ecarts += ($cle + " : livre=[" + $propLivre[$cle] + "] recompile=[" + $propRecomp[$cle] + "] - la commande Invoke-PS2EXE du README et celle de ce script doivent donner les memes options")
        }
    }
    Write-Host ("proprietes : " + $comparees + " champs compares, " + @($ecarts).Count + " ecart(s)")
    foreach ($e in $ecarts) { Write-Host ("  ECHEC  " + $e) -ForegroundColor Red }
}
catch {
    $ecarts += ("recompilation impossible : " + $_.Exception.Message)
    Write-Host ("ECHEC : " + $_.Exception.Message) -ForegroundColor Red
}
finally {
    Remove-Item -LiteralPath $dossier -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
if ($ecarts.Count -eq 0) {
    Write-Host "RESULTAT : l'EXE livre embarque Apercu-Maquette.ps1 du depot et porte les proprietes de la commande documentee." -ForegroundColor Green
    exit 0
}
Write-Host ("RESULTAT : " + $ecarts.Count + " ecart(s) entre l'EXE livre et une recompilation du script actuel.") -ForegroundColor Red
exit 1
