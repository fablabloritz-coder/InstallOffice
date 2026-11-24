<#
    Version de l'outil : une seule source.

    version.txt porte la version. Les autres emplacements la recopient, parce qu'ils ne
    peuvent pas la lire a l'execution : le badge est du XAML statique, le README du texte,
    et Compare-ExeRecompilation la passe a PS2EXE au moment de la compilation. Ce script
    les tient alignes.

    Sans argument : verification seule. Code 0 si tout est aligne, 1 sinon, avec le
    fichier, la ligne et la valeur trouvee.
    Avec -Version : verification d'abord (refus si un emplacement diverge), puis ecriture
    dans version.txt et dans toutes les copies.

    Le test lit version.txt : il ne doit contenir aucune version en dur, et ce script
    refuse s'il en trouve une.

    Apres un bump, Apercu-Maquette.ps1 change : recompiler l'EXE et regenerer les captures.
    Le script rappelle les deux commandes.
#>
param(
    [string]$Version = "",
    [string]$Racine = ""
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Racine)) { $Racine = $PSScriptRoot }
$fichierSource = Join-Path $Racine "version.txt"

$motifVersion = '[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+'

# Emplacements qui recopient la version : un seul groupe, la valeur.
$copies = @(
    @{ Fichier = "Apercu-Maquette.ps1" ; Nom = "constante du script" ;
       Regex = '\$script:versionOutil\s*=\s*"(' + $motifVersion + ')"' }
    @{ Fichier = "Maquette\Principal.xaml" ; Nom = "badge de la fenetre" ;
       Regex = '<TextBlock Text="(' + $motifVersion + ')"\s+FontSize="11"' }
    @{ Fichier = "Compare-ExeRecompilation.ps1" ; Nom = "Compare-ExeRecompilation" ;
       Regex = 'Version\s*=\s*"(' + $motifVersion + ')"' }
)
$fichierReadme = "README.md"
$fichierTest = "Test-InstallOffice.ps1"

function Get-LigneDe {
    param([string]$Texte, [int]$Index)
    if ($Index -le 0) { return 1 }
    return (@(($Texte.Substring(0, $Index) -split "`n")).Count)
}

function Write-Texte {
    param([string]$Chemin, [string]$Texte)
    # Le BOM d'origine est conserve : reecrire un script sans lui changerait ses octets.
    $octets = [System.IO.File]::ReadAllBytes($Chemin)
    $bom = ($octets.Length -ge 3 -and $octets[0] -eq 0xEF -and $octets[1] -eq 0xBB -and $octets[2] -eq 0xBF)
    [System.IO.File]::WriteAllText($Chemin, $Texte, (New-Object System.Text.UTF8Encoding($bom)))
}

# --- Source ----------------------------------------------------------------
$source = ""
try {
    if (-not (Test-Path -LiteralPath $fichierSource)) { throw "version.txt absent" }
    $source = ([string](Get-Content -LiteralPath $fichierSource -Raw)).Trim()
    if ($source -notmatch ('^' + $motifVersion + '$')) { throw ("version.txt ne porte pas N.N.N.N : '" + $source + "'") }
} catch {
    Write-Host ("ERREUR : " + $_.Exception.Message)
    exit 1
}

# --- Etat de chaque emplacement -------------------------------------------
$etat = New-Object System.Collections.ArrayList
foreach ($c in $copies) {
    $chemin = Join-Path $Racine $c.Fichier
    $valeur = "" ; $ligne = 0 ; $probleme = ""
    if (-not (Test-Path -LiteralPath $chemin)) {
        $probleme = "fichier absent"
    } else {
        $texte = [System.IO.File]::ReadAllText($chemin)
        $m = [regex]::Matches($texte, $c.Regex)
        if ($m.Count -eq 0) { $probleme = "emplacement introuvable" }
        elseif ($m.Count -gt 1) { $probleme = "plusieurs occurrences (" + $m.Count + ")" }
        else { $valeur = $m[0].Groups[1].Value ; $ligne = Get-LigneDe $texte $m[0].Groups[1].Index }
    }
    [void]$etat.Add([pscustomobject]@{ Fichier = $c.Fichier ; Nom = $c.Nom ; Valeur = $valeur ; Ligne = $ligne ; Probleme = $probleme })
}

# README : toutes les mentions "version X" doivent porter la meme valeur.
$cheminReadme = Join-Path $Racine $fichierReadme
if (-not (Test-Path -LiteralPath $cheminReadme)) {
    [void]$etat.Add([pscustomobject]@{ Fichier = $fichierReadme ; Nom = "README" ; Valeur = "" ; Ligne = 0 ; Probleme = "fichier absent" })
} else {
    $texte = [System.IO.File]::ReadAllText($cheminReadme)
    $m = [regex]::Matches($texte, ('(?i)version\s*["`]?(' + $motifVersion + ')'))
    $valeurs = @($m | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    $probleme = ""
    if ($m.Count -eq 0) { $probleme = "aucune mention de version" }
    elseif ($valeurs.Count -gt 1) { $probleme = "mentions divergentes : " + ($valeurs -join ", ") }
    $ligne = if ($m.Count -gt 0) { Get-LigneDe $texte $m[0].Groups[1].Index } else { 0 }
    [void]$etat.Add([pscustomobject]@{ Fichier = $fichierReadme ; Nom = "README (toutes les mentions)" ; Valeur = ($valeurs -join ", ") ; Ligne = $ligne ; Probleme = $probleme })
}

# Le test ne recopie pas la version : il lit version.txt. Une valeur en dur y serait une
# seconde source, donc un emplacement qui divergera au prochain bump.
$cheminTest = Join-Path $Racine $fichierTest
if (Test-Path -LiteralPath $cheminTest) {
    $texte = [System.IO.File]::ReadAllText($cheminTest)
    $m = [regex]::Match($texte, ('(?i)version\s*["`]?(' + $motifVersion + ')'))
    if ($m.Success) {
        [void]$etat.Add([pscustomobject]@{ Fichier = $fichierTest ; Nom = "constante du test" ; Valeur = $m.Groups[1].Value ; Ligne = (Get-LigneDe $texte $m.Groups[1].Index) ; Probleme = "version en dur : le test doit lire version.txt" })
    }
}

# --- Verdict ---------------------------------------------------------------
$ecarts = @($etat | Where-Object { $_.Probleme -ne "" -or $_.Valeur -ne $source })
if ($ecarts.Count -gt 0) {
    Write-Host ("Version divergente : version.txt annonce " + $source + ", mais :")
    foreach ($e in $ecarts) {
        $vu = if ($e.Probleme -ne "") { $e.Probleme } else { "trouve " + $e.Valeur }
        Write-Host ("  " + $e.Fichier + " ligne " + $e.Ligne + " (" + $e.Nom + ") : " + $vu + " - attendu " + $source)
    }
    Write-Host "Rien n'a ete ecrit. Aligner ces emplacements, puis relancer."
    exit 1
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    Write-Host ("Version unique : " + $source + " dans " + @($etat).Count + " emplacements verifies.")
    exit 0
}

# --- Ecriture --------------------------------------------------------------
$nouvelle = $Version.Trim()
if ($nouvelle -notmatch ('^' + $motifVersion + '$')) {
    Write-Host ("Version refusee : '" + $nouvelle + "' (attendu N.N.N.N)")
    exit 1
}
if ($nouvelle -eq $source) {
    Write-Host ("Deja en " + $source + " : rien a faire.")
    exit 0
}

[System.IO.File]::WriteAllText($fichierSource, $nouvelle + "`n", (New-Object System.Text.UTF8Encoding($false)))

foreach ($c in $copies) {
    $chemin = Join-Path $Racine $c.Fichier
    $texte = [System.IO.File]::ReadAllText($chemin)
    $m = [regex]::Match($texte, $c.Regex)
    if (-not $m.Success) { Write-Host ("ERREUR : emplacement perdu dans " + $c.Fichier) ; exit 1 }
    $g = $m.Groups[1]
    $texte = $texte.Substring(0, $g.Index) + $nouvelle + $texte.Substring($g.Index + $g.Length)
    Write-Texte -Chemin $chemin -Texte $texte
}

$texte = [System.IO.File]::ReadAllText($cheminReadme)
$texte = [regex]::Replace($texte, ('(?<![0-9.])' + [regex]::Escape($source) + '(?![0-9.])'), $nouvelle)
Write-Texte -Chemin $cheminReadme -Texte $texte

Write-Host ("Version " + $source + " -> " + $nouvelle + " : version.txt et " + @($etat).Count + " emplacements ecrits.")
Write-Host ""
Write-Host "Apercu-Maquette.ps1 a change : recompiler l'EXE, puis regenerer les captures."
Write-Host '  Invoke-PS2EXE -InputFile "Apercu-Maquette.ps1" -OutputFile "Install-Office2024-KMS.exe" -NoConsole -RequireAdmin -x64 -Title "Deploiement Office 2024" -Product "InstallOffice" -Description "Installation Office LTSC 2024 et activation KMS" -Company "Lycee Loritz" -Version "' + $nouvelle + '"'
Write-Host '  powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1 -EnregistrerCaptures'
exit 0
