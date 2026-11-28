#Requires -Version 5.1
<#
    Controles automatises de l'interface livree (Apercu-Maquette.ps1, compilee dans
    Install-Office2024-KMS.exe) et du moteur dont le worker est extrait
    (Install-Office2024-KMS.ps1, livre a cote de l'EXE)

    Rejoue les cas cles SANS toucher au systeme :
      - SourcePath impose/verifie au telechargement et a l'installation
      - destination des sources et verdict d'espace disque
      - codes de sortie ODT (32, 30066, 17002, 17029, absence de code)
      - echecs d'activation KMS (0xC004F074, 0x80070005, 0xC004F038, reponse inconnue, ospp.vbs absent)
      - depot invalide, proxy injoignable, bilan final

    setup.exe n'est JAMAIS execute : un faux binaire compile ici enregistre sa
    ligne de commande et renvoie le code voulu. Les echecs KMS utilisent un
    ospp.vbs factice dont le chemin remplace celui du script (seul ecart).

    Usage :
      powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1
      powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1 -Fast

    -Fast : section 1 seulement (syntaxe du script, empreinte et proprietes de
    l'EXE), quelques secondes. C'est le controle du hook pre-commit ; sans ce
    commutateur la suite complete tourne, comme sur GitHub.
    Code de sortie : 0 si tous les controles passent, 1 sinon.
#>

param([switch]$Fast, [switch]$EnregistrerCaptures)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing | Out-Null

# L'interface livree est Apercu-Maquette.ps1, compilee dans Install-Office2024-KMS.exe.
# Le moteur Install-Office2024-KMS.ps1 reste livre a cote de l'EXE : c'est de lui que le
# worker de deploiement est extrait, il fait donc partie du produit.
$MainScript   = Join-Path $PSScriptRoot "Apercu-Maquette.ps1"
$EngineScript = Join-Path $PSScriptRoot "Install-Office2024-KMS.ps1"
$ExePath      = Join-Path $PSScriptRoot "Install-Office2024-KMS.exe"
# PowerShell 5.1 du systeme : c'est lui qui rend les captures et rejoue la section 1 dans
# l'archive du depot, comme le ferait un poste livre.
$psHost       = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$Work       = Join-Path $env:TEMP ("ctl_office_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
$Fail       = 0
$Pass       = 0

function Check {
    param([string]$Name, [bool]$Ok, [string]$Detail = "")
    if ($Ok) { $script:Pass++ ; Write-Host ("  PASS   " + $Name) }
    else     { $script:Fail++ ; Write-Host ("  ECHEC  " + $Name + "   -> " + $Detail) -ForegroundColor Red }
}

function Section { param([string]$T) Write-Host "" ; Write-Host ("=== " + $T + " ===") }

function Get-EmpreinteSources {
    # SHA-256 des cinq XAML et du script, fins de ligne normalisees : le meme sur tout
    # checkout, quel que soit core.autocrlf. C'est la preuve independante du rendu.
    param([string]$Racine)
    $fichiers = @((Join-Path $Racine "Apercu-Maquette.ps1")) + @(Get-ChildItem -LiteralPath (Join-Path $Racine "Maquette") -Filter "*.xaml" | Sort-Object Name | ForEach-Object { $_.FullName })
    $flux = New-Object System.IO.MemoryStream
    foreach ($f in $fichiers) {
        $texte = (Split-Path $f -Leaf) + "`n" + (([System.IO.File]::ReadAllText($f)) -replace "`r`n", "`n") + "`n"
        $octets = [System.Text.Encoding]::UTF8.GetBytes($texte)
        $flux.Write($octets, 0, $octets.Length)
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = [System.BitConverter]::ToString($sha.ComputeHash($flux.ToArray())).Replace("-", "")
    $flux.Dispose()
    return $hash
}

function Get-EmpreinteRendu {
    # Polices que les captures utilisent, resolues sur ce poste : la comparaison au pixel
    # n'a de sens que si ce sont celles qui ont produit les references.
    Add-Type -AssemblyName PresentationCore
    $lignes = @()
    foreach ($nom in @("Segoe UI Variable Text, Segoe UI", "Segoe MDL2 Assets, Segoe UI", "Cascadia Mono, Consolas")) {
        $tf = New-Object System.Windows.Media.Typeface((New-Object System.Windows.Media.FontFamily($nom)), [System.Windows.FontStyles]::Normal, [System.Windows.FontWeights]::Normal, [System.Windows.FontStretches]::Normal)
        $gt = $null
        $uri = if ($tf.TryGetGlyphTypeface([ref]$gt)) { $gt.FontUri.ToString().ToLowerInvariant() } else { "(absent)" }
        $lignes += ($nom + "=" + $uri)
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    return ([System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes(($lignes -join "`n")))).Replace("-", ""))
}

# ---------------------------------------------------------------- extraction
function Get-WorkerBody {
    param([string[]]$Lines)
    $s = -1 ; $e = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($s -lt 0 -and $Lines[$i] -match '^\$script:BgDoWorkHandler = \{\s*$') { $s = $i ; continue }
        if ($s -ge 0 -and $Lines[$i] -match '^\}$') { $e = $i ; break }
    }
    if ($s -lt 0 -or $e -le $s) { throw "worker introuvable dans le script" }
    return ($Lines[($s + 1)..($e - 1)] -join "`n")
}

function Get-FunctionText {
    param([string[]]$Lines, [string]$Name)
    $start = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match ('^\s*function ' + [regex]::Escape($Name) + ' \{')) { $start = $i ; break }
    }
    if ($start -lt 0) { throw ("fonction " + $Name + " introuvable") }
    $depth = 0
    $hereString = $false
    for ($i = $start; $i -lt $Lines.Count; $i++) {
        $ligne = $Lines[$i]
        # Le code C# embarque (relais proxy) contient des accolades : sans ce saut,
        # le comptage s'arreterait au milieu du here-string et la fonction serait tronquee.
        if ($hereString) { if ($ligne -match "^'@") { $hereString = $false } ; continue }
        if ($ligne -match "@'\s*$") { $hereString = $true ; continue }
        $depth += ([regex]::Matches($ligne, '\{')).Count
        $depth -= ([regex]::Matches($ligne, '\}')).Count
        if ($depth -eq 0) { return ($Lines[$start..$i] -join "`n") }
    }
    throw ("fin de " + $Name + " introuvable")
}

function Get-ShownBlock {
    param([string[]]$Lines)
    $start = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match '^\$form\.Add_Shown\(\{') { $start = $i ; break }
    }
    if ($start -lt 0) { throw "bloc Add_Shown introuvable dans le script" }
    $depth = 0
    for ($i = $start; $i -lt $Lines.Count; $i++) {
        $depth += ([regex]::Matches($Lines[$i], '\{')).Count
        $depth -= ([regex]::Matches($Lines[$i], '\}')).Count
        if ($depth -eq 0) { return ($Lines[($start + 1)..($i - 1)] -join "`n") }
    }
    throw "fin du bloc Add_Shown introuvable"
}

# ---------------------------------------------------------------- fixtures
function New-FakeSetup {
    param([string]$Path)
    $code = 'using System; using System.IO; public class FakeOdt { public static int Main() { string a = Environment.GetEnvironmentVariable("FAKE_ODT_ARGS"); if (!String.IsNullOrEmpty(a)) { File.WriteAllText(a, Environment.CommandLine); } string v = Environment.GetEnvironmentVariable("FAKE_ODT_CODE"); int c; return Int32.TryParse(v, out c) ? c : 0; } }'
    Add-Type -TypeDefinition $code -OutputAssembly $Path -OutputType ConsoleApplication
}

function New-XmlFixture {
    param([string]$Path, [string]$SourcePath = "")
    $att = if ([string]::IsNullOrWhiteSpace($SourcePath)) { "" } else { ' SourcePath="' + $SourcePath + '"' }
    $xml = "<Configuration>`r`n  <Add OfficeClientEdition=`"64`" Channel=`"PerpetualVL2024`"$att>`r`n    <Product ID=`"ProPlus2024Volume`"><Language ID=`"fr-fr`" /></Product>`r`n  </Add>`r`n  <Display Level=`"None`" AcceptEULA=`"TRUE`" />`r`n</Configuration>"
    Set-Content -Path $Path -Value $xml -Encoding UTF8
}

function New-SyncHash {
    @{ LogQueue      = (New-Object System.Collections.Concurrent.ConcurrentQueue[object])
       StepState     = [string[]]@("waiting","waiting","waiting","waiting","waiting")
       StatusText    = "" ; ActiveStep = -1 ; StepStartTime = (Get-Date)
       HasError      = $false ; FinalMsg = "" ; Done = $false
       ShowPopup     = $false ; PopupTitle = "" ; PopupMsg = "" ; PopupIcon = "" }
}

function New-Cfg {
    param($Sync, [string]$Setup, [string]$Local,
          [string]$Depot = "", [string]$XmlPath = "",
          [bool]$Install = $false, [bool]$Activate = $false, [bool]$Proxy = $false)
    @{ SyncHash      = $Sync
       DepotEnabled  = (-not [string]::IsNullOrWhiteSpace($Depot))
       DepotPath     = $Depot
       UseCustomXML  = (-not [string]::IsNullOrWhiteSpace($XmlPath))
       CustomXMLPath = $XmlPath
       ActualXMLPath = ""
       SetupExePath  = $Setup
       LocalDataPath = $Local
       InstallFolder = (Join-Path $Local "install")
       SourcePath    = "" ; DepotCopy = $false
       LanguageID    = "fr-fr" ; ExcludedApps = @("Outlook")
       RemoveOld     = $false ; DoInstall = $Install ; DoActivate = $Activate ; ActivateOnly = $Activate
       ApplyProductKey = $false ; ProductKey = ""
       KmsHost       = "127.0.0.1" ; KmsPort = 1688 ; KmsSecondary = $null
       DiskMinGB     = 4
       ProxyEnabled  = $Proxy
       ProxyHost     = $(if ($Proxy) { "127.0.0.1" } else { "" })
       ProxyPort     = $(if ($Proxy) { "9" } else { "" })
       ProxyBypass   = "<local>" }
}

function Invoke-Deploy {
    param($Worker, $Cfg, [string]$ArgsFile)
    if (Test-Path $ArgsFile) { Remove-Item $ArgsFile -Force }
    $env:FAKE_ODT_ARGS = $ArgsFile
    $sync = $Cfg.SyncHash
    $ev = New-Object System.ComponentModel.DoWorkEventArgs($Cfg)
    $null = & $Worker $null $ev
    $logs = New-Object System.Collections.Generic.List[string]
    $item = $null
    while ($sync.LogQueue.TryDequeue([ref]$item)) { $logs.Add([string]$item.Text) }
    $args = if (Test-Path $ArgsFile) { [System.IO.File]::ReadAllText($ArgsFile).Trim() } else { "" }
    @{ Sync = $sync ; Logs = $logs ; Args = $args }
}

function Solution { param($Res) (@($Res.Logs) | Where-Object { $_ -match 'Solution :' } | Select-Object -Last 1) }

# Depot git jetable pret a tester le hook : fichiers copies, configuration fixee,
# etat de reference commite. Les cas A, B et C partent tous d'ici.
function New-HookRepo {
    param(
        [string]$Path,
        [string[]]$Files,
        [string[]]$Dossiers = @(),
        [string]$Autocrlf = "false",
        [scriptblock]$Avant = $null)
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
    foreach ($f in $Files) {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $f) -Destination $Path -Force
    }
    foreach ($d in $Dossiers) {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $d) -Destination $Path -Recurse -Force
    }
    # $Avant permet de preparer un etat non standard (script en LF) avant le commit.
    if ($null -ne $Avant) { & $Avant $Path $PSScriptRoot }
    $null = & git -C $Path init -q 2>&1
    $null = & git -C $Path config core.autocrlf $Autocrlf
    $null = & git -C $Path config core.eol native
    $null = & git -C $Path config core.hooksPath .githooks
    $null = & git -C $Path config commit.gpgsign false
    $null = & git -C $Path config user.email "controle@example.invalid"
    $null = & git -C $Path config user.name "Controle automatise"
    $null = & git -C $Path add -A 2>&1
    # --no-verify : sinon le hook du depot jetable relance la suite (recursion).
    $null = & git -C $Path commit -q --no-verify -m "etat de reference" 2>&1
}

# ------------------------------------------------------------- sh du hook
# Le hook est un script POSIX : il faut sh. Sur un poste ou PowerShell est lance
# hors de Git Bash, sh n'est pas sur le PATH alors que git, lui, est installe. Le
# resoudre par git evite un echec de harnais qui sauterait toute la section 10 ;
# s'il reste introuvable, les controles qui en dependent echouent en le nommant.
function Get-ShExe {
    $direct = Get-Command sh -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($direct) { return $direct.Source }
    $git = Get-Command git -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $git) { return $null }
    # Deux points de depart : le dossier de git.exe et celui que git declare comme
    # libexec (<git>/<arch>/libexec/git-core). sh vit a la racine de l'installation
    # (bin/ ou usr/bin/), donc on remonte jusqu'a quatre niveaux avant de chercher.
    $departs = @()
    $execPath = (& git --exec-path 2>$null)
    if (($LASTEXITCODE -eq 0) -and $execPath) { $departs += (Split-Path $execPath -Parent) }
    $departs += (Split-Path $git.Source -Parent)
    $relatifs = @("sh.exe", "bin\sh.exe", "usr\bin\sh.exe", "mingw64\bin\sh.exe", "ucrt64\bin\sh.exe")
    foreach ($depart in $departs) {
        $dossier = $depart
        for ($niveau = 0; ($niveau -lt 4) -and $dossier; $niveau++) {
            foreach ($relatif in $relatifs) {
                $candidat = Join-Path $dossier $relatif
                if (Test-Path -LiteralPath $candidat) { return $candidat }
            }
            $dossier = Split-Path $dossier -Parent
        }
    }
    return $null
}

# Jamais levante : un sh absent rend un code 127 et un motif lisible, pas une
# exception qui emporterait la section entiere. L'appelant a deja change de dossier.
function Invoke-HookSh {
    param([string]$ShExe, [string]$Hook)
    $script:HookShCode = 127
    if (-not $ShExe) { return "sh introuvable : ni sur le PATH, ni dans l'installation de git" }
    $eapHook = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $texte = (& $ShExe $Hook 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
        $script:HookShCode = $LASTEXITCODE
        return $texte
    } catch {
        $script:HookShCode = 127
        return ("sh inutilisable (" + $ShExe + ") : " + $_.Exception.Message)
    } finally {
        $ErrorActionPreference = $eapHook
    }
}

# ---------------------------------------------------------------- execution
if (-not (Test-Path $MainScript)) { Write-Host ("Interface introuvable : " + $MainScript) ; exit 2 }
if (-not (Test-Path $EngineScript)) { Write-Host ("Moteur introuvable : " + $EngineScript + " (le worker en est extrait)") ; exit 2 }
New-Item -ItemType Directory -Path $Work -Force | Out-Null
if ($Fast) { Write-Host "Mode rapide : section 1 seulement (script et proprietes de l'EXE)." }

# Enregistrement des empreintes des references, a lancer apres avoir regenere les quatre
# captures : la suite les relit, elle ne les ecrit jamais d'elle-meme.
if ($EnregistrerCaptures) {
    $fiche = Join-Path $PSScriptRoot "Maquette\captures.txt"
    $contenu = @(
        "# Empreintes enregistrees avec les quatre captures de Maquette (Test-InstallOffice.ps1)."
        "# sources : les cinq XAML et Apercu-Maquette.ps1, fins de ligne normalisees."
        "# rendu   : polices resolues sur le poste qui a produit les PNG."
        ("sources " + (Get-EmpreinteSources -Racine $PSScriptRoot))
        ("rendu " + (Get-EmpreinteRendu))
    )
    [System.IO.File]::WriteAllLines($fiche, $contenu, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host ("Captures : empreintes ecrites dans " + $fiche)
    exit 0
}

try {
    $src    = [System.IO.File]::ReadAllText($MainScript)
    $moteur = [System.IO.File]::ReadAllText($EngineScript)
    $lines  = [System.IO.File]::ReadAllLines($EngineScript)

    Section "1. Interface livree et proprietes de l'EXE"
    $tok = $null ; $err = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($src, [ref]$tok, [ref]$err)
    Check "analyse syntaxique" ((-not $err) -or ($err.Count -eq 0)) (($err | ForEach-Object { "L" + $_.Extent.StartLineNumber + ": " + $_.Message }) -join " ; ")
    $tokM = $null ; $errM = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($moteur, [ref]$tokM, [ref]$errM)
    Check "moteur : analyse syntaxique" ((-not $errM) -or ($errM.Count -eq 0)) (($errM | ForEach-Object { "L" + $_.Extent.StartLineNumber + ": " + $_.Message }) -join " ; ")
    # 'ProxyPassword' n'est plus un motif obsolete : c'est desormais une cle du fichier
    # de configuration du mode silencieux, lue et jamais ecrite. Ce qui reste interdit
    # est une affectation dans Save-ConfigPoste, verifiee juste apres sur son seul bloc.
    $obsoletes = @('bgWorker', 'RunWorkerAsync', 'DEFAULT_XML_PATH', 'proxyUserBox', 'proxyPassBox')
    $trouves = @($obsoletes | Where-Object { ([regex]::Matches($src, $_)).Count -gt 0 })
    Check "aucun motif obsolete" ($trouves.Count -eq 0) ($trouves -join ", ")
    # Les champs vivent dans le XAML de la fenetre principale, le script les remplit :
    # sans eux, un proxy qui repond 407 reste sans solution.
    $xamlPrincipal = Join-Path $PSScriptRoot "Maquette\Principal.xaml"
    $xaml = if (Test-Path -LiteralPath $xamlPrincipal) { [System.IO.File]::ReadAllText($xamlPrincipal) } else { "" }
    $champsProxy = (($xaml -match '<TextBox x:Name="ChampProxyUser"') -and ($xaml -match '<PasswordBox x:Name="ChampProxyPass"'))
    Check "proxy authentifie : champs utilisateur et mot de passe" ($champsProxy -and ($src -match '"ChampProxyUser"') -and ($src -match '"ChampProxyPass"')) "champs d'identifiants absents de l'interface : un proxy qui repond 407 reste alors sans solution"
    # La config enregistree porte l'utilisateur, jamais le mot de passe : le relais
    # le garde en memoire, il doit etre ressaisi a chaque lancement.
    $enregistreUser = ($src -match 'ProxyUser\s*=\s*\$Fenetre\.FindName\("ChampProxyUser"\)')
    # Le mot de passe peut apparaitre ailleurs (cle lue dans le fichier du mode
    # silencieux) : ce qui est interdit est de l'ecrire dans la configuration du poste.
    $blocConfigPoste = ""
    if ($src -match '(?s)function Save-ConfigPoste \{.*?\r?\n\}') { $blocConfigPoste = $matches[0] }
    $enregistrePass = [regex]::Matches($blocConfigPoste, 'ProxyPass\w*\s*=').Count
    Check "proxy authentifie : mot de passe jamais enregistre" ($enregistreUser -and ($enregistrePass -eq 0)) ("ProxyUser enregistre=" + $enregistreUser + " affectations de mot de passe=" + $enregistrePass)

    # Le dossier de travail n'est plus a la racine de C: : journaux, marqueur de relais et
    # sources vivent sous ProgramData. L'interface et le moteur ne doivent plus nommer
    # l'ancien dossier, sinon un poste deja configure y retourne.
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Get-DossierTravail")
    $dossierTravail = Get-DossierTravail
    $attenduTravail = Join-Path $env:ProgramData "Office2024Deploy"
    Check "dossier de travail : %ProgramData%\Office2024Deploy cote machine" ($dossierTravail -eq $attenduTravail) ("rendu=" + $dossierTravail + " attendu=" + $attenduTravail)
    $legacyMoteur = [regex]::Matches($moteur, 'C:\\InstallOffice2024').Count
    $legacyInterface = [regex]::Matches($src, 'C:\\InstallOffice2024').Count
    Check "chemins : plus aucun C:\InstallOffice2024 dans le moteur ni l'interface" (($legacyMoteur -eq 0) -and ($legacyInterface -eq 0)) ("moteur=" + $legacyMoteur + " interface=" + $legacyInterface + " - retirer le repli vers l'ancien dossier")

    # La case proxy etait desactivee des qu'un depot etait coche : elle devenait
    # incliquable, ce qui ressemblait a une panne. Le proxy ne sert qu'au telechargement,
    # mais son reglage doit rester accessible.
    $proxyCliquable = ($src -notmatch 'FindName\("ChkProxy"\)\.IsEnabled') -and ($moteur -notmatch '\$chkUseProxy\.Enabled')
    Check "proxy : la case reste cliquable quand un depot est choisi" $proxyCliquable "desactivation du controle proxy encore presente (interface ou moteur)"

    # Le champ setup affiche "setup.exe" : un chemin relatif se lit a cote de l'EXE, pas
    # dans le dossier courant du processus, qui depend de l'appelant (GPO, raccourci...).
    $lignesInterface = @($src -split "`r?`n")
    $ici = $PSScriptRoot
    Invoke-Expression (Get-FunctionText -Lines $lignesInterface -Name "Resolve-SetupPath")
    $setupRelatif = Resolve-SetupPath "setup.exe"
    $setupAbsolu  = Resolve-SetupPath "D:\outils\setup.exe"
    $setupVide    = Resolve-SetupPath ""
    $setupLegacy  = Resolve-SetupPath "C:\InstallOffice2024\setup.exe"
    $attenduSetup = Join-Path $PSScriptRoot "setup.exe"
    Check "setup : relatif et ancien emplacement resolus a cote de l'EXE, absolu conserve" (($setupRelatif -eq $attenduSetup) -and ($setupVide -eq $attenduSetup) -and ($setupLegacy -eq $attenduSetup) -and ($setupAbsolu -eq "D:\outils\setup.exe")) ("relatif=" + $setupRelatif + " vide=" + $setupVide + " legacy=" + $setupLegacy + " absolu=" + $setupAbsolu)

    # Contrat des noms : tout nom que le script demande a la fenetre (FindName) doit exister
    # dans l'un des cinq XAML. Un nom disparu ne se voit ni a la compilation ni au chargement
    # du XAML : la ligne d'interface reste muette, ou le script jette au moment de l'afficher.
    # Les noms construits (Ico$Ligne, Etape$Index, Glyphe$($script:etapeEncours)...) sont
    # verifies par leur prefixe, seul element stable d'un nom assemble.
    # Le dossier est teste avant d'etre lu : sans lui, Get-ChildItem leve et le harnais
    # s'arrete sur une erreur qui ne dit rien du contrat des noms.
    $dossierXaml = Join-Path $PSScriptRoot "Maquette"
    $nomsXaml = @()
    if (Test-Path -LiteralPath $dossierXaml -PathType Container) {
        foreach ($fXaml in @(Get-ChildItem -LiteralPath $dossierXaml -Filter "*.xaml" | Sort-Object Name)) {
            foreach ($m in [regex]::Matches([System.IO.File]::ReadAllText($fXaml.FullName), 'x:Name\s*=\s*([''"])([^''"]+)\1')) { $nomsXaml += $m.Groups[2].Value }
        }
    }
    $nomsXaml = @($nomsXaml | Sort-Object -Unique)
    # Les espaces et les guillemets simples sont toleres : FindName( "X" ) est la meme chose
    # que FindName("X"), et un nom ecrit ainsi ne doit pas echapper au contrat.
    $refsXaml = @([regex]::Matches($src, 'FindName\(\s*([''"])([^''"]+)\1\s*\)') | ForEach-Object { $_.Groups[2].Value } | Sort-Object -Unique)
    $nomsManquants = @() ; $prefixesManquants = @()
    foreach ($r in $refsXaml) {
        if ($r -match '\$') {
            $prefixe = ($r -split '\$')[0]
            if (@($nomsXaml | Where-Object { $_.StartsWith($prefixe) }).Count -eq 0) { $prefixesManquants += $prefixe }
        } elseif ($nomsXaml -notcontains $r) { $nomsManquants += $r }
    }
    $nomsManquants = @($nomsManquants | Sort-Object -Unique)
    $prefixesManquants = @($prefixesManquants | Sort-Object -Unique)
    # Un motif casse rendrait ce controle vide : il passerait en ne verifiant rien. Le
    # plancher sur le nombre de references lues l'interdit (63 aujourd'hui).
    $detailNoms = "absents des XAML=" + ($nomsManquants -join ", ") + " | prefixes sans aucun nom=" + ($prefixesManquants -join ", ") + " | references lues=" + $refsXaml.Count + ", noms XAML=" + $nomsXaml.Count
    if (-not (Test-Path -LiteralPath $dossierXaml)) { $detailNoms += " | dossier Maquette absent" }
    if ($refsXaml.Count -lt 20) { $detailNoms += " | moins de 20 FindName lus : motif de lecture casse, le controle ne verifierait rien" }
    Check "noms XAML : chaque FindName du script existe dans les cinq XAML" (($nomsManquants.Count -eq 0) -and ($prefixesManquants.Count -eq 0) -and ($refsXaml.Count -ge 20) -and ($nomsXaml.Count -gt 0)) ($detailNoms + " - renommer le XAML ou corriger le FindName")

    # Fins de ligne : l'EXE embarque le script tel qu'il est sur le poste, en
    # CRLF, et le controle compare des octets. Un clone en LF doit donc etre
    # signale ici, avec sa cause, plutot que par un faux "RECOMPILER" plus bas.
    $crlf = ([regex]::Matches($src, "`r`n")).Count
    $lfSeul = ([regex]::Matches($src, "(?<!`r)`n")).Count
    Check "script principal : fins de ligne CRLF (comme l'EXE embarque)" (($crlf -gt 0) -and ($lfSeul -eq 0)) ("CRLF=" + $crlf + " LF isole=" + $lfSeul + " - clone en LF plutot qu'une derive : obtenir un checkout CRLF (.gitattributes text eol=crlf) avant de conclure a une recompilation")
    $crlfM = ([regex]::Matches($moteur, "`r`n")).Count
    $lfSeulM = ([regex]::Matches($moteur, "(?<!`r)`n")).Count
    Check "moteur : fins de ligne CRLF" (($crlfM -gt 0) -and ($lfSeulM -eq 0)) ("CRLF=" + $crlfM + " LF isole=" + $lfSeulM)
    $attrPath = Join-Path $PSScriptRoot ".gitattributes"
    $attr = if (Test-Path -LiteralPath $attrPath) { [System.IO.File]::ReadAllText($attrPath) } else { "" }
    $declareCrlf = ($attr -match '(?m)^\s*\*\.ps1\s+text\s+eol=crlf\s*$')
    Check "fins de ligne : .gitattributes impose CRLF aux .ps1" $declareCrlf ("declaration absente de .gitattributes : un clone avec core.autocrlf=false produirait un faux RECOMPILER (ajouter : *.ps1 text eol=crlf)")
    Write-Host ("         interface " + $src.Length + " octets, moteur " + $lines.Count + " lignes")

    # L'EXE distribue embarque le script : son empreinte doit suivre le fichier .ps1.
    $exePath = $ExePath
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hScript = [System.BitConverter]::ToString($sha.ComputeHash([System.IO.File]::ReadAllBytes($MainScript))).Replace("-", "")
    $recompiler = "RECOMPILER : Invoke-PS2EXE -InputFile Apercu-Maquette.ps1 ... (README, section Compilation de l'EXE)"
    if (-not (Test-Path -LiteralPath $exePath)) {
        Check "EXE distribue present" $false ((Split-Path $exePath -Leaf) + " absent - " + $recompiler)
    } else {
        $exeBytes = [System.IO.File]::ReadAllBytes($exePath)
        $embarque = $null
        try {
            $asm = [System.Reflection.Assembly]::Load($exeBytes)
            $nom = @($asm.GetManifestResourceNames() | Where-Object { $_ -like "*.ps1" }) | Select-Object -First 1
            if ($nom) {
                $flux = $asm.GetManifestResourceStream($nom)
                $mem = New-Object System.IO.MemoryStream
                $flux.CopyTo($mem)
                $embarque = $mem.ToArray()
            }
        } catch { }
        if ($null -eq $embarque) {
            # Pas une derive : le binaire n'expose pas le script en clair (format different).
            Check "EXE : script embarque lisible" $false ("aucune ressource .ps1 en clair dans " + (Split-Path $exePath -Leaf) + " (binaire compresse ou construit autrement) : empreinte non verifiable, donc aucune recompilation a prescrire - verifier le format d'embarquement")
        } else {
            $hExe = [System.BitConverter]::ToString($sha.ComputeHash($embarque)).Replace("-", "")
            $okExe = ($hExe -eq $hScript)
            $detail = if ($okExe) { "sha256=" + $hScript.Substring(0, 12) + " (" + $embarque.Length + " octets embarques)" }
                      else {
                          # Meme contenu, seules les fins de ligne changent : ce n'est pas une derive.
                          $memeContenu = ((([System.Text.Encoding]::ASCII.GetString($embarque)) -replace "`r`n", "`n") -eq ($src -replace "`r`n", "`n"))
                          if ($memeContenu) {
                              "Contenu identique mais fins de ligne differentes (" + $embarque.Length + " octets embarques, script " + (Get-Item -LiteralPath $MainScript).Length + " octets) : le script est en LF ici, pas une derive - obtenir un checkout CRLF (.gitattributes, * text eol=crlf) et recompiler seulement si le contenu differe encore"
                          } else {
                              "L'EXE ne correspond plus au script (" + $hScript.Substring(0, 12) + " vs " + $hExe.Substring(0, 12) + ") - " + $recompiler
                          }
                      }
            Check "EXE a jour : empreinte du script embarque" $okExe $detail
        }

        # Autres proprietes du binaire livre : meme commande de compilation que le README.
        $vi = (Get-Item -LiteralPath $exePath).VersionInfo
        $cons = " - RECOMPILER avec Invoke-PS2EXE (README, section Compilation de l EXE)"
        # version.txt est la source unique : tout le reste s'y compare, EXE compris.
        $vSource = ""
        $cheminVersion = Join-Path $PSScriptRoot "version.txt"
        if (Test-Path -LiteralPath $cheminVersion) { $vSource = ([string](Get-Content -LiteralPath $cheminVersion -Raw)).Trim() }
        Check ("EXE : version " + $vSource) (($vSource -ne "") -and ($vi.FileVersion -eq $vSource) -and ($vi.ProductVersion -eq $vSource)) ("FileVersion=" + $vi.FileVersion + " ProductVersion=" + $vi.ProductVersion + " (version.txt=" + $vSource + " / -Version)" + $cons)

        # Une seule version : le script la porte pour le rapport du mode silencieux, le
        # badge l'affiche, le README et Compare-ExeRecompilation la repetent. Elles
        # doivent etre egales, sinon un bump oublie quelque part part en livraison.
        $vScript = ""
        if ($src -match '(?m)^\$script:versionOutil\s*=\s*"([0-9.]+)"') { $vScript = $matches[1] }
        $vBadge = ""
        $texteXaml = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "Maquette\Principal.xaml"))
        if ($texteXaml -match 'Text="([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)"') { $vBadge = $matches[1] }
        $vCompare = ""
        $texteCompare = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "Compare-ExeRecompilation.ps1"))
        if ($texteCompare -match 'Version\s*=\s*"([0-9.]+)"') { $vCompare = $matches[1] }
        $vReadme = ""
        $texteReadmeV = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "README.md"))
        if ($texteReadmeV -match '-Version\s+"([0-9.]+)"') { $vReadme = $matches[1] }
        $versionsV = @($vScript, $vBadge, $vCompare, $vReadme)
        $ecartsV = @($versionsV | Where-Object { $_ -ne $vSource })
        Check "version : script, badge, README et recompilation annoncent celle de version.txt" (($ecartsV.Count -eq 0) -and ($versionsV -notcontains "") -and ($vSource -ne "")) ("script=" + $vScript + " badge=" + $vBadge + " README=" + $vReadme + " compare=" + $vCompare + " version.txt=" + $vSource)
        # Le badge de la fenetre annonce la version du binaire : deux sources, donc un
        # controle. Le badge est lu dans le XAML livre, pas dans une constante du test.
        $badge = if ($xaml -match '<TextBlock Text="([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)"\s+FontSize="11"') { $matches[1] } else { "(absent)" }
        Check "XAML : badge de version = version de version.txt" ($badge -eq $vSource) ("badge=" + $badge + " version.txt=" + $vSource + " - aligner le badge de Maquette\Principal.xaml sur -Version")
        # Le README annonce la version du programme a deux endroits : la commande de
        # compilation (-Version "...") et la prose ("version `...`, titre"). Toutes doivent
        # valoir celle du binaire livre. Les autres nombres a quatre parties du README ne
        # sont pas des versions du programme : l'ODT (16.0.x) et l'adresse du relais.
        $readmeCompil = if (Test-Path -LiteralPath (Join-Path $PSScriptRoot "README.md")) { [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "README.md")) } else { "" }
        $versionsReadme = @([regex]::Matches($readmeCompil, '(?i)version\s*["`]?([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $ecartsVersion = @($versionsReadme | Where-Object { $_ -ne $vSource })
        Check "README : versions annoncees = version de version.txt" (($versionsReadme.Count -gt 0) -and ($ecartsVersion.Count -eq 0)) ("README annonce=" + ($versionsReadme -join ", ") + " version.txt=" + $vSource + " - aligner le README (commande Invoke-PS2EXE et prose)")

        # Le nom du journal etait fige dans le XAML (Deploiement_20261122.log) : il
        # annoncait une date passee et aucun fichier reellement ecrit. Le champ porte un
        # nom et suit le fichier ecrit par le deploiement ou par l'export.
        $nomFige = [regex]::Match($texteXaml, 'Text="[^"]*Deploiement_[0-9]{8}\.log"').Success
        $champJournal = [regex]::Match($texteXaml, 'x:Name="JournalFichier"').Success
        $misesAJour = ([regex]::Matches($src, 'JournalFichier')).Count
        Check "journal : le nom de fichier n'est plus fige dans le XAML" ((-not $nomFige) -and $champJournal -and ($misesAJour -ge 2)) ("fige=" + $nomFige + " champ=" + $champJournal + " mises a jour=" + $misesAJour)

        # --- Une seule source : version.txt, tenue par Set-Version.ps1 --------------
        # Le script verifie l'alignement (code 0) et refuse d'ecrire des qu'un emplacement
        # diverge (code 1). Les deux sont eprouves sur une copie, jamais sur le depot.
        $bump = Join-Path $PSScriptRoot "Set-Version.ps1"
        $codeVerif = -1
        if (Test-Path -LiteralPath $bump) {
            $null = & $psHost -NoProfile -ExecutionPolicy Bypass -File $bump -Racine $PSScriptRoot 2>&1
            $codeVerif = $LASTEXITCODE
        }
        Check "version.txt : source unique, Set-Version confirme l'alignement" (($codeVerif -eq 0) -and ($vSource -ne "")) ("code=" + $codeVerif + " version.txt=" + $vSource)

        if (-not $Fast) {
            # La cible du bump est calculee : une version en dur ici ferait de ce test
            # une seconde source, et Set-Version la refuserait.
            $morceaux = $vSource.Split('.')
            $morceaux[3] = [string]([int]$morceaux[3] + 1)
            $vCible = $morceaux -join '.'

            $copieBump = Join-Path $Work "bump"
            if (Test-Path -LiteralPath $copieBump) { Remove-Item -LiteralPath $copieBump -Recurse -Force }
            New-Item -ItemType Directory -Path (Join-Path $copieBump "Maquette") -Force | Out-Null
            foreach ($f in @("version.txt", "Apercu-Maquette.ps1", "README.md", "Compare-ExeRecompilation.ps1", "Test-InstallOffice.ps1")) {
                Copy-Item -LiteralPath (Join-Path $PSScriptRoot $f) -Destination $copieBump
            }
            Copy-Item -LiteralPath (Join-Path $PSScriptRoot "Maquette\Principal.xaml") -Destination (Join-Path $copieBump "Maquette")

            # 1) Un emplacement divergent : refus, et le message nomme fichier, ligne, valeur.
            $xamlCopie = Join-Path $copieBump "Maquette\Principal.xaml"
            $texteX = [System.IO.File]::ReadAllText($xamlCopie)
            [System.IO.File]::WriteAllText($xamlCopie, ($texteX -replace 'Text="[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+"', 'Text="9.9.9.9"'))
            $sortieRefus = (& $psHost -NoProfile -ExecutionPolicy Bypass -File $bump -Racine $copieBump 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
            $codeRefus = $LASTEXITCODE
            Check "version : un emplacement divergent fait refuser le bump" (($codeRefus -eq 1) -and ($sortieRefus -match 'Principal\.xaml') -and ($sortieRefus -match '9\.9\.9\.9') -and ($sortieRefus -match 'ligne \d+')) ((('code=' + $codeRefus + " | " + $sortieRefus) -replace "\s+", " ").Trim())

            # 2) Aligne a nouveau, un bump ecrit partout.
            Copy-Item -LiteralPath (Join-Path $PSScriptRoot "Maquette\Principal.xaml") -Destination (Join-Path $copieBump "Maquette") -Force
            $null = & $psHost -NoProfile -ExecutionPolicy Bypass -File $bump -Racine $copieBump -Version $vCible 2>&1
            $codeBump = $LASTEXITCODE
            $vues = @(
                [regex]::Match([System.IO.File]::ReadAllText((Join-Path $copieBump "version.txt")), '[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+').Value
                [regex]::Match([System.IO.File]::ReadAllText((Join-Path $copieBump "Apercu-Maquette.ps1")), 'versionOutil\s*=\s*"([0-9.]+)"').Groups[1].Value
                [regex]::Match([System.IO.File]::ReadAllText((Join-Path $copieBump "Maquette\Principal.xaml")), 'Text="([0-9.]+)"').Groups[1].Value
                [regex]::Match([System.IO.File]::ReadAllText((Join-Path $copieBump "Compare-ExeRecompilation.ps1")), 'Version\s*=\s*"([0-9.]+)"').Groups[1].Value
                [regex]::Match([System.IO.File]::ReadAllText((Join-Path $copieBump "README.md")), '-Version\s+"([0-9.]+)"').Groups[1].Value
            )
            Check "version : le bump ecrit version.txt et les quatre copies" (($codeBump -eq 0) -and (@($vues | Where-Object { $_ -ne $vCible }).Count -eq 0)) ("code=" + $codeBump + " cible=" + $vCible + " vues=" + ($vues -join ", "))

            # L'exemple de tag du README suit la meme regle que le workflow de
            # publication : la version sans son dernier ".0" (v8.5.3 pour 8.5.3.0).
            # Un bump qui le laisse perime ferait poser un tag que la publication
            # refuserait, apres coup et sans que la suite locale l'ait vu.
            $tagAttendu = "v" + ($vCible -replace '\.0$', '')
            $tagEcrit = [regex]::Match([System.IO.File]::ReadAllText((Join-Path $copieBump "README.md")), '`(v[0-9]+\.[0-9]+\.[0-9]+(?:\.[0-9]+)?)` pour ').Groups[1].Value
            Check "version : le bump met aussi a jour l'exemple de tag du README" (($codeBump -eq 0) -and ($tagEcrit -eq $tagAttendu)) ("code=" + $codeBump + " tag=" + $tagEcrit + " attendu=" + $tagAttendu)
        }
        # --- Demarrage de l'EXE : chemin du script nul ---------------------------
        # PS2EXE ne renseigne pas $MyInvocation.MyCommand.Path. Split-Path -Parent $null
        # levait "Impossible de lier l'argument au parametre Path" au demarrage : la
        # resolution est isolee dans Get-DossierScript, qui ne jette plus.
        $apercuTexte = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "Apercu-Maquette.ps1"))
        $lignesApercu = @($apercuTexte -split "`r?`n")
        # CurrentMajorVersionNumber vaut 10 sur Windows 11 : le nom affiche doit venir de
        # la build, sinon un poste 26300 s'annonce "Windows 10 26H2 (26300)".
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Get-NomWindows")
        $os11 = Get-NomWindows -Majeure 10 -Build "26300" -Affichage "26H2"
        $os10 = Get-NomWindows -Majeure 10 -Build "19045" -Affichage "22H2"
        $osVieux = Get-NomWindows -Majeure 0 -Build "" -Affichage ""
        Check "pre-vol : Windows 11 est nomme 11 (la build tranche, pas CurrentMajorVersionNumber)" (($os11.Court -eq "11 26H2 (26300)") -and ($os11.Libelle -eq "Windows 11 26H2 (26300)") -and ($os11.Ok) -and ($os10.Court -eq "10 22H2 (19045)") -and (-not $osVieux.Ok) -and ($osVieux.Court -eq "inconnu")) ("11=" + $os11.Court + " | 10=" + $os10.Court + " | vieux=" + $osVieux.Court)

        # Decision : la joignabilite du KMS est un avertissement, jamais un refus. Un
        # poste hors du reseau du KMS doit pouvoir installer ; l'activation se rejoue
        # ensuite ("Activer seulement"). Seule une activation sans adresse bloque.
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Get-RefusDeploiement")
        $cfgKms = [pscustomobject]@{ ActivateOnly = $false ; DoActivate = $true ; KmsHost = "kms.ac-nancy.fr" ; UseCustomXML = $false ; CustomXMLPath = "" ; DepotEnabled = $false ; DepotPath = "" }
        $pvKms = [pscustomobject]@{ SetupPresent = $true ; AssezDEspace = $true ; LibreGo = 100 ; Os = @{ Ok = $true } ; Admin = $true ; Kms = @{ Test = @{ Ok = $false } } }
        $refusInjoignable = @(Get-RefusDeploiement $cfgKms $pvKms)
        $cfgSansAdresse = [pscustomobject]@{ ActivateOnly = $false ; DoActivate = $true ; KmsHost = "" ; UseCustomXML = $false ; CustomXMLPath = "" ; DepotEnabled = $false ; DepotPath = "" }
        $refusSansAdresse = @(Get-RefusDeploiement $cfgSansAdresse $pvKms)
        Check "pre-vol : un KMS injoignable avertit (orange) mais ne bloque pas le depart" (($refusInjoignable.Count -eq 0) -and (($refusSansAdresse -join " ") -match "sans adresse") -and ($src -match 'Set-LignePreVol \$Fenetre "Kms"\s+"avr"') -and ($src -match 'Add-Journal \$Fenetre "avr" \("Test TCP du KMS')) ("refus injoignable=" + $refusInjoignable.Count + " refus sans adresse=[" + ($refusSansAdresse -join " ; ") + "]")
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Get-DossierScript")
        $dNul = Get-DossierScript -Chemin ""
        $dVide = Get-DossierScript -Chemin $null
        $dReel = Get-DossierScript -Chemin (Join-Path $PSScriptRoot "Apercu-Maquette.ps1")
        Check "demarrage : un chemin de script nul ne leve plus, le dossier du binaire prend le relais" (($dNul -ne "") -and ($dNul -eq $dVide) -and ($dReel -eq $PSScriptRoot)) ("nul=" + $dNul + " vide=" + $dVide + " reel=" + $dReel)
        Check "demarrage : Split-Path ne recoit plus le chemin brut de MyInvocation" (-not ($apercuTexte -match 'Split-Path\s+-Parent\s+\$MyInvocation')) "le garde doit preceder l'appel, pas le suivre"

        # --- Journal et polices : un demarrage ne peut plus rester muet -----------
        # Une police qui n'a pas pu etre creee rendait $null, et l'affectation a
        # FontFamily levait "Exception lors de la definition de FontFamily" : la ligne
        # n'entrait jamais au journal, qui restait vide. Les deux gardes ci-dessous
        # rendent une police de repli et n'affectent jamais $null.
        Add-Type -AssemblyName PresentationFramework | Out-Null
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "New-PoliceSure")
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Set-Police")
        # Nom vide : la source demandee ne peut pas etre retenue, la police vient donc
        # du repli. Sans ce repli, WPF rend une police de source vide et le glyphe tombe
        # sur la police par defaut, sans que rien ne le dise.
        $policeRepli = New-PoliceSure ""
        $policeReelle = New-PoliceSure "Segoe UI Variable Text, Segoe UI"
        $repliSource = if ($null -eq $policeRepli) { "nul" } else { [string]$policeRepli.Source }
        Check "journal : un nom de police vide rend une police de repli, jamais nulle" (($null -ne $policeRepli) -and ($policeRepli -is [System.Windows.Media.FontFamily]) -and ($repliSource -ne "") -and ($null -ne $policeReelle)) ("repli=[" + $repliSource + "] reelle=" + $(if ($null -eq $policeReelle) { "nul" } else { $policeReelle.Source }))
        $sansErreurPolice = $true
        try { Set-Police (New-Object System.Windows.Controls.TextBlock) $null } catch { $sansErreurPolice = $false }
        Check "journal : une police nulle n'est jamais affectee (pas d'exception FontFamily)" $sansErreurPolice "Set-Police a leve sur une police nulle"

        # Le journal doit s'ecrire meme si le panneau manque : c'est ce qui rendait un
        # "journal vide" sur une fenetre dont un nom avait change.
        $script:journalTexte = New-Object System.Collections.Generic.List[string]
        $script:journalLignes = 0
        $baseHorloge = $null
        $pinceaux = New-Object System.Windows.Media.BrushConverter
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Get-Pinceau")
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Add-Journal")
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Get-JournalTexte")
        $fausseFenetre = New-Object PSObject
        Add-Member -InputObject $fausseFenetre -MemberType ScriptMethod -Name FindName -Value { param($n) return $null }
        $journalOk = $true
        try { Add-Journal $fausseFenetre "info" "ligne de preuve" } catch { $journalOk = $false }
        $texteJournal = Get-JournalTexte $fausseFenetre
        Check "journal : une ligne s'ecrit sans panneau et l'export la reprend" ($journalOk -and ($texteJournal -match "ligne de preuve")) ("exception=" + (-not $journalOk) + " texte=[" + $texteJournal + "]")

        # Un etat d'etape qu'une autre version du worker nomme autrement tombait sur
        # $null : "expression Null" a chaque battement, et plus rien au journal.
        $script:etapeEncours = 0
        $policeTexte = $policeReelle
        $policeIcone = $policeReelle
        $renduEtapes = @{
            attente = @{ Fond = "#EDEDED" ; Texte = "#8A8A8A" }
            encours = @{ Fond = "#0F6CBD" ; Texte = "#FFFFFF" }
            termine = @{ Fond = "#DFF6DD" ; Texte = "#0F7B0F" }
            saute   = @{ Fond = "#F0F0F0" ; Texte = "#8A8A8A" }
            erreur  = @{ Fond = "#FDE7E9" ; Texte = "#C42B1C" }
        }
        $glyphes = @{ termine = "v" ; saute = "s" ; erreur = "!" }
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Set-Etape")
        $puce = New-Object System.Windows.Controls.Border
        $glyphe = New-Object System.Windows.Controls.TextBlock
        $titre = New-Object System.Windows.Controls.TextBlock
        $info = New-Object System.Windows.Controls.TextBlock
        $fausseEtape = New-Object PSObject
        Add-Member -InputObject $fausseEtape -MemberType ScriptMethod -Name FindName -Value { param($n) switch ($n) { "Etape1" { return $puce } "Glyphe1" { return $glyphe } "Titre1" { return $titre } "Etat1" { return $info } default { return $null } } }
        Add-Member -InputObject $fausseEtape -MemberType ScriptMethod -Name FindResource -Value { param($n) return "libelle" }
        $etapeOk = $true
        try { Set-Etape $fausseEtape 1 "etat-inconnu" } catch { $etapeOk = $false }
        Check "etape : un etat inconnu retombe sur attente au lieu de lever" ($etapeOk -and ($glyphe.Text -eq "1")) ("exception=" + (-not $etapeOk) + " texte=[" + $glyphe.Text + "]")

        # --- Chemins de travail : un dossier vide n'atteint plus Join-Path --------
        # C'est le parametre Path nul de la fenetre d'erreur : le marqueur de relais et
        # la configuration ne doivent jamais construire un chemin a partir de rien.
        Invoke-Expression (Get-FunctionText -Lines $lines -Name "Get-DossierTravail")
        Invoke-Expression (Get-FunctionText -Lines $lines -Name "Get-RelayMarkerPath")
        $LOG_FOLDER = ""
        $sauveLocal = $env:LOCALAPPDATA
        $env:LOCALAPPDATA = ""
        $marqueurVide = ""
        $marqueurOk = $true
        try { $marqueurVide = Get-RelayMarkerPath } catch { $marqueurOk = $false }
        $env:LOCALAPPDATA = $sauveLocal
        Check "reliquat : un dossier de travail vide ne lie plus un chemin nul" ($marqueurOk -and ($marqueurVide -ne "") -and ($marqueurVide -match 'relay-port\.txt$')) ("exception=" + (-not $marqueurOk) + " chemin=[" + $marqueurVide + "]")
        Invoke-Expression (Get-FunctionText -Lines $lines -Name "Get-Config")
        Invoke-Expression (Get-FunctionText -Lines $lines -Name "Save-Config")
        $CONFIG_FILE = ""
        $lectureVide = $null
        $configOk = $true
        try { $lectureVide = Get-Config ; Save-Config ([pscustomobject]@{ KmsServer = "x" }) } catch { $configOk = $false }
        Check "configuration : un chemin vide n'est ni lu ni ecrit, sans lever" ($configOk -and ($null -eq $lectureVide)) ("exception=" + (-not $configOk))

        # --- Cle produit : l'exemple du champ n'est pas une cle -----------------
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Get-ValeurSaisie")
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Get-CleSaisie")
        $exempleCle = "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
        Check "cle produit : l'exemple n'est jamais transmis, meme sans placeholder releve" (((Get-CleSaisie -Texte $exempleCle -Placeholder $exempleCle) -eq "") -and ((Get-CleSaisie -Texte $exempleCle -Placeholder "") -eq "")) ("avec=" + (Get-CleSaisie -Texte $exempleCle -Placeholder $exempleCle) + " sans=" + (Get-CleSaisie -Texte $exempleCle -Placeholder ""))
        Check "cle produit : une vraie cle reste intacte, espaces retires" ((Get-CleSaisie -Texte " abcde-12345-fghij-67890-klmno " -Placeholder $exempleCle) -eq "abcde-12345-fghij-67890-klmno") ("rendu=" + (Get-CleSaisie -Texte " abcde-12345-fghij-67890-klmno " -Placeholder $exempleCle))

        # Un exemple n'est ni une adresse ni une cle : il s'efface au focus, revient si
        # le champ reste vide, et n'est jamais transmis. Une vraie adresse identique au
        # texte de l'exemple reste une adresse des que le champ a ete touche.
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Get-ExempleChamp")
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Get-Pinceau")
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Register-ChampExemple")
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Set-ChampReel")
        if ($null -eq $pinceaux) { $pinceaux = New-Object System.Windows.Media.BrushConverter }
        $exempleKms = "kms.ac-nancy.fr"
        $champEssai = New-Object System.Windows.Controls.TextBox
        $champEssai.Text = $exempleKms
        Register-ChampExemple -Champ $champEssai
        $exempleJamais = ((Get-ValeurSaisie -Texte $exempleKms -Exemple (Get-ExempleChamp $champEssai)) -eq "")
        [void]$champEssai.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.UIElement]::GotFocusEvent)))
        $apresFocus = $champEssai.Text
        [void]$champEssai.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.UIElement]::LostFocusEvent)))
        $apresSortie = $champEssai.Text
        Set-ChampReel -Champ $champEssai -Texte $exempleKms
        $vraieIdentique = ((Get-ValeurSaisie -Texte $champEssai.Text -Exemple (Get-ExempleChamp $champEssai)) -eq $exempleKms)
        Check "champs KMS : l'exemple s'efface, revient et n'est jamais transmis" (($apresFocus -eq "") -and ($apresSortie -eq $exempleKms) -and $exempleJamais -and (([regex]::Matches($src, "Get-ValeurSaisie")).Count -ge 4)) ("focus=[" + $apresFocus + "] sortie=[" + $apresSortie + "] exemple transmis=" + (-not $exempleJamais))
        # Meme texte que l'exemple, mais venu d'une saisie ou de la configuration : le
        # champ est marque comme touche, donc ce n'est plus un exemple.
        Set-ChampReel -Champ $champEssai -Texte $exempleKms
        $vraieIdentique = ((Get-ValeurSaisie -Texte $champEssai.Text -Exemple (Get-ExempleChamp $champEssai)) -eq $exempleKms)
        Check "champs KMS : une vraie adresse identique a l'exemple reste valable" $vraieIdentique ("valeur=[" + $champEssai.Text + "] exemple=[" + (Get-ExempleChamp $champEssai) + "] rendu=[" + (Get-ValeurSaisie -Texte $champEssai.Text -Exemple (Get-ExempleChamp $champEssai)) + "]")

        # Premier lancement : sans adresse reelle, l'activation est decochee et le pied
        # ne promet pas un depart que le clic refuserait.
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Get-ExempleChamp")
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Disable-ActivationSansAdresse")
        Invoke-Expression (Get-FunctionText -Lines $lignesApercu -Name "Update-PiedStatut")
        $chkEssai = New-Object System.Windows.Controls.CheckBox ; $chkEssai.IsChecked = $true
        $tbEssai = New-Object System.Windows.Controls.TextBox ; $tbEssai.Text = "kms.ac-nancy.fr"
        $piedEssai = New-Object System.Windows.Controls.TextBlock
        if ($null -eq $script:exemples) { $script:exemples = @{} }
        $script:exemples[$tbEssai] = @{ Texte = "kms.ac-nancy.fr" }
        $fauxKms = New-Object PSObject
        Add-Member -InputObject $fauxKms -MemberType ScriptMethod -Name FindName -Value { param($n) if ($n -eq "ChkKms") { return $script:chkEssai } elseif ($n -eq "ChampKms") { return $script:tbEssai } elseif ($n -eq "PiedStatut") { return $script:piedEssai } else { return $null } }
        Add-Member -InputObject $fauxKms -MemberType ScriptMethod -Name FindResource -Value { param($n) if ($n -eq "PiedKmsAdresse") { return "KMS adresse requise" } else { return "Pret" } }
        Update-PiedStatut $fauxKms
        $piedAvant = $piedEssai.Text
        $decoche = Disable-ActivationSansAdresse $fauxKms
        $apresGuard = [bool]$chkEssai.IsChecked
        $tbEssai.Text = "10.0.0.5:1688"
        $chkEssai.IsChecked = $true
        $gardeReelle = Disable-ActivationSansAdresse $fauxKms
        $resteCoche = [bool]$chkEssai.IsChecked
        $journalKms = (@($script:journalTexte) -join "`n")
        $xamlDefaut = [regex]::Match($texteXaml, 'x:Name="ChkKms"[^>]*IsChecked="False"').Success
        Check "champs KMS : sans adresse reelle, l'activation est decochee au demarrage et le pied le dit" ($xamlDefaut -and ($piedAvant -eq "KMS adresse requise") -and $decoche -and (-not $apresGuard) -and (-not $gardeReelle) -and $resteCoche -and ($journalKms -match "Activation KMS decochee")) ("xaml=" + $xamlDefaut + " pied=[" + $piedAvant + "] decoche=" + $decoche + " reste coche=" + $resteCoche)
        # La vraie fenetre, hors ecran : focus sur le champ, collage, relecture. Le texte du
        # constat est ecrit par le script lui-meme, pas par le harnais.
        $preuveCle = Join-Path $Work "preuve_champ_cle.txt"
        $eapCle = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try { & $psHost -NoProfile -STA -ExecutionPolicy Bypass -File $MainScript -VerifierCle $preuveCle | Out-Null ; $codePreuveCle = $LASTEXITCODE }
        finally { $ErrorActionPreference = $eapCle }
        $textePreuveCle = if (Test-Path -LiteralPath $preuveCle) { [System.IO.File]::ReadAllText($preuveCle) } else { "" }
        Check "cle produit : la vraie fenetre efface l'exemple au focus et prend le collage" (($codePreuveCle -eq 0) -and ($textePreuveCle -match "texte apres focus=\[\]") -and ($textePreuveCle -match "texte apres collage=\[abcde-12345-fghij-67890-klmno\]")) ("code=" + $codePreuveCle + " ; " + ($textePreuveCle -replace "\s+", " "))

        # L'Office Deployment Tool a sa propre version, annoncee dans le README juste apres
        # son nom : elle doit valoir celle du setup.exe livre, sans quoi la prose decrit un
        # autre outil que le binaire joint au depot. La lecture suit le nom, pas la ligne
        # (une mention peut etre coupee par un retour a la ligne), et ne retient que le
        # premier nombre qui le suit : un autre nombre de la meme ligne (relais, version du
        # programme) n'est pas pris pour celui de l'ODT. Les deux mentions livrees annoncent
        # la version a 3 et 27 caracteres sans chiffre du nom ; 80 laisse passer un retour a
        # la ligne sans atteindre le nombre suivant d'une autre phrase.
        $setupPath = Join-Path $PSScriptRoot "setup.exe"
        $versionOdt = if (Test-Path -LiteralPath $setupPath) { (Get-Item -LiteralPath $setupPath).VersionInfo.FileVersion } else { "(setup.exe absent)" }
        $versionsOdt = @([regex]::Matches($readmeCompil, '(?i)Office\s+Deployment\s+Tool[^0-9]{0,80}?([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $ecartsOdt = @($versionsOdt | Where-Object { $_ -ne $versionOdt })
        Check "README : version ODT annoncee = version de setup.exe" (($versionsOdt.Count -gt 0) -and ($ecartsOdt.Count -eq 0)) ("README annonce=" + ($versionsOdt -join ", ") + " setup.exe=" + $versionOdt + " - aligner le README sur le setup.exe livre")
        # Complement des deux controles precedents : ils n'ancrent que les nombres annonces
        # comme versions. Tout nombre a quatre parties du README doit donc etre l'une de ces
        # deux versions ou une adresse de bouclage (le relais local 127.0.0.0/8) ; un nombre
        # ni l'un ni l'autre est une valeur que rien ne confronte a un fichier livre.
        # Chaque nombre hors controle est annonce avec sa ligne : une derive se corrige
        # sans rechercher la valeur dans le README. Un nombre present sur plusieurs lignes
        # est signale a chacune d'elles.
        $horsControle = @()
        $lignesReadme = $readmeCompil -split "`r?`n"
        for ($n = 0; $n -lt $lignesReadme.Count; $n++) {
            foreach ($m in [regex]::Matches($lignesReadme[$n], '[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+')) {
                if (($m.Value -ne $vi.FileVersion) -and ($m.Value -ne $versionOdt) -and ($m.Value -notmatch '^127\.')) {
                    $horsControle += ("ligne " + ($n + 1) + " : " + $m.Value)
                }
            }
        }
        Check "README : nombres a quatre parties = versions annoncees ou bouclage" ($horsControle.Count -eq 0) ("hors controle=" + ($horsControle -join ", ") + " (attendu : " + $vi.FileVersion + ", " + $versionOdt + " ou 127.0.0.0/8) - corriger la ligne ou retirer le nombre du README")
        Check "EXE : titre" ($vi.FileDescription -eq "Deploiement Office 2024") ("FileDescription=" + $vi.FileDescription + " (attendu : Deploiement Office 2024 / -Title)" + $cons)
        Check "EXE : produit et societe" (($vi.ProductName -eq "InstallOffice") -and ($vi.CompanyName -eq "Lycee Loritz")) ("ProductName=" + $vi.ProductName + " CompanyName=" + $vi.CompanyName + " (attendu : InstallOffice / -Product, Lycee Loritz / -Company)" + $cons)
        Check "EXE : OriginalFilename" ($vi.OriginalFilename -eq "Install-Office2024-KMS.exe") ("OriginalFilename=" + $vi.OriginalFilename + " (attendu : Install-Office2024-KMS.exe, donc compiler directement vers ce nom)" + $cons)
        Check "EXE : InternalName" ($vi.InternalName -eq "Install-Office2024-KMS.exe") ("InternalName=" + $vi.InternalName + " (attendu : Install-Office2024-KMS.exe, donc compiler directement vers ce nom)" + $cons)
        $asciiExe = [System.Text.Encoding]::ASCII.GetString($exeBytes)
        $niveau = if ($asciiExe -match 'requestedExecutionLevel[^>]*level\s*=\s*"(?<lvl>[A-Za-z]+)"') { $matches["lvl"] } else { "(absent)" }
        Check "EXE : manifeste requireAdmin" ($niveau -eq "requireAdministrator") ("requestedExecutionLevel=" + $niveau + " (attendu requireAdministrator / -RequireAdmin)" + $cons)
        # L'EXE dessine son interface depuis le dossier Maquette et extrait le worker du
        # moteur : les deux doivent etre livres avec lui, pas seulement presents ici.
        Check "livraison : fenetres XAML a cote de l'EXE" (Test-Path -LiteralPath $xamlPrincipal) ("Maquette\Principal.xaml absent : l'EXE ne peut pas dessiner son interface (livrer le dossier Maquette avec l'EXE)")
        # Meme regle pour l'Office Deployment Tool : sans lui, le programme s'ouvre puis
        # refuse tout deploiement. C'est ce qui manquait au ZIP telecharge.
        Check "livraison : setup.exe a cote de l'EXE" (Test-Path -LiteralPath $setupPath) "setup.exe absent : un dossier extrait sans l'Office Deployment Tool ne peut rien installer (livrer setup.exe avec l'EXE, README section Demarrage rapide)"
    }

    # Interface a l'execution : l'EXE livre (ou, sans droits administrateur, le script
    # qu'il embarque) ouvre les quatre fenetres et Styles.xaml hors ecran et resout
    # chaque nom que le script demande. La liste attendue vient du source : les noms
    # litteraux des FindName, plus les noms assembles dont les suffixes sont les
    # litteraux des appels (Ico/Val par ligne de pre-vol, Etape/Glyphe/Titre/Etat 1 a 5).
    # Un nom present dans le XAML mais place dans un modele ou un dictionnaire de
    # ressources ne se resout pas a l'execution : ni l'empreinte des sources ni le
    # controle statique ne peuvent le voir.
    $nomsAttendus = @($refsXaml | Where-Object { $_ -notmatch '\$' })
    $suffixesConnus = @{
        "Ico"    = @("Setup", "Disque", "Kms", "Windows")
        "Val"    = @("Setup", "Disque", "Kms", "Windows")
        "Etape"  = @("1", "2", "3", "4", "5")
        "Glyphe" = @("1", "2", "3", "4", "5")
        "Titre"  = @("1", "2", "3", "4", "5")
        "Etat"   = @("1", "2", "3", "4", "5")
    }
    $prefixesNonCouvers = @()
    foreach ($r in @($refsXaml | Where-Object { $_ -match '\$' })) {
        $p = ($r -split '\$')[0]
        if ($suffixesConnus.ContainsKey($p)) { foreach ($s in $suffixesConnus[$p]) { $nomsAttendus += ($p + $s) } }
        else { $prefixesNonCouvers += $p }
    }
    $nomsAttendus = @($nomsAttendus | Sort-Object -Unique)
    # La liste attendue est ecrite ici, le mode verificateur la lit et ecrit son rapport
    # dans un second fichier : les deux ne se recouvrent jamais.
    $listeInterface = Join-Path $Work "interface_noms.txt"
    $rapportInterface = Join-Path $Work "interface_rapport.txt"
    if (Test-Path -LiteralPath $rapportInterface) { Remove-Item -LiteralPath $rapportInterface -Force }
    [System.IO.File]::WriteAllLines($listeInterface, $nomsAttendus, (New-Object System.Text.UTF8Encoding($false)))
    $lignesInterface = @()
    $codeInterface = -1
    $lancePar = ""
    $noteIf = ""
    $estAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    # Le binaire livre porte un manifeste requireAdministrator : sans elevation, Windows
    # refuse de le demarrer (code 740). La CI et un poste administrateur lancent l'EXE ;
    # ailleurs, le script extrait du binaire est joue a sa place, memes XAML.
    # Lancement borne, sortie redirigee vers des fichiers. Le 24/11/2026, sur le runner
    # windows-latest, le lancement pipe (& ... 2>&1 | ForEach-Object) n'a jamais rendu la
    # main : le job a expire a 10 min sans dire ou il s'etait arrete. On attend le rapport
    # (ce que le controle verifie vraiment), avec une limite, et on arrete le processus qui
    # traine au lieu de laisser le job expirer.
    $delaiRapport = 180
    function Invoke-VerificationInterface {
        param([string]$Fichier, [string]$Arguments, [string]$Origine)
        # System.Diagnostics plutot que Start-Process : avec des redirections, Start-Process
        # rend un processus dont ExitCode est vide (verifie le 24/11/2026 en PowerShell 5.1).
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $Fichier
        $psi.Arguments = $Arguments
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        $p = [System.Diagnostics.Process]::Start($psi)
        # Lecture en tache de fond : un enfant bavard ne doit pas se bloquer sur un tube plein
        # pendant qu'on attend son rapport.
        $tacheSortie = $p.StandardOutput.ReadToEndAsync()
        $tacheErreur = $p.StandardError.ReadToEndAsync()
        $attente = 0.0
        while ((-not (Test-Path -LiteralPath $rapportInterface)) -and (-not $p.HasExited) -and ($attente -lt $delaiRapport)) {
            Start-Sleep -Milliseconds 500
            $attente += 0.5
        }
        $note = ""
        $code = -1
        if ($p.HasExited) {
            $code = $p.ExitCode
            $note = "sorti (code " + $code + ")"
        } elseif (Test-Path -LiteralPath $rapportInterface) {
            if ($p.WaitForExit(20000)) { $code = $p.ExitCode ; $note = "rapport ecrit, sorti (code " + $code + ")" }
            else { try { $p.Kill() } catch { } ; $note = "rapport ecrit, processus encore vivant apres 20 s, arrete" }
        } else {
            try { $p.Kill() } catch { }
            $note = "aucun rapport apres " + $delaiRapport + " s, processus arrete"
        }
        $txt = ""
        foreach ($t in @($tacheSortie, $tacheErreur)) {
            $v = ""
            try { $v = $t.Result.Trim() } catch { $v = "" }
            if ($v -ne "") { $txt += " " + $Origine + "=[" + $v + "]" }
        }
        return @{ Code = $code ; Note = $note ; Sortie = $txt }
    }
    if ($estAdmin -and (Test-Path -LiteralPath $exePath)) {
        $r = Invoke-VerificationInterface -Fichier $exePath -Arguments ("-VerifierInterface `"" + $listeInterface + "`" -RapportInterface `"" + $rapportInterface + "`"") -Origine "exe"
        $codeInterface = $r.Code
        $sortieIf = $r.Sortie
        $noteIf = $r.Note
        $lancePar = "EXE livre"
    }
    if (-not (Test-Path -LiteralPath $rapportInterface)) {
        $dossierVerif = Join-Path $Work "verif"
        New-Item -ItemType Directory -Path $dossierVerif -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot "Maquette") -Destination (Join-Path $dossierVerif "Maquette") -Recurse -Force
        $scriptVerif = Join-Path $dossierVerif "Apercu-Maquette.ps1"
        if ($null -ne $embarque) {
            [System.IO.File]::WriteAllBytes($scriptVerif, $embarque)
            $lancePar = "script embarque dans l'EXE"
        } else {
            Copy-Item -LiteralPath $MainScript -Destination $scriptVerif -Force
            $lancePar = "script de la copie de travail (EXE illisible)"
        }
        $r = Invoke-VerificationInterface -Fichier $psHost -Arguments ("-NoProfile -STA -ExecutionPolicy Bypass -File `"" + $scriptVerif + "`" -VerifierInterface `"" + $listeInterface + "`" -RapportInterface `"" + $rapportInterface + "`"") -Origine "script"
        $codeInterface = $r.Code
        $sortieIf = $r.Sortie
        $noteIf = $r.Note
    }
    if (Test-Path -LiteralPath $rapportInterface) { $lignesInterface = @([System.IO.File]::ReadAllLines($rapportInterface)) }
    $resumeIf = @($lignesInterface | Where-Object { $_ -match '^interface : (\d+)/(\d+) noms resolus' })
    $ecartsIf = @($lignesInterface | Where-Object { $_ -match '^(ABSENT|CHARGEMENT|OUVERTURE) ' })
    $fichiersIf = @($lignesInterface | Where-Object { $_ -match '^fichiers : 5/5 ' })
    $resolusIf = -1 ; $verifiesIf = -1
    if ($resumeIf.Count -eq 1) {
        if ($resumeIf[0] -match '^interface : (\d+)/(\d+) noms resolus') { $resolusIf = [int]$matches[1] ; $verifiesIf = [int]$matches[2] }
    }
    $okIf = ($codeInterface -eq 0) -and ($ecartsIf.Count -eq 0) -and ($fichiersIf.Count -eq 1) -and
            ($resolusIf -eq $verifiesIf) -and ($verifiesIf -eq $nomsAttendus.Count) -and
            ($nomsAttendus.Count -gt 0) -and ($prefixesNonCouvers.Count -eq 0)
    $detailIf = $lancePar + " (code " + $codeInterface + ") : " + $resolusIf + "/" + $verifiesIf + " noms resolus, " + $nomsAttendus.Count + " attendus ; cinq XAML charges=" + ($fichiersIf.Count -eq 1) + " ; rapport=" + $rapportInterface
    if ($ecartsIf.Count -gt 0) { $detailIf += " ; " + ($ecartsIf -join " ; ") }
    if ($prefixesNonCouvers.Count -gt 0) { $detailIf += " ; prefixe dynamique sans suffixe connu, etendre la liste du controle : " + ($prefixesNonCouvers -join ", ") }
    if ($noteIf -ne "") { $detailIf += " ; " + $noteIf }
    if (-not (Test-Path -LiteralPath $rapportInterface)) { $detailIf += " ; aucun rapport ecrit" + $sortieIf }
    Check "EXE livre : fenetres hors ecran, chaque nom d'interface resolu" $okIf $detailIf
    if ($lancePar -ne "EXE livre") {
        Write-Host ("         " + $lancePar + " joue a la place de l'EXE : sans droits administrateur, un binaire a manifeste requireAdministrator ne demarre pas (code 740). La CI et un poste administrateur lancent l'EXE.") -ForegroundColor Yellow
    }

    # Mode rapide : le hook pre-commit ne doit pas couter la suite entiere.
    # La section 1 est justement ce qui decide si l'EXE livre est a jour.
    if ($Fast) {
        Write-Host ""
        $totalRapide = $Pass + $Fail
        if ($Fail -eq 0) {
            Write-Host ("RESULTAT (rapide) : " + $Pass + "/" + $totalRapide + " controles passes.") -ForegroundColor Green
            exit 0
        }
        Write-Host ("RESULTAT (rapide) : " + $Fail + " echec(s) sur " + $totalRapide + " controles.") -ForegroundColor Red
        exit 1
    }

    # Les quatre captures livrees montrent le XAML livre. Deux preuves, dans un seul cas :
    #  - les sources (les cinq XAML et Apercu-Maquette.ps1) sont enregistrees a cote des PNG
    #    dans Maquette\captures.txt : une reference perimee est detectable meme la ou le
    #    rendu differe, aucune police n'entrant dans cette empreinte ;
    #  - les pixels doivent etre identiques, mais seulement la ou les polices resolues sont
    #    celles qui ont produit les references (rendu par le PowerShell 5.1 du systeme, en
    #    STA). Ailleurs la difference vient des polices, pas du XAML : elle est signalee
    #    sans faire echouer le controle.
    Section "1 bis. Captures de reference : identiques au pixel"
    $captures = @(
        @{ Nom = "apercu-principal.png"    ; Options = @() }
        @{ Nom = "apercu-dialogue.png"     ; Options = @("-Dialogue") }
        @{ Nom = "apercu-applications.png" ; Options = @("-Applications") }
        @{ Nom = "apercu-etats.png"        ; Options = @("-Etats") }
    )
    $empreinteSources = Get-EmpreinteSources -Racine $PSScriptRoot
    $empreinteRendu = Get-EmpreinteRendu
    $fiche = Join-Path $PSScriptRoot "Maquette\captures.txt"
    $enregistreSources = ""
    $enregistreRendu = ""
    if (Test-Path -LiteralPath $fiche) {
        foreach ($l in [System.IO.File]::ReadAllLines($fiche)) {
            if ($l -match '^sources\s+([0-9a-fA-F]{64})\s*$') { $enregistreSources = $matches[1].ToLowerInvariant() }
            elseif ($l -match '^rendu\s+([0-9a-fA-F]{64})\s*$') { $enregistreRendu = $matches[1].ToLowerInvariant() }
        }
    }
    $ecartsCapture = @()
    if ($enregistreSources -eq "") {
        $ecartsCapture += "Maquette\captures.txt absent ou sans empreinte des sources : les references ne peuvent pas etre datees"
    } elseif ($enregistreSources -ne $empreinteSources) {
        $ecartsCapture += ("sources modifiees depuis les references (enregistre " + $enregistreSources.Substring(0, 12) + ", actuel " + $empreinteSources.Substring(0, 12) + ")")
    }
    $renduComparable = ($enregistreRendu -ne "") -and ($enregistreRendu -eq $empreinteRendu)
    $rendusDifferents = @()
    foreach ($c in $captures) {
        $livre = Join-Path $PSScriptRoot ("Maquette\" + $c.Nom)
        if (-not (Test-Path -LiteralPath $livre)) { $ecartsCapture += ($c.Nom + " : PNG livre absent") ; continue }
        $rendu = Join-Path $Work ("rendu_" + $c.Nom)
        $arguments = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $MainScript, "-Capture", $rendu) + $c.Options
        # Comme pour le hook : le script de capture ecrit ses erreurs sur stderr, que
        # $ErrorActionPreference = Stop transformerait en erreur fatale du harnais (les
        # controles suivants ne seraient plus joues).
        $eapCapture = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        $sortie = (& $psHost @arguments 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
        $code = $LASTEXITCODE
        $ErrorActionPreference = $eapCapture
        if (-not (Test-Path -LiteralPath $rendu)) {
            # Rien n'a ete ecrit : le rendu a echoue (XAML ou script), ce n'est pas une
            # reference perimee et cela ne doit pas etre annonce comme tel.
            $derniere = @($sortie -split "`r?`n" | Where-Object { $_.Trim() -ne "" })[-1]
            $ecartsCapture += ($c.Nom + " : rendu impossible (code " + $code + " : " + $derniere + ")")
        } elseif ((Get-FileHash -LiteralPath $rendu -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $livre -Algorithm SHA256).Hash) {
            if ($renduComparable) { $ecartsCapture += ($c.Nom + " : different du rendu du XAML actuel") }
            else { $rendusDifferents += $c.Nom }
        }
    }
    if ($rendusDifferents.Count -gt 0) {
        Write-Host ("         rendu different sur ce poste (" + ($rendusDifferents -join ", ") + ") : polices autres que celles des references, comparaison au pixel non concluante") -ForegroundColor Yellow
    }
    $detailCaptures = if ($ecartsCapture.Count -eq 0) { "references a jour du XAML et rendu identique" }
                      else { ($ecartsCapture -join " ; ") + " - regenerer les quatre captures (memes commandes : -Capture, -Dialogue, -Applications, -Etats), lancer Test-InstallOffice.ps1 -EnregistrerCaptures, puis committer PNG et Maquette\captures.txt" }
    Check "captures : references a jour du XAML (empreinte) et rendu identique" ($ecartsCapture.Count -eq 0) $detailCaptures

    Section "1 ter. Configuration enregistree : secrets et relecture"
    # Les quatre fonctions sont extraites du moteur telles quelles, comme le worker : c'est
    # le code qui ecrit config.json sur un poste, pas une reecriture. $CONFIG_FILE est
    # redirige vers le dossier de travail : rien n'est ecrit dans le profil de l'utilisateur.
    # Tout se joue dans une portee enfant : les fonctions extraites et les variables d'essai
    # disparaissent a la sortie du bloc, sans laisser un Get-Config reel en portee de script
    # (la section 8 quater pose son propre Get-Config factice).
    & {
    $CONFIG_FILE = Join-Path $Work "appdata\Office2024Deploy\config.json"
    foreach ($nomFn in @("Get-Config", "Protect-SecretText", "Unprotect-SecretText", "Save-Config")) {
        Invoke-Expression (Get-FunctionText -Lines $lines -Name $nomFn)
    }
    $motDePasse = "MotDePasse-Essai-2024!"
    $cleProduit = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
    $kmsEssai   = "kms.exemple.invalid:1688"
    # Le dossier n'existe pas encore : Save-Config doit le creer, comme au premier lancement.
    Save-Config ([pscustomobject]@{
        KmsServer     = $kmsEssai
        ProductKeyEnc = (Protect-SecretText $cleProduit)
        ProxyUser     = "DOM\svc-install"
        ProxyPassEnc  = (Protect-SecretText $motDePasse)
    })
    $brutCfg = if (Test-Path -LiteralPath $CONFIG_FILE) { [System.IO.File]::ReadAllText($CONFIG_FILE) } else { "" }
    $reluCfg = Get-Config
    $mdpEnClair = $brutCfg.Contains($motDePasse)
    $cleEnClair = $brutCfg.Contains($cleProduit)
    $allerRetour = ($null -ne $reluCfg) -and ([string]$reluCfg.KmsServer -eq $kmsEssai) -and
                   ((Unprotect-SecretText ([string]$reluCfg.ProductKeyEnc)) -eq $cleProduit) -and
                   ((Unprotect-SecretText ([string]$reluCfg.ProxyPassEnc)) -eq $motDePasse)
    Check "config : aller-retour Save-Config/Protect-SecretText, aucun secret en clair" (($brutCfg.Length -gt 0) -and (-not $mdpEnClair) -and (-not $cleEnClair) -and $allerRetour) ("fichier=" + $CONFIG_FILE + " (" + $brutCfg.Length + " octets) mot de passe en clair=" + $mdpEnClair + " cle en clair=" + $cleEnClair + " relecture identique=" + $allerRetour)

    # Fichier tronque (ecriture interrompue, disque plein) : la relecture doit rendre null et
    # non lever, sinon l'interface ne demarre plus du tout pour un fichier de reglages.
    [System.IO.File]::WriteAllText($CONFIG_FILE, ('{"KmsServer": "' + $kmsEssai + '", '), (New-Object System.Text.UTF8Encoding($false)))
    $cfgCorrompu = Get-Config
    $renduCorrompu = if ($null -eq $cfgCorrompu) { "null" } else { $cfgCorrompu.GetType().Name }
    Check "config : fichier corrompu -> Get-Config rend null sans lever" ($null -eq $cfgCorrompu) ("fichier=" + $CONFIG_FILE + " tronque, rendu=" + $renduCorrompu + " (attendu null, sinon un fichier de reglages empeche le demarrage)")

    # Blob DPAPI d'un autre compte Windows (ou altere) : indechiffrable sous cette identite.
    # La relecture doit rendre une chaine vide, sans lever, et le reste du fichier rester lisible.
    $blobEtranger  = "01000000d08c9ddf0115d1118c7a00c04fc297eb" + ("0" * 192)
    $blobAleatoire = "0123456789abcdef" * 8
    $renduEtranger  = Unprotect-SecretText $blobEtranger
    $renduAleatoire = Unprotect-SecretText $blobAleatoire
    Save-Config ([pscustomobject]@{ KmsServer = $kmsEssai ; ProductKeyEnc = $blobEtranger })
    $reluEtranger = Get-Config
    $okEtranger = ($renduEtranger -eq "") -and ($renduAleatoire -eq "") -and ($null -ne $reluEtranger) -and ([string]$reluEtranger.KmsServer -eq $kmsEssai)
    Check "config : secret d'une autre identite -> relecture vide, config encore lisible" $okEtranger ("relecture=[" + $renduEtranger + "] blob aleatoire=[" + $renduAleatoire + "] KmsServer relu=[" + $(if ($null -eq $reluEtranger) { "(null)" } else { [string]$reluEtranger.KmsServer }) + "]")
    }

    # fixtures communes
    $fakeExe  = Join-Path $Work "fake_setup.exe"
    $argsFile = Join-Path $Work "args.txt"
    New-FakeSetup -Path $fakeExe
    $depot = Join-Path $Work "depot"
    New-Item -ItemType Directory -Path (Join-Path $depot "Office\Data") -Force | Out-Null
    Set-Content -Path (Join-Path $depot "Office\Data\v32.cab") -Value "x" -NoNewline
    $depotVide = Join-Path $Work "depot_vide"
    New-Item -ItemType Directory -Path $depotVide -Force | Out-Null
    $xmlSans = Join-Path $Work "xml_sans_sourcepath.xml"
    $xmlAvec = Join-Path $Work "xml_avec_sourcepath.xml"
    $destXml = Join-Path $Work "dest_xml"
    New-Item -ItemType Directory -Path (Join-Path $destXml "Office\Data") -Force | Out-Null
    Set-Content -Path (Join-Path $destXml "Office\Data\v64.cab") -Value "x" -NoNewline
    $xmlKo   = Join-Path $Work "xml_malforme.xml"
    New-XmlFixture -Path $xmlSans
    New-XmlFixture -Path $xmlAvec -SourcePath $destXml
    Set-Content -Path $xmlKo -Value "<Configuration><Add OfficeClientEdition=`"64`">" -Encoding UTF8
    $vbs = Join-Path $Work "ospp_factice.vbs"

    $body = Get-WorkerBody -Lines $lines
    $Worker = [scriptblock]::Create($body)
    $WorkerKms = [scriptblock]::Create($body.Replace('C:\Program Files\Microsoft Office\Office16\ospp.vbs', $vbs).Replace('C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs', $vbs))
    $env:FAKE_ODT_CODE = "0"
    $httpProxyAvant = [System.Environment]::GetEnvironmentVariable("HTTP_PROXY", "Process")

    Section "2. SourcePath au telechargement"
    $local = Join-Path $Work "dl_genere"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local) $argsFile
    $t = @($r.Logs) -join " "
    Check "XML genere : destination annoncee" ($t.Contains((Join-Path $local "Office"))) (($r.Logs | Where-Object { $_ -match 'Destination' }))
    Check "XML genere : /download recoit le XML genere" ($r.Args -match '/download .*OfficeConfig2024\.xml') $r.Args
    Check "XML genere : pas de copie inutile" (-not (Test-Path (Join-Path $local "OfficeConfig2024.sourcepath.xml"))) "copie creee"

    $local = Join-Path $Work "dl_sans"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -XmlPath $xmlSans) $argsFile
    $t = @($r.Logs) -join " "
    Check "XML sans SourcePath : copie avec SourcePath impose" (($t -match 'impose via une copie \(OfficeConfig2024\.sourcepath\.xml\) pour le telechargement') -and (Test-Path (Join-Path $local "OfficeConfig2024.sourcepath.xml"))) $t
    Check "XML sans SourcePath : /download recoit la copie" ($r.Args -match '/download .*OfficeConfig2024\.sourcepath\.xml') $r.Args
    Check "XML sans SourcePath : destination = dossier local" ($t.Contains((Join-Path $local "Office"))) (($r.Logs | Where-Object { $_ -match 'Destination' }))

    $local = Join-Path $Work "dl_avec"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -XmlPath $xmlAvec) $argsFile
    $t = @($r.Logs) -join " "
    Check "XML avec SourcePath : destination = celle du XML" ($t.Contains((Join-Path $destXml "Office"))) (($r.Logs | Where-Object { $_ -match 'Destination' }))
    Check "XML avec SourcePath : succes si les sources existent" (($r.Sync.StepState[2] -eq "done") -and ($t -match 'Sources telechargees avec succes')) ($r.Sync.StepState -join " ")

    Section "3. SourcePath a l'installation (mode depot)"
    $local = Join-Path $Work "in_sans"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -Depot $depot -XmlPath $xmlSans -Install $true) $argsFile
    $t = @($r.Logs) -join " "
    Check "XML sans SourcePath : copie pour l'installation" (($t -match "impose via une copie .* pour l'installation") -and (Test-Path (Join-Path $local "OfficeConfig2024.sourcepath.xml"))) (($r.Logs | Where-Object { $_ -match 'SourcePath' }))
    Check "XML sans SourcePath : /configure recoit la copie" ($r.Args -match '/configure .*OfficeConfig2024\.sourcepath\.xml') $r.Args
    Check "XML sans SourcePath : installation OK" ($r.Sync.StepState[3] -eq "done") ($r.Sync.StepState -join " ")

    $local = Join-Path $Work "in_genere"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -Depot $depot -Install $true) $argsFile
    Check "XML genere : /configure recoit le XML genere" ($r.Args -match '/configure .*OfficeConfig2024\.xml') $r.Args
    Check "XML genere : installation OK" ($r.Sync.StepState[3] -eq "done") ($r.Sync.StepState -join " ")

    # Le dossier choisi dans l'interface peut etre la racine du depot, le dossier Office
    # ou le dossier Data : ODT attend toujours <racine>\Office\Data. Un depot correct
    # choisi un niveau trop bas etait declare introuvable.
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "bg_ResolveDepotRoot")
    $dossierOffice = Join-Path $depot "Office"
    $dossierData   = Join-Path $depot "Office\Data"
    $rRacine = bg_ResolveDepotRoot $depot
    $rOffice = bg_ResolveDepotRoot $dossierOffice
    $rData   = bg_ResolveDepotRoot $dossierData
    $rVide   = bg_ResolveDepotRoot $depotVide
    $vue = { param($r) if ($null -eq $r) { "(nul)" } else { $r.Root } }
    $okResolu = ($null -ne $rRacine) -and ($null -ne $rOffice) -and ($null -ne $rData) -and ($null -eq $rVide) -and
                ($rRacine.Root -eq $depot) -and ($rOffice.Root -eq $depot) -and ($rData.Root -eq $depot)
    Check "depot : racine, dossier Office et dossier Data menent tous a Office\Data" $okResolu ("racine=" + (& $vue $rRacine) + " office=" + (& $vue $rOffice) + " data=" + (& $vue $rData) + " vide=" + (& $vue $rVide))

    $local = Join-Path $Work "in_dossier_office"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -Depot $dossierOffice -Install $true) $argsFile
    Check "depot choisi au niveau du dossier Office : installation OK" (($r.Sync.StepState[2] -eq "done") -and ($r.Sync.StepState[3] -eq "done")) ($r.Sync.StepState -join " ")

    # Partage UNC : le worker copie le dossier Office en local. La variante force la
    # branche "copie" sur un depot local (bg_TestDepotLocal rend false). Le XML genere a
    # l'etape 2 doit survivre a la copie : /MIR le supprimait avant l'installation.
    $WorkerCopie = [scriptblock]::Create($body.Replace('return ($di.DriveType -eq [System.IO.DriveType]::Fixed)', 'return $false'))
    $local = Join-Path $Work "in_copie"
    $r = Invoke-Deploy $WorkerCopie (New-Cfg (New-SyncHash) $fakeExe $local -Depot $depot -Install $true) $argsFile
    $xmlSurvit = Test-Path (Join-Path $local "OfficeConfig2024.xml")
    $journalCopie = @($r.Logs) -join " "
    # Les sources copiees sont nettoyees apres une installation reussie : le journal est
    # la preuve que la branche de copie a bien ete jouee et validee.
    $copieOk = ($journalCopie -match 'Copie du dossier Office du depot') -and ($journalCopie -match 'Sources validees : .*\\Office\\Data')
    Check "depot copie (partage) : le XML survit et l'installation le retrouve" ($copieOk -and $xmlSurvit -and ($r.Sync.StepState[2] -eq "done") -and ($r.Sync.StepState[3] -eq "done")) ("copie=" + $copieOk + " xml=" + $xmlSurvit + " etats=" + ($r.Sync.StepState -join " "))

    Section "4. Destination et espace disque"
    $fnSpace = Get-FunctionText -Lines $lines -Name "bg_SpaceVerdict"
    Invoke-Expression $fnSpace
    $insuff = bg_SpaceVerdict -FreeGB 2.4 -NeedGB 4 -Root "D:\"
    $attendu = "Espace insuffisant sur D:\ : 2.4 Go libres, environ 4 Go necessaires pour les sources. Le telechargement va echouer : liberez de l'espace ou choisissez un autre dossier de travail."
    Check "espace insuffisant : message exact" (($insuff.Level -eq "WARN") -and ($insuff.Msg -eq $attendu)) $insuff.Msg
    Check "espace limite (3.9 < 4)" ((bg_SpaceVerdict -FreeGB 3.9 -NeedGB 4 -Root "E:\").Level -eq "WARN") "OK attendu WARN"
    Check "espace egal (4 >= 4)" ((bg_SpaceVerdict -FreeGB 4 -NeedGB 4 -Root "C:\").Level -eq "OK") "OK attendu OK"
    Check "espace confortable" ((bg_SpaceVerdict -FreeGB 337.9 -NeedGB 4 -Root "C:\").Level -eq "OK") "OK attendu OK"
    $reseau = bg_SpaceVerdict -FreeGB $null -NeedGB 4 -Root $null
    Check "espace non verifiable (reseau)" (($reseau.Level -eq "WARN") -and ($reseau.Msg -match 'non verifiable')) $reseau.Msg

    $local = Join-Path $Work "es_reseau"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -XmlPath $xmlAvec) $argsFile
    $t = @($r.Logs) -join " "
    Check "worker : destination loguee + espace verifie" (($t -match 'Destination des sources :') -and ($t -match 'Espace ')) (($r.Logs | Where-Object { $_ -match 'Destination|Espace' }))

    Section "5. Codes de sortie ODT"
    $attenduSol = @(
        @{ Code = "32";    Motif = 'Fichier verrouille pendant l''installation' },
        @{ Code = "30066"; Motif = 'Une autre installation est en cours' },
        @{ Code = "17002"; Motif = 'Installation inachevee' },
        @{ Code = "17029"; Motif = 'Une version d''Office est deja installee' }
    )
    foreach ($cas in $attenduSol) {
        $env:FAKE_ODT_CODE = $cas.Code
        $local = Join-Path $Work ("code_" + $cas.Code)
        $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -Depot $depot -Install $true) $argsFile
        $sol = Solution $r
        Check ("code " + $cas.Code + " : message dedie") (($r.Sync.StepState[3] -eq "error") -and ($sol -match [regex]::Escape($cas.Motif)) -and ($r.Sync.PopupMsg -match [regex]::Escape($cas.Motif))) $sol
    }
    $env:FAKE_ODT_CODE = "0"
    $local = Join-Path $Work "code_aucun"
    $repertoire = Join-Path $Work "repertoire_setup"
    New-Item -ItemType Directory -Path $repertoire -Force | Out-Null
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $repertoire $local -Depot $depot -Install $true) $argsFile
    Check "setup.exe non demarrable : message dedie" ((($r.Sync.StepState[3] -eq "error")) -and ((Solution $r) -match "n'a pas pu demarrer")) (Solution $r)

    Section "6. Echecs d'activation KMS"
    $kms = @(
        @{ T = 'Error: 0xC004F074 No Key Management Service could be contacted'; M = 'Le serveur KMS ne repond pas' },
        @{ T = 'Error: 0x80070005 Access denied';                                  M = 'Acces refuse' },
        @{ T = 'Error: 0xC004F038 The count is not sufficient';                    M = 'seuil d''activation' },
        @{ T = 'Erreur inconnue du service de protection logicielle';              M = 'Erreur d''activation non identifiee' }
    )
    foreach ($cas in $kms) {
        Set-Content -Path $vbs -Value ('WScript.Echo "' + $cas.T + '"') -Encoding ASCII
        $local = Join-Path $Work ("kms_" + [guid]::NewGuid().ToString("N").Substring(0, 4))
        $r = Invoke-Deploy $WorkerKms (New-Cfg (New-SyncHash) $fakeExe $local -Activate $true) $argsFile
        $sol = Solution $r
        Check ("KMS : " + $cas.M) (($r.Sync.StepState[4] -eq "error") -and $r.Sync.HasError -and ($r.Sync.FinalMsg -match 'Activation KMS') -and ($sol -match [regex]::Escape($cas.M)) -and ($r.Sync.PopupMsg -match [regex]::Escape($cas.M))) $sol
    }
    $local = Join-Path $Work "kms_absent"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -Activate $true) $argsFile
    Check "KMS : ospp.vbs absent" (($r.Sync.StepState[4] -eq "error") -and ($r.Sync.PopupMsg -match 'ospp.vbs introuvable')) $r.Sync.PopupMsg

    Section "7. Garde-fous generaux"
    $local = Join-Path $Work "depot_ko"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -Depot $depotVide) $argsFile
    $msgDepotKo = ([string]$r.Sync.PopupMsg + " " + [string]$r.Sync.FinalMsg).Trim()
    Check "depot sans Office\Data : erreur, chemin et attente nommes, arret" (($r.Sync.StepState[2] -eq "error") -and ($r.Sync.FinalMsg -like "*interrompu*") -and ($msgDepotKo -match [regex]::Escape($depotVide)) -and ($msgDepotKo -match 'Office\\Data')) $msgDepotKo

    $local = Join-Path $Work "proxy_ko"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -Proxy $true) $argsFile
    $t = @($r.Logs) -join " "
    $netsh = @($r.Logs | Where-Object { $_ -match 'Proxy WinHTTP \(machine\) applique' }).Count -gt 0
    $envSame = ([System.Environment]::GetEnvironmentVariable("HTTP_PROXY", "Process") -eq $httpProxyAvant)
    Check "proxy injoignable : abandon avant netsh" (($r.Sync.StepState[2] -eq "error") -and (-not $netsh) -and $envSame) ("netsh=" + $netsh + " env=" + $envSame)
    Check "proxy injoignable : solution proxy/depot" (($t -match 'via le proxy') -and ($t -match 'Depot de sources')) (Solution $r)

    $local = Join-Path $Work "succes"
    $r = Invoke-Deploy $Worker (New-Cfg (New-SyncHash) $fakeExe $local -Depot $depot) $argsFile
    Check "depot valide : bilan de succes" ((-not $r.Sync.HasError) -and ($r.Sync.FinalMsg -like "*succes*")) $r.Sync.FinalMsg

    Section "8. Proxy applique puis restaure (machine simulee)"
    # Les trois fonctions sont extraites telles quelles ; seul leur acces systeme
    # (bg_ProxySink) est remplace : aucune ecriture sur le poste.
    $script:ProxyLogs = New-Object System.Collections.Generic.List[string]
    function bg_Log { param([string]$M,[string]$L = "INFO") $script:ProxyLogs.Add("[" + $L + "] " + $M) }
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "bg_ProxySink")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "bg_ApplyProxy")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "bg_RestoreProxy")

    $envInitial = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)
    $envInitial["HTTP_PROXY"] = "http://herite:8080" ; $envInitial["HTTPS_PROXY"] = ""
    $envInitial["http_proxy"] = "herite-min" ; $envInitial["https_proxy"] = ""
    $machine = @{ WinHttpDirect = $true ; WinHttpExists = $true ; WinHttpValue = [byte[]](7,7,7)
                  WinHttpServer = ""
                  WinInet = @{ ProxyEnable = 0 ; ProxyServer = "" ; ProxyOverride = "" }
                  Env     = $envInitial }
    function Copy-Etat { param($m)
        $envCopie = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)
        foreach ($n in @("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy")) { $envCopie[$n] = $m.Env[$n] }
        @{ WinHttpDirect = $m.WinHttpDirect ; WinHttpExists = $m.WinHttpExists ; WinHttpValue = $m.WinHttpValue
           WinInet = @{ ProxyEnable = $m.WinInet.ProxyEnable ; ProxyServer = $m.WinInet.ProxyServer ; ProxyOverride = $m.WinInet.ProxyOverride }
           Env = $envCopie } }
    $appels = New-Object System.Collections.Generic.List[string]
    $faux = @{
        Read = { Copy-Etat $machine }
        Apply = { param($srv,$byp)
                  $appels.Add("winhttp-set " + $srv + " bypass=" + $byp)
                  $machine.WinHttpServer = $srv
                  @{ Confirmed = $true ; Detail = "Proxy Server(s) : " + $srv ; Error = "" } }
        RestoreWinHttp = { param($st)
                  $appels.Add("winhttp-restore direct=" + $st.WinHttpDirect + " exists=" + $st.WinHttpExists + " valeur=" + (@($st.WinHttpValue) -join ","))
                  $machine.WinHttpServer = "" }
        WriteWinInet = { param($v)
                  $txt = "wininet"
                  foreach ($k in @($v.Keys | Sort-Object)) { $txt += " " + $k + "=" + $v[$k] }
                  $appels.Add($txt)
                  foreach ($k in @($v.Keys)) { $machine.WinInet[$k] = $v[$k] }
                  $true }
        WriteEnv = { param($v)
                  $txt = "env"
                  foreach ($k in @($v.Keys | Sort-Object)) { $txt += " " + $k + "=" + $v[$k] }
                  $appels.Add($txt)
                  foreach ($k in @($v.Keys)) { $machine.Env[$k] = $v[$k] } }
    }

    # etat reel du poste, releve avant tout
    $envReel = @{}
    foreach ($n in @("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy")) { $envReel[$n] = [System.Environment]::GetEnvironmentVariable($n, "Process") }
    $inetKeyReel = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
    function Get-InetReel {
        $v = @{}
        foreach ($p in @("ProxyEnable","ProxyServer","ProxyOverride")) { $v[$p] = $null ; try { $v[$p] = Get-ItemPropertyValue -Path $inetKeyReel -Name $p -ErrorAction Stop } catch {} }
        $v }
    $regAvant = Get-InetReel
    $netshAvant = ((& netsh winhttp show proxy 2>&1) -join " ").Trim()
    $avant = Copy-Etat $machine

    $cfgP = New-Cfg (New-SyncHash) $fakeExe (Join-Path $Work "proxy") -Proxy $true
    $cfgP.ProxySink = $faux
    $etat = bg_ApplyProxy $cfgP -ProxyHost "10.0.0.1" -ProxyPort "3128" -BypassList "<local>"
    $t = ($script:ProxyLogs -join " | ")

    # Les couches sont retrouvees par leur prefixe, jamais par leur rang : une etape
    # ajoutee dans bg_ApplyProxy ne decale plus ces trois controles (meme lecture que
    # pour la restauration plus bas). Un appel absent ou renomme reste signale.
    $aWh   = @($appels | Where-Object { $_ -like "winhttp-set*" }) | Select-Object -Last 1
    $aInet = @($appels | Where-Object { $_ -like "wininet*" })    | Select-Object -Last 1
    $aEnv  = @($appels | Where-Object { $_ -like "env *" })       | Select-Object -Last 1
    Check "apply : WinHTTP machine (netsh set proxy)" (($null -ne $aWh) -and ($aWh -eq "winhttp-set 10.0.0.1:3128 bypass=<local>")) $aWh
    Check "apply : WinINet session (ProxyEnable=1, serveur, exceptions)" (($null -ne $aInet) -and ($aInet -match 'ProxyEnable=1') -and ($aInet -match 'ProxyServer=10\.0\.0\.1:3128') -and ($aInet -match 'ProxyOverride=<local>')) $aInet
    Check "apply : 4 variables d'environnement" (($null -ne $aEnv) -and ($aEnv -match 'HTTP_PROXY=http://10\.0\.0\.1:3128') -and ($aEnv -match 'HTTPS_PROXY=http://10\.0\.0\.1:3128') -and ($aEnv -match 'http_proxy=http://10\.0\.0\.1:3128') -and ($aEnv -match 'https_proxy=http://10\.0\.0\.1:3128')) $aEnv
    Check "apply : journaux WinHTTP + WinINet" (($t -match 'Proxy WinHTTP \(machine\) applique : 10\.0\.0\.1:3128') -and ($t -match 'Proxy WinINet \(session\) applique')) $t
    Check "apply : les 3 couches sont effectivement modifiees" (($machine.WinHttpServer -eq "10.0.0.1:3128") -and ($machine.WinInet.ProxyEnable -eq 1) -and ($machine.WinInet.ProxyServer -eq "10.0.0.1:3128") -and ($machine.WinInet.ProxyOverride -eq "<local>") -and ($machine.Env.HTTP_PROXY -eq "http://10.0.0.1:3128") -and ($machine.Env.https_proxy -eq "http://10.0.0.1:3128")) ("winhttp=" + $machine.WinHttpServer + " inet=" + $machine.WinInet.ProxyEnable + " env=" + $machine.Env.HTTP_PROXY)

    bg_RestoreProxy $cfgP -State $etat
    $rEnv = @($appels | Where-Object { $_ -like "env *" }) | Select-Object -Last 1
    $rInet = @($appels | Where-Object { $_ -like "wininet*" }) | Select-Object -Last 1
    $rWh = @($appels | Where-Object { $_ -like "winhttp-restore*" }) | Select-Object -Last 1

    Check "restore : variables d'environnement d'origine" (($rEnv -match 'HTTP_PROXY=http://herite:8080') -and ($rEnv -match 'HTTPS_PROXY=( |$)') -and ($rEnv -match 'https_proxy=( |$)') -and ($rEnv -match 'http_proxy=herite-min')) $rEnv
    Check "restore : WinINet d'origine (ProxyEnable=0)" (($rInet -match 'ProxyEnable=0') -and ($rInet -match 'ProxyServer=( |$)') -and ($rInet -match 'ProxyOverride=( |$)')) $rInet
    Check "restore : WinHTTP remis en direct" (($rWh -match 'direct=True') -and ($rWh -match 'exists=True') -and ($rWh -match 'valeur=7,7,7')) $rWh
    Check "restore : etat simule identique a l'etat initial" ((Copy-Etat $machine | ConvertTo-Json -Compress -Depth 5) -eq ($avant | ConvertTo-Json -Compress -Depth 5)) (Copy-Etat $machine | ConvertTo-Json -Compress -Depth 5)
    Check "restore : journal de restauration" ((($script:ProxyLogs -join " | ") -match "Configuration proxy d'origine restauree")) ($script:ProxyLogs -join " | ")
    Check "restore : aucun appel netsh reel" (@($appels | Where-Object { $_ -match 'netsh' }).Count -eq 0) ($appels -join " // ")

    $regApres = Get-InetReel
    $envApres = @{}
    foreach ($n in @("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy")) { $envApres[$n] = [System.Environment]::GetEnvironmentVariable($n, "Process") }
    $netshApres = ((& netsh winhttp show proxy 2>&1) -join " ").Trim()
    $intact = (($regApres.ProxyEnable -eq $regAvant.ProxyEnable) -and ($regApres.ProxyServer -eq $regAvant.ProxyServer) -and ($regApres.ProxyOverride -eq $regAvant.ProxyOverride) -and ($envApres.HTTP_PROXY -eq $envReel.HTTP_PROXY) -and ($envApres.https_proxy -eq $envReel.https_proxy) -and ($netshApres -eq $netshAvant))
    Check "poste reel intact (HKCU, environnement, netsh)" $intact ("regle=" + $regApres.ProxyServer + " env=" + $envApres.HTTP_PROXY + " netsh=" + $netshApres)

    Section "8 bis. Relais proxy authentifie (proxy amont simule)"
    # Le service Click-to-Run ne peut pas repondre a un 407 : le relais ajoute
    # l'en-tete Proxy-Authorization et laisse passer le tunnel. On le teste contre
    # un proxy amont simule qui exige un Basic connu : aucune sortie reseau.
    $fakeProxyCs = @'
using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

public class FakeAuthProxy
{
    private readonly string _attendu;
    private TcpListener _l;
    private int _port;
    private int _refus;
    private int _tunnels;

    public FakeAuthProxy(string user, string password)
    {
        _attendu = "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(user + ":" + password));
    }

    // Arme, ce faux proxy ne propose que NTLM : le cas a nommer, pas a confondre
    // avec un mot de passe faux.
    public bool NegocieSeul;

    public int Port { get { return _port; } }
    public int Refus { get { return _refus; } }
    public int Tunnels { get { return _tunnels; } }

    public void Start()
    {
        _l = new TcpListener(IPAddress.Loopback, 0);
        _l.Start();
        _port = ((IPEndPoint)_l.LocalEndpoint).Port;
        Thread t = new Thread(Loop);
        t.IsBackground = true;
        t.Start();
    }

    public void Stop()
    {
        try { _l.Stop(); } catch { }
    }

    private void Loop()
    {
        while (true)
        {
            TcpClient c;
            try { c = _l.AcceptTcpClient(); } catch { return; }
            Thread t = new Thread(delegate() { Serve(c); });
            t.IsBackground = true;
            t.Start();
        }
    }

    private void Serve(TcpClient c)
    {
        try
        {
            NetworkStream s = c.GetStream();
            string entete = "";
            byte[] un = new byte[1];
            while (entete.Length < 8192)
            {
                int n = s.Read(un, 0, 1);
                if (n <= 0) { break; }
                entete += (char)un[0];
                if (entete.EndsWith("\r\n\r\n")) { break; }
            }
            bool ok = false;
            foreach (string ligne in entete.Split('\n'))
            {
                if (ligne.Trim().Equals("Proxy-Authorization: " + _attendu, StringComparison.OrdinalIgnoreCase)) { ok = true; }
            }
            if (!ok)
            {
                Interlocked.Increment(ref _refus);
                string defi = NegocieSeul ? "NTLM" : "Basic realm=\"proxy\"";
                byte[] ko = Encoding.ASCII.GetBytes("HTTP/1.1 407 Proxy Authentication Required\r\nProxy-Authenticate: " + defi + "\r\nContent-Length: 0\r\n\r\n");
                s.Write(ko, 0, ko.Length);
                s.Flush();
                return;
            }
            Interlocked.Increment(ref _tunnels);
            byte[] yy = Encoding.ASCII.GetBytes("HTTP/1.1 200 Connection established\r\n\r\n");
            s.Write(yy, 0, yy.Length);
            s.Flush();
            byte[] tampon = new byte[4096];
            int m;
            while ((m = s.Read(tampon, 0, tampon.Length)) > 0)
            {
                s.Write(tampon, 0, m);
                s.Flush();
            }
        }
        catch { }
        finally { try { c.Close(); } catch { } }
    }
}
'@
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Start-ProxyRelay")
    # Le mode silencieux passe par cette couche : sans elle, le refus viendrait d'une
    # fonction absente du harnais, pas du relais lui-meme.
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Start-AuthenticatedProxy")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Stop-ProxyRelay")
    $script:ProxyRelay = $null
    if (-not ("FakeAuthProxy" -as [type])) { Add-Type -TypeDefinition $fakeProxyCs -Language CSharp }

    function Read-EnteteProxy {
        param($Flux)
        $txt = ""
        $un = New-Object byte[] 1
        while ($txt.Length -lt 4096) {
            if ($Flux.Read($un, 0, 1) -le 0) { break }
            $txt += [char]$un[0]
            if ($txt.EndsWith("`r`n`r`n")) { break }
        }
        return $txt
    }
    function Connect-ViaRelais {
        param([int]$Port, [string]$Cible)
        $c = New-Object System.Net.Sockets.TcpClient
        $c.ReceiveTimeout = 8000
        $c.Connect("127.0.0.1", $Port)
        $s = $c.GetStream()
        $req = [System.Text.Encoding]::ASCII.GetBytes("CONNECT $Cible HTTP/1.1`r`nHost: $Cible`r`n`r`n")
        $s.Write($req, 0, $req.Length) ; $s.Flush()
        # Un CONNECT abouti compte : c'est la reference du nombre de tunnels dus.
        $rep = Read-EnteteProxy $s
        if ($rep -match ' 200 ') { $script:ConnectOk++ }
        @{ Client = $c ; Stream = $s ; Reponse = $rep }
    }

    $script:ConnectOk = 0
    $fauxAmont = New-Object FakeAuthProxy -ArgumentList @("eleve", "secret")
    $fauxAmont.Start()
    $relais = Start-ProxyRelay -UpstreamHost "127.0.0.1" -UpstreamPort $fauxAmont.Port -User "eleve" -Password "secret"
    $c1 = Connect-ViaRelais $relais.Port "officecdn.microsoft.com:443"
    $echange = ""
    if ($c1.Reponse -match ' 200 ') {
        $ping = [System.Text.Encoding]::ASCII.GetBytes("ping-relais")
        $c1.Stream.Write($ping, 0, $ping.Length) ; $c1.Stream.Flush()
        $buf = New-Object byte[] 64
        $n = 0
        try { $n = $c1.Stream.Read($buf, 0, $buf.Length) } catch {}
        if ($n -gt 0) { $echange = [System.Text.Encoding]::ASCII.GetString($buf, 0, $n) }
    }
    # Les tunnels ouverts chez le proxy amont doivent egaler les CONNECT que le harnais
    # a vus aboutir (200) : une fuite (tunnel ouvert sans reponse au client, ou tunnel
    # en double) se voit, sans figer de total en dur. Le plancher evite le cas 0 = 0.
    Check "relais proxy : CONNECT authentifie et tunnel bidirectionnel" (($c1.Reponse -match ' 200 ') -and ($echange -eq "ping-relais") -and ($script:ConnectOk -ge 1) -and ($fauxAmont.Tunnels -eq $script:ConnectOk)) (("reponse=" + ($c1.Reponse -replace "\s+", " ") + " echo=[" + $echange + "] tunnels=" + $fauxAmont.Tunnels + " connect_ok=" + $script:ConnectOk))
    try { $c1.Client.Close() } catch {}

    $relaisKo = Start-ProxyRelay -UpstreamHost "127.0.0.1" -UpstreamPort $fauxAmont.Port -User "eleve" -Password "faux"
    $c2 = Connect-ViaRelais $relaisKo.Port "officecdn.microsoft.com:443"
    Check "relais proxy : identifiants refuses (407 relaye)" (($c2.Reponse -match ' 407 ') -and ($relaisKo.Erreurs -ge 1)) (("reponse=" + ($c2.Reponse -replace "\s+", " ") + " erreurs=" + $relaisKo.Erreurs))
    try { $c2.Client.Close() } catch {}

    $portLibere = $relaisKo.Port
    Stop-ProxyRelay -Relais $relaisKo
    Stop-ProxyRelay -Relais $relais
    $fauxAmont.Stop()
    $ferme = $false
    try {
        $tc = New-Object System.Net.Sockets.TcpClient
        $tc.Connect("127.0.0.1", $portLibere)
        $ferme = -not $tc.Connected
        $tc.Close()
    } catch { $ferme = $true }
    Check "relais proxy : arret libere le port d'ecoute" $ferme ("port " + $portLibere + " encore ouvert=" + (-not $ferme))

    # --- Le controle precedent doit reellement mordre --------------------------
    # Variante fautive du relais, compilee ici : meme code, plus une seconde connexion
    # amont authentifiee par CONNECT client. Aucun reseau reel (le proxy amont simule
    # reste sur la loopback) et aucun total en dur : on exige seulement que l'egalite
    # tunnels = CONNECT du controle precedent echoue devant cette fuite.
    $fnFuite = (Get-FunctionText -Lines $lines -Name "Start-ProxyRelay")
    $fnFuite = $fnFuite.Replace("OfficeProxyRelay", "OfficeProxyRelayFuite").Replace("function Start-ProxyRelay", "function Start-ProxyRelayFuite")
    $ancreFuite = "            us.Flush();`n`n            string reponse = LireEntete(us);"
    $blocFuite = @(
        "            us.Flush();",
        "",
        "            // FUITE : une seconde connexion amont authentifiee par CONNECT client.",
        "            TcpClient fuite = new TcpClient();",
        "            fuite.Connect(_upHost, _upPort);",
        "            NetworkStream fs = fuite.GetStream();",
        "            fs.Write(rb, 0, rb.Length);",
        "            fs.Flush();",
        "            LireEntete(fs);",
        "",
        "            string reponse = LireEntete(us);"
    ) -join "`n"
    $ancreFuiteTrouvee = $fnFuite.Contains($ancreFuite)
    if ($ancreFuiteTrouvee) { Invoke-Expression ($fnFuite.Replace($ancreFuite, $blocFuite)) }

    $fauxFuite = New-Object FakeAuthProxy -ArgumentList @("eleve", "secret")
    $fauxFuite.Start()
    $script:ConnectOk = 0
    $cFuite = $null
    $relaisFuite = $null
    if ($ancreFuiteTrouvee) {
        $relaisFuite = Start-ProxyRelayFuite -UpstreamHost "127.0.0.1" -UpstreamPort $fauxFuite.Port -User "eleve" -Password "secret"
        $cFuite = Connect-ViaRelais $relaisFuite.Port "officecdn.microsoft.com:443"
    }
    $tunnelsFuite = $fauxFuite.Tunnels
    $connectFuite = $script:ConnectOk
    if ($null -ne $cFuite) { try { $cFuite.Client.Close() } catch {} }
    if ($null -ne $relaisFuite) { Stop-ProxyRelay -Relais $relaisFuite }
    $fauxFuite.Stop()
    Check "relais proxy : une fuite de tunnel fait echouer l'egalite tunnels/CONNECT" ($ancreFuiteTrouvee -and ($null -ne $cFuite) -and ($cFuite.Reponse -match ' 200 ') -and ($connectFuite -ge 1) -and ($tunnelsFuite -gt $connectFuite)) ("ancre=" + $ancreFuiteTrouvee + " tunnels=" + $tunnelsFuite + " connect_ok=" + $connectFuite)

    # --- Cablage cote deploiement : ProxyRelay transmis au worker -------------
    # Le worker teste le CDN AVANT de lancer setup.exe. Un proxy qui repond 407 doit
    # produire le message du relais quand la config porte ProxyRelay, et le message
    # historique sinon : c'est le champ de config qui choisit, pas une heuristique.
    $faux407 = New-Object FakeAuthProxy -ArgumentList @("eleve", "secret")
    $faux407.Start()

    $cfgRelais = New-Cfg (New-SyncHash) $fakeExe (Join-Path $Work "relais_407") -Proxy $true
    $cfgRelais.ProxyPort = [string]$faux407.Port
    $cfgRelais.ProxyRelay = $true
    $cfgRelais.ProxyRelayTarget = "10.0.0.1:3128"
    $r = Invoke-Deploy $Worker $cfgRelais $argsFile
    $t = @($r.Logs) -join " "
    $netsh = @($r.Logs | Where-Object { $_ -match 'Proxy WinHTTP \(machine\) applique' }).Count -gt 0
    Check "relais deploye : 407 -> message des identifiants, sans toucher au poste" (($r.Sync.StepState[2] -eq "error") -and ($t -match 'refuse les identifiants fournis \(407\)') -and ($t -match '10\.0\.0\.1:3128') -and ($t -match 'jamais enregistre') -and (-not $netsh)) ("netsh=" + $netsh + " | " + (Solution $r))

    $cfgSansRelais = New-Cfg (New-SyncHash) $fakeExe (Join-Path $Work "relais_407_sans") -Proxy $true
    $cfgSansRelais.ProxyPort = [string]$faux407.Port
    $r2 = Invoke-Deploy $Worker $cfgSansRelais $argsFile
    $t2 = @($r2.Logs) -join " "
    Check "relais deploye : sans ProxyRelay, message historique (LocalSystem)" (($r2.Sync.StepState[2] -eq "error") -and ($t2 -match 'exige une authentification \(407\)') -and ($t2 -match 'LocalSystem') -and ($t2 -notmatch 'refuse les identifiants fournis')) (Solution $r2)
    $faux407.Stop()

    # Un proxy qui n'offre que NTLM doit etre nomme : sinon l'administrateur cherche un
    # mot de passe faux alors que le relais ne peut pas negocier.
    # La copie d'interface est extraite comme Stop-ProxyRelay : le harnais ne charge
    # pas tout le moteur dans sa portee.
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Test-ProxyConnect")
    $fauxNtlm = New-Object FakeAuthProxy -ArgumentList @("eleve", "secret")
    $fauxNtlm.NegocieSeul = $true
    $fauxNtlm.Start()
    $sonde = Test-ProxyConnect -ProxyHost "127.0.0.1" -ProxyPort $fauxNtlm.Port -TargetHost "officecdn.microsoft.com" -TargetPort 443
    Check "proxy 407 NTLM : les schemas annonces sont lus" (($sonde.Code -eq 407) -and ($sonde.Schemes -match 'NTLM') -and ($sonde.Schemes -notmatch 'Basic')) ("code=" + $sonde.Code + " schemas=" + $sonde.Schemes)

    $cfgNtlm = New-Cfg (New-SyncHash) $fakeExe (Join-Path $Work "relais_ntlm") -Proxy $true
    $cfgNtlm.ProxyPort = [string]$fauxNtlm.Port
    $cfgNtlm.ProxyRelay = $true
    $cfgNtlm.ProxyRelayTarget = "10.0.0.1:3128"
    $r3 = Invoke-Deploy $Worker $cfgNtlm $argsFile
    $t3 = @($r3.Logs) -join " "
    Check "proxy 407 NTLM : le message nomme le schema et les issues" (($r3.Sync.StepState[2] -eq "error") -and ($t3 -match 'NTLM') -and ($t3 -match 'Depot de sources') -and ($t3 -notmatch 'refuse les identifiants fournis')) (Solution $r3)
    $fauxNtlm.Stop()

    Section "8 ter. Reliquat de relais au demarrage"
    # Un arret brutal laisse le relais loopback dans WinHTTP (machine) et/ou dans le
    # proxy de session (HKCU Internet Settings), alors que le port est mort. Le port du
    # relais est enregistre au lancement : seule cette valeur autorise une remise en
    # direct, pour ne jamais toucher un proxy local legitime. Seul l'acces systeme est
    # remplace (Get-ProxyRelaySink) ; le test de port reste un vrai test TCP loopback.
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Test-TcpPort")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Get-RelayMarkerPath")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Save-RelayPort")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Get-RelayPort")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Clear-RelayPort")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Get-ProxyRelaySink")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Get-WinHttpProxyAddress")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Get-InetProxyAddress")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Test-ProxyRelayResidue")
    Invoke-Expression (Get-FunctionText -Lines $lines -Name "Reset-ProxyRelayResidue")

    # Machine simulee : sortie netsh + etat HKCU, sans jamais toucher au poste.
    function New-SinkReliquat {
        param([string]$Netsh = "Acces direct (pas de serveur proxy).", [int]$ProxyEnable = 0,
              [string]$ProxyServer = "", [scriptblock]$Reset, [scriptblock]$Disable)
        $etat = @{ Netsh = $Netsh ; Inet = @{ ProxyEnable = $ProxyEnable ; ProxyServer = $ProxyServer ; ProxyOverride = "<local>" } }
        if ($null -eq $Reset) { $Reset = { $etat.Netsh = "Acces direct (pas de serveur proxy)." ; "Proxy reinitialise." }.GetNewClosure() }
        if ($null -eq $Disable) { $Disable = { $etat.Inet.ProxyEnable = 0 ; $true }.GetNewClosure() }
        @{ Sink = @{ Show = { $etat.Netsh }.GetNewClosure() ; Reset = $Reset ;
                     ReadInet = { $etat.Inet }.GetNewClosure() ; DisableInet = $Disable } ; Etat = $etat }
    }

    # Deux ports libres releves puis refermes : celui du relais et un autre.
    $sonde = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $sonde.Start()
    $portMort = $sonde.LocalEndpoint.Port
    $sonde.Stop()
    $sonde2 = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $sonde2.Start()
    $portAutre = $sonde2.LocalEndpoint.Port
    $sonde2.Stop()

    # 1. WinHTTP sur le port enregistre, qui ne repond plus.
    $s1 = New-SinkReliquat -Netsh "Proxy Server(s) :  127.0.0.1:$portMort"
    $r1 = Test-ProxyRelayResidue -Sink $s1.Sink -RelayPort $portMort -TimeoutMs 800
    Check "reliquat proxy : WinHTTP sur le port enregistre et mort detecte" ($r1.Found -and $r1.WinHttp -and (-not $r1.WinInet) -and ($r1.Port -eq $portMort)) ("found=" + $r1.Found + " winhttp=" + $r1.WinHttp + " wininet=" + $r1.WinInet + " port=" + $r1.Port)

    # 2. Port observe different du port enregistre : on ne touche a rien.
    $r2 = Test-ProxyRelayResidue -Sink $s1.Sink -RelayPort $portAutre -TimeoutMs 800
    Check "reliquat proxy : port different jamais signale" ((-not $r2.Found) -and ($r2.Reason -eq "port different") -and (-not $r2.WinHttp) -and (-not $r2.WinInet)) ("enregistre=" + $portAutre + " observe=" + $r2.Port + " -> found=" + $r2.Found + " raison=" + $r2.Reason)

    # 3. Aucun port enregistre : aucune proposition, meme devant un port mort.
    $r3 = Test-ProxyRelayResidue -Sink $s1.Sink -RelayPort 0 -TimeoutMs 500
    Check "reliquat proxy : sans port enregistre, aucun signalement" ((-not $r3.Found) -and ($r3.Reason -eq "aucun port enregistre")) ("found=" + $r3.Found + " raison=" + $r3.Reason)

    # 4. Relais encore vivant, ou acces direct : aucun signalement non plus.
    $vivant = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $vivant.Start()
    $portVivant = $vivant.LocalEndpoint.Port
    $sVivant = New-SinkReliquat -Netsh "Proxy Server(s) :  127.0.0.1:$portVivant"
    $r4 = Test-ProxyRelayResidue -Sink $sVivant.Sink -RelayPort $portVivant -TimeoutMs 800
    $vivant.Stop()
    $sDirect = New-SinkReliquat
    $r4b = Test-ProxyRelayResidue -Sink $sDirect.Sink -RelayPort $portMort -TimeoutMs 500
    Check "reliquat proxy : relais vivant ou acces direct non signale" ((-not $r4.Found) -and (-not $r4b.Found)) ("vivant=" + $portVivant + " -> " + $r4.Found + " | direct -> " + $r4b.Found)

    # 5. Le marqueur de port : ecrit au lancement, relu, puis efface.
    $marqueur = Join-Path $Work "relay-port.txt"
    Save-RelayPort -Port 4321 -Path $marqueur
    $lu = Get-RelayPort -Path $marqueur
    Clear-RelayPort -Path $marqueur
    $apres = Get-RelayPort -Path $marqueur
    Check "reliquat proxy : port du relais enregistre puis efface" (($lu -eq 4321) -and ($apres -eq 0) -and (-not (Test-Path -LiteralPath $marqueur))) ("relu=" + $lu + " apres effacement=" + $apres)

    # 6. Proxy de session seul (WinHTTP en direct), sur le port enregistre.
    $s6 = New-SinkReliquat -ProxyEnable 1 -ProxyServer "127.0.0.1:$portMort"
    $r6 = Test-ProxyRelayResidue -Sink $s6.Sink -RelayPort $portMort -TimeoutMs 800
    Check "reliquat proxy : proxy de session sur le port enregistre detecte" ($r6.Found -and $r6.WinInet -and (-not $r6.WinHttp) -and ($r6.Port -eq $portMort)) ("found=" + $r6.Found + " wininet=" + $r6.WinInet + " winhttp=" + $r6.WinHttp + " port=" + $r6.Port)

    # 7. Session sur un autre port loopback : jamais signalee.
    $s7 = New-SinkReliquat -ProxyEnable 1 -ProxyServer "127.0.0.1:$portAutre"
    $r7 = Test-ProxyRelayResidue -Sink $s7.Sink -RelayPort $portMort -TimeoutMs 800
    Check "reliquat proxy : proxy de session sur un autre port non signale" ((-not $r7.Found) -and (-not $r7.WinInet) -and ($r7.Reason -eq "port different")) ("enregistre=" + $portMort + " session=" + $portAutre + " -> found=" + $r7.Found + " raison=" + $r7.Reason)

    # 8. Session desactivee (ProxyEnable=0) : proxy ignore par Windows, rien a faire.
    $s8 = New-SinkReliquat -ProxyEnable 0 -ProxyServer "127.0.0.1:$portMort"
    $r8 = Test-ProxyRelayResidue -Sink $s8.Sink -RelayPort $portMort -TimeoutMs 800
    Check "reliquat proxy : proxy de session desactive non signale" ((-not $r8.Found) -and (-not $r8.WinInet)) ("found=" + $r8.Found + " raison=" + $r8.Reason)

    # 9. Remise en direct des deux couches a la fois.
    $s9 = New-SinkReliquat -Netsh "Proxy Server(s) :  127.0.0.1:$portMort" -ProxyEnable 1 -ProxyServer "127.0.0.1:$portMort"
    $fix9 = Reset-ProxyRelayResidue -Sink $s9.Sink -WinHttp $true -WinInet $true
    Check "reliquat proxy : remise en direct des deux couches" ($fix9.OK -and ($s9.Etat.Netsh -match 'Acces direct') -and ($s9.Etat.Inet.ProxyEnable -eq 0)) ("ok=" + $fix9.OK + " netsh=[" + $s9.Etat.Netsh + "] ProxyEnable=" + $s9.Etat.Inet.ProxyEnable)

    # 10. Une couche qui n'est pas en cause n'est pas touchee.
    $s10 = New-SinkReliquat -Netsh "Proxy Server(s) :  10.0.0.1:3128" -ProxyEnable 1 -ProxyServer "127.0.0.1:$portMort"
    $fix10 = Reset-ProxyRelayResidue -Sink $s10.Sink -WinHttp $false -WinInet $true
    Check "reliquat proxy : couche non concernee laissee intacte" ($fix10.OK -and ($s10.Etat.Netsh -match '10\.0\.0\.1:3128') -and ($s10.Etat.Inet.ProxyEnable -eq 0)) ("ok=" + $fix10.OK + " netsh=[" + $s10.Etat.Netsh + "] ProxyEnable=" + $s10.Etat.Inet.ProxyEnable)

    # 11. Remise en direct sans effet : l'echec est signale, jamais maquille en succes.
    $s11 = New-SinkReliquat -Netsh "Proxy Server(s) :  127.0.0.1:$portMort" -ProxyEnable 1 -ProxyServer "127.0.0.1:$portMort" -Reset { "acces refuse" } -Disable { $false }
    $fix11 = Reset-ProxyRelayResidue -Sink $s11.Sink -WinHttp $true -WinInet $true
    Check "reliquat proxy : echec de la remise en direct signale" ((-not $fix11.OK) -and ($fix11.Detail -match 'acces refuse')) ("ok=" + $fix11.OK + " detail=[" + $fix11.Detail + "]")

    # --- Lecture reelle, en lecture seule --------------------------------------
    # Tous les cas precedents remplacent le sink : une erreur dans la lecture reelle
    # passerait inapercue. On interroge donc le poste (netsh + HKCU), on confronte le
    # resultat a une lecture independante, et on verifie que rien n'a bouge.
    $sinkReel = Get-ProxyRelaySink
    $cleInet = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
    $regReel = @{ ProxyEnable = $null ; ProxyServer = $null ; ProxyOverride = $null }
    foreach ($p in @("ProxyEnable","ProxyServer","ProxyOverride")) {
        try { $regReel[$p] = Get-ItemPropertyValue -Path $cleInet -Name $p -ErrorAction Stop } catch {}
    }
    $avantNetsh = ([string](& $sinkReel.Show)).Trim()
    $avantInet = & $sinkReel.ReadInet
    # Appel netsh independant : la sortie du sink doit etre celle du poste, pas un texte
    # fabrique. Sur un poste en acces direct les valeurs sont des defauts, mais un sink
    # qui n'appellerait pas netsh, ou une cle lue au mauvais endroit, tombent ici.
    $netshDirect = (@(& netsh winhttp show proxy 2>&1) -join "`n").Trim()
    $okCleInet = ([string]$sinkReel.InetKey -eq $cleInet)

    # Appel sans -Sink : c'est le chemin par defaut du programme qui est exerce ici.
    $hReel = Get-WinHttpProxyAddress
    $iReel = Get-InetProxyAddress
    $hExplicite = Get-WinHttpProxyAddress -Sink $sinkReel
    $iExplicite = Get-InetProxyAddress -Sink $sinkReel

    $loopNetsh = ($avantNetsh -match '(?i)(127\.0\.0\.1|localhost)\s*:\s*\d')
    $okNetsh = ($avantNetsh.Length -gt 0) -and ($avantNetsh -eq $netshDirect)
    if ($loopNetsh) { $okNetsh = $okNetsh -and ($null -ne $hReel) -and ($avantNetsh -match (':' + $hReel.Port)) }
    else { $okNetsh = $okNetsh -and ($null -eq $hReel) }
    $memeWinHttp = (($null -eq $hReel) -and ($null -eq $hExplicite)) -or (($null -ne $hReel) -and ($null -ne $hExplicite) -and ($hReel.Port -eq $hExplicite.Port))
    $luWinHttp = if ($null -eq $hReel) { "aucun proxy loopback" } else { "$($hReel.Host):$($hReel.Port)" }
    Check "reliquat reel : WinHTTP lu sur le poste (netsh reel, aucun sink simule)" ($okNetsh -and $memeWinHttp) ("netsh sink=[" + ($avantNetsh -replace "\s+", " ") + "] identique au netsh direct=" + ($avantNetsh -eq $netshDirect) + " lu=" + $luWinHttp)

    $actifReel = ([int]$regReel.ProxyEnable -eq 1)
    $loopInet = ([string]$regReel.ProxyServer -match '(?i)(127\.0\.0\.1|localhost)\s*:\s*\d')
    $okInet = $okCleInet -and ([int]$avantInet.ProxyEnable -eq [int]$regReel.ProxyEnable) -and ([string]$avantInet.ProxyServer -eq [string]$regReel.ProxyServer) -and ([string]$avantInet.ProxyOverride -eq [string]$regReel.ProxyOverride)
    if ($actifReel -and $loopInet) { $okInet = $okInet -and ($null -ne $iReel) -and (([string]$regReel.ProxyServer) -match (':' + $iReel.Port)) }
    else { $okInet = $okInet -and ($null -eq $iReel) }
    $memeInet = (($null -eq $iReel) -and ($null -eq $iExplicite)) -or (($null -ne $iReel) -and ($null -ne $iExplicite) -and ($iReel.Port -eq $iExplicite.Port))
    $luInet = if ($null -eq $iReel) { "aucun proxy loopback" } else { "$($iReel.Host):$($iReel.Port)" }
    Check "reliquat reel : proxy de session lu dans HKCU (aucun sink simule)" ($okInet -and $memeInet) ("cle=" + $sinkReel.InetKey + " (attendue=" + $cleInet + ", identique=" + $okCleInet + ") ProxyEnable=" + $regReel.ProxyEnable + " ProxyServer=[" + $regReel.ProxyServer + "] lu=" + $luInet)

    $apresNetsh = ([string](& $sinkReel.Show)).Trim()
    $apresInet = & $sinkReel.ReadInet
    Check "reliquat reel : lectures sans aucune ecriture" (($apresNetsh -eq $avantNetsh) -and ([int]$apresInet.ProxyEnable -eq [int]$avantInet.ProxyEnable) -and ([string]$apresInet.ProxyServer -eq [string]$avantInet.ProxyServer) -and ([string]$apresInet.ProxyOverride -eq [string]$avantInet.ProxyOverride)) ("netsh identique=" + ($apresNetsh -eq $avantNetsh) + " ProxyEnable=" + $avantInet.ProxyEnable + "->" + $apresInet.ProxyEnable + " ProxyServer=[" + $apresInet.ProxyServer + "]")

    Section "8 quater. Bloc de demarrage : reliquat propose au demarrage"
    # Le bloc Add_Shown est extrait tel quel et execute : seules les fonctions qu'il
    # appelle et les controles WinForms qu'il touche sont remplaces. C'est donc le
    # cablage reel du demarrage qui est exerce (test du reliquat, puis remise en
    # direct seulement si l'utilisateur accepte), et non les fonctions seules.
    $corpsDemarrage = Get-ShownBlock -Lines $lines

    $script:JournalDemarrage = New-Object System.Collections.Generic.List[string]
    $fauxLogBox = New-Object PSObject
    Add-Member -InputObject $fauxLogBox -MemberType ScriptMethod -Name AppendText -Value { param($t) $script:JournalDemarrage.Add([string]$t) }
    Add-Member -InputObject $fauxLogBox -MemberType ScriptMethod -Name ScrollToCaret -Value { }
    $logBox = $fauxLogBox
    $chkUseDepot    = New-Object PSObject -Property @{ Checked = $false }
    $depotPathBox   = New-Object PSObject -Property @{ Enabled = $false }
    $depotBrowseBtn = New-Object PSObject -Property @{ Enabled = $false }

    $script:Demarrage = @{ Appels = $null ; Residu = $null ; Reponse = $true ; Port = 0 ; Message = "" }
    function Get-RelayPort { param([string]$Path) $script:Demarrage.Appels.Add("Get-RelayPort") ; return [int]$script:Demarrage.Port }
    function Test-ProxyRelayResidue { param($Sink, [int]$RelayPort = 0, [int]$TimeoutMs = 1500) $script:Demarrage.Appels.Add("Test-ProxyRelayResidue:" + $RelayPort) ; return $script:Demarrage.Residu }
    function Reset-ProxyRelayResidue { param($Sink, [bool]$WinHttp = $true, [bool]$WinInet = $true) $script:Demarrage.Appels.Add("Reset-ProxyRelayResidue:winhttp=" + $WinHttp + ",wininet=" + $WinInet) ; return @{ OK = $true ; Detail = "" } }
    function Clear-RelayPort { param([string]$Path) $script:Demarrage.Appels.Add("Clear-RelayPort") }
    function Confirm-ProxyRelayRestore { param([string]$Message) $script:Demarrage.Message = $Message ; $script:Demarrage.Appels.Add("Confirm-ProxyRelayRestore") ; return [bool]$script:Demarrage.Reponse }
    function Update-PreChecks { return $null }
    function Update-ProxyControlState { }
    function Get-Config { return $null }

    function Invoke-Demarrage {
        param([string]$Corps, $Residu, [bool]$Reponse = $true, [int]$Port = 4321)
        $script:Demarrage.Appels = New-Object System.Collections.Generic.List[string]
        $script:JournalDemarrage.Clear()
        $script:Demarrage.Residu  = $Residu
        $script:Demarrage.Reponse = $Reponse
        $script:Demarrage.Port    = $Port
        $script:Demarrage.Message = ""
        & ([scriptblock]::Create($Corps)) | Out-Null
        return $script:Demarrage.Appels
    }
    function Nb-Reset { param($Appels) @($Appels | Where-Object { $_ -like 'Reset-*' }).Count }

    # 1. Reliquat trouve et utilisateur d'accord : test, question, remise, marqueur efface.
    $residuTrouve = @{ Found = $true ; Host = "127.0.0.1" ; Port = 4321 ; Alive = $false ; Reason = "relais mort" ; WinHttp = $true ; WinInet = $false }
    $appels1 = @(Invoke-Demarrage -Corps $corpsDemarrage -Residu $residuTrouve -Reponse $true)
    $journal1 = ($script:JournalDemarrage -join " ").Trim()
    # La sequence attendue est verifiee sans figer le nombre total d'appels : un journal
    # ou une etape ajoutes dans le bloc ne doivent pas casser ce controle. Les seules
    # etapes surveillees sont retenues, dans leur ordre ; une remise ou un effacement
    # supplementaire reste donc visible.
    $attendu1 = @("Get-RelayPort", "Test-ProxyRelayResidue:4321", "Confirm-ProxyRelayRestore",
                  "Reset-ProxyRelayResidue:winhttp=True,wininet=False", "Clear-RelayPort")
    $suivi1 = @($appels1 | Where-Object { $v = [string]$_ ;
                  ($v -eq "Get-RelayPort") -or ($v -like "Test-ProxyRelayResidue:*") -or ($v -eq "Confirm-ProxyRelayRestore") -or
                  ($v -like "Reset-ProxyRelayResidue:*") -or ($v -eq "Clear-RelayPort") })
    $ok1 = (($suivi1 -join " > ") -eq ($attendu1 -join " > ")) -and
           ($journal1 -match 'Reliquat supprime') -and ($script:Demarrage.Message -match 'Windows \(WinHTTP\)')
    Check "demarrage : reliquat trouve -> test puis remise en direct" $ok1 ("suite=" + ($suivi1 -join " > ") + " | appels=" + ($appels1 -join " > ") + " | journal=" + $journal1)

    # 2. Utilisateur refuse : rien n'est modifie, le marqueur reste pour plus tard.
    $appels2 = @(Invoke-Demarrage -Corps $corpsDemarrage -Residu $residuTrouve -Reponse $false)
    $journal2 = ($script:JournalDemarrage -join " ").Trim()
    $ok2 = ($appels2 -contains "Confirm-ProxyRelayRestore") -and ((Nb-Reset $appels2) -eq 0) -and
           (-not ($appels2 -contains "Clear-RelayPort")) -and ($journal2 -match 'Reliquat conserve')
    Check "demarrage : refus de l'utilisateur -> aucune modification" $ok2 (($appels2 -join " > ") + " | journal=" + $journal2)

    # 3. Relais toujours vivant : aucune question, aucune remise en direct.
    $residuVivant = @{ Found = $false ; Host = "" ; Port = 0 ; Alive = $true ; Reason = "relais vivant" ; WinHttp = $false ; WinInet = $false }
    $appels3 = @(Invoke-Demarrage -Corps $corpsDemarrage -Residu $residuVivant -Reponse $true)
    $ok3 = ($appels3 -contains "Test-ProxyRelayResidue:4321") -and (-not ($appels3 -contains "Confirm-ProxyRelayRestore")) -and
           ((Nb-Reset $appels3) -eq 0) -and (-not ($appels3 -contains "Clear-RelayPort"))
    Check "demarrage : pas de reliquat -> ni question ni remise en direct" $ok3 ($appels3 -join " > ")

    # 4. Marqueur perime (le port enregistre n'est plus celui du relais) : on l'efface.
    $residuPerime = @{ Found = $false ; Host = "" ; Port = 0 ; Alive = $false ; Reason = "port different" ; WinHttp = $false ; WinInet = $false }
    $appels4 = @(Invoke-Demarrage -Corps $corpsDemarrage -Residu $residuPerime -Reponse $true)
    $ok4 = ($appels4 -contains "Clear-RelayPort") -and ((Nb-Reset $appels4) -eq 0) -and
           (-not ($appels4 -contains "Confirm-ProxyRelayRestore"))
    Check "demarrage : marqueur perime efface sans remise en direct" $ok4 ($appels4 -join " > ")

    # 5. Le bloc extrait est bien celui du script, et l'ordre du source est le bon.
    $iTest  = $corpsDemarrage.IndexOf("Test-ProxyRelayResidue")
    $iReset = $corpsDemarrage.IndexOf("Reset-ProxyRelayResidue")
    Check "demarrage : bloc Add_Shown reel (test avant la remise en direct)" (($iTest -ge 0) -and ($iReset -gt $iTest) -and ($corpsDemarrage -match 'Update-PreChecks')) ("test@" + $iTest + " reset@" + $iReset + " taille=" + $corpsDemarrage.Length)

    Section "9. Snapshot du proxy : copie, et echec netsh"
    # 1) l'etat retourne par bg_ApplyProxy doit etre une copie, pas une vue de la machine
    $machine.WinInet.ProxyEnable = 42
    $machine.WinInet.ProxyServer = "pollue"
    $machine.Env["HTTP_PROXY"] = "http://pollue:1"
    $machine.Env["http_proxy"] = "pollue"
    $machine.WinHttpServer = "pollue"
    $copieIntacte = (($etat.WinInet.ProxyEnable -eq $avant.WinInet.ProxyEnable) -and
                     ($etat.WinInet.ProxyServer -eq $avant.WinInet.ProxyServer) -and
                     ($etat.Env["HTTP_PROXY"] -eq $avant.Env["HTTP_PROXY"]) -and
                     ($etat.Env["http_proxy"] -eq $avant.Env["http_proxy"]) -and
                     ($etat.WinHttpDirect -eq $avant.WinHttpDirect))
    Check "snapshot : inchange malgre une mutation de la machine" $copieIntacte ("snapshot=" + $etat.WinInet.ProxyEnable + "/" + $etat.Env["HTTP_PROXY"] + "  machine=" + $machine.WinInet.ProxyEnable + "/" + $machine.Env["HTTP_PROXY"])

    # 2) la lecture reelle du poste (sans ecriture) renvoie des objets neufs a chaque appel
    $vraiSink = bg_ProxySink $null
    $lu1 = & $vraiSink.Read
    $lu1.WinInet["ProxyEnable"] = 999
    $lu1.Env["HTTP_PROXY"] = "pollue"
    $lu2 = & $vraiSink.Read
    Check "lecture reelle : etat neuf a chaque appel" (($lu2.WinInet.ProxyEnable -ne 999) -and ($lu2.Env["HTTP_PROXY"] -ne "pollue")) ("ProxyEnable=" + $lu2.WinInet.ProxyEnable)

    # 3) netsh en erreur : l'application continue (WinINet + environnement) et le dit
    $fauxKo = @{}
    foreach ($k in @($faux.Keys)) { $fauxKo[$k] = $faux[$k] }
    $fauxKo.Apply = { param($srv,$byp) $appels.Add("winhttp-set-echec " + $srv) ; @{ Confirmed = $false ; Detail = "" ; Error = "Access is denied." } }
    $machine.WinHttpServer = "" ; $machine.WinHttpDirect = $avant.WinHttpDirect ; $machine.WinHttpExists = $avant.WinHttpExists ; $machine.WinHttpValue = $avant.WinHttpValue
    foreach ($k in @($avant.WinInet.Keys)) { $machine.WinInet[$k] = $avant.WinInet[$k] }
    foreach ($n in @("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy")) { $machine.Env[$n] = $avant.Env[$n] }
    $script:ProxyLogs.Clear() ; $appels.Clear()
    $cfgP.ProxySink = $fauxKo
    $etatKo = bg_ApplyProxy $cfgP -ProxyHost "10.0.0.1" -ProxyPort "3128" -BypassList "<local>"
    $tKo = ($script:ProxyLogs -join " | ")
    Check "netsh en erreur : avertissement explicite" ($tKo -match 'Echec netsh winhttp set proxy : Access is denied') $tKo
    Check "netsh en erreur : WinINet et environnement appliques quand meme" (($machine.WinHttpServer -eq "") -and ($machine.WinInet.ProxyEnable -eq 1) -and ($machine.Env["HTTP_PROXY"] -eq "http://10.0.0.1:3128")) ("winhttp=[" + $machine.WinHttpServer + "] inet=" + $machine.WinInet.ProxyEnable + " env=" + $machine.Env["HTTP_PROXY"])
    bg_RestoreProxy $cfgP -State $etatKo
    Check "netsh en erreur : restauration complete malgre l'echec" ((Copy-Etat $machine | ConvertTo-Json -Compress -Depth 5) -eq ($avant | ConvertTo-Json -Compress -Depth 5)) (Copy-Etat $machine | ConvertTo-Json -Compress -Depth 5)
    Check "netsh en erreur : journal de restauration" ((($script:ProxyLogs -join " | ") -match "Configuration proxy d'origine restauree") -and (($script:ProxyLogs -join " | ") -match 'Proxy WinINet \(session\) applique')) ($script:ProxyLogs -join " | ")
    Section "10. Hook pre-commit : contenu indexe"
    # Le hook doit valider l'index (ce qui part dans le commit), pas l'arbre de travail.
    # Preuve : faire diverger les deux dans un depot jetable et exiger le bon verdict.
    $gitCmd = Get-Command git -ErrorAction SilentlyContinue
    $shExe = Get-ShExe
    $hookJouable = ($null -ne $shExe)
    $motifSh = if ($hookJouable) { "" } else { "sh introuvable (ni PATH ni installation de git), hook non execute : " }
    $hookSrc = Join-Path $PSScriptRoot ".githooks\pre-commit"
    if ($null -eq $gitCmd) {
        Check "hook : git disponible" $false "git introuvable : le hook pre-commit n'a pas pu etre verifie"
    }
    elseif (-not (Test-Path -LiteralPath $hookSrc)) {
        Check "hook : .githooks/pre-commit present" $false "hook absent : le controle automatique n'est pas installe"
    }
    elseif ($env:INSTALLOFFICE_HOOKTEST -eq "1") {
        Write-Host "  (suite lancee par le hook : section 10 ignoree pour ne pas recurser)"
    }
    else {
        # Le hook doit rester rapide : verifier que le mode -Fast existe, qu'il
        # passe sur l'arbre courant et qu'il s'arrete avant les sections lourdes.
        $eapRapide = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        $sortieRapide = (& powershell -NoProfile -STA -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "Test-InstallOffice.ps1") -Fast 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
        $codeRapide = $LASTEXITCODE
        $ErrorActionPreference = $eapRapide
        $rapideCourt = ($sortieRapide -replace "\s+", " ").Trim()
        Check "hook : mode rapide -Fast operationnel" (($codeRapide -eq 0) -and $sortieRapide.Contains("RESULTAT (rapide)")) ("code=" + $codeRapide + " sortie=" + $rapideCourt)
        Check "hook : mode rapide limite a la section 1" (($sortieRapide.Contains("=== 1. Interface livree et proprietes de l'EXE ===")) -and (-not $sortieRapide.Contains("=== 2. SourcePath"))) $rapideCourt

        $depot = Join-Path $Work "hookrepo"
        $avantTest = $env:INSTALLOFFICE_HOOKTEST
        New-Item -ItemType Directory -Path $depot -Force | Out-Null
        try {
            # Le script compile est Apercu-Maquette.ps1 ; le moteur et le dossier Maquette
            # voyagent avec lui, sans quoi la suite rapide ne peut pas juger la livraison.
            New-HookRepo -Path $depot -Autocrlf "false" -Dossiers @(".githooks", "Maquette") `
                -Files @("Apercu-Maquette.ps1", "Install-Office2024-KMS.ps1", "Install-Office2024-KMS.exe", "Test-InstallOffice.ps1")
            $scriptDepot = Join-Path $depot "Apercu-Maquette.ps1"
            $headHash = (Get-FileHash -LiteralPath $scriptDepot -Algorithm SHA256).Hash

            # Cas A : l'index derive, l'arbre de travail est propre.
            Copy-Item -LiteralPath $scriptDepot -Destination (Join-Path $depot "_pristine.ps1") -Force
            Add-Content -LiteralPath $scriptDepot -Value "# derive indexee uniquement" -Encoding ASCII
            $null = & git -C $depot add "Apercu-Maquette.ps1" 2>&1
            Copy-Item -LiteralPath (Join-Path $depot "_pristine.ps1") -Destination $scriptDepot -Force
            Remove-Item -LiteralPath (Join-Path $depot "_pristine.ps1") -Force
            & git -C $depot diff --quiet
            $arbreVsIndex = ($LASTEXITCODE -ne 0)
            & git -C $depot diff --cached --quiet
            $indexVsHead = ($LASTEXITCODE -ne 0)
            Check "hook : index different de l'arbre de travail" ($arbreVsIndex -and $indexVsHead) ("arbre/index diverge=" + $arbreVsIndex + " index/HEAD diverge=" + $indexVsHead)
            $arbreHash = (Get-FileHash -LiteralPath $scriptDepot -Algorithm SHA256).Hash
            Check "hook : arbre de travail laisse propre" ($arbreHash -eq $headHash) ("arbre=" + $arbreHash.Substring(0, 12) + " HEAD=" + $headHash.Substring(0, 12))

            $env:INSTALLOFFICE_HOOKTEST = "1"
            # Les messages d'echec du hook partent sur stderr : ne pas les traiter
            # comme des exceptions (le harnais tourne en ErrorActionPreference Stop).
            $eapAvant = $ErrorActionPreference
            $ErrorActionPreference = "Continue"
            Push-Location $depot
            $sortieA = Invoke-HookSh -ShExe $shExe -Hook ".githooks/pre-commit"
            $codeA = $script:HookShCode
            Pop-Location
            $ErrorActionPreference = $eapAvant
            $sortieAU = ($sortieA -replace "\s+", " ").Trim()
            Check "hook : refuse un script indexe desynchronise" ($hookJouable -and ($codeA -ne 0)) ($motifSh + "code=" + $codeA + " sortie=" + $sortieAU)
            Check "hook : refuse en parlant de recompilation" ($hookJouable -and $sortieA.Contains("ne correspond plus au script")) ($motifSh + $sortieAU)
            # L'empreinte citee dans l'echec doit etre celle du script indexe, pas celle de l'arbre.
            $vueHook = Join-Path $depot "_vue"
            New-Item -ItemType Directory -Path $vueHook -Force | Out-Null
            $null = & git -C $depot checkout-index --all --prefix="$vueHook/" 2>&1
            $hashIndexe = (Get-FileHash -LiteralPath (Join-Path $vueHook "Apercu-Maquette.ps1") -Algorithm SHA256).Hash
            Remove-Item -LiteralPath $vueHook -Recurse -Force -ErrorAction SilentlyContinue
            Check "hook : empreinte comparee = script indexe" ($hookJouable -and ($hashIndexe -ne $headHash) -and $sortieA.Contains($hashIndexe.Substring(0, 12))) ($motifSh + "index=" + $hashIndexe.Substring(0, 12) + " arbre=" + $headHash.Substring(0, 12) + " sortie=" + $sortieAU)

            # Cas B : l'arbre de travail derive, l'index est propre (suite temoin, rapide).
            $null = & git -C $depot reset -q 2>&1
            Add-Content -LiteralPath $scriptDepot -Value "# derive non indexee" -Encoding ASCII
            $sonde = Join-Path $depot "_sonde.txt"
            $stub = @(
                '$h = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot "Apercu-Maquette.ps1") -Algorithm SHA256).Hash',
                '[System.IO.File]::WriteAllText($env:HOOK_PROBE, $h)',
                'exit 0'
            ) -join "`r`n"
            [System.IO.File]::WriteAllText((Join-Path $depot "Test-InstallOffice.ps1"), $stub + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
            $null = & git -C $depot add "Test-InstallOffice.ps1" 2>&1
            $env:HOOK_PROBE = $sonde
            $ErrorActionPreference = "Continue"
            Push-Location $depot
            $sortieB = Invoke-HookSh -ShExe $shExe -Hook ".githooks/pre-commit"
            $codeB = $script:HookShCode
            Pop-Location
            $ErrorActionPreference = $eapAvant
            Check "hook : accepte un arbre derive non indexe" ($hookJouable -and ($codeB -eq 0)) ($motifSh + "code=" + $codeB + " sortie=" + (($sortieB -replace "\s+", " ").Trim()))
            $vu = if (Test-Path -LiteralPath $sonde) { [System.IO.File]::ReadAllText($sonde).Trim() } else { "(absent)" }
            $arbreHash2 = (Get-FileHash -LiteralPath $scriptDepot -Algorithm SHA256).Hash
            $vuCourt = if ($vu.Length -ge 12) { $vu.Substring(0, 12) } else { $vu }
            Check "hook : suite lancee sur le contenu indexe" ($hookJouable -and ($vu -eq $headHash) -and ($arbreHash2 -ne $headHash)) ($motifSh + "suite a vu=" + $vuCourt + " index=" + $headHash.Substring(0, 12) + " arbre=" + $arbreHash2.Substring(0, 12))

            # Cas C : la declaration de fins de ligne retiree sur un poste en
            # core.autocrlf=true, comme celui du mainteneur. Le commit ne touche
            # QUE .gitattributes : le hook doit quand meme se declencher, exporter
            # l'index sans conversion automatique, et refuser.
            $depotEol = Join-Path $Work "hookrepo_eol"
            $eapEol = $ErrorActionPreference
            # git add signale en stderr que LF sera converti en CRLF au prochain
            # checkout : ce n'est pas une erreur, et le harnais tourne en Stop.
            $ErrorActionPreference = "Continue"
            try {
                # Script ecrit en LF avant le commit, comme dans les blobs du depot.
                New-HookRepo -Path $depotEol -Autocrlf "true" -Dossiers @(".githooks", "Maquette") `
                    -Files @("Apercu-Maquette.ps1", "Install-Office2024-KMS.exe", "Test-InstallOffice.ps1", ".gitattributes") `
                    -Avant {
                        param($p, $src)
                        $txtLf = ([System.IO.File]::ReadAllText((Join-Path $src "Install-Office2024-KMS.ps1"))) -replace "`r`n", "`n"
                        [System.IO.File]::WriteAllText((Join-Path $p "Install-Office2024-KMS.ps1"), $txtLf, (New-Object System.Text.UTF8Encoding($false)))
                    }
                # Premisse du cas C : le script doit etre commite en LF. Si cette
                # ecriture regressait, le hook verrait du CRLF d'office et le
                # scenario refuserait pour une autre raison, sans rien prouver.
                # Taille attendue : le contenu ramene a des fins de ligne LF. Un
                # blob qui aurait garde ses CR pese exactement un octet de plus
                # par ligne, donc la comparaison suffit a prouver l'absence de CR.
                $contenuLf = ([System.IO.File]::ReadAllText((Join-Path $depotEol "Install-Office2024-KMS.ps1"))) -replace "`r`n", "`n"
                $attenduLf = [System.Text.Encoding]::ASCII.GetByteCount($contenuLf)
                # Une fin de ligne CRLF pese un octet de plus qu'une fin LF : le
                # nombre de lignes donne donc l'ecart, sans chiffre en dur.
                $nbLignes = ([regex]::Matches($contenuLf, "`n")).Count
                $tailleBlob = [int](& git -C $depotEol cat-file -s "HEAD:Install-Office2024-KMS.ps1")
                Check "hook : cas C - script commite en LF (blob sans CR)" (($tailleBlob -gt 0) -and ($tailleBlob -eq $attenduLf)) ("blob=" + $tailleBlob + " octets, attendu LF=" + $attenduLf + " : un blob en CRLF peserait " + ($attenduLf + $nbLignes) + " octets (" + $nbLignes + " lignes)")
                # La regression : uniquement la declaration de fins de ligne.
                $attrEol = Join-Path $depotEol ".gitattributes"
                $garde = @([System.IO.File]::ReadAllLines($attrEol) | Where-Object { $_ -notmatch '^\s*\*\.ps1\s+text\s+eol=crlf\s*$' })
                [System.IO.File]::WriteAllLines($attrEol, $garde, (New-Object System.Text.UTF8Encoding($false)))
                $null = & git -C $depotEol add ".gitattributes" 2>&1
                $indexeEol = @(& git -C $depotEol diff --cached --name-only 2>&1)
                # Tolerance : un fichier supplementaire stage ne fausse plus le verdict ;
                # ce qui compte est que .gitattributes soit la declaration en cause, sans
                # script stage a cote qui pourrait provoquer le refus a sa place.
                Check "hook : .gitattributes surveille" (($indexeEol -contains ".gitattributes") -and (@($indexeEol | Where-Object { $_ -like "*.ps1" }).Count -eq 0)) ($indexeEol -join ", ")
                $ErrorActionPreference = "Continue"
                Push-Location $depotEol
                $sortieC = Invoke-HookSh -ShExe $shExe -Hook ".githooks/pre-commit"
                $codeC = $script:HookShCode
                Pop-Location
                $ErrorActionPreference = $eapAvant
                $sortieCU = ($sortieC -replace "\s+", " ").Trim()
                Check "hook : refuse une declaration de fins de ligne retiree" ($hookJouable -and ($codeC -ne 0)) ($motifSh + "code=" + $codeC + " sortie=" + $sortieCU)
                Check "hook : export independant de core.autocrlf" ($hookJouable -and ($sortieC -match "CRLF=0 LF isole=")) ($motifSh + $sortieCU)
            }
            finally {
                $ErrorActionPreference = $eapEol
                if (Test-Path -LiteralPath $depotEol) { Remove-Item -LiteralPath $depotEol -Recurse -Force -ErrorAction SilentlyContinue }
            }
        }
        finally {
            $env:INSTALLOFFICE_HOOKTEST = $avantTest
            $env:HOOK_PROBE = $null
            if (Test-Path -LiteralPath $depot) { Remove-Item -LiteralPath $depot -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }

    # Pre-livraison : le ZIP telecharge depuis GitHub est l'archive du depot (git archive
    # HEAD). On l'exporte ici, on verifie qu'elle contient ce sans quoi l'EXE ne peut pas
    # fonctionner, puis on y rejoue la section 1 : le dossier extrait est un poste comme un
    # autre, il ne voit ni .git ni l'arbre de travail.
    Section "11. Pre-livraison : archive du depot"
    $archiveZip = Join-Path $Work "depot.zip"
    $archiveDir = Join-Path $Work "depot"
    $problemesArchive = @()
    # Racine de l'archive exportee, quand il y en a une : c'est la que l'empreinte des
    # captures est recalculee. Une copie sans .git n'a pas d'archive a inspecter.
    $racinePartante = $null
    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot ".git"))) {
        # Copie extraite d'un ZIP (Code -> Download ZIP) : il n'y a pas de depot a exporter,
        # donc rien a recouper ici - la section 1 bis vient de verifier l'empreinte des
        # captures sur cette meme copie, la rejouer n'apprendrait rien. Reste la completude
        # du dossier, que seule cette section regarde. Un clone et la CI, eux, exportent et
        # verifient l'archive.
        $manquants = @()
        foreach ($attendu in @("Maquette\Principal.xaml", "Install-Office2024-KMS.ps1", "setup.exe")) {
            if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $attendu))) { $manquants += $attendu }
        }
        if ($manquants.Count -gt 0) {
            $problemesArchive += ("copie incomplete : manquants=" + ($manquants -join ", "))
        } else {
            Write-Host "         copie sans .git : archive non exportable ici (clone et CI la verifient), fichiers livres presents"
        }
    } else {
    $eapArchive = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    $sortieArchive = (& git archive --format=zip -o $archiveZip HEAD 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
    $codeArchive = $LASTEXITCODE
    $ErrorActionPreference = $eapArchive
    if (($codeArchive -ne 0) -or -not (Test-Path -LiteralPath $archiveZip)) {
        $problemesArchive += ("git archive HEAD a echoue (code " + $codeArchive + " : " + $sortieArchive + ")")
    } else {
        try {
            Expand-Archive -LiteralPath $archiveZip -DestinationPath $archiveDir -Force
            $racinePartante = $archiveDir
        } catch { $problemesArchive += ("archive illisible : " + $_.Exception.Message) }
        $manquants = @()
        foreach ($attendu in @("Maquette\Principal.xaml", "Install-Office2024-KMS.ps1", "setup.exe")) {
            if (-not (Test-Path -LiteralPath (Join-Path $archiveDir $attendu))) { $manquants += $attendu }
        }
        if ($manquants.Count -gt 0) { $problemesArchive += ("archive incomplete : manquants=" + ($manquants -join ", ")) }
        $suiteArchive = Join-Path $archiveDir "Test-InstallOffice.ps1"
        if (Test-Path -LiteralPath $suiteArchive) {
            $eapArchive = $ErrorActionPreference
            $ErrorActionPreference = "Continue"
            $sortieInterne = (& $psHost -NoProfile -STA -ExecutionPolicy Bypass -File $suiteArchive -Fast 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
            $codeInterne = $LASTEXITCODE
            $ErrorActionPreference = $eapArchive
            # Code 0 ne suffit pas : la ligne RESULTAT prouve que la section 1 a bien ete
            # jouee, et N/N qu'aucun de ses controles n'a echoue.
            $resultatInterne = @($sortieInterne -split "`r?`n" | Where-Object { $_ -match 'RESULTAT \(rapide\)' })[-1]
            $okInterne = ($codeInterne -eq 0) -and ($resultatInterne -match 'RESULTAT \(rapide\) : (\d+)/(\d+) controles passes') -and ($matches[1] -eq $matches[2]) -and ([int]$matches[1] -gt 0)
            if (-not $okInterne) {
                # Les controles en echec, la ligne RESULTAT, ou a defaut la premiere ligne
                # ecrite : un processus qui meurt avant de jouer la section 1 doit laisser
                # une trace lisible, et jamais des parentheses vides.
                $detailsInterne = @($sortieInterne -split "`r?`n" | Where-Object { $_ -match 'ECHEC|ERREUR DU HARNAIS' })
                if ($resultatInterne) { $detailsInterne += $resultatInterne }
                if ($detailsInterne.Count -eq 0) { $detailsInterne = @(@($sortieInterne -split "`r?`n" | Where-Object { $_.Trim() -ne "" }) | Select-Object -First 1) }
                $suffixeInterne = if ($detailsInterne.Count -gt 0) { " (" + ($detailsInterne -join " | ") + ")" } else { "" }
                $problemesArchive += ("section 1 dans l'archive : code " + $codeInterne + $suffixeInterne)
            }
        } else {
            $problemesArchive += "section 1 non rejouee : Test-InstallOffice.ps1 absent de l'archive"
        }
        # L'archive est ce qui part : son Maquette\captures.txt doit decrire le XAML qu'elle
        # contient. Un XAML commite sans regenerer les captures laisse une reference perimee
        # dans le ZIP, que la section 1 bis ne voit pas (elle juge la copie de travail).
        # La liste des PNG vient de la section 1 bis : une seule copie de cette liste.
        if ($null -ne $racinePartante) {
            $sourcesArchive = ""
            $ficheArchive = Join-Path $racinePartante "Maquette\captures.txt"
            if (Test-Path -LiteralPath $ficheArchive) {
                foreach ($l in [System.IO.File]::ReadAllLines($ficheArchive)) {
                    if ($l -match '^sources\s+([0-9a-fA-F]{64})\s*$') { $sourcesArchive = $matches[1].ToLowerInvariant() }
                }
            }
            $empreinteArchive = ""
            if (Test-Path -LiteralPath (Join-Path $racinePartante "Apercu-Maquette.ps1") -PathType Leaf) {
                $empreinteArchive = Get-EmpreinteSources -Racine $racinePartante
            }
            $pngAbsents = @($captures | ForEach-Object { $_.Nom } |
                Where-Object { -not (Test-Path -LiteralPath (Join-Path $racinePartante ("Maquette\" + $_))) })
            if ($sourcesArchive -eq "") {
                $problemesArchive += "captures de l'archive : Maquette\captures.txt absent ou sans ligne sources"
            } elseif ($empreinteArchive -eq "") {
                $problemesArchive += ("captures de l'archive : empreinte incalculable, Apercu-Maquette.ps1 absent de l'archive (enregistre=" + $sourcesArchive.Substring(0, 12) + ")")
            } elseif ($sourcesArchive -ne $empreinteArchive) {
                $problemesArchive += ("captures de l'archive perimees : enregistre=" + $sourcesArchive.Substring(0, 12) + " recalcule=" + $empreinteArchive.Substring(0, 12) + " - regenerer les quatre captures, lancer Test-InstallOffice.ps1 -EnregistrerCaptures, puis committer PNG et Maquette\captures.txt")
            }
            if ($pngAbsents.Count -gt 0) { $problemesArchive += ("captures de l'archive : PNG absents=" + ($pngAbsents -join ", ")) }
        }
    }
    }
    Check "pre-livraison : l'archive du depot est complete, ses captures sont a jour et elle passe la section 1" ($problemesArchive.Count -eq 0) ($problemesArchive -join " ; ")


    Section "12. Livraison : signature, empreinte publiee, ZIP en liste blanche"
    # setup.exe part tel quel sur les postes : sa signature est la seule preuve que c'est
    # bien l'Office Deployment Tool de Microsoft et pas un binaire substitue.
    $sigSetup = $null
    if (Test-Path -LiteralPath $setupPath) { $sigSetup = Get-AuthenticodeSignature -LiteralPath $setupPath }
    $sujetSetup = if (($null -ne $sigSetup) -and ($null -ne $sigSetup.SignerCertificate)) { $sigSetup.SignerCertificate.Subject } else { "" }
    $etatSetup = if ($null -ne $sigSetup) { [string]$sigSetup.Status } else { "(absent)" }
    Check "setup.exe : signature Authenticode valide emise par Microsoft" (($etatSetup -eq "Valid") -and ($sujetSetup -match 'O=Microsoft Corporation')) ("Status=" + $etatSetup + " sujet=[" + $sujetSetup + "] - un setup.exe non signe ou d'un autre editeur ne doit pas partir")

    # L'empreinte du binaire est publiee dans le README : c'est ce qu'un poste peut
    # comparer sans git, sans suite de tests et sans rien installer.
    $hashExeLivre = (Get-FileHash -LiteralPath $ExePath -Algorithm SHA256).Hash.ToUpperInvariant()
    $texteReadme = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "README.md"))
    $hashAnnonce = ""
    if ($texteReadme -match '(?m)^SHA256 Install-Office2024-KMS\.exe\s*=\s*([0-9a-fA-F]{64})\s*$') { $hashAnnonce = $matches[1].ToUpperInvariant() }
    $detailHash = if ($hashAnnonce -ne "") { $hashAnnonce } else { "(aucune ligne 'SHA256 Install-Office2024-KMS.exe = ...')" }
    Check "README : empreinte SHA-256 de l'EXE annoncee = fichier livre" (($hashAnnonce -ne "") -and ($hashAnnonce -eq $hashExeLivre)) ("README annonce=" + $detailHash + " / fichier=" + $hashExeLivre + " - mettre le README a jour apres chaque recompilation")

    # Le ZIP se construit ici, dans le dossier de travail, avec la liste blanche de
    # New-Livraison.ps1 : la suite inspecte exactement ce qui partirait.
    $zipLivraison = Join-Path $Work "Livraison.zip"
    $eapLiv = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    $sortieLiv = (& $psHost -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "New-Livraison.ps1") -Destination $zipLivraison -Racine $PSScriptRoot 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
    $codeLiv = $LASTEXITCODE
    $ErrorActionPreference = $eapLiv
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $entrees = @()
    if (Test-Path -LiteralPath $zipLivraison) {
        $zipLu = [System.IO.Compression.ZipFile]::OpenRead($zipLivraison)
        try { $entrees = @($zipLu.Entries | ForEach-Object { $_.FullName }) } finally { $zipLu.Dispose() }
    }
    $requisZip = @("Install-Office2024-KMS.exe", "Install-Office2024-KMS.ps1", "setup.exe", "README.md",
                   "SHA256SUMS", "Verifier-Livraison.ps1",
                   "Maquette/Principal.xaml", "Maquette/Styles.xaml", "Maquette/Applications.xaml",
                   "Maquette/Etats.xaml", "Maquette/ReliquatProxy.xaml")
    $manquantsZip = @($requisZip | Where-Object { $entrees -notcontains $_ })
    $detailManquants = if ($manquantsZip.Count -gt 0) { $manquantsZip -join ", " } else { "aucun" }
    Check "Livraison.zip : les elements livres sont dedans (liste blanche)" (($codeLiv -eq 0) -and ($manquantsZip.Count -eq 0)) ("code=" + $codeLiv + " entrees=" + $entrees.Count + " manquants=" + $detailManquants + " ; sortie=[" + ($sortieLiv -replace "\s+", " ") + "]")

    # Ce qui ne doit jamais partir : reglages du poste, journaux, sources Office, suite de
    # tests. Un config.json dans le ZIP livrerait le serveur KMS et les chemins du poste.
    $interditsZip = @($entrees | Where-Object {
        $_ -match '(^|/)config\.json$' -or $_ -match '\.log$' -or
        $_ -match '(^|/)Office/' -or $_ -match '(^|/)\.git' -or
        $_ -match 'Test-InstallOffice\.ps1$' -or $_ -match 'Livraison\.zip$' -or
        $_ -match '(^|/)_[^/]*$' })
    Check "Livraison.zip : aucun fichier local (config.json, journal, sources Office, suite de tests)" ($interditsZip.Count -eq 0) ("trouves : " + ($interditsZip -join ", ") + " - retirer ces fichiers de la liste blanche de New-Livraison.ps1")

    # SHA256SUMS doit decrire exactement ce que contient le ZIP, empreintes comprises :
    # c'est ce que le verificateur du poste recompte.
    $sommesZip = @()
    $entreesZip = @{}
    $problemesSommes = @()
    if ($entrees -contains "SHA256SUMS") {
        $zipLu = [System.IO.Compression.ZipFile]::OpenRead($zipLivraison)
        try {
            $lecteur = New-Object System.IO.StreamReader($zipLu.GetEntry("SHA256SUMS").Open())
            try { $texteSommes = $lecteur.ReadToEnd() } finally { $lecteur.Dispose() }
            foreach ($l in ($texteSommes -split "`r?`n")) {
                if ($l -match '^([0-9a-fA-F]{64})\s+\*?(.+)$') { $sommesZip += @{ Hash = $matches[1].ToUpperInvariant() ; Chemin = $matches[2].Trim() } }
            }
            $shaZip = [System.Security.Cryptography.SHA256]::Create()
            foreach ($e in @($zipLu.Entries | Where-Object { $_.FullName -ne "SHA256SUMS" })) {
                $flux = $e.Open()
                try { $h = [System.BitConverter]::ToString($shaZip.ComputeHash($flux)).Replace("-", "").ToUpperInvariant() } finally { $flux.Dispose() }
                $entreesZip[$e.FullName] = $h
            }
        } finally { $zipLu.Dispose() }
    }
    foreach ($s in $sommesZip) {
        if (-not $entreesZip.ContainsKey($s.Chemin)) { $problemesSommes += ($s.Chemin + " : annonce mais absent du ZIP") }
        elseif ($entreesZip[$s.Chemin] -ne $s.Hash) { $problemesSommes += ($s.Chemin + " : annonce=" + $s.Hash.Substring(0, 12) + " trouve=" + $entreesZip[$s.Chemin].Substring(0, 12)) }
    }
    $nomsSommes = @($sommesZip | ForEach-Object { $_.Chemin })
    foreach ($n in @($entreesZip.Keys | Where-Object { $nomsSommes -notcontains $_ })) { $problemesSommes += ($n + " : present dans le ZIP sans ligne dans SHA256SUMS") }
    Check "Livraison.zip : SHA256SUMS couvre exactement les fichiers et leurs empreintes" (($sommesZip.Count -gt 0) -and ($problemesSommes.Count -eq 0)) ("lignes=" + $sommesZip.Count + " fichiers=" + $entreesZip.Count + " ; " + ($problemesSommes -join " ; "))

    # Le verificateur est livre dans le ZIP : la suite le rejoue sur l'archive extraite,
    # comme le ferait un enseignant sur un poste sans git.
    $dossierLiv = Join-Path $Work "livraison"
    if (Test-Path -LiteralPath $dossierLiv) { Remove-Item -LiteralPath $dossierLiv -Recurse -Force }
    if (Test-Path -LiteralPath $zipLivraison) { Expand-Archive -LiteralPath $zipLivraison -DestinationPath $dossierLiv -Force }
    $sortieVerif = ""
    $codeVerif = -1
    $verifPoste = Join-Path $dossierLiv "Verifier-Livraison.ps1"
    if (Test-Path -LiteralPath $verifPoste) {
        $eapVerif = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        $sortieVerif = (& $psHost -NoProfile -ExecutionPolicy Bypass -File $verifPoste 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
        $codeVerif = $LASTEXITCODE
        $ErrorActionPreference = $eapVerif
    }
    $resumeVerif = @($sortieVerif -split "`r?`n" | Where-Object { $_ -match 'RESULTAT :' })
    $echecsVerif = @($sortieVerif -split "`r?`n" | Where-Object { $_ -match 'ECHEC' })
    $ligneVerif = if ($resumeVerif.Count -gt 0) { $resumeVerif[-1].Trim() } else { "(aucune ligne RESULTAT)" }
    $detailVerif = "code=" + $codeVerif + " ; " + $ligneVerif
    if ($echecsVerif.Count -gt 0) { $detailVerif += " ; " + (($echecsVerif | ForEach-Object { $_.Trim() }) -join " ; ") }
    Check "Livraison.zip : le verificateur du poste (sans git) passe sur l'archive extraite" (($codeVerif -eq 0) -and ($ligneVerif -match 'RESULTAT : livraison conforme')) $detailVerif


    Section "13. Mode silencieux : configuration, rapport, codes de sortie"
    # Le mode silencieux est lance comme le ferait une GPO : processus separe, code de
    # sortie du processus, rapport relu sur disque. Les refus constates sont ceux du
    # script lui-meme, pas une simulation du harnais.
    function Invoke-Silencieux {
        param([string]$FichierConfig, [string]$FichierRapport)
        $eap = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try {
            $sortie = (& $psHost -NoProfile -STA -ExecutionPolicy Bypass -File $MainScript -Silent -Config $FichierConfig -Rapport $FichierRapport 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
            $code = $LASTEXITCODE
        } finally { $ErrorActionPreference = $eap }
        return @{ Code = $code ; Sortie = $sortie }
    }

    function New-ConfigSilencieuse {
        param([string]$Chemin, $Champs)
        $obj = [ordered]@{}
        foreach ($k in $Champs.Keys) { $obj[$k] = $Champs[$k] }
        [System.IO.File]::WriteAllText($Chemin, ($obj | ConvertTo-Json -Depth 4), (New-Object System.Text.UTF8Encoding($false)))
        return $Chemin
    }

    function Read-Rapport {
        param([string]$Chemin)
        if (-not (Test-Path -LiteralPath $Chemin)) { return $null }
        try { return (Get-Content -LiteralPath $Chemin -Raw | ConvertFrom-Json) } catch { return $null }
    }

    $jetonsSil = $null ; $errSil = $null
    $astSil = [System.Management.Automation.Language.Parser]::ParseFile($MainScript, [ref]$jetonsSil, [ref]$errSil)
    $paramsSil = @()
    if ($null -ne $astSil.ParamBlock) { $paramsSil = @($astSil.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath }) }
    $manquantsParams = @(@("Silent", "Config", "Rapport") | Where-Object { $paramsSil -notcontains $_ })
    Check "silencieux : -Silent, -Config et -Rapport declares" ($manquantsParams.Count -eq 0) ("parametres = " + ($paramsSil -join ", "))

    $defsSil = @{}
    foreach ($d in $astSil.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        if (-not $defsSil.ContainsKey($d.Name)) { $defsSil[$d.Name] = $d.Extent.Text }
    }
    # Les fonctions de lancement sont partagees avec l'ecran : le mode sans fenetre les
    # extrait comme le reste, sinon il tourne avec un lancement a lui.
    $requisesSil = @("Get-CodeSortie", "New-EtapesRapport", "New-RapportPoste", "Write-RapportPoste",
                     "New-SyncDeploiement", "Start-RelaisDeploiement", "Start-RelaisSilencieux",
                     "Start-WorkerRunspace",
                     "Invoke-DeploiementSilencieux")
    $absentesSil = @($requisesSil | Where-Object { -not $defsSil.ContainsKey($_) })
    Check "silencieux : les fonctions du mode sans fenetre existent" ($absentesSil.Count -eq 0) ("absentes : " + ($absentesSil -join ", "))
    . ([scriptblock]::Create((@($requisesSil | ForEach-Object { $defsSil[$_] }) -join "`r`n`r`n")))

    # 1. Fichier absent : refus avant tout depart, rapport ecrit quand meme.
    $rapAbsent = Join-Path $Work "silencieux_absent.json"
    $r1 = Invoke-Silencieux (Join-Path $Work "config_inexistant.json") $rapAbsent
    $rap1 = Read-Rapport $rapAbsent
    $codeRap1 = "(aucun)"
    if ($null -ne $rap1) { $codeRap1 = [string]$rap1.Code }
    Check "silencieux : config absente -> code 10 et rapport ecrit" (($r1.Code -eq 10) -and ($null -ne $rap1) -and ($rap1.Code -eq 10) -and ($rap1.Resultat -eq "Echec")) ("code=" + $r1.Code + " rapport=" + $codeRap1 + " ; " + ($r1.Sortie -replace "\s+", " "))

    # 2. JSON illisible : refuse en nommant le fichier, pas une exception brute.
    $cfgCassee = Join-Path $Work "config_cassee.json"
    Set-Content -LiteralPath $cfgCassee -Value "{ ceci n'est pas du JSON" -Encoding ASCII
    $r2 = Invoke-Silencieux $cfgCassee (Join-Path $Work "silencieux_cassee.json")
    Check "silencieux : JSON illisible -> code 10" (($r2.Code -eq 10) -and ($r2.Sortie -match "JSON illisible")) ("code=" + $r2.Code + " ; " + ($r2.Sortie -replace "\s+", " "))

    # 3. Un secret en clair dans un fichier qui part sur un partage serait lu par tous.
    $cfgClair = New-ConfigSilencieuse (Join-Path $Work "config_clair.json") ([ordered]@{
        SetupExePath = (Join-Path $Work "absent_setup.exe")
        KmsServer = "127.0.0.1:1688"
        AutoInstall = $true
        ProductKey = "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
    })
    $r3 = Invoke-Silencieux $cfgClair (Join-Path $Work "silencieux_clair.json")
    Check "silencieux : secret en clair refuse en le nommant" (($r3.Code -eq 10) -and ($r3.Sortie -match "secret en clair refuse : ProductKey")) ("code=" + $r3.Code + " ; " + ($r3.Sortie -replace "\s+", " "))

    # 4. Pre-vol non conforme : refus 10, et le rapport dit que rien n'a tourne.
    $cfgRefus = New-ConfigSilencieuse (Join-Path $Work "config_refus.json") ([ordered]@{
        SetupExePath = (Join-Path $Work "absent_setup.exe")
        KmsServer = ""
        AutoInstall = $true
        AutoActivate = $true
    })
    $rapRefus = Join-Path $Work "silencieux_refus.json"
    $r4 = Invoke-Silencieux $cfgRefus $rapRefus
    $rap4 = Read-Rapport $rapRefus
    $etapesAttente = -1
    if ($null -ne $rap4) { $etapesAttente = @($rap4.Etapes | Where-Object { $_.Etat -eq "waiting" }).Count }
    Check "silencieux : pre-vol non conforme -> code 10, cinq etapes en attente" (($r4.Code -eq 10) -and ($null -ne $rap4) -and ($etapesAttente -eq 5) -and ($rap4.Message -match "activation KMS demandee sans adresse")) ("code=" + $r4.Code + " etapes en attente=" + $etapesAttente + " ; " + ($r4.Sortie -replace "\s+", " "))

    # 4b. Relais impossible : le code 50 est exerce, pas seulement annonce. Un port
    #     amont invalide fait echouer Start-AuthenticatedProxy, comme un relais qui ne
    #     peut pas demarrer ; le contexte ne doit annoncer aucun proxy.
    $cfgRelais = New-Cfg (New-SyncHash) $fakeExe (Join-Path $Work "sil_relais") -Proxy $true
    $cfgRelais.ProxyPort = "0"
    $ctxRelais = @{ Proxy = "" }
    $relais50 = Start-RelaisSilencieux -Cfg $cfgRelais -Valeurs @{ ProxyUser = "eleve" ; ProxyPassword = "secret" } -Contexte $ctxRelais
    # Le motif doit venir du relais : "Port proxy invalide" est leve par
    # Start-AuthenticatedProxy, pas par une fonction absente du harnais.
    Check "silencieux : relais proxy impossible -> code 50" (($relais50.Code -eq 50) -and ($relais50.Motif -match "relais proxy impossible : Port proxy invalide") -and ($ctxRelais.Proxy -eq "")) ("code=" + $relais50.Code + " motif=" + $relais50.Motif)

    # 4c. Configuration partielle : les cles absentes prennent leur defaut, et le refus
    #     nomme ce qui manque au lieu de lever sur une cle absente.
    $cfgPartiel = New-ConfigSilencieuse (Join-Path $Work "config_partiel.json") ([ordered]@{
        AutoInstall  = $true
        SetupExePath = (Join-Path $Work "absent_partiel.exe")
    })
    $rapPartiel = Join-Path $Work "silencieux_partiel.json"
    $r4c = Invoke-Silencieux $cfgPartiel $rapPartiel
    $rap4c = Read-Rapport $rapPartiel
    Check "silencieux : config partielle -> defauts et refus nomme" (($r4c.Code -eq 10) -and ($r4c.Sortie -match "setup.exe absent") -and ($null -ne $rap4c) -and ($rap4c.Code -eq 10)) ("code=" + $r4c.Code + " ; " + ($r4c.Sortie -replace "\s+", " "))

    # 4d. Chemin de rapport invalide : le processus rend son code, dit que le rapport
    #     manque, et n'echoue pas sur une exception apres coup.
    $blocage = Join-Path $Work "blocage_rapport.txt"
    Set-Content -LiteralPath $blocage -Value "x"
    $r4d = Invoke-Silencieux $cfgPartiel (Join-Path $blocage "rapport.json")
    Check "silencieux : chemin de rapport invalide -> code rendu et erreur dite" (($r4d.Code -eq 10) -and ($r4d.Sortie -match "rapport non ecrit")) ("code=" + $r4d.Code + " ; " + ($r4d.Sortie -replace "\s+", " "))

    # 5. Codes de sortie : chaque etape en echec a le sien, et la desinstallation
    #    compte comme une installation (30).
    $codesEtape = @{}
    foreach ($cas in @(@(0, 30), @(2, 20), @(3, 30), @(4, 40))) {
        $syncCas = @{ StepState = [string[]]@("waiting", "waiting", "waiting", "waiting", "waiting") ; HasError = $false }
        $syncCas.StepState[$cas[0]] = "error"
        $codesEtape[$cas[0]] = Get-CodeSortie $syncCas
    }
    Check "silencieux : chaque etape en echec donne son code (20/30/40)" (($codesEtape[0] -eq 30) -and ($codesEtape[2] -eq 20) -and ($codesEtape[3] -eq 30) -and ($codesEtape[4] -eq 40)) ("desinstallation=" + $codesEtape[0] + " sources=" + $codesEtape[2] + " installation=" + $codesEtape[3] + " activation=" + $codesEtape[4])

    # 6. Un worker qui se declare en echec sans etape en erreur reste un echec (60).
    #    Le succes 0 et l'echec "sans conclure" sont exerces en vrai plus bas : ici,
    #    seule la correspondance propre a HasError est verifiee.
    $syncHors = @{ StepState = [string[]]@("done", "done", "done", "done", "skipped") ; HasError = $true }
    $cHors = Get-CodeSortie $syncHors
    Check "silencieux : echec declare hors etapes -> code 60" ($cHors -eq 60) ("hors etapes=" + $cHors)

    # 7. Le rapport : forme unique, relue en JSON, etapes comprises.
    $rapForme = Join-Path $Work "rapport_forme.json"
    $ctxForme = @{ Version = "9.9.9.9" ; Mode = "Deploiement" ; Debut = (Get-Date).AddSeconds(-12) ; Kms = "kms.exemple:1688" ; Proxy = "" }
    $syncForme = @{ StepState = [string[]]@("done", "done", "done", "error", "skipped") ; HasError = $true }
    $rapportForme = New-RapportPoste $ctxForme 30 "echec simule" (New-EtapesRapport $syncForme) ($rapForme -replace "\.json$", ".log")
    Write-RapportPoste $rapForme $rapportForme
    $relu = Read-Rapport $rapForme
    $champsOk = ($null -ne $relu) -and ($relu.Outil -eq "InstallOffice") -and ($relu.Version -eq "9.9.9.9") -and
                ($relu.Poste -eq $env:COMPUTERNAME) -and ($relu.Code -eq 30) -and ($relu.Resultat -eq "Echec") -and
                (@($relu.Etapes).Count -eq 5) -and ($relu.Etapes[3].Etat -eq "error") -and ($relu.DureeSecondes -ge 10)
    Check "silencieux : rapport complet et relu (version, poste, code, etapes)" $champsOk ("champs=" + $(if ($null -ne $relu) { "code=" + $relu.Code + " etapes=" + @($relu.Etapes).Count + " duree=" + $relu.DureeSecondes } else { "(rapport illisible)" }))

    # 8. Le worker du moteur, lance sans fenetre, rend le code de son resultat. Le
    # depot est cree ici : la section 10 reutilise le nom $depot pour son depot git,
    # et un depot sans 'Office\Data' ferait echouer l'etape des sources.
    $depotSil = Join-Path $Work "depot_silencieux"
    New-Item -ItemType Directory -Path (Join-Path $depotSil "Office\Data") -Force | Out-Null
    Set-Content -Path (Join-Path $depotSil "Office\Data\v64.cab") -Value "x" -NoNewline
    $env:FAKE_ODT_CODE = "0"
    $rapOk = Join-Path $Work "silencieux_ok.json"
    $ctxSil = @{ Version = "9.9.9.9" ; Mode = "Deploiement" ; Debut = (Get-Date) ; Kms = "" ; Proxy = "" }
    $codeOk = Invoke-DeploiementSilencieux -Cfg (New-Cfg (New-SyncHash) $fakeExe (Join-Path $Work "sil_ok") -Depot $depotSil -Install $true) `
                -Texte $Worker -CheminRapport $rapOk -CheminJournal ($rapOk -replace "\.json$", ".log") -Contexte $ctxSil -DelaiMs 50 -SansConsole
    $rapOkLu = Read-Rapport $rapOk
    $etapesDone = -1
    if ($null -ne $rapOkLu) { $etapesDone = @($rapOkLu.Etapes | Where-Object { $_.Etat -eq "done" }).Count }
    Check "silencieux : worker rejoue sans fenetre -> code 0 et rapport" (($codeOk -eq 0) -and ($null -ne $rapOkLu) -and ($rapOkLu.Code -eq 0) -and ($etapesDone -ge 3)) ("code=" + $codeOk + " etapes faites=" + $etapesDone)

    # 9. Echec ODT : le code dit la phase, pas seulement "echec".
    $env:FAKE_ODT_CODE = "30066"
    $codeOdt = Invoke-DeploiementSilencieux -Cfg (New-Cfg (New-SyncHash) $fakeExe (Join-Path $Work "sil_odt") -Depot $depotSil -Install $true) `
                -Texte $Worker -CheminRapport (Join-Path $Work "silencieux_odt.json") -CheminJournal (Join-Path $Work "silencieux_odt.log") -Contexte $ctxSil -DelaiMs 50 -SansConsole
    Check "silencieux : echec ODT -> code 30 (installation)" ($codeOdt -eq 30) ("code=" + $codeOdt + " (attendu 30)")

    # 9b. Worker qui rend la main sans conclure : 60, et le rapport dit pourquoi.
    $rapNu = Join-Path $Work "silencieux_nu.json"
    $codeNu = Invoke-DeploiementSilencieux -Cfg (New-Cfg (New-SyncHash) $fakeExe (Join-Path $Work "sil_nu")) `
                -Texte "return" -CheminRapport $rapNu -CheminJournal ($rapNu -replace "\.json$", ".log") -Contexte $ctxSil -DelaiMs 20 -SansConsole
    $rapNuLu = Read-Rapport $rapNu
    Check "silencieux : worker muet -> code 60 et rapport qui le dit" (($codeNu -eq 60) -and ($null -ne $rapNuLu) -and ($rapNuLu.Code -eq 60) -and ($rapNuLu.Message -match "sans conclure")) ("code=" + $codeNu + " message=" + $(if ($null -ne $rapNuLu) { $rapNuLu.Message } else { "(aucun)" }))

    # 10. Activation demandee sans adresse : code 40, sans toucher a ospp.vbs.
    $env:FAKE_ODT_CODE = "0"
    $cfgAct = New-Cfg (New-SyncHash) $fakeExe (Join-Path $Work "sil_act") -Depot $depotSil -Install $true -Activate $true
    $cfgAct.KmsHost = ""
    $cfgAct.ActivateOnly = $false
    $codeAct = Invoke-DeploiementSilencieux -Cfg $cfgAct -Texte $Worker `
                -CheminRapport (Join-Path $Work "silencieux_act.json") -CheminJournal (Join-Path $Work "silencieux_act.log") -Contexte $ctxSil -DelaiMs 50 -SansConsole
    Check "silencieux : activation sans adresse -> code 40" ($codeAct -eq 40) ("code=" + $codeAct + " (attendu 40)")

    # 11. Le journal accompagne le rapport : c'est ce qu'un enseignant envoie.
    $journalOk = Join-Path $Work "silencieux_ok.log"
    $lignesJournal = 0
    if (Test-Path -LiteralPath $journalOk) { $lignesJournal = @(Get-Content -LiteralPath $journalOk).Count }
    Check "silencieux : journal ecrit a cote du rapport" ($lignesJournal -gt 5) ("lignes=" + $lignesJournal + " dans " + $journalOk)

    # 12. Ce que la documentation doit dire pour qu'un deploiement scripte soit possible.
    $texteReadmeSil = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "README.md"))
    $docsSil = ($texteReadmeSil -match '-Silent') -and ($texteReadmeSil -match '-Rapport') -and
               ($texteReadmeSil -match 'ProxyPasswordEnv') -and ($texteReadmeSil -match 'Code de sortie')
    Check "README : mode silencieux, secrets et codes de sortie documentes" $docsSil ("-Silent=" + ($texteReadmeSil -match '-Silent') + " ProxyPasswordEnv=" + ($texteReadmeSil -match 'ProxyPasswordEnv') + " codes=" + ($texteReadmeSil -match 'Code de sortie'))


    Section "14. Banc de preuve du proxy authentifie (vrai proxy local)"
    # Le banc monte un vrai proxy amont (Basic exige, 407, journal par requete), le
    # relais de l'outil devant, et une cible locale. Il est lance comme le ferait un
    # utilisateur : processus separe, rapport relu sur disque.
    $bancScript = Join-Path $PSScriptRoot "Banc-ProxyAuthentifie.ps1"
    $rapBanc = Join-Path $Work "banc.json"
    $codeBanc = -1
    $sortieBanc = ""
    if (Test-Path -LiteralPath $bancScript) {
        $eapBanc = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try {
            $sortieBanc = (& $psHost -NoProfile -STA -ExecutionPolicy Bypass -File $bancScript -SansConsole -Rapport $rapBanc 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
            $codeBanc = $LASTEXITCODE
        } finally { $ErrorActionPreference = $eapBanc }
    }
    $banc = $null
    if (Test-Path -LiteralPath $rapBanc) { try { $banc = Get-Content -LiteralPath $rapBanc -Raw | ConvertFrom-Json } catch { } }
    $journalBanc = @()
    $testsBanc = @()
    if ($null -ne $banc) {
        $journalBanc = @($banc.JournalProxy | ForEach-Object { [string]$_ })
        $testsBanc = @($banc.Tests)
    }
    Check "banc proxy : huit preuves hors ligne, code 0" (($codeBanc -eq 0) -and ($null -ne $banc) -and ($banc.Ok -eq $true) -and ($testsBanc.Count -eq 8)) ("code=" + $codeBanc + " tests=" + $testsBanc.Count + " ; " + ($sortieBanc -replace "\s+", " "))
    Check "banc proxy : le proxy amont voit le Basic du relais sur le tunnel" (@($journalBanc | Where-Object { $_ -match 'CONNECT 127\.0\.0\.1:\d+ auth=oui user=' }).Count -ge 1) ($journalBanc -join " | ")
    Check "banc proxy : requete HTTP en clair transmise avec le Basic (defaut corrige)" (@($journalBanc | Where-Object { $_ -match 'GET http://.* auth=oui user=' }).Count -ge 1) ($journalBanc -join " | ")
    Check "banc proxy : connexion reutilisee - deux requetes authentifiees, aucune sans en-tete" ((@($journalBanc | Where-Object { $_ -match 'GET http://.* auth=oui user=' }).Count -ge 2) -and (@($journalBanc | Where-Object { $_ -match 'GET http://.* auth=non' }).Count -eq 0)) ($journalBanc -join " | ")
    Check "banc proxy : mauvais mot de passe -> 407 journalise" ((@($journalBanc | Where-Object { $_ -match 'auth=non' }).Count -ge 1) -and (@($journalBanc | Where-Object { $_ -match '^[0-9:.]+ 407 ' }).Count -ge 1)) ($journalBanc -join " | ")
    $jsonBanc = ""
    Check "banc proxy : le proxy Negocier a authentifie le compte Windows" (@($testsBanc | Where-Object { $_.Nom -match 'Negocier : le relais' -and $_.Ok -eq $true }).Count -eq 1) ("le banc doit prouver l'identite du compte, pas seulement un 200")
    if ($null -ne $banc) { $jsonBanc = ConvertTo-Json -InputObject $banc -Depth 6 }
    Check "banc proxy : le mot de passe ne figure pas dans le rapport" (($jsonBanc -ne "") -and ($jsonBanc -notmatch [regex]::Escape("secret-banc"))) "le rapport ne doit porter que l'utilisateur"

    Section "15. Mode diagnostic (lecture seule)"
    # Le moteur est lance comme un utilisateur le ferait, sans droits administrateur :
    # le mode doit rendre ses six points et un verdict sans ouvrir de fenetre ni toucher
    # au poste. Un seul cas compare des valeurs mesurees ici (espace libre, etat admin)
    # et l'absence d'ecriture (config, HKCU, netsh, environnement) : un rapport fige
    # ou un diagnostic qui ecrit echoue.
    $cfgReelDiag = Join-Path $env:APPDATA "Office2024Deploy\config.json"
    $dossierCfgDiag = Split-Path -Parent $cfgReelDiag
    function Get-EtatCfgDiag {
        $dossier = Test-Path -LiteralPath $dossierCfgDiag
        $fichier = if (Test-Path -LiteralPath $cfgReelDiag) {
            $f = Get-Item -LiteralPath $cfgReelDiag
            (Get-FileHash -LiteralPath $cfgReelDiag -Algorithm SHA256).Hash + "|" + $f.LastWriteTimeUtc.Ticks + "|" + $f.Length
        } else { "(absent)" }
        return @{ Fichier = $fichier ; Dossier = $dossier }
    }
    $etatCfgAvant = Get-EtatCfgDiag
    $cleInetDiag = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
    $regDiagAvant = @{}
    foreach ($p in @("ProxyEnable","ProxyServer","ProxyOverride")) {
        $regDiagAvant[$p] = $null
        try { $regDiagAvant[$p] = Get-ItemPropertyValue -Path $cleInetDiag -Name $p -ErrorAction Stop } catch {}
    }
    $netshDiagAvant = ((& netsh winhttp show proxy 2>&1) -join "`n").Trim()
    $envDiagAvant = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)
    foreach ($n in @("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy")) {
        foreach ($portee in @("Process","User","Machine")) {
            $envDiagAvant[$portee + ":" + $n] = [System.Environment]::GetEnvironmentVariable($n, $portee)
        }
    }
    $driveAvant = Get-PSDrive C -ErrorAction SilentlyContinue
    $espaceDiagAvant = if ($driveAvant) { [math]::Round($driveAvant.Free / 1GB, 1) } else { -1 }
    $estAdminDiag = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $adminDiagAttendu = if ($estAdminDiag) { "session administrateur" } else { "session standard" }
    $sortieDiag = ""
    $codeDiag = -1
    if (Test-Path -LiteralPath $EngineScript) {
        $eapDiag = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try {
            $sortieDiag = (& $psHost -NoProfile -ExecutionPolicy Bypass -File $EngineScript -Diagnostic 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
            $codeDiag = $LASTEXITCODE
        } finally { $ErrorActionPreference = $eapDiag }
    }
    $etatCfgApres = Get-EtatCfgDiag
    $regDiagApres = @{}
    foreach ($p in @("ProxyEnable","ProxyServer","ProxyOverride")) {
        $regDiagApres[$p] = $null
        try { $regDiagApres[$p] = Get-ItemPropertyValue -Path $cleInetDiag -Name $p -ErrorAction Stop } catch {}
    }
    $netshDiagApres = ((& netsh winhttp show proxy 2>&1) -join "`n").Trim()
    $envDiagApres = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)
    foreach ($n in @("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy")) {
        foreach ($portee in @("Process","User","Machine")) {
            $envDiagApres[$portee + ":" + $n] = [System.Environment]::GetEnvironmentVariable($n, $portee)
        }
    }
    $driveApres = Get-PSDrive C -ErrorAction SilentlyContinue
    $espaceDiagApres = if ($driveApres) { [math]::Round($driveApres.Free / 1GB, 1) } else { -1 }
    $espaceReelMin = [math]::Min($espaceDiagAvant, $espaceDiagApres)
    $espaceReelMax = [math]::Max($espaceDiagAvant, $espaceDiagApres)
    $mEspace = [regex]::Match($sortieDiag, 'Espace disque C:\s*:\s*([0-9]+(?:[.,][0-9]+)?)\s*Go libres\s*;\s*minimum\s*([0-9]+(?:[.,][0-9]+)?)\s*Go')
    $espaceDiag = if ($mEspace.Success) { [double]($mEspace.Groups[1].Value -replace ',', '.') } else { -1 }
    $minimumDiag = if ($mEspace.Success) { [double]($mEspace.Groups[2].Value -replace ',', '.') } else { -1 }
    $espaceOk = ($espaceDiag -ge ($espaceReelMin - 0.5)) -and ($espaceDiag -le ($espaceReelMax + 0.5))
    $mAdmin = [regex]::Match($sortieDiag, 'Droits administrateur\s*:\s*(session administrateur|session standard)')
    $adminDiag = if ($mAdmin.Success) { $mAdmin.Groups[1].Value } else { "" }
    $adminOk = ($adminDiag -eq $adminDiagAttendu)
    $cfgInchange = ($etatCfgAvant.Fichier -eq $etatCfgApres.Fichier) -and ($etatCfgAvant.Dossier -eq $etatCfgApres.Dossier)
    $regInchange = ($regDiagAvant.ProxyEnable -eq $regDiagApres.ProxyEnable) -and ($regDiagAvant.ProxyServer -eq $regDiagApres.ProxyServer) -and ($regDiagAvant.ProxyOverride -eq $regDiagApres.ProxyOverride)
    $netshInchange = ($netshDiagAvant -eq $netshDiagApres)
    $envInchange = $true
    foreach ($cleEnv in @($envDiagAvant.Keys)) {
        if ($envDiagAvant[$cleEnv] -ne $envDiagApres[$cleEnv]) { $envInchange = $false }
    }
    $pointsDiag = @("Espace disque", "Office deja present", "Droits administrateur", "Proxy joignable", "DNS KMS", "Horloge", "Verdict")
    $manquantsDiag = @($pointsDiag | Where-Object { $sortieDiag -notmatch [regex]::Escape($_) })
    $lectureSeuleOk = ($sortieDiag -match 'Aucune installation, aucun reglage modifie') -and $cfgInchange -and $regInchange -and $netshInchange -and $envInchange
    Check "diagnostic : valeurs reelles, six points et aucune ecriture" (($codeDiag -eq 0 -or $codeDiag -eq 1) -and ($manquantsDiag.Count -eq 0) -and $espaceOk -and ($minimumDiag -eq 4) -and $adminOk -and $lectureSeuleOk) ("code=" + $codeDiag + " manquants=" + ($manquantsDiag -join ", ") + " espace rapport=" + $espaceDiag + " reel=" + $espaceDiagAvant + ".." + $espaceDiagApres + " min=" + $minimumDiag + " admin rapport=[" + $adminDiag + "] reel=[" + $adminDiagAttendu + "] lecture seule=" + $lectureSeuleOk + " (config=" + $cfgInchange + " HKCU=" + $regInchange + " netsh=" + $netshInchange + " env=" + $envInchange + ") ; " + ($sortieDiag -replace "\s+", " "))

    # Configuration isolee, jamais le profil reel : le moteur relit le fichier donne par
    # -ConfigFile. Les valeurs posees ici n'existent nulle part ailleurs (port du proxy
    # attribue par le systeme, port KMS 1699) : les voir dans le rapport prouve une lecture
    # du fichier, pas des intitules ni des valeurs par defaut.
    $dossierIso = Join-Path $Work "config_diag"
    New-Item -ItemType Directory -Path $dossierIso -Force | Out-Null
    $cfgIso = Join-Path $dossierIso "config.json"
    $cfgProfilReel = Join-Path $env:APPDATA "Office2024Deploy\config.json"
    $profilAvant = if (Test-Path -LiteralPath $cfgProfilReel) { (Get-FileHash -LiteralPath $cfgProfilReel -Algorithm SHA256).Hash } else { "(absent)" }
    $proxyDiag = New-Object FakeAuthProxy -ArgumentList @("diag", "secret")
    $proxyDiag.Start()
    # Le fichier vient de Save-ConfigPoste, la fonction qui enregistre les champs de
    # l'ecran : le nom du champ KMS est donc celui que l'interface ecrit reellement.
    # Le test retrouve ce champ par sa valeur unique, puis exige que le diagnostic
    # lise cette valeur : un renommage non suivi retombe sur l'adresse d'exemple.
    $champsPoste = @{
        ChampSetup      = @{ Text = "C:\fake\setup.exe" }
        ChampKms        = @{ Text = "127.0.0.1:1699" }
        LangueTexte     = @{ Text = "fr-fr" }
        ChampProxy      = @{ Text = "127.0.0.1" }
        ChampPort       = @{ Text = [string]$proxyDiag.Port }
        ChampBypass     = @{ Text = "<local>" }
        ChampProxyUser  = @{ Text = "" }
        ChampDepot      = @{ Text = "" }
        ChampKmsSecours = @{ Text = "" }
        ChampCle        = @{ Text = "" }
    }
    $fenetrePoste = New-Object PSObject
    $fenetrePoste | Add-Member -MemberType NoteProperty -Name Champs -Value $champsPoste
    $fenetrePoste | Add-Member -MemberType ScriptMethod -Name FindName -Value { param($n) return $this.Champs[$n] }
    $cfgPoste = [pscustomobject]@{
        RemoveOld       = $false
        DoInstall       = $false
        DoActivate      = $true
        UseCustomXML    = $false
        CustomXMLPath   = ""
        ProxyEnabled    = $true
        DepotEnabled    = $false
        ApplyProductKey = $false
        ProductKey      = ""
    }
    $apercuTexteDiag = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "Apercu-Maquette.ps1"))
    $lignesApercuDiag = @($apercuTexteDiag -split "`r?`n")
    $lignesMoteurDiag = [System.IO.File]::ReadAllLines($EngineScript)
    & {
        param($Chemin, $Fenetre, $Cfg, $LignesMoteur, $LignesInterface)
        $CONFIG_FILE = $Chemin
        Invoke-Expression (Get-FunctionText -Lines $LignesMoteur -Name "Save-Config")
        Invoke-Expression (Get-FunctionText -Lines $LignesMoteur -Name "Protect-SecretText")
        Invoke-Expression (Get-FunctionText -Lines $LignesInterface -Name "Get-ValeurSaisie")
        Invoke-Expression (Get-FunctionText -Lines $LignesInterface -Name "Get-ExempleChamp")
        Invoke-Expression (Get-FunctionText -Lines $LignesInterface -Name "Save-ConfigPoste")
        Save-ConfigPoste $Fenetre $Cfg
    } $cfgIso $fenetrePoste $cfgPoste $lignesMoteurDiag $lignesApercuDiag
    $cfgEcrit = Get-Content -LiteralPath $cfgIso -Raw | ConvertFrom-Json
    $champKmsPoste = @($cfgEcrit.PSObject.Properties | Where-Object { [string]$_.Value -eq "127.0.0.1:1699" } | Select-Object -First 1)
    $nomChampKmsPoste = if ($champKmsPoste.Count -gt 0) { [string]$champKmsPoste[0].Name } else { "" }
    $sortieIso = ""
    $codeIso = -1
    $eapIso = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $sortieIso = (& $psHost -NoProfile -ExecutionPolicy Bypass -File $EngineScript -Diagnostic -ConfigFile $cfgIso 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
        $codeIso = $LASTEXITCODE
    } finally { $ErrorActionPreference = $eapIso }
    $proxyDiag.Stop()
    $profilApres = if (Test-Path -LiteralPath $cfgProfilReel) { (Get-FileHash -LiteralPath $cfgProfilReel -Algorithm SHA256).Hash } else { "(absent)" }
    $attenduProxy = "proxy 127.0.0.1:" + $proxyDiag.Port + " joignable ; 407 (authentification : Basic)"
    $attenduKms   = "DNS KMS : 127.0.0.1 -> 127.0.0.1 (port 1699)"
    $lectureIso = ($sortieIso -match [regex]::Escape($attenduProxy)) -and ($sortieIso -match [regex]::Escape($attenduKms)) -and
                  (($codeIso -eq 0) -or ($codeIso -eq 1)) -and ($cfgIso.StartsWith($Work)) -and ($profilAvant -eq $profilApres) -and
                  ($nomChampKmsPoste -eq "KmsServer")
    Check "diagnostic : proxy et KMS lus depuis la configuration ecrite par l'interface (profil reel intact)" $lectureIso ("code=" + $codeIso + " champ KMS interface=[" + $nomChampKmsPoste + "] attendu proxy=[" + $attenduProxy + "] attendu KMS=[" + $attenduKms + "] profil " + $profilAvant + " -> " + $profilApres + " ; " + ($sortieIso -replace "\s+", " "))

    # Deuxieme configuration isolee : le diagnostic doit suivre ProxyEnabled et le port
    # KMS du fichier, pas ceux du premier essai ni les valeurs par defaut.
    [ordered]@{
        KmsServer    = "127.0.0.1:1700"
        ProxyEnabled = $false
        ProxyHost    = ""
        ProxyPort    = ""
    } | ConvertTo-Json | Set-Content -LiteralPath $cfgIso -Encoding UTF8
    $sortieIso2 = ""
    $codeIso2 = -1
    $eapIso2 = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $sortieIso2 = (& $psHost -NoProfile -ExecutionPolicy Bypass -File $EngineScript -Diagnostic -ConfigFile $cfgIso -SystemProxy direct 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
        $codeIso2 = $LASTEXITCODE
    } finally { $ErrorActionPreference = $eapIso2 }
    $attenduKms2 = "DNS KMS : 127.0.0.1 -> 127.0.0.1 (port 1700)"
    Check "diagnostic : configuration isolee sans proxy -> acces direct et port KMS du fichier" (($codeIso2 -eq 0 -or $codeIso2 -eq 1) -and ($sortieIso2 -match [regex]::Escape("aucun proxy dans la configuration ; acces direct")) -and ($sortieIso2 -match [regex]::Escape($attenduKms2)) -and ($sortieIso2 -notmatch [regex]::Escape("1699"))) ("code=" + $codeIso2 + " ; " + ($sortieIso2 -replace "\s+", " "))

    # Proxy systeme : sans proxy dans la configuration, le diagnostic doit tester
    # WinHTTP (machine) puis WinINet (session) et nommer celui qu'il utilise. Le
    # proxy local est injecte par -SystemProxy : aucun reglage du poste n'est touche.
    function Invoke-DiagOptions {
        param([string]$Source = "", [string]$Proxy = "", [string]$ProxySource = "")
        $arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $EngineScript, "-Diagnostic", "-ConfigFile", $cfgIso)
        if ($Source -ne "") { $arguments += @("-TimeSource", $Source) }
        if ($Proxy -ne "") { $arguments += @("-SystemProxy", $Proxy) }
        if ($ProxySource -ne "") { $arguments += @("-SystemProxySource", $ProxySource) }
        $eap = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try {
            $out = (& $psHost @arguments 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
            $code = $LASTEXITCODE
        } finally { $ErrorActionPreference = $eap }
        return @{ Code = $code ; Out = $out }
    }

    $proxySys = New-Object FakeAuthProxy -ArgumentList @("sys", "secret")
    $proxySys.Start()
    $sondeUdp = [System.Net.Sockets.UdpClient]::new([System.Net.IPAddress]::Loopback, 0)
    $portFerme = ([System.Net.IPEndPoint]$sondeUdp.Client.LocalEndPoint).Port
    $sondeUdp.Close()
    try {
        $h4 = Invoke-DiagOptions -Source ("127.0.0.1:" + $portFerme) -Proxy ("127.0.0.1:" + $proxySys.Port) -ProxySource "WinHTTP"
        $attendu4 = "proxy systeme WinHTTP 127.0.0.1:" + $proxySys.Port + " joignable ; 407 (authentification : Basic)"
        Check "diagnostic : sans proxy d'outil, le proxy systeme WinHTTP est teste" (($h4.Code -eq 0) -and ($h4.Out -match [regex]::Escape($attendu4)) -and ($h4.Out -notmatch 'acces direct')) ("code=" + $h4.Code + " attendu=" + $attendu4 + " ; " + ($h4.Out -replace "\s+", " "))
        $h5 = Invoke-DiagOptions -Source ("127.0.0.1:" + $portFerme) -Proxy ("127.0.0.1:" + $proxySys.Port) -ProxySource "WinINet"
        $attendu5 = "proxy systeme WinINet 127.0.0.1:" + $proxySys.Port + " joignable ; 407 (authentification : Basic)"
        Check "diagnostic : sans proxy d'outil, le proxy systeme WinINet est teste" (($h5.Code -eq 0) -and ($h5.Out -match [regex]::Escape($attendu5)) -and ($h5.Out -notmatch 'acces direct')) ("code=" + $h5.Code + " attendu=" + $attendu5 + " ; " + ($h5.Out -replace "\s+", " "))
    } finally {
        $proxySys.Stop()
    }

    # Lecture des deux formats systeme, sans toucher au registre ni a netsh : les
    # fonctions du moteur sont jouees en portee enfant sur des sorties representatives.
    & {
        foreach ($fn in @("ConvertFrom-ProxyAdresse", "ConvertFrom-ProxyWinHttp", "ConvertFrom-ProxyWinInet")) {
            Invoke-Expression (Get-FunctionText -Lines $lines -Name $fn)
        }
        $wh1 = ConvertFrom-ProxyWinHttp -Texte "Proxy Server(s) : 10.1.2.3:3128`r`nBypass List : <local>"
        $wh2 = ConvertFrom-ProxyWinHttp -Texte "Direct access (no proxy server)."
        $wh3 = ConvertFrom-ProxyWinHttp -Texte "Serveur(s) proxy : proxy.local:8080"
        $okWh = ($null -ne $wh1) -and ($wh1.Host -eq "10.1.2.3") -and ($wh1.Port -eq 3128) -and ($null -eq $wh2) -and ($null -ne $wh3) -and ($wh3.Port -eq 8080)
        Check "diagnostic : lecture WinHTTP d'un proxy ou d'un acces direct" $okWh ("proxy=" + $(if ($null -eq $wh1) { "(null)" } else { $wh1.Host + ":" + $wh1.Port }) + " direct=" + $(if ($null -eq $wh2) { "null" } else { "lu" }) + " fr=" + $(if ($null -eq $wh3) { "(null)" } else { $wh3.Host + ":" + $wh3.Port }))
        $in1 = ConvertFrom-ProxyWinInet -Valeur "http=proxy.local:8080;https=proxy.local:8443"
        $in2 = ConvertFrom-ProxyWinInet -Valeur "proxy.local:3128"
        $in3 = ConvertFrom-ProxyWinInet -Valeur ""
        $okIn = ($null -ne $in1) -and ($in1.Host -eq "proxy.local") -and ($in1.Port -eq 8443) -and ($null -ne $in2) -and ($in2.Port -eq 3128) -and ($null -eq $in3)
        Check "diagnostic : lecture WinINet d'une liste http/https ou d'une adresse simple" $okIn ("https=" + $(if ($null -eq $in1) { "(null)" } else { $in1.Host + ":" + $in1.Port }) + " simple=" + $(if ($null -eq $in2) { "(null)" } else { $in2.Host + ":" + $in2.Port }))
    }

    # Horloge : un faux serveur SNTP local repond avec l'heure du poste + un decalage
    # connu. Le diagnostic tourne en processus separe, comme un utilisateur, et ne
    # touche ni au service W32Time ni a la configuration du poste.
    $fakeSntpCs = @'
using System;
using System.Net;
using System.Net.Sockets;
using System.Threading;

public class FakeSntp
{
    private UdpClient _u;
    private Thread _t;
    private int _port;
    private int _offset;

    public FakeSntp(int offsetSeconds) { _offset = offsetSeconds; }
    public int Port { get { return _port; } }

    public void Start()
    {
        _u = new UdpClient(new IPEndPoint(IPAddress.Loopback, 0));
        _port = ((IPEndPoint)_u.Client.LocalEndPoint).Port;
        _t = new Thread(Loop);
        _t.IsBackground = true;
        _t.Start();
    }

    public void Stop()
    {
        try { _u.Close(); } catch { }
    }

    private static void Put(byte[] b, int i, DateTime utc)
    {
        double s = (utc - new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
        uint sec = (uint)Math.Floor(s);
        uint frac = (uint)((s - sec) * 4294967296.0);
        b[i] = (byte)(sec >> 24); b[i + 1] = (byte)(sec >> 16); b[i + 2] = (byte)(sec >> 8); b[i + 3] = (byte)sec;
        b[i + 4] = (byte)(frac >> 24); b[i + 5] = (byte)(frac >> 16); b[i + 6] = (byte)(frac >> 8); b[i + 7] = (byte)frac;
    }

    private void Loop()
    {
        while (true)
        {
            IPEndPoint remote = null;
            byte[] req;
            try { req = _u.Receive(ref remote); } catch { return; }
            if (req == null || req.Length < 48) { continue; }
            byte[] rep = new byte[48];
            Array.Copy(req, rep, 48);
            DateTime now = DateTime.UtcNow.AddSeconds(_offset);
            rep[0] = 0x24;
            rep[1] = 1;
            rep[2] = 4;
            rep[3] = 0xEC;
            Put(rep, 16, now);
            for (int i = 24; i < 32; i++) { rep[i] = req[i]; }
            Put(rep, 32, now);
            Put(rep, 40, now);
            try { _u.Send(rep, rep.Length, remote); } catch { return; }
        }
    }
}
'@
    if (-not ("FakeSntp" -as [type])) { Add-Type -TypeDefinition $fakeSntpCs -Language CSharp }

    function Invoke-DiagHorloge {
        param([string]$Source)
        $eap = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try {
            $out = (& $psHost -NoProfile -ExecutionPolicy Bypass -File $EngineScript -Diagnostic -ConfigFile $cfgIso -TimeSource $Source -SystemProxy direct 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
            $code = $LASTEXITCODE
        } finally { $ErrorActionPreference = $eap }
        return @{ Code = $code ; Out = $out }
    }

    $sntpJuste = New-Object FakeSntp -ArgumentList @(60)
    $sntpFaux  = New-Object FakeSntp -ArgumentList @(18000)
    try {
        $sntpJuste.Start() ; $sntpFaux.Start()
        $h1 = Invoke-DiagHorloge ("127.0.0.1:" + $sntpJuste.Port)
        Check "diagnostic : ecart d'horloge de 1 min mesure et non bloquant" (($h1.Code -eq 0) -and ($h1.Out -match 'ecart mesure : \+1 min') -and ($h1.Out -match 'tolerance KMS 4 h') -and ($h1.Out -notmatch '\[BLOQUANT\]\s+Horloge')) ("code=" + $h1.Code + " ; " + ($h1.Out -replace "\s+", " "))
        $h2 = Invoke-DiagHorloge ("127.0.0.1:" + $sntpFaux.Port)
        Check "diagnostic : ecart d'horloge de 5 h -> BLOQUANT et code 1" (($h2.Code -eq 1) -and ($h2.Out -match '\[BLOQUANT\]\s+Horloge') -and ($h2.Out -match 'ecart mesure : \+5 h') -and ($h2.Out -match '0xC004F06C')) ("code=" + $h2.Code + " ; " + ($h2.Out -replace "\s+", " "))
    } finally {
        $sntpJuste.Stop() ; $sntpFaux.Stop()
    }

    $h3 = Invoke-DiagHorloge ("127.0.0.1:" + $portFerme)
    Check "diagnostic : source de temps injoignable -> avertissement, pas de blocage" (($h3.Code -eq 0) -and ($h3.Out -match '\[ATTENTION\]\s+Horloge') -and ($h3.Out -match 'aucune source de temps joignable') -and ($h3.Out -match 'non bloquant')) ("code=" + $h3.Code + " ; " + ($h3.Out -replace "\s+", " "))
}
catch {
    Write-Host ("ERREUR DU HARNAIS : " + $_.Exception.Message) -ForegroundColor Red
    $Fail++
}
finally {
    Remove-Item $Work -Recurse -Force -ErrorAction SilentlyContinue
    $env:FAKE_ODT_CODE = $null
    $env:FAKE_ODT_ARGS = $null
}

Write-Host ""
# Dernier controle : le nombre de cas annonce dans le README doit etre exactement
# celui qui vient d'etre joue. Ce chiffre a deja derive trois fois (60, 67, 69) :
# un cas ajoute sans mettre la documentation a jour echoue desormais tout seul.
$totalVise = $Pass + $Fail + 1   # ce controle compte lui aussi
$readme = Join-Path $PSScriptRoot "README.md"
$annonces = @()
if (Test-Path -LiteralPath $readme) {
    foreach ($ligneReadme in [System.IO.File]::ReadAllLines($readme)) {
        foreach ($m in [regex]::Matches($ligneReadme, '(\d+)\s+cas')) { $annonces += [int]$m.Groups[1].Value }
    }
}
$ecarts = @($annonces | Where-Object { $_ -ne $totalVise })
Check "README : nombre de cas annonce = total reellement execute" (($annonces.Count -gt 0) -and ($ecarts.Count -eq 0)) ("README annonce=" + (($annonces | Sort-Object -Unique) -join ", ") + " / execute=" + $totalVise + " - corriger README.md")

$total = $Pass + $Fail
if ($Fail -eq 0) {
    Write-Host ("RESULTAT : " + $Pass + "/" + $total + " controles passes.") -ForegroundColor Green
    exit 0
}
Write-Host ("RESULTAT : " + $Fail + " echec(s) sur " + $total + " controles.") -ForegroundColor Red
exit 1