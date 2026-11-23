# ================================================================
#  Apercu de la maquette XAML (style Fluent Windows 11)
# ================================================================
#  Charge les XAML de Maquette/ avec XamlReader - exactement le mecanisme
#  qu'utiliserait le script de production - et les anime avec des donnees
#  simulees. Le formulaire reprend les options du script : setup.exe, langue,
#  depot de sources, XML et applications, KMS (principal, secours, cle produit),
#  proxy (adresse, port, identifiants, exceptions).
#
#  En interactif, le pre-vol mesure reellement le poste (setup.exe, espace de C:,
#  version de Windows), les boutons de test ouvrent une connexion TCP vers le KMS ou
#  le proxy, et le bouton "Installer Office 2024" lance le vrai deroule : le worker
#  du moteur (Install-Office2024-KMS.ps1) est extrait par analyse syntaxique et
#  execute tel quel, sans passer par son interface WinForms. Un deploiement reel
#  demande les droits administrateur et refuse de partir si le pre-vol n'a pas
#  mesure un poste pret.
#
#  En capture, rien n'est deploye : le scenario simule fournit les valeurs de
#  reference, annoncees comme simulees dans le journal, pour que les PNG restent
#  reproductibles au pixel.
#
#  Les points de controle ne reprennent que les lignes que le worker marque OK ou
#  ERR : un point affiche est donc un resultat constate, jamais une intention.
#
#  Usage :
#    powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Apercu-Maquette.ps1
#    ... -Dialogue                       (ouvre directement le reliquat de proxy)
#    ... -Applications                   (ouvre directement les applications)
#    ... -Capture C:\tmp\ecran.png       (rend l'ecran en PNG et sort)
#    ... -Capture ... -Etape 6           (etat simule : index du scenario)
#    ... -Etats                         (planche des etats : boutons, champs, cases, radios)
#    ... -Capture ... -Etats             (rend la planche en PNG)
#    ... -Capture ... -Dialogue | -Applications
#    ... -Capture ... -Focus ChkProxy     (met ce controle en focus avant le rendu)
#    ... -Capture ... -Horloge 16:30:22   (heure du journal ; figee d'office en capture)
#    ... -Capture ... -Mesure            (mesure le poste au lieu des valeurs de reference)
# ================================================================
param(
    [string]$Capture = "",
    [switch]$Dialogue,
    [switch]$Applications,
    [switch]$Etats,
    [string]$Focus = "",
    [ValidatePattern("^(\d{2}:\d{2}:\d{2})?$")]
    [string]$Horloge = "",
    [switch]$Mesure,
    [int]$Etape = 7
)

# Journal : en capture l'heure est figee (16:30:22 par defaut, plus 7 s par ligne,
# ce qui fait tomber "Office installe : 16h32" a 16:32:00) pour que deux rendus de
# la maquette soient identiques au pixel ; en interactif elle suit l'horloge du poste.
$baseHorloge = $null
$heureJournal = if ($Horloge -ne "") { $Horloge } elseif ($Capture -ne "") { "16:30:22" } else { "" }
if ($heureJournal -ne "") {
    try {
        $baseHorloge = [DateTime]::ParseExact($heureJournal, "HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        throw ("Horloge illisible : -Horloge attend HH:mm:ss, recu '" + $heureJournal + "'.")
    }
}
$script:journalLignes = 0

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

$ici      = Split-Path -Parent $MyInvocation.MyCommand.Path
$dossier  = Join-Path $ici "Maquette"
$pinceaux = New-Object System.Windows.Media.BrushConverter
$policeTexte = New-Object System.Windows.Media.FontFamily("Segoe UI Variable Text, Segoe UI")
$policeIcone = New-Object System.Windows.Media.FontFamily("Segoe MDL2 Assets, Segoe UI")

# Exclusions par defaut : celles qu'annonce deja l'ecran (Teams, OneDrive, Skype).
# Ce sont les identifiants ODT, ceux que le moteur ecrit dans le XML.
$script:appsExclues = @("Teams", "OneDrive", "Lync")

# Seuil d'espace du moteur ($DISK_MIN_GB) : une seule copie ici, pour que le pre-vol,
# le refus de depart et la configuration envoyee au worker annoncent le meme chiffre.
$script:disqueMinGo = 4

$APPS = @(
    @{ Nom = "Word" ; Cle = "AppWord" ; Libelle = "Word" }
    @{ Nom = "Excel" ; Cle = "AppExcel" ; Libelle = "Excel" }
    @{ Nom = "PowerPoint" ; Cle = "AppPowerPoint" ; Libelle = "PowerPoint" }
    @{ Nom = "Outlook" ; Cle = "AppOutlook" ; Libelle = "Outlook" }
    @{ Nom = "OneNote" ; Cle = "AppOneNote" ; Libelle = "OneNote" }
    @{ Nom = "Access" ; Cle = "AppAccess" ; Libelle = "Access" }
    @{ Nom = "Teams" ; Cle = "AppTeams" ; Libelle = "Teams" }
    @{ Nom = "OneDrive" ; Cle = "AppOneDrive" ; Libelle = "OneDrive" }
    @{ Nom = "Lync" ; Cle = "AppLync" ; Libelle = "Skype Entreprise" }
)

function Import-Xaml {
    param([string]$Nom)
    $chemin = Join-Path $dossier $Nom
    if (-not (Test-Path -LiteralPath $chemin)) { throw "Fichier introuvable : $chemin" }
    # Les fenetres fusionnent Styles.xaml : une URI de base est necessaire pour que
    # Source="Styles.xaml" se resolve depuis un XAML non compile.
    $contexte = New-Object System.Windows.Markup.ParserContext
    $contexte.BaseUri = New-Object System.Uri($chemin)
    $flux = [System.IO.File]::OpenRead($chemin)
    try { return [System.Windows.Markup.XamlReader]::Load($flux, $contexte) } finally { $flux.Close() }
}

function Get-Pinceau {
    param([string]$Couleur)
    return $pinceaux.ConvertFromString($Couleur)
}

# Etats d'une etape : fond de la puce, couleur du glyphe, glyphe des etats finis.
# Le libelle accentue est porte par le XAML (cle Etat<Etat>), le script reste ASCII.
$renduEtapes = @{
    attente = @{ Fond = "#EDEDED" ; Texte = "#8A8A8A" }
    encours = @{ Fond = "#0F6CBD" ; Texte = "#FFFFFF" }
    termine = @{ Fond = "#DFF6DD" ; Texte = "#0F7B0F" }
    saute   = @{ Fond = "#F0F0F0" ; Texte = "#8A8A8A" }
    erreur  = @{ Fond = "#FDE7E9" ; Texte = "#C42B1C" }
}
$glyphes = @{
    termine = [string][char]0xE73E ; saute = [string][char]0xE738 ; erreur = "!"
}

# Etat du worker -> etat de la puce : une seule table, lue a chaque battement.
$etatsEtapes = @{ waiting = "attente" ; running = "encours" ; done = "termine" ; error = "erreur" ; skipped = "saute" }

# --- Pre-vol : mesures reelles ---------------------------------------------
# Meme logique que le moteur (Install-Office2024-KMS.ps1, Get-PreChecks) : le setup
# est cherche la ou le champ le dit, sinon a cote du script ; l'espace est celui de
# C: ; la version de Windows vient du registre. Le KMS n'est annonce joignable
# qu'apres une connexion TCP reussie.
function Resolve-SetupPath {
    param([string]$Saisi)
    if (-not [string]::IsNullOrWhiteSpace($Saisi)) { return $Saisi.Trim() }
    return (Join-Path $ici "setup.exe")
}

function Split-HotePort {
    # "kms.domaine.fr" prend le port par defaut ; "hote:1700" le sien ; "[::1]:1700".
    param([string]$Adresse, [int]$PortDefaut)
    $a = $Adresse.Trim()
    if ($a -eq "") { return $null }
    if ($a -match '^\[([^\]]+)\](?::(\d+))?$') {
        $port = if ($Matches[2]) { [int]$Matches[2] } else { $PortDefaut }
        return @{ Hote = $Matches[1] ; Port = $port }
    }
    if ($a -match '^(.+):(\d+)$') { return @{ Hote = $Matches[1] ; Port = [int]$Matches[2] } }
    return @{ Hote = $a ; Port = $PortDefaut }
}

function Test-PortTcp {
    # Connexion TCP reelle, duree mesuree. Un echec porte son motif : le texte
    # affiche ne peut donc jamais etre plus sur que la mesure elle-meme.
    param([string]$Hote, [int]$Port, [int]$Delai = 1500)
    $resultat = @{ Ok = $false ; Ms = 0 ; Motif = "" }
    # Adresse et port sont deja valides par les appelants (Split-HotePort rend $null sur
    # une adresse vide) : inutile de rejouer ce controle ici.
    $client = New-Object System.Net.Sockets.TcpClient
    $chrono = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $essai = $client.ConnectAsync($Hote, $Port)
        if (-not $essai.Wait($Delai)) { $resultat.Motif = ("delai depasse (" + $Delai + " ms)") }
        elseif ($client.Connected) { $resultat.Ok = $true }
        else { $resultat.Motif = "connexion refusee" }
    } catch {
        $resultat.Motif = $_.Exception.GetBaseException().Message
    } finally {
        $chrono.Stop()
        $resultat.Ms = [int]$chrono.ElapsedMilliseconds
        try { $client.Close() } catch { }
    }
    $resultat
}

function Get-VersionWindows {
    # Libelle court : "Windows 11 24H2 (26100)". Le moteur lit deja
    # CurrentMajorVersionNumber pour son seuil ; la build et la version d'affichage
    # etaient lues sans etre montrees, d'ou l'impression d'une version fausse.
    $majeure = 0 ; $build = "" ; $affichage = ""
    try {
        $cle = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop
        $majeure = [int]$cle.CurrentMajorVersionNumber
        $build = [string]$cle.CurrentBuildNumber
        $affichage = [string]$cle.DisplayVersion
    } catch { }
    $morceaux = @()
    if ($affichage -ne "") { $morceaux += $affichage }
    if ($build -ne "") { $morceaux += ("(" + $build + ")") }
    # Deux formes : Libelle pour le journal et le pied de carte, Court pour la ligne
    # de la carte, dont l'etiquette porte deja le mot Windows.
    $court = if ($majeure -ge 10) { (([string]$majeure + " " + ($morceaux -join " ")).Trim()) } else { "inconnu" }
    return @{ Ok = ($majeure -ge 10) ; Court = $court ; Libelle = ("Windows " + $court) }
}

function Get-PreVol {
    param([string]$SetupSaisi, [string]$KmsSaisi, [bool]$KmsActif)
    $setup  = Resolve-SetupPath $SetupSaisi
    $disque = Get-PSDrive C -ErrorAction SilentlyContinue
    $libre  = if ($disque) { [math]::Round($disque.Free / 1GB, 1) } else { 0 }
    $kms = $null
    if ($KmsActif) {
        $cible = Split-HotePort $KmsSaisi 1688
        if ($null -ne $cible) { $kms = @{ Cible = $cible ; Test = (Test-PortTcp $cible.Hote $cible.Port) } }
    }
    return @{
        Setup        = $setup
        # Resolve-SetupPath ne rend jamais une chaine vide : il retombe sur setup.exe
        # a cote du script.
        SetupPresent = (Test-Path -LiteralPath $setup)
        LibreGo      = $libre
        AssezDEspace = ($libre -ge $script:disqueMinGo)
        Os           = (Get-VersionWindows)
        Reseau       = [System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()
        Kms          = $kms
        Admin        = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
        Poste        = $env:COMPUTERNAME
    }
}

# Etats d'une ligne de pre-vol, comme ceux des etapes.
$etatsPreVol = @{
    ok      = @{ Glyphe = [string][char]0xE73E ; Couleur = "#0F7B0F" }
    encours = @{ Glyphe = [string][char]0xE895 ; Couleur = "#0F6CBD" }
    attente = @{ Glyphe = [string][char]0xE738 ; Couleur = "#8A8A8A" }
    erreur  = @{ Glyphe = [string][char]0xE711 ; Couleur = "#C42B1C" }
}

function Set-LignePreVol {
    param($Fenetre, [string]$Ligne, [string]$Etat, [string]$Valeur)
    $icone = $Fenetre.FindName("Ico$Ligne")
    if ($null -eq $icone) { throw "Ligne pre-vol '$Ligne' introuvable dans la fenetre." }
    $e = $etatsPreVol[$Etat]
    $icone.Text = $e.Glyphe
    $icone.Foreground = Get-Pinceau $e.Couleur
    # Nom distinct de $Valeur : PowerShell ne distingue pas la casse, et $valeur
    # ecraserait le parametre avant la ligne suivante.
    $champ = $Fenetre.FindName("Val$Ligne")
    if ($null -eq $champ) { throw "Valeur pre-vol '$Ligne' introuvable dans la fenetre." }
    $champ.Text = $Valeur
    # La valeur reste discrete quand tout va bien ; sinon elle prend la couleur de
    # l'erreur, pour que le probleme se lise sans chercher la ligne.
    $champ.Foreground = Get-Pinceau $(if ($Etat -eq "ok" -or $Etat -eq "attente") { "#5D5D5D" } else { $e.Couleur })
}

function Update-PreVol {
    # Remplit la carte pre-vol : mesures reelles par defaut ; valeurs de reference
    # avec -Simule, quand une capture doit rester reproductible au pixel.
    param($Fenetre, [switch]$Simule)
    if ($Simule) {
        # La parenthese autour de FindResource n'est pas decorative : sans elle, le
        # transtypage s'applique a la methode et la valeur affichee est son nom.
        Set-LignePreVol $Fenetre "Setup"   "ok" ([string]$Fenetre.FindResource("PreVolPresent"))
        Set-LignePreVol $Fenetre "Disque"  "ok" "212 Go libres"
        Set-LignePreVol $Fenetre "Kms"     "ok" "kms.ac-nancy.fr:1688, 42 ms"
        Set-LignePreVol $Fenetre "Windows" "ok" "11 24H2 (26100)"
        $Fenetre.FindName("InfoPoste").Text = "Administrateur, poste LORITZ-PC042, Windows 11 24H2"
        Add-Journal $Fenetre "avr" ("Pre-vol : " + [string]$Fenetre.FindResource("PreVolSimule"))
        return
    }
    $activerKms = [bool]$Fenetre.FindName("ChkKms").IsChecked
    $c = Get-PreVol -SetupSaisi $Fenetre.FindName("ChampSetup").Text `
                    -KmsSaisi $Fenetre.FindName("ChampKms").Text `
                    -KmsActif $activerKms

    if ($c.SetupPresent) { Set-LignePreVol $Fenetre "Setup" "ok" ([string]$Fenetre.FindResource("PreVolPresent")) }
    else { Set-LignePreVol $Fenetre "Setup" "erreur" ([string]$Fenetre.FindResource("PreVolAbsent")) }

    # Culture invariante : sinon un poste francais ecrirait "344,7" et la meme
    # mesure changerait de forme d'un poste a l'autre.
    $go = $c.LibreGo.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    if ($c.AssezDEspace) { Set-LignePreVol $Fenetre "Disque" "ok" ($go + " Go libres") }
    else { Set-LignePreVol $Fenetre "Disque" "erreur" ($go + " Go libres, 4 Go requis") }

    if ($null -eq $c.Kms) { Set-LignePreVol $Fenetre "Kms" "attente" ([string]$Fenetre.FindResource("PreVolNonTeste")) }
    elseif ($c.Kms.Test.Ok) { Set-LignePreVol $Fenetre "Kms" "ok" ($c.Kms.Cible.Hote + ":" + $c.Kms.Cible.Port + ", " + $c.Kms.Test.Ms + " ms") }
    else { Set-LignePreVol $Fenetre "Kms" "erreur" ($c.Kms.Cible.Hote + ":" + $c.Kms.Cible.Port + ", " + [string]$Fenetre.FindResource("PreVolInjoignable")) }

    if ($c.Os.Ok) { Set-LignePreVol $Fenetre "Windows" "ok" $c.Os.Court }
    else { Set-LignePreVol $Fenetre "Windows" "erreur" ($c.Os.Court + ", Windows 10+ requis") }

    $pied = @($(if ($c.Admin) { "Administrateur" } else { "Sans droits administrateur" }), ("poste " + $c.Poste), $c.Os.Libelle)
    if (-not $c.Reseau) { $pied += "aucune liaison reseau" }
    $Fenetre.FindName("InfoPoste").Text = ($pied -join ", ")

    Add-Journal $Fenetre "info" ("Pre-vol mesure : " + $c.Os.Libelle + ", " + $go + " Go libres sur C:.")
    Add-Journal $Fenetre $(if ($c.SetupPresent) { "ok" } else { "err" }) ("setup.exe " + $(if ($c.SetupPresent) { "present" } else { "absent" }) + " : " + $c.Setup)
    if ($null -eq $c.Kms) {
        Add-Journal $Fenetre "avr" "KMS non teste : adresse vide ou activation decochee."
    } elseif ($c.Kms.Test.Ok) {
        Add-Journal $Fenetre "ok" ("Test TCP du KMS " + $c.Kms.Cible.Hote + ":" + $c.Kms.Cible.Port + " : joignable en " + $c.Kms.Test.Ms + " ms.")
    } else {
        Add-Journal $Fenetre "err" ("Test TCP du KMS " + $c.Kms.Cible.Hote + ":" + $c.Kms.Cible.Port + " : injoignable (" + $c.Kms.Test.Motif + ").")
    }
    # Les mesures sont rendues a l'appelant : le bouton Installer refuse de partir sur
    # autre chose que ce qui vient d'etre constate.
    return $c
}

# --- Deploiement reel : le worker du moteur existant ------------------------
# Le moteur tient ses cinq etapes dans un scriptblock autonome ($script:BgDoWorkHandler)
# qui ne parle a son interface que par une table synchronisee. On l'extrait ici par
# analyse syntaxique, sans executer le moteur : c'est le meme code qui tourne, pas une
# reecriture. L'ancienne interface reste intacte jusqu'a ce que ce portage soit verifie.
$cheminMoteur = Join-Path $ici "Install-Office2024-KMS.ps1"
$script:deploiement = $null
$script:pompeMoteur = $null
$script:derniersFaits = @()

function Get-ScriptblockMoteur {
    # Le scriptblock de deploiement, extrait de la source du moteur. Une source qui
    # aurait change de forme s'arrete ici en le disant, plutot que de lancer autre chose.
    param([string]$Chemin)
    if (-not (Test-Path -LiteralPath $Chemin)) { throw ("Moteur introuvable : " + $Chemin) }
    $jetons = $null ; $erreurs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Chemin, [ref]$jetons, [ref]$erreurs)
    $trouves = @($ast.FindAll({
        param($n)
        ($n -is [System.Management.Automation.Language.AssignmentStatementAst]) -and
        ($n.Left.Extent.Text -eq '$script:BgDoWorkHandler')
    }, $true))
    if ($trouves.Count -ne 1) {
        throw ("Moteur : " + $trouves.Count + " affectation(s) de BgDoWorkHandler trouvee(s) dans " + $Chemin + ", une seule est attendue.")
    }
    # La valeur est un scriptblock, eventuellement enveloppe d'une expression : on le
    # cherche dans l'arbre plutot que de deviner la forme exacte de l'affectation.
    $bloc = $trouves[0].Right.Find({
        param($n)
        $n -is [System.Management.Automation.Language.ScriptBlockExpressionAst]
    }, $true)
    if ($null -eq $bloc) {
        throw ("Moteur : BgDoWorkHandler n'est plus un scriptblock dans " + $Chemin + ".")
    }
    # Corps seul : [scriptblock]::Create recevrait sinon une accolade qui ne fait que
    # rendre un scriptblock imbrique, sans jamais executer le worker. Un
    # ScriptBlockExpressionAst porte toujours ses deux accolades.
    $t = $bloc.Extent.Text.Trim()
    return $t.Substring(1, $t.Length - 2)
}

function Get-ConstanteMoteur {
    # Constante du moteur lue dans sa source (LOG_FOLDER...), pour ne pas la recopier.
    # Le fichier est deja valide par Get-ScriptblockMoteur : une constante absente est
    # une source qui a change de forme, et elle s'arrete ici en la nommant plutot que
    # de laisser l'apercu travailler sur un chemin invente.
    param([string]$Chemin, [string]$Nom)
    foreach ($l in [System.IO.File]::ReadAllLines($Chemin)) {
        if ($l -match ('^\$' + $Nom + '\s*=\s*"([^"]+)"')) { return $Matches[1] }
    }
    throw ("Moteur : constante $" + $Nom + " introuvable dans " + $Chemin + ".")
}

function Get-ConfigDeploiement {
    # Traduit les champs de l'ecran en l'objet que le worker du moteur attend. Les deux
    # valeurs mesurees par le pre-vol (setup.exe, adresse KMS) sont reprises telles
    # quelles : le worker recoit ce qui vient d'etre constate, pas une seconde lecture.
    param($Fenetre, $Sync, [string]$CheminMoteur, $PreVol)
    $dossier = Get-ConstanteMoteur $CheminMoteur "LOG_FOLDER"
    $mode = Get-SourceMode $Fenetre
    $xml = Join-Path $dossier "OfficeConfig2024.xml"
    return [pscustomobject]@{
        SyncHash        = $Sync
        ActivateOnly    = $false
        SetupExePath    = $PreVol.Setup
        InstallFolder   = $dossier
        ActualXMLPath   = $xml
        UseCustomXML    = [bool]$Fenetre.FindName("RadXmlFichier").IsChecked
        CustomXMLPath   = $Fenetre.FindName("ChampXml").Text.Trim()
        ExcludedApps    = @($script:appsExclues)
        LanguageID      = $Fenetre.FindName("LangueTexte").Text
        RemoveOld       = [bool]$Fenetre.FindName("ChkRemoveOld").IsChecked
        DoInstall       = [bool]$Fenetre.FindName("ChkInstall").IsChecked
        DoActivate      = [bool]$Fenetre.FindName("ChkKms").IsChecked
        KmsHost         = $(if ($null -ne $PreVol.Kms) { $PreVol.Kms.Cible.Hote } else { "" })
        KmsPort         = $(if ($null -ne $PreVol.Kms) { [string]$PreVol.Kms.Cible.Port } else { "1688" })
        KmsSecondary    = $Fenetre.FindName("ChampKmsSecours").Text.Trim()
        ProxyEnabled    = $mode.Proxy
        ProxyHost       = $Fenetre.FindName("ChampProxy").Text.Trim()
        ProxyPort       = $Fenetre.FindName("ChampPort").Text.Trim()
        ProxyRelay      = $false
        ProxyRelayTarget = ""
        ProxyBypass     = $Fenetre.FindName("ChampBypass").Text.Trim()
        DepotEnabled    = $mode.Depot
        DepotPath       = $Fenetre.FindName("ChampDepot").Text.Trim()
        SourcePath      = $dossier
        LocalDataPath   = $dossier
        DepotCopy       = $false
        ApplyProductKey = [bool]$Fenetre.FindName("ChkCle").IsChecked
        ProductKey      = $Fenetre.FindName("ChampCle").Text.Trim()
        DiskMinGB       = $script:disqueMinGo
    }
}

function Write-LigneMoteur {
    # Une ligne du worker, ramenee au journal de l'ecran : le worker ecrit
    # "[HH:mm:ss] [OK]  texte", l'horodatage et le niveau sont donc relus ici.
    param($Fenetre, [string]$Texte)
    $t = $Texte -replace '^\[\d\d:\d\d:\d\d\]', ''
    $niveau = if ($t -match '^---') { "etape" }
              elseif ($t -match '^\s*\[OK\]') { "ok" }
              elseif ($t -match '^\s*\[ERR\]') { "err" }
              elseif ($t -match '^\s*\[!\]') { "avr" }
              else { "info" }
    $t = ($t -replace '^(---\s*|\s*\[(OK|ERR|!)\]\s*)', '').Trim()
    Add-Journal $Fenetre $niveau $t
    if ($niveau -eq "etape") {
        # Une ligne ETAPE ouvre une etape : le worker l'ecrit avant ses premieres mesures,
        # donc les points de la precedente ne suivent pas celle-ci.
        $script:derniersFaits = @()
        Set-PointsEtape $Fenetre @()
    } elseif ($niveau -eq "ok" -or $niveau -eq "err") {
        # Points de controle : seulement les lignes que le worker a marquees OK ou ERR,
        # donc un resultat constate. Une ligne INFO ou [!] reste au journal : elle
        # annonce une intention, une progression ou une consigne, pas une mesure.
        $court = if ($t.Length -gt 42) { $t.Substring(0, 39) + "..." } else { $t }
        $etat = if ($niveau -eq "ok") { "ok" } else { "erreur" }
        $script:derniersFaits += ($etat + "|" + $court)
        if ($script:derniersFaits.Count -gt 3) { $script:derniersFaits = @($script:derniersFaits[-3..-1]) }
        Set-PointsEtape $Fenetre $script:derniersFaits
    }
}

function Stop-PompeMoteur {
    param($Fenetre)
    if ($null -ne $script:pompeMoteur) { $script:pompeMoteur.Stop() ; $script:pompeMoteur = $null }
    if ($null -ne $script:deploiement) {
        try { $script:deploiement.PS.Dispose() } catch { }
        try { $script:deploiement.Runspace.Close() ; $script:deploiement.Runspace.Dispose() } catch { }
        $script:deploiement = $null
    }
    $Fenetre.FindName("BtnInstaller").IsEnabled = $true
    Set-Activite $Fenetre $false
}

function Start-Deploiement {
    # Lance le worker du moteur, si et seulement si le pre-vol vient de mesurer un poste
    # pret : rien ne part sur une valeur supposee.
    param($Fenetre, [string]$CheminMoteur)
    if ($null -ne $script:deploiement) { Add-Journal $Fenetre "avr" "Un deploiement est deja en cours." ; return }

    $c = Update-PreVol $Fenetre
    # Table partagee avec le worker : seules les cles que l'ecran lit sont declarees, le
    # worker ajoute celles qu'il ecrit lui-meme (PopupIcon, StepStartTime).
    $sync = [hashtable]::Synchronized(@{
        LogQueue   = New-Object System.Collections.Concurrent.ConcurrentQueue[object]
        StepState  = [string[]]@("waiting","waiting","waiting","waiting","waiting")
        ActiveStep = -1
        StatusText = "Demarrage..."
        Done       = $false
        HasError   = $false
        FinalMsg   = ""
        PopupTitle = ""
        PopupMsg   = ""
        ShowPopup  = $false
    })

    $rs = $null ; $ps = $null
    try {
        $texte = Get-ScriptblockMoteur $CheminMoteur
        $cfg = Get-ConfigDeploiement $Fenetre $sync $CheminMoteur $c

        # Les refus portent sur la configuration qui part reellement : setup.exe et
        # l'adresse KMS sont celles que le pre-vol vient de constater.
        $refus = @()
        if (-not $c.SetupPresent) { $refus += ("setup.exe absent : " + $c.Setup) }
        if (-not $c.Admin) { $refus += "droits administrateur absents" }
        if (-not $c.AssezDEspace) { $refus += ($c.LibreGo.ToString([System.Globalization.CultureInfo]::InvariantCulture) + " Go libres, " + $script:disqueMinGo + " Go requis") }
        if (-not $c.Os.Ok) { $refus += "Windows 10+ requis" }
        if ($cfg.ProxyEnabled -and $Fenetre.FindName("ChampProxyUser").Text.Trim() -ne "") {
            $refus += "proxy authentifie : le relais local n'est pas encore porte (etape 2)"
        }
        if ($cfg.UseCustomXML -and $cfg.CustomXMLPath -eq "") { $refus += "fichier XML choisi mais chemin vide" }
        if ($refus.Count -gt 0) {
            Add-Journal $Fenetre "err" ("Deploiement refuse : " + ($refus -join " ; ") + ".")
            $Fenetre.FindName("StatutTexte").Text = "Deploiement refuse"
            $Fenetre.FindName("StatutDetail").Text = "Pre-vol non conforme"
            $Fenetre.FindName("PiedStatut").Text = "Deploiement refuse"
            return
        }

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.Open()
        $rs.SessionStateProxy.SetVariable("cfg", $cfg)
        $rs.SessionStateProxy.SetVariable("workerText", $texte)
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript({
            Add-Type -AssemblyName System.Drawing
            $worker = [scriptblock]::Create($workerText)
            $e = [System.ComponentModel.DoWorkEventArgs]::new($cfg)
            & $worker $null $e
        })
        $handle = $ps.BeginInvoke()
    } catch {
        # Rien n'a demarre : on referme ce qui a pu s'ouvrir avant de le dire.
        if ($null -ne $ps) { try { $ps.Dispose() } catch { } }
        if ($null -ne $rs) { try { $rs.Close() ; $rs.Dispose() } catch { } }
        Add-Journal $Fenetre "err" ("Deploiement impossible : " + $_.Exception.Message)
        return
    }

    $script:derniersFaits = @()
    # Etats : Reset-Scene vient de rendre "attente" sur les cinq puces, inutile de les
    # repeindre au premier battement.
    $script:deploiement = @{ PS = $ps ; Runspace = $rs ; Sync = $sync ; Handle = $handle ; Etats = @("waiting","waiting","waiting","waiting","waiting") ; Fin = $false ; Pourcent = -1 ; Actif = $true }
    $Fenetre.FindName("Journal").Children.Clear()
    Reset-Scene $Fenetre
    $Fenetre.FindName("BtnInstaller").IsEnabled = $false
    Set-Activite $Fenetre $true
    Add-Journal $Fenetre "info" "Deploiement lance : worker du moteur, cinq etapes."

    $script:pompeMoteur = New-Object System.Windows.Threading.DispatcherTimer
    $script:pompeMoteur.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:pompeMoteur.Add_Tick({
        $fen = $script:fenetre
        $d = $script:deploiement
        if ($null -eq $d) { return }
        $sync = $d.Sync

        # Les etats d'abord : la puce de l'etape est a jour avant que ses lignes
        # n'entrent au journal.
        $faits = 0 ; $actif = $false
        for ($i = 0; $i -lt 5; $i++) {
            $etat = [string]$sync.StepState[$i]
            if ($etat -ne $d.Etats[$i]) {
                Set-Etape $fen ($i + 1) $etatsEtapes[$etat]
                $d.Etats[$i] = $etat
            }
            if ($etat -eq "running") { $actif = $true }
            # Une etape en erreur n'est pas une etape faite : elle ne compte pas dans la jauge.
            if ($etat -eq "done" -or $etat -eq "skipped") { $faits++ }
        }

        $entree = $null ; $n = 0
        while ($sync.LogQueue.TryDequeue([ref]$entree) -and $n -lt 40) { Write-LigneMoteur $fen $entree.Text ; $n++ }
        # Une seule bascule par changement : relancer la minuterie a chaque battement
        # remettrait la bande et l'anneau sur place.
        if ($actif -ne $d.Actif) { $d.Actif = $actif ; Set-Activite $fen $actif }

        $pct = [int](100 * $faits / 5)
        if ($sync.Done -and -not $sync.HasError) { $pct = 100 }
        if ($pct -ne $d.Pourcent) {
            $d.Pourcent = $pct
            $fen.FindName("Jauge").Value = $pct
            $fen.FindName("EtapePourcent").Text = ("{0} %" -f $pct)
        }
        if ($sync.StatusText) {
            $fen.FindName("StatutTexte").Text = [string]$sync.StatusText
            $fen.FindName("PiedStatut").Text = [string]$sync.StatusText
        }
        $fen.FindName("StatutDetail").Text = if ([int]$sync.ActiveStep -ge 0) { "Etape " + ([int]$sync.ActiveStep + 1) + " sur 5" } else { ($faits.ToString() + " etapes sur 5") }

        if ($sync.ShowPopup) {
            Add-Journal $fen "avr" (([string]$sync.PopupTitle) + " : " + ([string]$sync.PopupMsg))
            $sync.ShowPopup = $false
        }

        if (-not $sync.Done -and $d.Handle.IsCompleted) {
            $sync.HasError = $true
            $sync.FinalMsg = "Worker interrompu sans message."
            $sync.Done = $true
        }
        if ($sync.Done -and -not $d.Fin) {
            $d.Fin = $true
            if ($sync.HasError) {
                $final = [string]$sync.FinalMsg
                if ($final -eq "") { $final = "Deploiement interrompu sans message du moteur." }
                Add-Journal $fen "err" $final
                $fen.FindName("StatutTexte").Text = "Echec"
                $fen.FindName("PiedStatut").Text = "Echec"
            } else {
                if ([string]$sync.FinalMsg -ne "") { Add-Journal $fen "ok" ([string]$sync.FinalMsg) }
                $fen.FindName("StatutTexte").Text = "Deploiement termine"
                $fen.FindName("PiedStatut").Text = "Deploiement termine"
                $fen.FindName("Jauge").Value = 100
                $fen.FindName("EtapePourcent").Text = "100 %"
            }
            Stop-PompeMoteur $fen
        }
    })
    $script:pompeMoteur.Start()
}

# --- Animations : seulement en interactif ----------------------------------
# En capture, aucune animation n'est demarree : la reference reste identique au
# pixel d'un rendu a l'autre.
$script:minuterieActivite = $null
$script:etapeEncours = 0
# Seul l'interactif anime : la capture doit rester reproductible au pixel.
$script:animer = ($Capture -eq "")

function Set-Activite {
    # Barre d'activite et anneau de l'etape en cours : deux mouvements, une seule
    # minuterie. Elle est creee a la premiere activation, et rien n'est demarre en
    # capture, ou l'image doit rester identique d'un rendu a l'autre.
    param($Fenetre, [bool]$Actif)
    $Fenetre.FindName("BandeActivite").Visibility = if ($Actif) { "Visible" } else { "Collapsed" }
    if (-not $script:animer) { return }
    if ($null -eq $script:minuterieActivite) {
        if (-not $Actif) { return }
        $script:minuterieActivite = New-Object System.Windows.Threading.DispatcherTimer
        $script:minuterieActivite.Interval = [TimeSpan]::FromMilliseconds(40)
        $script:minuterieActivite.Add_Tick({
            $fen = $script:fenetre
            if ($null -eq $fen) { return }
            $segment = $fen.FindName("SegmentActivite")
            $course = [math]::Max(1, $fen.FindName("BandeActivite").ActualWidth - $segment.Width)
            $segment.RenderTransform.X = (($segment.RenderTransform.X + 16) % $course)
            $glyphe = $fen.FindName("Glyphe$($script:etapeEncours)")
            if ($null -ne $glyphe -and $null -ne $glyphe.RenderTransform) { $glyphe.RenderTransform.Angle = ($glyphe.RenderTransform.Angle + 30) % 360 }
        })
    }
    if ($Actif) { $script:minuterieActivite.Start() } else { $script:minuterieActivite.Stop() }
}

function Set-PointsEtape {
    # Points de verification de l'etape en cours : un glyphe par point, vert quand
    # il est franchi, bleu pendant la mesure, gris avant.
    param($Fenetre, [string[]]$Points)
    $zone = $Fenetre.FindName("PointsEtape")
    $zone.Children.Clear()
    foreach ($p in $Points) {
        $bout = ([string]$p).Split("|", 2)
        if ($bout.Count -lt 2) { throw "Point de controle illisible : '$p' (attendu etat|texte)." }
        $e = $etatsPreVol[$bout[0]]
        $bloc = New-Object System.Windows.Controls.StackPanel
        $bloc.Orientation = "Horizontal"
        $bloc.Margin = New-Object System.Windows.Thickness(0, 0, 14, 0)
        $icone = New-Object System.Windows.Controls.TextBlock
        $icone.Text = $e.Glyphe
        $icone.FontFamily = $policeIcone
        $icone.FontSize = 11
        $icone.Foreground = Get-Pinceau $e.Couleur
        $icone.VerticalAlignment = "Center"
        $texte = New-Object System.Windows.Controls.TextBlock
        $texte.Text = $bout[1]
        $texte.FontSize = 11
        $texte.Margin = New-Object System.Windows.Thickness(5, 0, 0, 0)
        $texte.Foreground = Get-Pinceau $(if ($bout[0] -eq "attente") { "#8A8A8A" } else { $e.Couleur })
        $texte.VerticalAlignment = "Center"
        [void]$bloc.Children.Add($icone)
        [void]$bloc.Children.Add($texte)
        [void]$zone.Children.Add($bloc)
    }
}

function Set-Etape {
    param($Fenetre, [int]$Index, [string]$Etat)
    $e = $renduEtapes[$Etat]
    $puce   = $Fenetre.FindName("Etape$Index")
    $glyphe = $Fenetre.FindName("Glyphe$Index")
    $titre  = $Fenetre.FindName("Titre$Index")
    $info   = $Fenetre.FindName("Etat$Index")
    $puce.Background = Get-Pinceau $e.Fond
    $glyphe.Foreground = Get-Pinceau $e.Texte
    if ($Etat -eq "encours") {
        # Anneau de chargement dans la puce : le glyphe tourne pendant l'etape.
        $glyphe.Text = [string][char]0xE895
        $glyphe.FontFamily = $policeIcone
        $glyphe.FontWeight = "Normal"
        $glyphe.FontSize = 12
        $glyphe.RenderTransformOrigin = New-Object System.Windows.Point(0.5, 0.5)
        $glyphe.RenderTransform = New-Object System.Windows.Media.RotateTransform
        $script:etapeEncours = $Index
    } elseif ($Etat -eq "attente") {
        $glyphe.Text = [string]$Index
        $glyphe.FontFamily = $policeTexte
        $glyphe.FontWeight = "SemiBold"
        $glyphe.FontSize = 12
    } else {
        $glyphe.Text = $glyphes[$Etat]
        $glyphe.FontFamily = $policeIcone
        $glyphe.FontWeight = "Normal"
        $glyphe.FontSize = 11
    }
    # Une puce qui quitte l'etat "encours" abandonne sa rotation : sans cela le
    # glyphe termine resterait incline de l'angle ou l'animation l'a laisse, et la
    # minuterie continuerait de tourner sur une etape qui n'avance plus.
    if ($Etat -ne "encours") {
        $glyphe.RenderTransform = $null
        if ($script:etapeEncours -eq $Index) { $script:etapeEncours = 0 }
    }
    $titre.Foreground = Get-Pinceau $(if ($Etat -eq "attente" -or $Etat -eq "saute") { "#1B1B1B" } else { $e.Texte })
    $cle = "Etat" + $Etat.Substring(0, 1).ToUpper() + $Etat.Substring(1)
    $info.Text = [string]$Fenetre.FindResource($cle)
    $info.Foreground = Get-Pinceau $(if ($Etat -eq "attente" -or $Etat -eq "saute") { "#5D5D5D" } else { $e.Texte })
}

function Add-Journal {
    param($Fenetre, [string]$Niveau, [string]$Texte)
    $couleur = switch ($Niveau) {
        "ok"    { "#0F7B0F" }
        "avr"   { "#9D5D00" }
        "err"   { "#C42B1C" }
        "etape" { "#115EA3" }
        default { "#3C3C3C" }
    }
    $ligne = New-Object System.Windows.Controls.TextBlock
    $ligne.FontSize = 12
    $ligne.FontFamily = New-Object System.Windows.Media.FontFamily("Cascadia Mono, Consolas")
    $ligne.Foreground = Get-Pinceau $couleur
    $ligne.Margin = New-Object System.Windows.Thickness(0, 1, 0, 1)
    $horodatage = Get-Date
    if ($null -ne $baseHorloge) {
        $horodatage = $baseHorloge.AddSeconds(7 * $script:journalLignes)
        $script:journalLignes++
    }
    $heure = $horodatage.ToString("HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
    $ligne.Text = ("{0}  {1}" -f $heure, $Texte)
    [void]$Fenetre.FindName("Journal").Children.Add($ligne)
    $Fenetre.FindName("Defilement").ScrollToEnd()
}

function Set-Actif {
    param($Fenetre, [string[]]$Noms, [bool]$Actif)
    foreach ($n in $Noms) { $Fenetre.FindName($n).IsEnabled = $Actif }
}

function Get-SourceMode {
    # Regle de l'ecran : un depot de sources supprime le telechargement, donc le proxy.
    # Elle est lue aux trois endroits qui en dependent : les champs actifs, la
    # configuration envoyee au worker et les refus de depart.
    param($Fenetre)
    $depot = [bool]$Fenetre.FindName("ChkDepot").IsChecked
    return @{ Depot = $depot ; Proxy = (-not $depot) -and [bool]$Fenetre.FindName("ChkProxy").IsChecked }
}

function Update-Options {
    param($Fenetre)
    $mode = Get-SourceMode $Fenetre
    Set-Actif $Fenetre @("ChampDepot", "BtnDepot") $mode.Depot
    $Fenetre.FindName("ChkProxy").IsEnabled = -not $mode.Depot
    Set-Actif $Fenetre @("ChampProxy", "ChampPort", "ChampProxyUser", "ChampProxyPass", "ChampBypass", "BtnTestProxy") $mode.Proxy
    $kms = [bool]$Fenetre.FindName("ChkKms").IsChecked
    Set-Actif $Fenetre @("ChampKms", "ChampKmsSecours", "BtnTestKms") $kms
    Set-Actif $Fenetre @("ChampCle") ([bool]$Fenetre.FindName("ChkCle").IsChecked)
    $auto = [bool]$Fenetre.FindName("RadXmlAuto").IsChecked
    Set-Actif $Fenetre @("BtnApps", "LblApps") $auto
    Set-Actif $Fenetre @("ChampXml", "BtnXml") (-not $auto)
}

# Scenario simule. Chaque scene ne porte que ce qu'elle change : les etapes qui
# basculent (P), la jauge, le statut, le detail et les lignes "niveau|texte".
$scenario = @(
    @{ P = @{ 1 = "encours" } ; J = 4 ; S = "Recherche d'une installation existante..." ; D = "Etape 1 sur 5" ;
       C = @("encours|ospp.vbs /dstatus", "attente|produits detectes", "attente|desinstallation") ;
       L = @("etape|Etape 1/5 - Desinstallation", "info|Detection des produits Office installes (ospp.vbs /dstatus).") }
    @{ P = @{ 1 = "saute" ; 2 = "encours" } ; J = 14 ; S = "Verifications pre-vol" ; D = "Etape 2 sur 5" ;
       C = @("ok|setup.exe", "ok|espace disque", "encours|KMS", "attente|Windows") ;
       L = @("ok|Aucune version precedente a retirer : etape ignoree.", "etape|Etape 2/5 - Preparation",
             "info|Pre-vol : setup.exe present, 212 Go libres, Windows 11 24H2 (26100).",
             "info|Test TCP du KMS kms.ac-nancy.fr:1688 : joignable en 42 ms.") }
    @{ P = @{ 2 = "termine" ; 3 = "encours" } ; J = 28 ; S = "Copie des sources depuis le depot" ; D = "Etape 3 sur 5" ;
       C = @("ok|dossier cree", "encours|copie", "attente|6,8 Go verifies") ;
       L = @("ok|Configuration ODT generee (ProPlus2024Volume, PerpetualVL2024).", "etape|Etape 3/5 - Sources", "info|robocopy \\serveur\office -> C:\Office2024\sources") }
    @{ J = 46 ; S = "Copie des sources depuis le depot" ; D = "Etape 3 sur 5" ;
       L = @("info|4,1 Go copies sur 6,8 Go (60 %).") }
    @{ P = @{ 3 = "termine" ; 4 = "encours" } ; J = 66 ; S = "Installation Click-to-Run en cours" ; D = "Etape 4 sur 5" ;
       C = @("ok|Click-to-Run", "encours|applications", "attente|configuration") ;
       L = @("ok|Sources pretes : 6,8 Go dans C:\Office2024\sources.", "etape|Etape 4/5 - Installation", "info|setup.exe /configure OfficeConfig2024.xml (silencieux).") }
    @{ J = 82 ; S = "Installation Click-to-Run en cours" ; D = "Etape 4 sur 5" ;
       L = @("info|Click-to-Run : installation des applications Office...") }
    @{ P = @{ 4 = "termine" ; 5 = "encours" } ; J = 94 ; S = "Activation KMS" ; D = "Etape 5 sur 5" ;
       C = @("encours|ospp.vbs /sethst", "attente|activation", "attente|licence") ;
       L = @("ok|Office installe : 16h32.", "etape|Etape 5/5 - Activation KMS", "info|ospp.vbs /sethst:kms.ac-nancy.fr:1688") }
    @{ P = @{ 5 = "termine" } ; J = 100 ; S = "Deploiement termine" ; D = "5 etapes sur 5" ;
       C = @("ok|ospp.vbs /sethst", "ok|activation", "ok|licence") ;
       L = @("ok|Activation KMS reussie : licence perpetuelle accordee.", "ok|Configuration proxy restauree (acces direct), port du relais libere.", "ok|Deploiement termine sans erreur.") }
)

function Set-Scene {
    param($Fenetre, [int]$Index)
    if ($Index -lt 0 -or $Index -ge $scenario.Count) { return }
    $s = $scenario[$Index]
    $actif = $false
    if ($s.ContainsKey("P")) {
        foreach ($n in @($s.P.Keys)) {
            Set-Etape $Fenetre $n $s.P[$n]
            if ($s.P[$n] -eq "encours") { $actif = $true }
        }
    }
    Set-Activite $Fenetre $actif
    if ($s.ContainsKey("C")) { Set-PointsEtape $Fenetre $s.C }
    $Fenetre.FindName("Jauge").Value = $s.J
    $Fenetre.FindName("EtapePourcent").Text = ("{0} %" -f $s.J)
    $Fenetre.FindName("StatutTexte").Text = $s.S
    $Fenetre.FindName("StatutDetail").Text = $s.D
    $Fenetre.FindName("PiedStatut").Text = $s.S
    foreach ($l in $s.L) {
        $p = ([string]$l).Split("|", 2)
        Add-Journal $Fenetre $p[0] $p[1]
    }
}

function Reset-Scene {
    param($Fenetre)
    for ($i = 1; $i -le 5; $i++) { Set-Etape $Fenetre $i "attente" }
    $Fenetre.FindName("Jauge").Value = 0
    $Fenetre.FindName("EtapePourcent").Text = "0 %"
    Set-Activite $Fenetre $false
    Set-PointsEtape $Fenetre @()
}

function Get-TailleCapture {
    # Taille de la capture, lue dans le XAML de la fenetre ; la planche n'y
    # declare pas de hauteur (SizeToContent), elle se mesure alors sur le
    # contenu, a la largeur de la fenetre, ses colonnes etant proportionnelles.
    # Arrondi au pixel superieur : une fraction de pixel de plus rognerait le bas.
    param($Fenetre)
    if ([double]::IsNaN($Fenetre.Width)) {
        throw "Capture sans largeur : la fenetre doit declarer Width."
    }
    $largeur = [int][Math]::Ceiling($Fenetre.Width)
    $hauteur = $Fenetre.Height
    if ([double]::IsNaN($hauteur)) {
        $Fenetre.Content.Measure((New-Object System.Windows.Size($largeur, [double]::PositiveInfinity)))
        $hauteur = $Fenetre.Content.DesiredSize.Height
    }
    return (New-Object System.Windows.Size($largeur, [int][Math]::Ceiling($hauteur)))
}

function Export-Fenetre {
    # Rendu hors ecran de la fenetre ; la taille vient de Get-TailleCapture.
    param($Fenetre, [string]$Chemin)
    $taille  = Get-TailleCapture $Fenetre
    $largeur = [int]$taille.Width
    $hauteur = [int]$taille.Height
    $contenu = $Fenetre.Content
    $contenu.Measure($taille)
    $contenu.Arrange((New-Object System.Windows.Rect(0, 0, $largeur, $hauteur)))
    $contenu.UpdateLayout()
    $bmp = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($largeur, $hauteur, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $bmp.Render($contenu)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bmp))
    $flux = [System.IO.File]::Open($Chemin, [System.IO.FileMode]::Create)
    try { $enc.Save($flux) } finally { $flux.Close() }
}

function Show-Reliquat {
    param($Fenetre)
    $dlg = Import-Xaml "ReliquatProxy.xaml"
    $script:dlgReliquat = $dlg
    $dlg.Owner = $Fenetre
    $dlg.FindName("BtnRestaurer").Add_Click({
        Add-Journal $script:fenetre "ok" "Reliquat supprime : acces direct retabli (WinHTTP, session)."
        $script:dlgReliquat.DialogResult = $true
    })
    $dlg.FindName("BtnConserver").Add_Click({
        Add-Journal $script:fenetre "avr" "Reliquat conserve : proxy 127.0.0.1:52144 (port mort)."
        $script:dlgReliquat.DialogResult = $false
    })
    [void]$dlg.ShowDialog()
    $script:dlgReliquat = $null
}

# Exclues selon les cases du dialogue des applications (6 incluses, 3 exclues au depart).
function Get-ApplicationsExclues {
    param($Dlg)
    $exclues = @()
    foreach ($a in $APPS) {
        if (-not $Dlg.FindName($a.Cle).IsChecked) { $exclues += $a.Nom }
    }
    return $exclues
}

function Show-Applications {
    param($Fenetre)
    $dlg = Import-Xaml "Applications.xaml"
    $script:dlgApps = $dlg
    $dlg.Owner = $Fenetre
    $maj = {
        $exclues = @(Get-ApplicationsExclues $script:dlgApps)
        $dlg2 = $script:dlgApps
        $dlg2.FindName("ResumeApps").Text = ("{0} applications incluses, {1} exclues" -f ($APPS.Count - $exclues.Count), $exclues.Count)
    }
    foreach ($a in $APPS) { $dlg.FindName($a.Cle).Add_Click($maj) }
    & $maj
    $dlg.FindName("BtnAppliquer").Add_Click({
        $exclues = @(Get-ApplicationsExclues $script:dlgApps)
        # Les identifiants ODT pilotent le XML ; l'etiquette, elle, montre les libelles.
        $script:appsExclues = $exclues
        $libelles = @($APPS | Where-Object { $exclues -contains $_.Nom } | ForEach-Object { $_.Libelle })
        $texte = if ($exclues.Count -eq 0) { "Toutes les applications incluses" } else { "Exclues : " + ($libelles -join ", ") }
        $script:fenetre.FindName("LblApps").Text = $texte
        Add-Journal $script:fenetre "info" ("Applications retenues - " + $(if ($exclues.Count -eq 0) { "aucune exclusion" } else { "exclues : " + ($exclues -join ", ") }))
        $script:dlgApps.DialogResult = $true
    })
    $dlg.FindName("BtnAnnuler").Add_Click({ $script:dlgApps.DialogResult = $false })
    [void]$dlg.ShowDialog()
    $script:dlgApps = $null
}

# IsMouseOver et IsPressed sont des proprietes en lecture seule, exposees avec leur
# cle : les forcer fait evaluer les vrais declencheurs des styles partages.
$drapeauxStatiques = [System.Reflection.BindingFlags]"NonPublic,Public,Static"

function Get-CleEtat {
    # WPF ne publie la cle d'une propriete en lecture seule que par un champ interne
    # du type qui la declare. Rien d'autre que la cle de cette propriete precise
    # n'est accepte : une cle nulle ou etrangere n'atteint jamais SetValue.
    param([System.Windows.DependencyProperty]$Propriete, [string]$Champ)
    $type = $Propriete.OwnerType
    $reflexion = $type.GetField($Champ, $drapeauxStatiques)
    if ($null -eq $reflexion) {
        throw ("Planche des etats impossible : le champ interne '$Champ' est introuvable sur " +
               "$($type.FullName). Ce runtime WPF ne permet pas de forcer l'etat des boutons ; " +
               "lancer l'apercu sans -Etats.")
    }
    $cle = $reflexion.GetValue($null)
    if ($cle -isnot [System.Windows.DependencyPropertyKey]) {
        $obtenu = if ($null -eq $cle) { "nul" } else { $cle.GetType().FullName }
        throw ("Planche des etats impossible : '$Champ' sur $($type.FullName) ne contient pas " +
               "une DependencyPropertyKey (obtenu : $obtenu).")
    }
    if ($cle.DependencyProperty -ne $Propriete) {
        throw ("Planche des etats impossible : la cle '$Champ' porte sur " +
               "$($cle.DependencyProperty.Name), pas sur $($Propriete.Name).")
    }
    $cle
}

# Chaque etat se force sur une propriete en lecture seule precise : c'est le seul
# endroit ou les noms de champs internes de WPF apparaissent.
$definitionCles = @{
    IsMouseOver           = @{ Propriete = [System.Windows.UIElement]::IsMouseOverProperty ; Champ = "IsMouseOverPropertyKey" }
    IsPressed             = @{ Propriete = [System.Windows.Controls.Primitives.ButtonBase]::IsPressedProperty ; Champ = "IsPressedPropertyKey" }
    IsKeyboardFocusWithin = @{ Propriete = [System.Windows.UIElement]::IsKeyboardFocusWithinProperty ; Champ = "IsKeyboardFocusWithinPropertyKey" }
    IsKeyboardFocused     = @{ Propriete = [System.Windows.UIElement]::IsKeyboardFocusedProperty ; Champ = "IsKeyboardFocusedPropertyKey" }
}
$clesResolues = @{}

function Get-CleEtatNommee {
    # Resout une cle a la premiere demande et la garde : un etat qu'aucune variante
    # ne porte ne declenche aucune lecture de champ interne.
    param([string]$Etat)
    if (-not $clesResolues.ContainsKey($Etat)) {
        $definition = $definitionCles[$Etat]
        if ($null -eq $definition) {
            throw "Etat '$Etat' inconnu : la table de la planche demande un etat absent de definitionCles."
        }
        $clesResolues[$Etat] = Get-CleEtat $definition.Propriete $definition.Champ
    }
    $clesResolues[$Etat]
}

# Une seule table decrit la planche : par variante, l'etat d'appui (boutons
# seulement) et l'etat de focus que surveille son gabarit. Une valeur vide dit que
# la cellule n'existe pas dans Etats.xaml.
# Une cellule "desactivee" ne force aucune propriete : ce marqueur la distingue.
$desactive = "eteint"

# Une seule table decrit la planche : par variante, chaque cellule porte son suffixe
# et les etats a y forcer. Les suffixes sont donc exactement les cles de ces lignes,
# et le garde comme le forcage les lisent ici.
$tablePlanche = @(
    @{ Nom = "Sec"    ; Survol = "IsMouseOver" ; Appui = "IsMouseOver,IsPressed" ; Focus = "IsKeyboardFocused"    ; Off = $desactive }
    @{ Nom = "Pri"    ; Survol = "IsMouseOver" ; Appui = "IsMouseOver,IsPressed" ; Focus = "IsKeyboardFocused"    ; Off = $desactive }
    @{ Nom = "DlgSec" ; Survol = "IsMouseOver" ; Appui = "IsMouseOver,IsPressed" ; Focus = "IsKeyboardFocused"    ; Off = $desactive }
    @{ Nom = "DlgPri" ; Survol = "IsMouseOver" ; Appui = "IsMouseOver,IsPressed" ; Focus = "IsKeyboardFocused"    ; Off = $desactive }
    @{ Nom = "Fen"    ; Survol = "IsMouseOver" ; Appui = "IsMouseOver,IsPressed" ; Focus = "IsKeyboardFocused" ; Off = $desactive }
    @{ Nom = "Sai"    ; Survol = "IsMouseOver" ; Focus = "IsKeyboardFocusWithin" ; Off = $desactive }
    @{ Nom = "Secr"   ; Survol = "IsMouseOver" ; Focus = "IsKeyboardFocusWithin" ; Off = $desactive }
    @{ Nom = "Chk"    ; Survol = "IsMouseOver" ; Focus = "IsKeyboardFocused"      ; Off = $desactive }
    @{ Nom = "ChkC"   ; Survol = "IsMouseOver" ; Focus = "IsKeyboardFocused"      ; Off = $desactive }
    @{ Nom = "Opt"    ; Survol = "IsMouseOver" ; Focus = "IsKeyboardFocused"      ; Off = $desactive }
    @{ Nom = "OptC"   ; Survol = "IsMouseOver" ; Focus = "IsKeyboardFocused"      ; Off = $desactive }
)

function Get-Cellule {
    # Une cellule renommee dans Etats.xaml s'arrete ici, en la nommant, et non
    # plus loin sur un appel de methode a $null.
    param($Fenetre, [string]$Nom)
    $cellule = $Fenetre.FindName($Nom)
    if ($null -eq $cellule) {
        throw "Cellule '$Nom' introuvable dans la planche : ce nom a-t-il change dans Etats.xaml ?"
    }
    $cellule
}

function Get-NomsSousArbre {
    # Noms portes par les descendants d'un element, parent exclu.
    param($Element)
    $noms = @()
    foreach ($enfant in [System.Windows.LogicalTreeHelper]::GetChildren($Element)) {
        if ($enfant -isnot [System.Windows.FrameworkElement]) { continue }
        if ($enfant.Name) { $noms += $enfant.Name }
        $noms += Get-NomsSousArbre $enfant
    }
    $noms
}

function Set-PlancheEtats {
    param($Fenetre)
    # Cellules que la table decrit, validees avant tout forcage : une cellule
    # renommee s'arrete sur un message qui la nomme.
    $attendues = @()
    $suffixes = @()
    foreach ($ligne in $tablePlanche) {
        foreach ($suffixe in $ligne.Keys) {
            if ($suffixe -eq "Nom") { continue }
            $attendues += "$($ligne.Nom)_$suffixe"
            if ($suffixes -notcontains $suffixe) { $suffixes += $suffixe }
        }
    }
    foreach ($nom in $attendues) { [void](Get-Cellule $Fenetre $nom) }
    # Et ce que la grille porte en plus : un nom d'une autre forme (un libelle) n'est
    # pas une cellule et ne bloque rien.
    $grille = $Fenetre.FindName("GrilleCellules")
    if ($null -eq $grille) {
        throw "Grille 'GrilleCellules' introuvable : son nom a-t-il change dans Etats.xaml ?"
    }
    $motif = "_($($suffixes -join '|'))$"
    $orphelines = @(Get-NomsSousArbre $grille |
        Where-Object { $_ -match $motif -and $attendues -notcontains $_ })
    if ($orphelines.Count -gt 0) {
        throw ("Cellule(s) de la planche hors table : " + ($orphelines -join ", ") +
               ". Ajouter la variante a tablePlanche, sinon son etat ne serait jamais force.")
    }
    foreach ($ligne in $tablePlanche) {
        foreach ($suffixe in $ligne.Keys) {
            if ($suffixe -eq "Nom") { continue }
            $cellule = Get-Cellule $Fenetre "$($ligne.Nom)_$suffixe"
            $etats = $ligne[$suffixe]
            if ($etats -eq $desactive) {
                $cellule.IsEnabled = $false
                continue
            }
            foreach ($etat in $etats -split ",") {
                $cellule.SetValue((Get-CleEtatNommee $etat), $true)
            }
        }
    }
}

function Set-FocusCapture {
    # Met un controle nomme en focus avant une capture, pour voir l'anneau en
    # contexte. Le nom est verifie : une faute de frappe s'arrete ici.
    param($Fenetre, [string]$Nom)
    $controle = $Fenetre.FindName($Nom)
    if ($null -eq $controle) {
        throw "Controle '$Nom' introuvable dans la fenetre capturee : -Focus demande un x:Name existant."
    }
    $controle.SetValue((Get-CleEtatNommee "IsKeyboardFocusWithin"), $true)
    $controle.SetValue((Get-CleEtatNommee "IsKeyboardFocused"), $true)
}

# --- Mode capture : rendu hors ecran, aucun affichage ----------------------
if ($Capture -ne "") {
    if ($Dialogue) {
        $dlg = Import-Xaml "ReliquatProxy.xaml"
        if ($Focus -ne "") { Set-FocusCapture $dlg $Focus }
        Export-Fenetre $dlg $Capture
    } elseif ($Applications) {
        $dlgApps = Import-Xaml "Applications.xaml"
        if ($Focus -ne "") { Set-FocusCapture $dlgApps $Focus }
        Export-Fenetre $dlgApps $Capture
    } elseif ($Etats) {
        $planche = Import-Xaml "Etats.xaml"
        Set-PlancheEtats $planche
        Export-Fenetre $planche $Capture
    } else {
        $win = Import-Xaml "Principal.xaml"
        # Capture : valeurs de reference, sauf -Mesure qui interroge vraiment le poste.
        if ($Mesure) { Update-PreVol $win } else { Update-PreVol $win -Simule }
        for ($k = 0; $k -le $Etape; $k++) { Set-Scene $win $k }
        if ($Focus -ne "") { Set-FocusCapture $win $Focus }
        Export-Fenetre $win $Capture
    }
    Write-Host ("Capture ecrite : " + $Capture)
    return
}

# --- Mode interactif -------------------------------------------------------
if ($Etats) {
    $planche = Import-Xaml "Etats.xaml"
    Set-PlancheEtats $planche
    [void]$planche.ShowDialog()
    return
}

$script:fenetre = Import-Xaml "Principal.xaml"
$win = $script:fenetre
$win.Add_MouseLeftButtonDown({ try { $win.DragMove() } catch {} })
$win.FindName("BtnFermer").Add_Click({ $win.Close() })
$win.FindName("BtnReduire").Add_Click({ $win.WindowState = "Minimized" })
$win.FindName("BtnEffacer").Add_Click({
    $win.FindName("Journal").Children.Clear()
    Add-Journal $win "info" "Journal efface."
})
$win.FindName("BtnRafraichir").Add_Click({ Update-PreVol $win })
$win.FindName("BtnReliquat").Add_Click({ Show-Reliquat $win })
$win.FindName("BtnApps").Add_Click({ Show-Applications $win })

# Selecteur de langue : l'apercu fait defiler les deux valeurs du script.
$win.FindName("LangueChoix").Add_MouseLeftButtonDown({
    $t = $win.FindName("LangueTexte")
    $t.Text = if ($t.Text -eq "fr-fr") { "en-us" } else { "fr-fr" }
    Add-Journal $win "info" ("Langue d'installation : " + $t.Text)
})
foreach ($n in @("ChkDepot", "ChkProxy", "ChkKms", "ChkCle", "RadXmlAuto", "RadXmlFichier")) {
    $win.FindName($n).Add_Click({ Update-Options $win })
}
Update-Options $win
# Pre-vol mesure des l'ouverture : l'ecran ne montre que ce qui a ete constate.
Update-PreVol $win

# Parcourir : vraies boites de dialogue systeme, aucune ecriture.
$win.FindName("BtnSetup").Add_Click({
    $d = New-Object Microsoft.Win32.OpenFileDialog
    $d.Title = "Selectionner setup.exe (Office Deployment Tool)"
    $d.Filter = "Executables (*.exe)|*.exe|Tous les fichiers (*.*)|*.*"
    if ($d.ShowDialog()) { $win.FindName("ChampSetup").Text = $d.FileName }
})
$win.FindName("BtnXml").Add_Click({
    $d = New-Object Microsoft.Win32.OpenFileDialog
    $d.Title = "Selectionner un fichier de configuration XML (ODT)"
    $d.Filter = "Fichiers XML (*.xml)|*.xml|Tous les fichiers (*.*)|*.*"
    if ($d.ShowDialog()) { $win.FindName("ChampXml").Text = $d.FileName }
})
$win.FindName("BtnDepot").Add_Click({
    $d = New-Object System.Windows.Forms.FolderBrowserDialog
    $d.Description = "Dossier contenant le sous-dossier 'Office' produit par setup.exe /download"
    if ($d.ShowDialog() -eq "OK") { $win.FindName("ChampDepot").Text = $d.SelectedPath }
})
# Tests reels : une connexion TCP de mesure, dont le resultat porte la duree ou le
# motif de l'echec. Rien n'est affirme sans que la connexion ait abouti.
$win.FindName("BtnTestProxy").Add_Click({
    $hote = $win.FindName("ChampProxy").Text.Trim()
    $port = 0
    [void][int]::TryParse($win.FindName("ChampPort").Text, [ref]$port)
    if ($hote -eq "" -or $port -le 0) {
        Add-Journal $win "avr" "Test du proxy impossible : renseignez l'adresse et le port."
        return
    }
    Add-Journal $win "info" ("Test TCP du proxy " + $hote + ":" + $port + "...")
    $t = Test-PortTcp $hote $port
    if ($t.Ok) { Add-Journal $win "ok" ("Proxy joignable (TCP) : " + $hote + ":" + $port + " en " + $t.Ms + " ms. Le tunnel CONNECT n'est pas teste ici.") }
    else { Add-Journal $win "err" ("Proxy injoignable : " + $hote + ":" + $port + " (" + $t.Motif + ").") }
})
# Un seul chemin pour le pre-vol : le bouton KMS relance la meme mesure que
# Rafraichir. Il n'a pas son propre journal, sinon un seul test produisait trois
# lignes et deux connexions.
$win.FindName("BtnTestKms").Add_Click({ Update-PreVol $win })

# Le bouton lance le worker du moteur. Le scenario simule ne sert plus qu'aux
# captures : l'interactif deploye pour de vrai, ou refuse en nommant ce qui manque.
$win.FindName("BtnInstaller").Add_Click({ Start-Deploiement $win $cheminMoteur })

if ($Dialogue) { Show-Reliquat $win }
if ($Applications) { Show-Applications $win }
[void]$win.ShowDialog()
