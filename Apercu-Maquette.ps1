# ================================================================
#  Apercu de la maquette XAML (style Fluent Windows 11)
# ================================================================
#  Charge les XAML de Maquette/ avec XamlReader - exactement le mecanisme
#  qu'utiliserait le script de production - et les anime avec des donnees
#  simulees. Le formulaire reprend les options du script : setup.exe, langue,
#  depot de sources, XML et applications, KMS (principal, secours, cle produit),
#  proxy (adresse, port, identifiants, exceptions).
#
#  Aucun acces reseau, aucune ecriture sur le poste, aucune activation :
#  cet apercu ne touche ni au moteur ni aux controles.
#
#  Usage :
#    powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Apercu-Maquette.ps1
#    ... -Dialogue                       (ouvre directement le reliquat de proxy)
#    ... -Applications                   (ouvre directement les applications)
#    ... -Capture C:\tmp\ecran.png       (rend l'ecran en PNG et sort)
#    ... -Capture ... -Etape 6           (etat simule : index du scenario)
#    ... -Etats                         (planche des etats de boutons)
#    ... -Capture ... -Etats             (rend la planche en PNG)
#    ... -Capture ... -Dialogue | -Applications
#    ... -Capture ... -Focus ChkProxy     (met ce controle en focus avant le rendu)
# ================================================================
param(
    [string]$Capture = "",
    [switch]$Dialogue,
    [switch]$Applications,
    [switch]$Etats,
    [string]$Focus = "",
    [int]$Etape = 7
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

$ici      = Split-Path -Parent $MyInvocation.MyCommand.Path
$dossier  = Join-Path $ici "Maquette"
$pinceaux = New-Object System.Windows.Media.BrushConverter
$policeTexte = New-Object System.Windows.Media.FontFamily("Segoe UI Variable Text, Segoe UI")
$policeIcone = New-Object System.Windows.Media.FontFamily("Segoe MDL2 Assets, Segoe UI")

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

function Set-Etape {
    param($Fenetre, [int]$Index, [string]$Etat)
    $e = $renduEtapes[$Etat]
    $puce   = $Fenetre.FindName("Etape$Index")
    $glyphe = $Fenetre.FindName("Glyphe$Index")
    $titre  = $Fenetre.FindName("Titre$Index")
    $info   = $Fenetre.FindName("Etat$Index")
    $puce.Background = Get-Pinceau $e.Fond
    $glyphe.Foreground = Get-Pinceau $e.Texte
    if ($Etat -eq "attente" -or $Etat -eq "encours") {
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
    $ligne.Text = ("{0}  {1}" -f (Get-Date).ToString("HH:mm:ss"), $Texte)
    [void]$Fenetre.FindName("Journal").Children.Add($ligne)
    $Fenetre.FindName("Defilement").ScrollToEnd()
}

function Set-Actif {
    param($Fenetre, [string[]]$Noms, [bool]$Actif)
    foreach ($n in $Noms) { $Fenetre.FindName($n).IsEnabled = $Actif }
}

# Regle de l'ecran : un depot de sources supprime le telechargement, donc le proxy.
function Update-Options {
    param($Fenetre)
    $depot = [bool]$Fenetre.FindName("ChkDepot").IsChecked
    Set-Actif $Fenetre @("ChampDepot", "BtnDepot") $depot
    $Fenetre.FindName("ChkProxy").IsEnabled = -not $depot
    $proxy = (-not $depot) -and [bool]$Fenetre.FindName("ChkProxy").IsChecked
    Set-Actif $Fenetre @("ChampProxy", "ChampPort", "ChampProxyUser", "ChampProxyPass", "ChampBypass", "BtnTestProxy") $proxy
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
       L = @("etape|Etape 1/5 - Desinstallation", "info|Detection des produits Office installes (ospp.vbs /dstatus).") }
    @{ P = @{ 1 = "saute" ; 2 = "encours" } ; J = 14 ; S = "Verifications pre-vol" ; D = "Etape 2 sur 5" ;
       L = @("ok|Aucune version precedente a retirer : etape ignoree.", "etape|Etape 2/5 - Preparation", "info|setup.exe present, 212 Go libres, KMS joignable.") }
    @{ P = @{ 2 = "termine" ; 3 = "encours" } ; J = 28 ; S = "Copie des sources depuis le depot" ; D = "Etape 3 sur 5" ;
       L = @("ok|Configuration ODT generee (ProPlus2024Volume, PerpetualVL2024).", "etape|Etape 3/5 - Sources", "info|robocopy \\serveur\office -> C:\Office2024\sources") }
    @{ J = 46 ; S = "Copie des sources depuis le depot" ; D = "Etape 3 sur 5" ;
       L = @("info|4,1 Go copies sur 6,8 Go (60 %).") }
    @{ P = @{ 3 = "termine" ; 4 = "encours" } ; J = 66 ; S = "Installation Click-to-Run en cours" ; D = "Etape 4 sur 5" ;
       L = @("ok|Sources pretes : 6,8 Go dans C:\Office2024\sources.", "etape|Etape 4/5 - Installation", "info|setup.exe /configure OfficeConfig2024.xml (silencieux).") }
    @{ J = 82 ; S = "Installation Click-to-Run en cours" ; D = "Etape 4 sur 5" ;
       L = @("info|Click-to-Run : installation des applications Office...") }
    @{ P = @{ 4 = "termine" ; 5 = "encours" } ; J = 94 ; S = "Activation KMS" ; D = "Etape 5 sur 5" ;
       L = @("ok|Office installe : 16h32.", "etape|Etape 5/5 - Activation KMS", "info|ospp.vbs /sethst:kms.ac-nancy.fr:1688") }
    @{ P = @{ 5 = "termine" } ; J = 100 ; S = "Deploiement termine" ; D = "5 etapes sur 5" ;
       L = @("ok|Activation KMS reussie : licence perpetuelle accordee.", "ok|Configuration proxy restauree (acces direct), port du relais libere.", "ok|Deploiement termine sans erreur.") }
)

function Set-Scene {
    param($Fenetre, [int]$Index)
    if ($Index -lt 0 -or $Index -ge $scenario.Count) { return }
    $s = $scenario[$Index]
    if ($s.ContainsKey("P")) { foreach ($n in @($s.P.Keys)) { Set-Etape $Fenetre $n $s.P[$n] } }
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
}

function Export-Visuel {
    param($Fenetre, [string]$Chemin, [int]$Largeur, [int]$Hauteur)
    $contenu = $Fenetre.Content
    $taille  = New-Object System.Windows.Size($Largeur, $Hauteur)
    $contenu.Measure($taille)
    $contenu.Arrange((New-Object System.Windows.Rect(0, 0, $Largeur, $Hauteur)))
    $contenu.UpdateLayout()
    $bmp = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($Largeur, $Hauteur, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
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
        if (-not $Dlg.FindName($a.Cle).IsChecked) { $exclues += $a.Libelle }
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
        $texte = if ($exclues.Count -eq 0) { "Toutes les applications incluses" } else { "Exclues : " + ($exclues -join ", ") }
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
        Export-Visuel $dlg $Capture 560 392
    } elseif ($Applications) {
        $dlgApps = Import-Xaml "Applications.xaml"
        if ($Focus -ne "") { Set-FocusCapture $dlgApps $Focus }
        Export-Visuel $dlgApps $Capture 620 560
    } elseif ($Etats) {
        $planche = Import-Xaml "Etats.xaml"
        Set-PlancheEtats $planche
        Export-Visuel $planche $Capture 1216 595
    } else {
        $win = Import-Xaml "Principal.xaml"
        for ($k = 0; $k -le $Etape; $k++) { Set-Scene $win $k }
        if ($Focus -ne "") { Set-FocusCapture $win $Focus }
        Export-Visuel $win $Capture 1240 1000
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
$win.FindName("BtnRafraichir").Add_Click({
    Add-Journal $win "info" "Verifications pre-vol relancees : setup.exe, espace disque, reseau, Windows."
})
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
$win.FindName("BtnTestProxy").Add_Click({
    Add-Journal $win "info" ("Test du proxy " + $win.FindName("ChampProxy").Text + ":" + $win.FindName("ChampPort").Text + "...")
    Add-Journal $win "ok" "Proxy joignable : 200 Connection established en 42 ms."
})
$win.FindName("BtnTestKms").Add_Click({
    Add-Journal $win "info" ("Connexion KMS " + $win.FindName("ChampKms").Text + ":1688...")
    Add-Journal $win "ok" "Connexion KMS reussie sur kms.ac-nancy.fr:1688."
})

$script:i = 0
$script:minuterie = New-Object System.Windows.Threading.DispatcherTimer
$script:minuterie.Interval = [TimeSpan]::FromMilliseconds(1300)
$script:minuterie.Add_Tick({
    Set-Scene $script:fenetre $script:i
    $script:i++
    if ($script:i -ge $scenario.Count) { $script:minuterie.Stop() }
})
$win.FindName("BtnInstaller").Add_Click({
    $script:fenetre.FindName("Journal").Children.Clear()
    Reset-Scene $script:fenetre
    $script:i = 0
    $script:minuterie.Start()
})

if ($Dialogue) { Show-Reliquat $win }
if ($Applications) { Show-Applications $win }
[void]$win.ShowDialog()
