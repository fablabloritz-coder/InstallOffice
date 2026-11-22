# ================================================================
#  Apercu de la maquette XAML (style Fluent Windows 11)
# ================================================================
#  Charge Maquette/Principal.xaml et Maquette/ReliquatProxy.xaml avec
#  XamlReader - exactement le mecanisme qu'utiliserait le script de
#  production - et les anime avec des donnees simulees.
#
#  Aucun acces reseau, aucune ecriture sur le poste, aucune activation :
#  cet apercu ne touche ni au moteur ni aux controles.
#
#  Usage :
#    powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Apercu-Maquette.ps1
#    ... -Dialogue                       (ouvre directement le dialogue)
#    ... -Capture C:\tmp\ecran.png       (rend l'ecran en PNG et sort)
#    ... -Capture ... -Etape 6           (etat simule : index du scenario)
#    ... -Capture ... -Dialogue          (rend le dialogue en PNG)
# ================================================================
param(
    [string]$Capture = "",
    [switch]$Dialogue,
    [int]$Etape = 6
)

Add-Type -AssemblyName PresentationFramework

$ici      = Split-Path -Parent $MyInvocation.MyCommand.Path
$dossier  = Join-Path $ici "Maquette"
$pinceaux = New-Object System.Windows.Media.BrushConverter
$policeTexte = New-Object System.Windows.Media.FontFamily("Segoe UI Variable Text, Segoe UI")
$policeIcone = New-Object System.Windows.Media.FontFamily("Segoe MDL2 Assets, Segoe UI")

function Import-Xaml {
    param([string]$Nom)
    $chemin = Join-Path $dossier $Nom
    if (-not (Test-Path -LiteralPath $chemin)) { throw "Fichier introuvable : $chemin" }
    [System.Windows.Markup.XamlReader]::Parse([System.IO.File]::ReadAllText($chemin, [System.Text.Encoding]::UTF8))
}

function Get-Pinceau {
    param([string]$Couleur)
    return $pinceaux.ConvertFromString($Couleur)
}

# Etats d'une etape : fond de la puce, couleur du glyphe, glyphe des etats finis.
# Le libelle accentue est porte par le XAML (cle Etat<Etat>), le script reste ASCII.
$etats = @{
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
    $e = $etats[$Etat]
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
    $info.Text = [string]$Fenetre.Resources[$cle]
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

# --- Mode capture : rendu hors ecran, aucun affichage ----------------------
if ($Capture -ne "") {
    if ($Dialogue) {
        Export-Visuel (Import-Xaml "ReliquatProxy.xaml") $Capture 560 392
    } else {
        $win = Import-Xaml "Principal.xaml"
        for ($k = 0; $k -le $Etape; $k++) { Set-Scene $win $k }
        Export-Visuel $win $Capture 1120 828
    }
    Write-Host ("Capture ecrite : " + $Capture)
    return
}

# --- Mode interactif -------------------------------------------------------
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
$win.FindName("BtnTester").Add_Click({
    Add-Journal $win "info" "Test du proxy : CONNECT officecdn.microsoft.com:443 via le relais 127.0.0.1:52144..."
    Add-Journal $win "ok" "Proxy joignable : 200 Connection established en 38 ms."
})
$win.FindName("BtnReliquat").Add_Click({ Show-Reliquat $win })

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
[void]$win.ShowDialog()
