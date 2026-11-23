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

param([switch]$Fast)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing | Out-Null

# L'interface livree est Apercu-Maquette.ps1, compilee dans Install-Office2024-KMS.exe.
# Le moteur Install-Office2024-KMS.ps1 reste livre a cote de l'EXE : c'est de lui que le
# worker de deploiement est extrait, il fait donc partie du produit.
$MainScript   = Join-Path $PSScriptRoot "Apercu-Maquette.ps1"
$EngineScript = Join-Path $PSScriptRoot "Install-Office2024-KMS.ps1"
$ExePath      = Join-Path $PSScriptRoot "Install-Office2024-KMS.exe"
$Work       = Join-Path $env:TEMP ("ctl_office_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
$Fail       = 0
$Pass       = 0

function Check {
    param([string]$Name, [bool]$Ok, [string]$Detail = "")
    if ($Ok) { $script:Pass++ ; Write-Host ("  PASS   " + $Name) }
    else     { $script:Fail++ ; Write-Host ("  ECHEC  " + $Name + "   -> " + $Detail) -ForegroundColor Red }
}

function Section { param([string]$T) Write-Host "" ; Write-Host ("=== " + $T + " ===") }

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

# ---------------------------------------------------------------- execution
if (-not (Test-Path $MainScript)) { Write-Host ("Interface introuvable : " + $MainScript) ; exit 2 }
if (-not (Test-Path $EngineScript)) { Write-Host ("Moteur introuvable : " + $EngineScript + " (le worker en est extrait)") ; exit 2 }
New-Item -ItemType Directory -Path $Work -Force | Out-Null
if ($Fast) { Write-Host "Mode rapide : section 1 seulement (script et proprietes de l'EXE)." }

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
    $obsoletes = @('bgWorker', 'RunWorkerAsync', 'DEFAULT_XML_PATH', 'proxyUserBox', 'proxyPassBox', 'ProxyPassword')
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
    $enregistrePass = [regex]::Matches($src, 'ProxyPass\w*\s*=').Count
    Check "proxy authentifie : mot de passe jamais enregistre" ($enregistreUser -and ($enregistrePass -eq 0)) ("ProxyUser enregistre=" + $enregistreUser + " affectations de mot de passe=" + $enregistrePass)

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
        Check "EXE : version 8.3.0.0" (($vi.FileVersion -eq "8.3.0.0") -and ($vi.ProductVersion -eq "8.3.0.0")) ("FileVersion=" + $vi.FileVersion + " ProductVersion=" + $vi.ProductVersion + " (attendu 8.3.0.0 / -Version)" + $cons)
        # Le badge de la fenetre annonce la version du binaire : deux sources, donc un
        # controle. Le badge est lu dans le XAML livre, pas dans une constante du test.
        $badge = if ($xaml -match '<TextBlock Text="([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)"\s+FontSize="11"') { $matches[1] } else { "(absent)" }
        Check "XAML : badge de version = version de l'EXE" ($badge -eq $vi.FileVersion) ("badge=" + $badge + " EXE=" + $vi.FileVersion + " - aligner le badge de Maquette\Principal.xaml sur -Version")
        # Le README annonce la version du programme a deux endroits : la commande de
        # compilation (-Version "...") et la prose ("version `...`, titre"). Toutes doivent
        # valoir celle du binaire livre. Les autres nombres a quatre parties du README ne
        # sont pas des versions du programme : l'ODT (16.0.x) et l'adresse du relais.
        $readmeCompil = if (Test-Path -LiteralPath (Join-Path $PSScriptRoot "README.md")) { [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "README.md")) } else { "" }
        $versionsReadme = @([regex]::Matches($readmeCompil, '(?i)version\s*["`]?([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $ecartsVersion = @($versionsReadme | Where-Object { $_ -ne $vi.FileVersion })
        Check "README : versions annoncees = version de l'EXE" (($versionsReadme.Count -gt 0) -and ($ecartsVersion.Count -eq 0)) ("README annonce=" + ($versionsReadme -join ", ") + " EXE=" + $vi.FileVersion + " - aligner le README (commande Invoke-PS2EXE et prose)")
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
        Check "livraison : setup.exe a cote de l'EXE" (Test-Path -LiteralPath (Join-Path $PSScriptRoot "setup.exe")) "setup.exe absent : un dossier extrait sans l'Office Deployment Tool ne peut rien installer (livrer setup.exe avec l'EXE, README section Demarrage rapide)"
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
    Check "depot sans Office\Data : erreur + arret" (($r.Sync.StepState[2] -eq "error") -and ($r.Sync.FinalMsg -like "*interrompu*")) $r.Sync.FinalMsg

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
                byte[] ko = Encoding.ASCII.GetBytes("HTTP/1.1 407 Proxy Authentication Required\r\nProxy-Authenticate: Basic realm=\"proxy\"\r\nContent-Length: 0\r\n\r\n");
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
            $sortieA = (& sh ".githooks/pre-commit" 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
            $codeA = $LASTEXITCODE
            Pop-Location
            $ErrorActionPreference = $eapAvant
            $sortieAU = ($sortieA -replace "\s+", " ").Trim()
            Check "hook : refuse un script indexe desynchronise" ($codeA -ne 0) ("code=" + $codeA + " sortie=" + $sortieAU)
            Check "hook : refuse en parlant de recompilation" ($sortieA.Contains("ne correspond plus au script")) $sortieAU
            # L'empreinte citee dans l'echec doit etre celle du script indexe, pas celle de l'arbre.
            $vueHook = Join-Path $depot "_vue"
            New-Item -ItemType Directory -Path $vueHook -Force | Out-Null
            $null = & git -C $depot checkout-index --all --prefix="$vueHook/" 2>&1
            $hashIndexe = (Get-FileHash -LiteralPath (Join-Path $vueHook "Apercu-Maquette.ps1") -Algorithm SHA256).Hash
            Remove-Item -LiteralPath $vueHook -Recurse -Force -ErrorAction SilentlyContinue
            Check "hook : empreinte comparee = script indexe" (($hashIndexe -ne $headHash) -and $sortieA.Contains($hashIndexe.Substring(0, 12))) ("index=" + $hashIndexe.Substring(0, 12) + " arbre=" + $headHash.Substring(0, 12) + " sortie=" + $sortieAU)

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
            $sortieB = (& sh ".githooks/pre-commit" 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
            $codeB = $LASTEXITCODE
            Pop-Location
            $ErrorActionPreference = $eapAvant
            Check "hook : accepte un arbre derive non indexe" ($codeB -eq 0) ("code=" + $codeB + " sortie=" + (($sortieB -replace "\s+", " ").Trim()))
            $vu = if (Test-Path -LiteralPath $sonde) { [System.IO.File]::ReadAllText($sonde).Trim() } else { "(absent)" }
            $arbreHash2 = (Get-FileHash -LiteralPath $scriptDepot -Algorithm SHA256).Hash
            $vuCourt = if ($vu.Length -ge 12) { $vu.Substring(0, 12) } else { $vu }
            Check "hook : suite lancee sur le contenu indexe" (($vu -eq $headHash) -and ($arbreHash2 -ne $headHash)) ("suite a vu=" + $vuCourt + " index=" + $headHash.Substring(0, 12) + " arbre=" + $arbreHash2.Substring(0, 12))

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
                $sortieC = (& sh ".githooks/pre-commit" 2>&1 | ForEach-Object { $_.ToString() }) -join "`n"
                $codeC = $LASTEXITCODE
                Pop-Location
                $ErrorActionPreference = $eapAvant
                $sortieCU = ($sortieC -replace "\s+", " ").Trim()
                Check "hook : refuse une declaration de fins de ligne retiree" ($codeC -ne 0) ("code=" + $codeC + " sortie=" + $sortieCU)
                Check "hook : export independant de core.autocrlf" ($sortieC -match "CRLF=0 LF isole=") $sortieCU
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