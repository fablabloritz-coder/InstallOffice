#Requires -Version 5.1
<#
    Controles automatises de Install-Office2024-KMS.ps1

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

$MainScript = Join-Path $PSScriptRoot "Install-Office2024-KMS.ps1"
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
        if ($Lines[$i] -match ('^\s{4}function ' + [regex]::Escape($Name) + ' \{')) { $start = $i ; break }
    }
    if ($start -lt 0) { throw ("fonction " + $Name + " introuvable") }
    $depth = 0
    for ($i = $start; $i -lt $Lines.Count; $i++) {
        $depth += ([regex]::Matches($Lines[$i], '\{')).Count
        $depth -= ([regex]::Matches($Lines[$i], '\}')).Count
        if ($depth -eq 0) { return ($Lines[$start..$i] -join "`n") }
    }
    throw ("fin de " + $Name + " introuvable")
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
if (-not (Test-Path $MainScript)) { Write-Host ("Script principal introuvable : " + $MainScript) ; exit 2 }
New-Item -ItemType Directory -Path $Work -Force | Out-Null
if ($Fast) { Write-Host "Mode rapide : section 1 seulement (script et proprietes de l'EXE)." }

try {
    $src   = [System.IO.File]::ReadAllText($MainScript)
    $lines = [System.IO.File]::ReadAllLines($MainScript)

    Section "1. Script principal"
    $tok = $null ; $err = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($src, [ref]$tok, [ref]$err)
    Check "analyse syntaxique" ((-not $err) -or ($err.Count -eq 0)) (($err | ForEach-Object { "L" + $_.Extent.StartLineNumber + ": " + $_.Message }) -join " ; ")
    $obsoletes = @('bgWorker', 'RunWorkerAsync', 'DEFAULT_XML_PATH', 'proxyUserBox', 'proxyPassBox', 'ProxyPassword')
    $trouves = @($obsoletes | Where-Object { ([regex]::Matches($src, $_)).Count -gt 0 })
    Check "aucun motif obsolete" ($trouves.Count -eq 0) ($trouves -join ", ")

    # Fins de ligne : l'EXE embarque le script tel qu'il est sur le poste, en
    # CRLF, et le controle compare des octets. Un clone en LF doit donc etre
    # signale ici, avec sa cause, plutot que par un faux "RECOMPILER" plus bas.
    $crlf = ([regex]::Matches($src, "`r`n")).Count
    $lfSeul = ([regex]::Matches($src, "(?<!`r)`n")).Count
    Check "script principal : fins de ligne CRLF (comme l'EXE embarque)" (($crlf -gt 0) -and ($lfSeul -eq 0)) ("CRLF=" + $crlf + " LF isole=" + $lfSeul + " - clone en LF plutot qu'une derive : obtenir un checkout CRLF (.gitattributes text eol=crlf) avant de conclure a une recompilation")
    $attrPath = Join-Path $PSScriptRoot ".gitattributes"
    $attr = if (Test-Path -LiteralPath $attrPath) { [System.IO.File]::ReadAllText($attrPath) } else { "" }
    $declareCrlf = ($attr -match '(?m)^\s*\*\.ps1\s+text\s+eol=crlf\s*$')
    Check "fins de ligne : .gitattributes impose CRLF aux .ps1" $declareCrlf ("declaration absente de .gitattributes : un clone avec core.autocrlf=false produirait un faux RECOMPILER (ajouter : *.ps1 text eol=crlf)")
    Write-Host ("         " + $lines.Count + " lignes, " + $src.Length + " octets")

    # L'EXE distribue embarque le script : son empreinte doit suivre le fichier .ps1.
    $exePath = Join-Path $PSScriptRoot ([System.IO.Path]::GetFileNameWithoutExtension($MainScript) + ".exe")
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hScript = [System.BitConverter]::ToString($sha.ComputeHash([System.IO.File]::ReadAllBytes($MainScript))).Replace("-", "")
    $recompiler = "RECOMPILER : Invoke-PS2EXE -InputFile Install-Office2024-KMS.ps1 ... (README, section Compilation de l'EXE)"
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
        Check "EXE : version 8.1.0.0" (($vi.FileVersion -eq "8.1.0.0") -and ($vi.ProductVersion -eq "8.1.0.0")) ("FileVersion=" + $vi.FileVersion + " ProductVersion=" + $vi.ProductVersion + " (attendu 8.1.0.0 / -Version)" + $cons)
        Check "EXE : titre" ($vi.FileDescription -eq "Deploiement Office 2024") ("FileDescription=" + $vi.FileDescription + " (attendu : Deploiement Office 2024 / -Title)" + $cons)
        Check "EXE : produit et societe" (($vi.ProductName -eq "InstallOffice") -and ($vi.CompanyName -eq "Lycee Loritz")) ("ProductName=" + $vi.ProductName + " CompanyName=" + $vi.CompanyName + " (attendu : InstallOffice / -Product, Lycee Loritz / -Company)" + $cons)
        Check "EXE : OriginalFilename" ($vi.OriginalFilename -eq "Install-Office2024-KMS.exe") ("OriginalFilename=" + $vi.OriginalFilename + " (attendu : Install-Office2024-KMS.exe, donc compiler directement vers ce nom)" + $cons)
        Check "EXE : InternalName" ($vi.InternalName -eq "Install-Office2024-KMS.exe") ("InternalName=" + $vi.InternalName + " (attendu : Install-Office2024-KMS.exe, donc compiler directement vers ce nom)" + $cons)
        $asciiExe = [System.Text.Encoding]::ASCII.GetString($exeBytes)
        $niveau = if ($asciiExe -match 'requestedExecutionLevel[^>]*level\s*=\s*"(?<lvl>[A-Za-z]+)"') { $matches["lvl"] } else { "(absent)" }
        Check "EXE : manifeste requireAdmin" ($niveau -eq "requireAdministrator") ("requestedExecutionLevel=" + $niveau + " (attendu requireAdministrator / -RequireAdmin)" + $cons)
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

    Check "apply : WinHTTP machine (netsh set proxy)" ($appels[0] -eq "winhttp-set 10.0.0.1:3128 bypass=<local>") $appels[0]
    Check "apply : WinINet session (ProxyEnable=1, serveur, exceptions)" (($appels[1] -match 'ProxyEnable=1') -and ($appels[1] -match 'ProxyServer=10\.0\.0\.1:3128') -and ($appels[1] -match 'ProxyOverride=<local>')) $appels[1]
    Check "apply : 4 variables d'environnement" (($appels[2] -match 'HTTP_PROXY=http://10\.0\.0\.1:3128') -and ($appels[2] -match 'HTTPS_PROXY=http://10\.0\.0\.1:3128') -and ($appels[2] -match 'http_proxy=http://10\.0\.0\.1:3128') -and ($appels[2] -match 'https_proxy=http://10\.0\.0\.1:3128')) $appels[2]
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
        Check "hook : mode rapide limite a la section 1" (($sortieRapide.Contains("=== 1. Script principal ===")) -and (-not $sortieRapide.Contains("=== 2. SourcePath"))) $rapideCourt

        $depot = Join-Path $Work "hookrepo"
        $avantTest = $env:INSTALLOFFICE_HOOKTEST
        New-Item -ItemType Directory -Path $depot -Force | Out-Null
        try {
            New-HookRepo -Path $depot -Autocrlf "false" -Dossiers @(".githooks") `
                -Files @("Install-Office2024-KMS.ps1", "Install-Office2024-KMS.exe", "Test-InstallOffice.ps1")
            $scriptDepot = Join-Path $depot "Install-Office2024-KMS.ps1"
            $headHash = (Get-FileHash -LiteralPath $scriptDepot -Algorithm SHA256).Hash

            # Cas A : l'index derive, l'arbre de travail est propre.
            Copy-Item -LiteralPath $scriptDepot -Destination (Join-Path $depot "_pristine.ps1") -Force
            Add-Content -LiteralPath $scriptDepot -Value "# derive indexee uniquement" -Encoding ASCII
            $null = & git -C $depot add "Install-Office2024-KMS.ps1" 2>&1
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
            $hashIndexe = (Get-FileHash -LiteralPath (Join-Path $vueHook "Install-Office2024-KMS.ps1") -Algorithm SHA256).Hash
            Remove-Item -LiteralPath $vueHook -Recurse -Force -ErrorAction SilentlyContinue
            Check "hook : empreinte comparee = script indexe" (($hashIndexe -ne $headHash) -and $sortieA.Contains($hashIndexe.Substring(0, 12))) ("index=" + $hashIndexe.Substring(0, 12) + " arbre=" + $headHash.Substring(0, 12) + " sortie=" + $sortieAU)

            # Cas B : l'arbre de travail derive, l'index est propre (suite temoin, rapide).
            $null = & git -C $depot reset -q 2>&1
            Add-Content -LiteralPath $scriptDepot -Value "# derive non indexee" -Encoding ASCII
            $sonde = Join-Path $depot "_sonde.txt"
            $stub = @(
                '$h = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot "Install-Office2024-KMS.ps1") -Algorithm SHA256).Hash',
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
                New-HookRepo -Path $depotEol -Autocrlf "true" -Dossiers @(".githooks") `
                    -Files @("Install-Office2024-KMS.exe", "Test-InstallOffice.ps1", ".gitattributes") `
                    -Avant {
                        param($p, $src)
                        $txtLf = ([System.IO.File]::ReadAllText((Join-Path $src "Install-Office2024-KMS.ps1"))) -replace "`r`n", "`n"
                        [System.IO.File]::WriteAllText((Join-Path $p "Install-Office2024-KMS.ps1"), $txtLf, (New-Object System.Text.UTF8Encoding($false)))
                    }
                # La regression : uniquement la declaration de fins de ligne.
                $attrEol = Join-Path $depotEol ".gitattributes"
                $garde = @([System.IO.File]::ReadAllLines($attrEol) | Where-Object { $_ -notmatch '^\s*\*\.ps1\s+text\s+eol=crlf\s*$' })
                [System.IO.File]::WriteAllLines($attrEol, $garde, (New-Object System.Text.UTF8Encoding($false)))
                $null = & git -C $depotEol add ".gitattributes" 2>&1
                $indexeEol = @(& git -C $depotEol diff --cached --name-only 2>&1)
                Check "hook : .gitattributes surveille" (($indexeEol.Count -eq 1) -and ($indexeEol[0] -eq ".gitattributes")) ($indexeEol -join ", ")
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
