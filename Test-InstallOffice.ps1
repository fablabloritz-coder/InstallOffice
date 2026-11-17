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
    Code de sortie : 0 si tous les controles passent, 1 sinon.
#>

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

# ---------------------------------------------------------------- execution
if (-not (Test-Path $MainScript)) { Write-Host ("Script principal introuvable : " + $MainScript) ; exit 2 }
New-Item -ItemType Directory -Path $Work -Force | Out-Null

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
    Write-Host ("         " + $lines.Count + " lignes, " + $src.Length + " octets")

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
$total = $Pass + $Fail
if ($Fail -eq 0) {
    Write-Host ("RESULTAT : " + $Pass + "/" + $total + " controles passes.") -ForegroundColor Green
    exit 0
}
Write-Host ("RESULTAT : " + $Fail + " echec(s) sur " + $total + " controles.") -ForegroundColor Red
exit 1
