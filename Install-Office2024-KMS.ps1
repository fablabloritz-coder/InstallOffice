# ================================================================
#  Install-Office2024-KMS.ps1  |  v8
#
#  NOUVEAUTES v8 :
#  + PROXY : la configuration proxy est desormais appliquee a WINDOWS (WinHTTP
#      machine + WinINet session), et plus seulement en variables d'environnement.
#      Aucun composant Office ne lit HTTP_PROXY : le telechargement est realise
#      par le service 'Microsoft Office Click-to-Run' (compte LocalSystem).
#  + PROXY : test CONNECT vers officecdn.microsoft.com AVANT le telechargement
#      -> diagnostic explicite (proxy injoignable / destination bloquee / 407).
#  + PROXY : sauvegarde puis restauration exacte de la configuration d'origine.
#  + PROXY AUTHENTIFIE : un relais local (loopback) ajoute l'en-tete
#      Proxy-Authorization : le service Click-to-Run, incapable de repondre a un
#      407, voit un proxy ordinaire. Identifiants saisis dans Options, mot de
#      passe jamais enregistre.
#  + MODE 'DEPOT DE SOURCES' : installation + activation KMS sur le reseau
#      pedagogique SANS telechargement -> plus de changement de reseau par poste.
#  + Journal : correction des lignes {0:mm}m{0:ss}s non formatees.
#
#  HERITE v7 :
#  + Selecteur setup.exe : champ + Parcourir, chemin libre
#  + KMS host:port : le champ accepte "kms.serveur.fr:1688"
#  + Configuration XML :
#      - Generation auto avec selection des applications a inclure
#      - OU import d'un fichier XML existant (ODT personnalise)
#  + Dialog "Personnaliser les apps" avec apercu XML en temps reel
#
#  HERITE v6 :
#  NotifyIcon, Sons, Config persistante, Confirmation, Test KMS async,
#  Win10+ check, Temps restant estime, Copier journal, Avertissement
#  Office2024 deja installe, RunWorkerCompleted, Form.Shown, IsBusy
# ================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# System.Media est fourni par System.dll sur Windows PowerShell; charge optionnel pour compatibilite.
try { Add-Type -AssemblyName System -ErrorAction Stop } catch {}

# Calcul du repertoire du script : compatible PS.exe, EXE compile (PS2EXE), et appel direct
$script:ScriptRoot = $null
if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $script:ScriptRoot = $PSScriptRoot
} elseif (-not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)) {
    $script:ScriptRoot = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
} elseif (-not [string]::IsNullOrWhiteSpace([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)) {
    $script:ScriptRoot = Split-Path -Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Parent
}
if ([string]::IsNullOrWhiteSpace($script:ScriptRoot)) { $script:ScriptRoot = $env:TEMP }

# --- Elevation automatique (PS1 uniquement - l'EXE compile a deja le manifest UAC requireAdmin) ---
$isCompiledExe = ($MyInvocation.MyCommand.CommandType -eq 'Application') -or [string]::IsNullOrWhiteSpace($PSCommandPath)
if (-not $isCompiledExe) {
    If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]"Administrator")) {
        Start-Process powershell "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs ; Exit
    }
}

# ================================================================
#  CONFIGURATION  (adaptez selon votre environnement)
# ================================================================
# setup.exe est d'abord cherche a cote du script (ou de l'EXE compile), puis a l'ancien emplacement.
$DEFAULT_SETUP_PATH = Join-Path -Path $script:ScriptRoot -ChildPath "setup.exe"
if (-not (Test-Path $DEFAULT_SETUP_PATH)) { $DEFAULT_SETUP_PATH = "C:\InstallOffice2024\setup.exe" }
$PROXY_BYPASS_DEF   = '<local>;*.local'
$KMS_SERVEUR        = "KMS.VOTRE-SERVEUR.FR"   # accepte aussi "host:port"
$KMS_PORT           = "1688"
$DISK_MIN_GB        = 4
$LOG_FOLDER         = "C:\InstallOffice2024"
$CONFIG_FILE        = "$env:APPDATA\Office2024Deploy\config.json"
$VERSION            = "v8.2"
$STEP_ESTIMATES     = @{ 0=3; 1=1; 2=18; 3=12; 4=1 }

# Liste des applications Office 2024 installables
# Default=$true  -> incluse par defaut
# Default=$false -> exclue par defaut (non installee)
$OFFICE_APPS = @(
    @{ ID="Word";       Label="Word";             Desc="Traitement de texte"         ; Default=$true  },
    @{ ID="Excel";      Label="Excel";            Desc="Tableur"                     ; Default=$true  },
    @{ ID="PowerPoint"; Label="PowerPoint";       Desc="Presentations"               ; Default=$true  },
    @{ ID="Outlook";    Label="Outlook";          Desc="Messagerie et agenda"        ; Default=$true  },
    @{ ID="OneNote";    Label="OneNote";          Desc="Prise de notes"              ; Default=$true  },
    @{ ID="Access";     Label="Access";           Desc="Base de donnees"             ; Default=$true  },
    @{ ID="Teams";      Label="Teams";            Desc="Collaboration (optionnel)"   ; Default=$false },
    @{ ID="OneDrive";   Label="OneDrive";         Desc="Synchronisation cloud"       ; Default=$false },
    @{ ID="Lync";       Label="Skype Entreprise"; Desc="UC legacy (obsolete)"        ; Default=$false }
)

# ================================================================
#  ETAT XML (script-scope - modifiable depuis les dialogues UI)
# ================================================================
$script:xmlExcludedApps  = ($OFFICE_APPS | Where-Object { -not $_.Default } | ForEach-Object { $_.ID })
$script:xmlUseCustomFile = $false
$script:xmlCustomFilePath = ""

# ================================================================
#  COULEURS & POLICES
# ================================================================
$clrBlueDark  = [System.Drawing.Color]::FromArgb(0,   74, 140)
$clrBlueLight = [System.Drawing.Color]::FromArgb(0,  120, 215)
$clrGreen     = [System.Drawing.Color]::FromArgb(16, 137,  62)
$clrRed       = [System.Drawing.Color]::FromArgb(196,  43,  28)
$clrOrange    = [System.Drawing.Color]::FromArgb(202,  80,  16)
$clrGray      = [System.Drawing.Color]::FromArgb(160, 160, 160)
$clrBg        = [System.Drawing.Color]::FromArgb(245, 247, 250)
$clrLogBg     = [System.Drawing.Color]::FromArgb(30,  30,  30)
$clrWhite     = [System.Drawing.Color]::White

$fontTitle    = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$fontSection  = New-Object System.Drawing.Font("Segoe UI",  9, [System.Drawing.FontStyle]::Bold)
$fontNormal   = New-Object System.Drawing.Font("Segoe UI",  9)
$fontSmall    = New-Object System.Drawing.Font("Segoe UI",  8)
$fontStep     = New-Object System.Drawing.Font("Segoe UI",  8, [System.Drawing.FontStyle]::Bold)
$fontStepNum  = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$fontMono     = New-Object System.Drawing.Font("Consolas", 8.5)

# ================================================================
#  ETAT PARTAGE THREAD-SAFE
# ================================================================
$syncHash = [hashtable]::Synchronized(@{
    LogQueue   = New-Object System.Collections.Concurrent.ConcurrentQueue[object]
    StepState  = [string[]]@("waiting","waiting","waiting","waiting","waiting")
    ActiveStep = -1
    StatusText = "Pret."
    Done       = $false
    HasError   = $false
    FinalMsg   = ""
    StartTime      = $null
    StepStartTime  = $null
    PopupTitle = ""
    PopupMsg   = ""
    PopupIcon  = "Information"
    ShowPopup  = $false
})

# ================================================================
#  FONCTIONS UTILITAIRES
# ================================================================

# --- Parse adresse KMS : accepte "host" ou "host:port" ---
function ConvertTo-KmsAddress {
    param([string]$Raw)
    $s = $Raw.Trim()
    if ($s -match '^(.+):(\d{1,5})$') {
        return @{ Host = $matches[1].Trim() ; Port = $matches[2].Trim() }
    }
    return @{ Host = $s ; Port = $KMS_PORT }
}

# --- Resout un chemin relatif depuis le dossier du script ---
function Resolve-DeploymentPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    try {
        if ([System.IO.Path]::IsPathRooted($Path)) { return $Path }
        return [System.IO.Path]::GetFullPath((Join-Path -Path $script:ScriptRoot -ChildPath $Path))
    } catch {
        return $Path
    }
}

# --- Normalise une cle produit : espaces/tirets speciaux/collage sans tirets ---
function Normalize-ProductKey {
    param([string]$RawKey)

    if ($null -eq $RawKey) { return "" }

    $k = $RawKey.Trim().ToUpper()
    if ([string]::IsNullOrWhiteSpace($k)) { return "" }

    # Uniformiser les differents caracteres de tiret puis retirer les espaces.
    $k = $k -replace '[\u2010\u2011\u2012\u2013\u2014\u2212\uFE58\uFE63\uFF0D]', '-'
    $k = $k -replace '\s', ''

    # Si la cle est fournie en 25 caracteres sans separateur, la reformatter en 5x5.
    $compact = $k -replace '[^A-Z0-9]', ''
    if ($compact.Length -eq 25) {
        return (($compact -split '(.{5})' | Where-Object { $_ }) -join '-')
    }

    return $k
}

# --- Test TCP (pour bouton "Tester" KMS) ---
function Test-TcpPort {
    param([string]$Server, [int]$Port, [int]$TimeoutMs = 3000)
    $tcp = $null
    try {
        $tcp    = New-Object System.Net.Sockets.TcpClient
        $result = $tcp.BeginConnect($Server, $Port, $null, $null)
        if (-not $result.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { return $false }
        # EndConnect leve une exception si la connexion est refusee : sans lui,
        # WaitOne() renvoie True meme quand le port est ferme.
        $tcp.EndConnect($result)
        return $tcp.Connected
    } catch { return $false }
    finally { if ($tcp) { try { $tcp.Close() } catch {} } }
}

# ================================================================
#  PROXY SYSTEME  (WinHTTP machine + WinINet session)
#
#  setup.exe /download ne lit PAS HTTP_PROXY. Le telechargement est realise
#  par le service "Microsoft Office Click-to-Run" (LocalSystem), qui utilise la
#  configuration WinHTTP de la machine ; les processus de la session utilisent
#  en plus les reglages WinINet. On applique donc ces deux couches (+ variables
#  d'environnement) avant le telechargement, puis on restaure l'etat d'origine.
# ================================================================

# --- Test reel du proxy : CONNECT HTTP vers les endpoints Microsoft ---
#     Copie volontaire de bg_TestProxyConnect : le worker tourne dans un runspace
#     separe et ne voit donc pas les fonctions de la portee principale.
#     Ne pas supprimer l'une des deux copies sans refondre ce mecanisme.
function Test-ProxyConnect {
    param([string]$ProxyHost,[int]$ProxyPort,[string]$TargetHost,[int]$TargetPort,[int]$TimeoutMs = 6000)

    $res = @{ Name = "${TargetHost}:${TargetPort}" ; OK = $false ; Code = 0 ; Detail = "" }
    $tcp = $null
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $iar = $tcp.BeginConnect($ProxyHost, $ProxyPort, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            $res.Detail = "timeout : proxy ${ProxyHost}:${ProxyPort} injoignable"
            return $res
        }
        $tcp.EndConnect($iar)
        $tcp.ReceiveTimeout = $TimeoutMs

        $stream = $tcp.GetStream()
        $sw = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::ASCII)
        $sw.NewLine = "`r`n"
        $sw.AutoFlush = $true
        $sw.WriteLine("CONNECT ${TargetHost}:${TargetPort} HTTP/1.1")
        $sw.WriteLine("Host: ${TargetHost}:${TargetPort}")
        $sw.WriteLine("")

        $sr = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::ASCII)
        $line = $sr.ReadLine()
        if ($line -match 'HTTP/\d(\.\d)?\s+(\d{3})') {
            $res.Code = [int]$matches[2]
            $res.OK   = ($res.Code -ge 200 -and $res.Code -lt 300)
            $res.Detail = "reponse proxy : $($line.Trim())"
        } else {
            $res.Detail = "reponse proxy inattendue : $line"
        }
    } catch {
        $res.Detail = $_.Exception.Message
    } finally {
        if ($tcp) { try { $tcp.Close() } catch {} }
    }
    return $res
}

function Test-ProxyReadiness {
    param([string]$ProxyHost,[int]$ProxyPort)

    $out = New-Object System.Collections.ArrayList
    $probe = Test-TcpPort -Server $ProxyHost -Port $ProxyPort -TimeoutMs 4000
    $detail = if ($probe) { "port ouvert" } else { "port ferme ou filtre" }
    [void]$out.Add(@{ Name = "proxy ${ProxyHost}:${ProxyPort}" ; OK = $probe ; Code = 0 ; Detail = $detail })
    if (-not $probe) { return $out }

    foreach ($t in @(@{ H = "officecdn.microsoft.com" ; P = 443 }, @{ H = "config.office.com" ; P = 443 })) {
        [void]$out.Add((Test-ProxyConnect -ProxyHost $ProxyHost -ProxyPort $ProxyPort -TargetHost $t.H -TargetPort $t.P))
    }
    return $out
}

# --- Relais proxy local avec authentification (loopback) -------------------
#  Le service Click-to-Run tourne en LocalSystem : il ne peut pas repondre a un
#  challenge 407. Ce relais ecoute sur la loopback, ajoute l'en-tete
#  Proxy-Authorization (Basic) vers le proxy reel et laisse passer le tunnel.
#  Windows est ensuite pointe sur 127.0.0.1:<port> : pour Office, ce relais est
#  un proxy ordinaire, sans authentification.
function Start-ProxyRelay {
    param([string]$UpstreamHost, [int]$UpstreamPort, [string]$User, [string]$Password)
    $code = @'
using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

public class OfficeProxyRelay
{
    private readonly string _upHost;
    private readonly int _upPort;
    private readonly string _auth;
    private TcpListener _listener;
    private Thread _accept;
    private volatile bool _stopping;
    private int _port;
    private int _tunnels;
    private int _erreurs;
    private string _derniereErreur = "";

    public OfficeProxyRelay(string upstreamHost, int upstreamPort, string user, string password)
    {
        _upHost = upstreamHost;
        _upPort = upstreamPort;
        if (!string.IsNullOrEmpty(user))
        {
            _auth = "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(user + ":" + password));
        }
    }

    public int Port { get { return _port; } }
    public int Tunnels { get { return _tunnels; } }
    public int Erreurs { get { return _erreurs; } }
    public string DerniereErreur { get { return _derniereErreur; } }

    public void Start()
    {
        _listener = new TcpListener(IPAddress.Loopback, 0);
        _listener.Start();
        _port = ((IPEndPoint)_listener.LocalEndpoint).Port;
        _accept = new Thread(AcceptLoop);
        _accept.IsBackground = true;
        _accept.Start();
    }

    public void Stop()
    {
        _stopping = true;
        try { if (_listener != null) { _listener.Stop(); } } catch { }
        try { if (_accept != null) { _accept.Join(2000); } } catch { }
    }

    private void AcceptLoop()
    {
        while (!_stopping)
        {
            TcpClient client;
            try { client = _listener.AcceptTcpClient(); }
            catch { return; }
            Thread t = new Thread(delegate() { Serve(client); });
            t.IsBackground = true;
            t.Start();
        }
    }

    private void Serve(TcpClient client)
    {
        TcpClient up = null;
        try
        {
            client.NoDelay = true;
            NetworkStream cs = client.GetStream();
            string entete = LireEntete(cs);
            string[] lignes = entete.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            string requete = lignes.Length > 0 ? lignes[0] : "";
            string[] parts = requete.Split(' ');
            if (parts.Length < 2 || !parts[0].Equals("CONNECT", StringComparison.OrdinalIgnoreCase))
            {
                Signaler("methode non tunnelisable : " + requete);
                Ecrire(cs, "HTTP/1.1 405 Method Not Allowed\r\nContent-Length: 0\r\n\r\n");
                return;
            }
            string cible = parts[1];
            up = new TcpClient();
            up.NoDelay = true;
            up.Connect(_upHost, _upPort);
            NetworkStream us = up.GetStream();
            StringBuilder r = new StringBuilder();
            r.Append("CONNECT " + cible + " HTTP/1.1\r\n");
            r.Append("Host: " + cible + "\r\n");
            r.Append("User-Agent: InstallOffice2024\r\n");
            if (_auth != null) { r.Append("Proxy-Authorization: " + _auth + "\r\n"); }
            r.Append("\r\n");
            byte[] rb = Encoding.ASCII.GetBytes(r.ToString());
            us.Write(rb, 0, rb.Length);
            us.Flush();

            string reponse = LireEntete(us);
            int code = 0;
            int sp = reponse.IndexOf(' ');
            if (sp > 0 && reponse.Length >= sp + 4) { int.TryParse(reponse.Substring(sp + 1, 3), out code); }
            if (code < 200 || code > 299)
            {
                Signaler("proxy amont : " + reponse.Replace("\r\n", " ").Trim());
                Ecrire(cs, reponse);
                return;
            }
            Interlocked.Increment(ref _tunnels);
            Ecrire(cs, "HTTP/1.1 200 Connection established\r\n\r\n");
            Tunnel(cs, us);
        }
        catch (Exception ex) { Signaler(ex.Message); }
        finally
        {
            try { if (up != null) { up.Close(); } } catch { }
            try { client.Close(); } catch { }
        }
    }

    private static void Tunnel(NetworkStream a, NetworkStream b)
    {
        Thread t = new Thread(delegate() { Pomper(a, b); });
        t.IsBackground = true;
        t.Start();
        Pomper(b, a);
        try { t.Join(2000); } catch { }
    }

    private static void Pomper(NetworkStream de, NetworkStream vers)
    {
        byte[] tampon = new byte[32768];
        try
        {
            int n;
            while ((n = de.Read(tampon, 0, tampon.Length)) > 0)
            {
                vers.Write(tampon, 0, n);
                vers.Flush();
            }
        }
        catch { }
        finally
        {
            try { de.Close(); } catch { }
            try { vers.Close(); } catch { }
        }
    }

    private static string LireEntete(NetworkStream s)
    {
        StringBuilder sb = new StringBuilder();
        byte[] un = new byte[1];
        while (sb.Length < 16384)
        {
            int n = s.Read(un, 0, 1);
            if (n <= 0) { break; }
            sb.Append((char)un[0]);
            if (sb.Length >= 4 && sb[sb.Length - 1] == '\n' && sb[sb.Length - 2] == '\r' && sb[sb.Length - 3] == '\n' && sb[sb.Length - 4] == '\r') { break; }
            if (sb.Length >= 2 && sb[sb.Length - 1] == '\n' && sb[sb.Length - 2] == '\n') { break; }
        }
        return sb.ToString();
    }

    private static void Ecrire(NetworkStream s, string texte)
    {
        try
        {
            byte[] b = Encoding.ASCII.GetBytes(texte);
            s.Write(b, 0, b.Length);
            s.Flush();
        }
        catch { }
    }

    private void Signaler(string message)
    {
        _derniereErreur = message;
        Interlocked.Increment(ref _erreurs);
    }
}
'@
    if (-not ("OfficeProxyRelay" -as [type])) {
        Add-Type -TypeDefinition $code -Language CSharp -ErrorAction Stop
    }
    $relais = New-Object -TypeName OfficeProxyRelay -ArgumentList @($UpstreamHost, $UpstreamPort, $User, $Password)
    $relais.Start()
    $script:ProxyRelay = $relais
    return $relais
}

function Stop-ProxyRelay {
    param($Relais)
    foreach ($r in @($Relais, $script:ProxyRelay)) {
        if ($null -ne $r) { try { $r.Stop() } catch {} }
    }
    $script:ProxyRelay = $null
}

# --- Genere le contenu XML ODT selon les exclusions ---
function Get-XmlContent {
    param(
        [string[]]$ExcludedApps = @(),
        [string]$Language = "fr-fr"
    )
    $lines = $ExcludedApps | ForEach-Object { "      <ExcludeApp ID=`"$_`" />" }
    $excl  = if ($lines) { "`n" + ($lines -join "`n") } else { "" }
    return @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume">
      <Language ID="$Language" />$excl
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
</Configuration>
"@
}

# --- Resume des apps installees (pour affichage UI) ---
function Get-AppsLabel {
    $excluded = $OFFICE_APPS | Where-Object { $script:xmlExcludedApps -contains  $_.ID }
    if ($excluded.Count -eq 0) { return "Toutes les applications incluses" }
    return "Exclues : " + (($excluded | ForEach-Object { $_.Label }) -join ", ")
}

# --- Pre-checks ---
function Get-PreChecks {
    param([string]$SetupPath = $DEFAULT_SETUP_PATH)
    $resolvedSetupPath = Resolve-DeploymentPath $SetupPath
    $drive  = Get-PSDrive C -ErrorAction SilentlyContinue
    $freeGB = if ($drive) { [math]::Round($drive.Free / 1GB, 1) } else { 0 }
    # Une seule lecture du registre pour les deux informations.
    $officeEntries = @(Get-ItemProperty `
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" `
        -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -and (
            $_.DisplayName -like "Microsoft Office*" -or $_.DisplayName -like "Microsoft 365*" -or
            $_.DisplayName -like "*Office*2024*" -or $_.DisplayName -like "*2024*Volume*"
        ) })
    $existing = $officeEntries | Select-Object -First 1
    $has2024  = $null -ne ($officeEntries | Where-Object {
        $_.DisplayName -like "*Office*2024*" -or $_.DisplayName -like "*2024*Volume*"
    } | Select-Object -First 1)
    # Evite les faux resultats dans un EXE compile (manifest absent => version virtualisee).
    $osCaption = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
    $isWin10Plus = $false
    try {
        $major = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop).CurrentMajorVersionNumber
        if ($major -ge 10) { $isWin10Plus = $true }
    } catch {
        $verStr = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Version
        $verObj = $null
        if ([version]::TryParse($verStr, [ref]$verObj)) {
            $isWin10Plus = ($verObj.Major -ge 10)
        }
    }
    return @{
        FreeGB         = $freeGB
        HasEnoughSpace = ($freeGB -ge $DISK_MIN_GB)
        # Test-Path refuse une chaine vide : sans ce garde-fou, Get-PreChecks leve et
        # le panneau pre-vol garde ses anciennes couleurs (vert perime).
        HasSetup       = (-not [string]::IsNullOrWhiteSpace($resolvedSetupPath)) -and (Test-Path -LiteralPath $resolvedSetupPath)
        ResolvedSetup  = $resolvedSetupPath
        HasNet         = [System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()
        IsWin10Plus    = $isWin10Plus
        OsCaption      = $osCaption
        ExistingOffice = if ($existing) { "$($existing.DisplayName) $($existing.DisplayVersion)".Trim() } else { $null }
        HasOffice2024  = $has2024
        Hostname       = $env:COMPUTERNAME
        Username       = $env:USERNAME
        IsAdmin        = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    }
}

# --- Config persistante ---
function Get-Config {
    if (Test-Path $CONFIG_FILE) {
        try { return Get-Content $CONFIG_FILE -Raw | ConvertFrom-Json } catch {}
    }
    return $null
}

function Protect-SecretText {
    param([string]$PlainText)
    if ([string]::IsNullOrWhiteSpace($PlainText)) { return "" }
    try {
        $secure = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
        return ($secure | ConvertFrom-SecureString)
    } catch { return "" }
}

function Unprotect-SecretText {
    param([string]$EncryptedText)
    if ([string]::IsNullOrWhiteSpace($EncryptedText)) { return "" }
    try {
        $secure = ConvertTo-SecureString -String $EncryptedText
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } catch { return "" }
}

function Save-Config {
    param($Cfg)
    try {
        $dir = Split-Path $CONFIG_FILE
        if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $Cfg | ConvertTo-Json -Depth 3 | Out-File -FilePath $CONFIG_FILE -Encoding UTF8 -Force
    } catch {}
}

# --- Confirmation avant lancement ---
function Show-LaunchConfirmation {
    param([bool]$ActivateOnly, [hashtable]$Opts, [string]$KmsRaw, [string]$KmsSecondaryRaw = "")
    $addr = ConvertTo-KmsAddress $KmsRaw
    $kmsLine = "$($addr.Host):$($addr.Port)"
    if (-not [string]::IsNullOrWhiteSpace($KmsSecondaryRaw)) {
        $addr2 = ConvertTo-KmsAddress $KmsSecondaryRaw
        $kmsLine += " (secours: $($addr2.Host):$($addr2.Port))"
    }
    if ($ActivateOnly) {
        $msg = "MODE : ACTIVATION KMS UNIQUEMENT`n" + ("=" * 44) + "`n"
        $msg += "  Poste       : $env:COMPUTERNAME`n"
        $msg += "  Utilisateur : $env:USERNAME`n`n"
        $msg += "  Serveur KMS : $kmsLine`n"
        if ($Opts.ApplyProductKey) { $msg += "  Cle produit : appliquee avant activation`n" }
        $msg += "  Estimation  : ~2 minutes`n`n  Confirmer ?"
    } else {
        $total = 0
        $msg = "MODE : DEPLOIEMENT COMPLET`n" + ("=" * 44) + "`n"
        $msg += "  Poste       : $env:COMPUTERNAME`n"
        $msg += "  Utilisateur : $env:USERNAME`n`n"
        $msg += "  Actions prevues :`n"
        if ($Opts.RemoveOld)    { $msg += "    [x] Desinstallation ancienne version  (~3 min)`n"  ; $total += 3  }
        else                    { $msg += "    [ ] Desinstallation (ignoree)`n" }
        $xmlDesc = if ($script:xmlUseCustomFile) {
                       if ([string]::IsNullOrWhiteSpace($script:xmlCustomFilePath)) {
                           "fichier : non renseigne"
                       } else {
                           "fichier : $(Split-Path -Path $script:xmlCustomFilePath -Leaf)"
                       }
                   } else {
                       "$($OFFICE_APPS.Count - $script:xmlExcludedApps.Count) app(s) incluses"
                   }
        if ($Opts.UseDepot) {
            $msg += "    [x] Sources locales (depot) : aucun telechargement`n" ; $total += 2
            $msg += "        depot : $($Opts.DepotPath)`n"
        } else {
            $msg += "    [x] Telechargement Office 2024 ($($Opts.Language)) ($xmlDesc)  (~18 min)`n" ; $total += 18
            if ($Opts.UseProxy) { $msg += "        via proxy $($Opts.ProxyHost):$($Opts.ProxyPort) (applique au niveau systeme)`n" }
        }
        if ($Opts.AutoInstall)  { $msg += "    [x] Installation silencieuse  (~12 min)`n"              ; $total += 12 }
        else                    { $msg += "    [ ] Installation (differee)`n" }
        if ($Opts.AutoInstall -and $Opts.AutoActivate) {
            $msg += "    [x] Activation KMS : $kmsLine`n"                           ; $total += 1  }
        else                    { $msg += "    [ ] Activation KMS (ignoree)`n" }
        if ($Opts.AutoInstall -and $Opts.AutoActivate -and $Opts.ApplyProductKey) {
            $msg += "    [x] Cle produit appliquee avant activation`n"
        }
        $msg += "`n  Estimation totale : ~$total minutes`n`n  Lancer le deploiement ?"
    }
    return ([System.Windows.Forms.MessageBox]::Show(
        $msg, "Confirmation du deploiement", "YesNo",
        [System.Windows.Forms.MessageBoxIcon]::Question
    ) -eq "Yes")
}

# --- Dialog personnalisation des applications ---
function Show-AppCustomizationDialog {
    param([string]$CurrentLanguage = "fr-fr")

    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Personnalisation des applications Office 2024"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox     = $false
    $dlg.StartPosition   = "CenterParent"
    $dlg.BackColor       = $clrBg

    # Titre
    $lTitle           = New-Object System.Windows.Forms.Label
    $lTitle.Location  = New-Object System.Drawing.Point(14, 12)
    $lTitle.Size      = New-Object System.Drawing.Size(420, 20)
    $lTitle.Text      = "Choisir les applications a installer"
    $lTitle.Font      = $fontSection
    $lTitle.ForeColor = $clrBlueDark
    $dlg.Controls.Add($lTitle)

    $lDesc           = New-Object System.Windows.Forms.Label
    $lDesc.Location  = New-Object System.Drawing.Point(14, 36)
    $lDesc.Size      = New-Object System.Drawing.Size(420, 28)
    $lDesc.Text      = "Applications non cochees : exclues de l'installation.`nRecommande en lycee : decocher Teams, OneDrive, Skype Entreprise."
    $lDesc.Font      = $fontSmall
    $lDesc.ForeColor = $clrGray
    $dlg.Controls.Add($lDesc)

    # Checkboxes par application
    $appCbx = @{}
    $y = 72
    foreach ($app in $OFFICE_APPS) {
        $chk           = New-Object System.Windows.Forms.CheckBox
        $chk.Location  = New-Object System.Drawing.Point(16, $y)
        $chk.Size      = New-Object System.Drawing.Size(190, 22)
        $chk.Text      = $app.Label
        $chk.Font      = $fontNormal
        $chk.Checked   = ($script:xmlExcludedApps -notcontains $app.ID)
        $dlg.Controls.Add($chk)

        $lbl           = New-Object System.Windows.Forms.Label
        $lbl.Location  = New-Object System.Drawing.Point(210, ($y + 4))
        $lbl.Size      = New-Object System.Drawing.Size(220, 16)
        $lbl.Text      = $app.Desc
        $lbl.Font      = $fontSmall
        $lbl.ForeColor = $clrGray
        $dlg.Controls.Add($lbl)

        $appCbx[$app.ID] = $chk
        $y += 26
    }

    # Separateur
    $sep           = New-Object System.Windows.Forms.Label
    $sep.Location  = New-Object System.Drawing.Point(14, ($y + 4))
    $sep.Size      = New-Object System.Drawing.Size(420, 1)
    $sep.BorderStyle = "Fixed3D"
    $dlg.Controls.Add($sep)

    # Apercu XML
    $lPrev           = New-Object System.Windows.Forms.Label
    $lPrev.Location  = New-Object System.Drawing.Point(14, ($y + 12))
    $lPrev.Size      = New-Object System.Drawing.Size(420, 16)
    $lPrev.Text      = "Apercu du fichier XML qui sera genere :"
    $lPrev.Font      = $fontSection
    $lPrev.ForeColor = $clrBlueDark
    $dlg.Controls.Add($lPrev)

    $xmlPrev              = New-Object System.Windows.Forms.RichTextBox
    $xmlPrev.Location     = New-Object System.Drawing.Point(14, ($y + 32))
    $xmlPrev.Size         = New-Object System.Drawing.Size(420, 108)
    $xmlPrev.ReadOnly     = $true
    $xmlPrev.Font         = New-Object System.Drawing.Font("Consolas", 7.5)
    $xmlPrev.BackColor    = [System.Drawing.Color]::FromArgb(238, 242, 246)
    $xmlPrev.BorderStyle  = "FixedSingle"
    $xmlPrev.ScrollBars   = "Vertical"
    $dlg.Controls.Add($xmlPrev)

    # Mise a jour de l'apercu
    $refreshPreview = {
        $excl = @()
        foreach ($app in $OFFICE_APPS) { if (-not $appCbx[$app.ID].Checked) { $excl += $app.ID } }
        $xmlPrev.Text = Get-XmlContent -ExcludedApps $excl -Language $CurrentLanguage
    }
    foreach ($app in $OFFICE_APPS) {
        $appCbx[$app.ID].Add_CheckedChanged($refreshPreview)
    }
    & $refreshPreview

    # Boutons
    $btnY   = $y + 148
    $btnOk  = New-Object System.Windows.Forms.Button
    $btnOk.Location  = New-Object System.Drawing.Point(256, $btnY)
    $btnOk.Size      = New-Object System.Drawing.Size(88, 28)
    $btnOk.Text      = "Appliquer"
    $btnOk.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOk.BackColor = $clrBlueLight
    $btnOk.ForeColor = $clrWhite
    $btnOk.FlatStyle = "Flat"
    $btnOk.FlatAppearance.BorderSize = 0
    $btnOk.DialogResult = "OK"
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk

    $btnCancel  = New-Object System.Windows.Forms.Button
    $btnCancel.Location  = New-Object System.Drawing.Point(352, $btnY)
    $btnCancel.Size      = New-Object System.Drawing.Size(82, 28)
    $btnCancel.Text      = "Annuler"
    $btnCancel.Font      = $fontNormal
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.DialogResult = "Cancel"
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    $dlg.ClientSize = New-Object System.Drawing.Size(448, ($btnY + 44))

    $res = $dlg.ShowDialog()
    if ($res -eq "OK") {
        $excl = @()
        foreach ($app in $OFFICE_APPS) { if (-not $appCbx[$app.ID].Checked) { $excl += $app.ID } }
        return $excl
    }
    return $null
}

# ================================================================
#  FORMULAIRE PRINCIPAL
# ================================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Deploiement Office 2024 - Lycee  [$VERSION]"
$form.ClientSize      = New-Object System.Drawing.Size(520, 982)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.BackColor       = $clrBg

$form.GetType().GetProperty(
    "DoubleBuffered", [System.Reflection.BindingFlags]"NonPublic,Instance"
).SetValue($form, $true, $null)

# ----------------------------------------------------------------
#  EN-TETE
# ----------------------------------------------------------------
$headerPanel           = New-Object System.Windows.Forms.Panel
$headerPanel.Location  = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size      = New-Object System.Drawing.Size(520, 58)
$headerPanel.BackColor = $clrBlueDark
$form.Controls.Add($headerPanel)

$titleLabel            = New-Object System.Windows.Forms.Label
$titleLabel.Location   = New-Object System.Drawing.Point(16, 10)
$titleLabel.Size       = New-Object System.Drawing.Size(400, 28)
$titleLabel.Text       = "  Deploiement Microsoft Office 2024"
$titleLabel.Font       = $fontTitle
$titleLabel.ForeColor  = $clrWhite
$headerPanel.Controls.Add($titleLabel)

$versionLabel          = New-Object System.Windows.Forms.Label
$versionLabel.Location = New-Object System.Drawing.Point(440, 14)
$versionLabel.Size     = New-Object System.Drawing.Size(68, 16)
$versionLabel.Text     = $VERSION
$versionLabel.Font     = $fontSmall
$versionLabel.ForeColor = [System.Drawing.Color]::FromArgb(140, 180, 220)
$versionLabel.TextAlign = "MiddleRight"
$headerPanel.Controls.Add($versionLabel)

$subLabel              = New-Object System.Windows.Forms.Label
$subLabel.Location     = New-Object System.Drawing.Point(18, 38)
$subLabel.Size         = New-Object System.Drawing.Size(490, 14)
$subLabel.Text         = "  Telechargement  .  Installation  .  Activation KMS  -  Lycee Loritz, Grand Est"
$subLabel.Font         = $fontSmall
$subLabel.ForeColor    = [System.Drawing.Color]::FromArgb(160, 200, 235)
$headerPanel.Controls.Add($subLabel)

# ----------------------------------------------------------------
#  BARRE D'ETAPES
# ----------------------------------------------------------------
$stepsData   = @(
    @{ Label="Desinstall."; Short="1" },
    @{ Label="Preparation"; Short="2" },
    @{ Label="Sources";     Short="3" },
    @{ Label="Install.";    Short="4" },
    @{ Label="Activation";  Short="5" }
)
$stepBubbles = @()
$stepLabels  = @()

$stepBarPanel           = New-Object System.Windows.Forms.Panel
$stepBarPanel.Location  = New-Object System.Drawing.Point(0, 58)
$stepBarPanel.Size      = New-Object System.Drawing.Size(520, 72)
$stepBarPanel.BackColor = $clrWhite
$form.Controls.Add($stepBarPanel)

for ($i = 0; $i -lt $stepsData.Count; $i++) {
    $sp = New-Object System.Windows.Forms.Panel
    $sp.Location  = New-Object System.Drawing.Point((10 + $i * 100), 5)
    $sp.Size      = New-Object System.Drawing.Size(98, 62)
    $sp.BackColor = $clrWhite
    $stepBarPanel.Controls.Add($sp)

    $bubble = New-Object System.Windows.Forms.Label
    $bubble.Location  = New-Object System.Drawing.Point(31, 4)
    $bubble.Size      = New-Object System.Drawing.Size(36, 36)
    $bubble.Text      = $stepsData[$i].Short
    $bubble.Font      = $fontStepNum
    $bubble.TextAlign = "MiddleCenter"
    $bubble.BackColor = $clrGray
    $bubble.ForeColor = $clrWhite
    $sp.Controls.Add($bubble)
    $stepBubbles += $bubble

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location  = New-Object System.Drawing.Point(0, 42)
    $lbl.Size      = New-Object System.Drawing.Size(98, 18)
    $lbl.Text      = $stepsData[$i].Label
    $lbl.Font      = $fontStep
    $lbl.ForeColor = $clrGray
    $lbl.TextAlign = "MiddleCenter"
    $sp.Controls.Add($lbl)
    $stepLabels += $lbl

    if ($i -lt ($stepsData.Count - 1)) {
        $ln = New-Object System.Windows.Forms.Label
        $ln.Location  = New-Object System.Drawing.Point(82, 21)
        $ln.Size      = New-Object System.Drawing.Size(18, 2)
        $ln.BackColor = $clrGray
        $sp.Controls.Add($ln)
    }
}

function Set-StepState {
    param([int]$Index, [string]$State)
    switch ($State) {
        "waiting" { $stepBubbles[$Index].BackColor=$clrGray  ; $stepBubbles[$Index].Text=$stepsData[$Index].Short ; $stepLabels[$Index].ForeColor=$clrGray }
        "done"    { $stepBubbles[$Index].BackColor=$clrGreen ; $stepBubbles[$Index].Text="OK"  ; $stepLabels[$Index].ForeColor=$clrGreen }
        "skipped" { $stepBubbles[$Index].BackColor=$clrGray  ; $stepBubbles[$Index].Text="-"   ; $stepLabels[$Index].ForeColor=$clrGray }
        "error"   { $stepBubbles[$Index].BackColor=$clrRed   ; $stepBubbles[$Index].Text="ERR" ; $stepLabels[$Index].ForeColor=$clrRed }
    }
}

# ----------------------------------------------------------------
#  PANNEAU PRE-CHECKS
# ----------------------------------------------------------------
$checkPanel             = New-Object System.Windows.Forms.Panel
$checkPanel.Location    = New-Object System.Drawing.Point(12, 140)
$checkPanel.Size        = New-Object System.Drawing.Size(492, 110)
$checkPanel.BackColor   = $clrWhite
$checkPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($checkPanel)

$checkTitle             = New-Object System.Windows.Forms.Label
$checkTitle.Location    = New-Object System.Drawing.Point(10, 8)
$checkTitle.Size        = New-Object System.Drawing.Size(300, 20)
$checkTitle.Text        = "Verifications pre-vol"
$checkTitle.Font        = $fontSection
$checkTitle.ForeColor   = $clrBlueDark
$checkPanel.Controls.Add($checkTitle)

$refreshBtn             = New-Object System.Windows.Forms.Button
$refreshBtn.Location    = New-Object System.Drawing.Point(390, 5)
$refreshBtn.Size        = New-Object System.Drawing.Size(88, 22)
$refreshBtn.Text        = "Rafraichir"
$refreshBtn.Font        = $fontSmall
$refreshBtn.FlatStyle   = "Flat"
$refreshBtn.FlatAppearance.BorderColor = $clrGray
$checkPanel.Controls.Add($refreshBtn)

$checkItems = @(
    @{ Key="setup"; X=10;  Label="setup.exe" },
    @{ Key="disk";  X=130; Label="Espace disque" },
    @{ Key="net";   X=260; Label="Reseau" },
    @{ Key="os";    X=374; Label="Windows" }
)
$checkIcons  = @{}
$checkLabels = @{}
foreach ($ci in $checkItems) {
    $icon = New-Object System.Windows.Forms.Label
    $icon.Location  = New-Object System.Drawing.Point($ci.X, 36)
    $icon.Size      = New-Object System.Drawing.Size(20, 20)
    $icon.Text      = "?"
    $icon.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $icon.ForeColor = $clrGray
    $icon.TextAlign = "MiddleCenter"
    $checkPanel.Controls.Add($icon)
    $checkIcons[$ci.Key] = $icon

    $clbl = New-Object System.Windows.Forms.Label
    $clbl.Location  = New-Object System.Drawing.Point(($ci.X + 22), 38)
    $clbl.Size      = New-Object System.Drawing.Size(100, 16)
    $clbl.Text      = $ci.Label
    $clbl.Font      = $fontSmall
    $clbl.ForeColor = $clrGray
    $checkPanel.Controls.Add($clbl)
    $checkLabels[$ci.Key] = $clbl
}

$sysInfoLabel           = New-Object System.Windows.Forms.Label
$sysInfoLabel.Location  = New-Object System.Drawing.Point(10, 64)
$sysInfoLabel.Size      = New-Object System.Drawing.Size(470, 16)
$sysInfoLabel.Font      = $fontSmall
$sysInfoLabel.ForeColor = $clrGray
$checkPanel.Controls.Add($sysInfoLabel)

$office2024WarnLabel        = New-Object System.Windows.Forms.Label
$office2024WarnLabel.Location = New-Object System.Drawing.Point(10, 82)
$office2024WarnLabel.Size   = New-Object System.Drawing.Size(470, 16)
$office2024WarnLabel.Font   = $fontSmall
$office2024WarnLabel.ForeColor = $clrOrange
$checkPanel.Controls.Add($office2024WarnLabel)

# ----------------------------------------------------------------
#  PANNEAU OPTIONS  (h=262 - inclut setup.exe + XML config)
# ----------------------------------------------------------------
$optPanel             = New-Object System.Windows.Forms.Panel
$optPanel.Location    = New-Object System.Drawing.Point(12, 260)
$optPanel.Size        = New-Object System.Drawing.Size(492, 500)
$optPanel.BackColor   = $clrWhite
$optPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($optPanel)

$optTitle             = New-Object System.Windows.Forms.Label
$optTitle.Location    = New-Object System.Drawing.Point(10, 8)
$optTitle.Size        = New-Object System.Drawing.Size(460, 20)
$optTitle.Text        = "Options"
$optTitle.Font        = $fontSection
$optTitle.ForeColor   = $clrBlueDark
$optPanel.Controls.Add($optTitle)

# ---- Ligne 1 : setup.exe ----
$setupLabel           = New-Object System.Windows.Forms.Label
$setupLabel.Location  = New-Object System.Drawing.Point(10, 34)
$setupLabel.Size      = New-Object System.Drawing.Size(106, 22)
$setupLabel.Text      = "Chemin setup.exe :"
$setupLabel.Font      = $fontSmall
$setupLabel.TextAlign = "MiddleLeft"
$optPanel.Controls.Add($setupLabel)

$setupPathBox         = New-Object System.Windows.Forms.TextBox
$setupPathBox.Location = New-Object System.Drawing.Point(118, 32)
$setupPathBox.Size    = New-Object System.Drawing.Size(252, 22)
$setupPathBox.Text    = $DEFAULT_SETUP_PATH
$setupPathBox.Font    = $fontSmall
$optPanel.Controls.Add($setupPathBox)

$setupBrowseBtn       = New-Object System.Windows.Forms.Button
$setupBrowseBtn.Location = New-Object System.Drawing.Point(374, 31)
$setupBrowseBtn.Size  = New-Object System.Drawing.Size(104, 24)
$setupBrowseBtn.Text  = "Parcourir..."
$setupBrowseBtn.Font  = $fontSmall
$setupBrowseBtn.FlatStyle = "Flat"
$setupBrowseBtn.FlatAppearance.BorderColor = $clrGray
$optPanel.Controls.Add($setupBrowseBtn)

$setupBrowseBtn.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title  = "Selectionner setup.exe (Office Deployment Tool)"
    $ofd.Filter = "Executables (*.exe)|*.exe|Tous les fichiers (*.*)|*.*"
    $ofd.FileName = "setup.exe"
    if (Test-Path (Split-Path $setupPathBox.Text -Parent -ErrorAction SilentlyContinue)) {
        $ofd.InitialDirectory = Split-Path $setupPathBox.Text -Parent
    }
    if ($ofd.ShowDialog() -eq "OK") {
        $setupPathBox.Text = $ofd.FileName
        Update-PreChecks | Out-Null
    }
})

# ---- Separateur ----
$sep1           = New-Object System.Windows.Forms.Label
$sep1.Location  = New-Object System.Drawing.Point(10, 60)
$sep1.Size      = New-Object System.Drawing.Size(468, 1)
$sep1.BorderStyle = "Fixed3D"
$optPanel.Controls.Add($sep1)

# ---- Ligne 2 : Langue ----
$langLabel            = New-Object System.Windows.Forms.Label
$langLabel.Location   = New-Object System.Drawing.Point(10, 68)
$langLabel.Size       = New-Object System.Drawing.Size(130, 22)
$langLabel.Text       = "Langue d'installation :"
$langLabel.Font       = $fontNormal
$optPanel.Controls.Add($langLabel)

$langCombo            = New-Object System.Windows.Forms.ComboBox
$langCombo.Location   = New-Object System.Drawing.Point(144, 65)
$langCombo.Size       = New-Object System.Drawing.Size(130, 22)
$langCombo.DropDownStyle = "DropDownList"
$langCombo.Font       = $fontNormal
$langCombo.Items.AddRange(@("fr-fr", "en-us"))
$langCombo.SelectedIndex = 0
$optPanel.Controls.Add($langCombo)

# ---- Ligne 3 : Desinstaller ----
$chkRemoveOld         = New-Object System.Windows.Forms.CheckBox
$chkRemoveOld.Location = New-Object System.Drawing.Point(10, 92)
$chkRemoveOld.Size    = New-Object System.Drawing.Size(468, 22)
$chkRemoveOld.Text    = "Desinstaller l'ancienne version d'Office avant installation"
$chkRemoveOld.Font    = $fontNormal
$optPanel.Controls.Add($chkRemoveOld)

# ---- Ligne 4 : Installer auto ----
$chkInstall           = New-Object System.Windows.Forms.CheckBox
$chkInstall.Location  = New-Object System.Drawing.Point(10, 116)
$chkInstall.Size      = New-Object System.Drawing.Size(468, 22)
$chkInstall.Text      = "Lancer l'installation automatiquement apres le telechargement"
$chkInstall.Font      = $fontNormal
$chkInstall.Checked   = $true
$optPanel.Controls.Add($chkInstall)

# ---- Ligne 5 : KMS (accepte host ou host:port) ----
$chkActivate          = New-Object System.Windows.Forms.CheckBox
$chkActivate.Location = New-Object System.Drawing.Point(10, 140)
$chkActivate.Size     = New-Object System.Drawing.Size(128, 22)
$chkActivate.Text     = "Activer via KMS"
$chkActivate.Font     = $fontNormal
$chkActivate.Checked  = $true
$optPanel.Controls.Add($chkActivate)

$kmsTextBox           = New-Object System.Windows.Forms.TextBox
$kmsTextBox.Location  = New-Object System.Drawing.Point(142, 140)
$kmsTextBox.Size      = New-Object System.Drawing.Size(184, 22)
$kmsTextBox.Text      = $KMS_SERVEUR
$kmsTextBox.Font      = $fontNormal
$kmsTextBox.ForeColor = $clrGray
$optPanel.Controls.Add($kmsTextBox)

$kmsTextBox.Add_Enter({
    if ($kmsTextBox.Text -eq $KMS_SERVEUR -and $kmsTextBox.ForeColor -eq $clrGray) {
        $kmsTextBox.Text = ""
        $kmsTextBox.ForeColor = [System.Drawing.Color]::Black
    }
})

$kmsTextBox.Add_KeyPress({
    if ($kmsTextBox.Text -eq $KMS_SERVEUR -and $kmsTextBox.ForeColor -eq $clrGray -and -not [char]::IsControl($_.KeyChar)) {
        $kmsTextBox.Text = $_.KeyChar
        $kmsTextBox.SelectionStart = $kmsTextBox.Text.Length
        $kmsTextBox.ForeColor = [System.Drawing.Color]::Black
        $_.Handled = $true
    }
})

$kmsTextBox.Add_Leave({
    if ([string]::IsNullOrWhiteSpace($kmsTextBox.Text)) {
        $kmsTextBox.Text = $KMS_SERVEUR
        $kmsTextBox.ForeColor = $clrGray
    } else {
        $kmsTextBox.ForeColor = [System.Drawing.Color]::Black
    }
})

$kmsTestBtn           = New-Object System.Windows.Forms.Button
$kmsTestBtn.Location  = New-Object System.Drawing.Point(330, 139)
$kmsTestBtn.Size      = New-Object System.Drawing.Size(148, 24)
$kmsTestBtn.Text      = "Tester la connexion"
$kmsTestBtn.Font      = $fontSmall
$kmsTestBtn.FlatStyle = "Flat"
$kmsTestBtn.FlatAppearance.BorderColor = $clrGray
$optPanel.Controls.Add($kmsTestBtn)
$script:kmsTestContext = $null
$script:deployWorkerContext = $null

# Test KMS asynchrone (Runspace)
$kmsTestBtn.Add_Click({
    if ($kmsTestBtn.Tag -eq "running") {
        $ctx = $script:kmsTestContext
        try { if ($ctx -and $ctx.Timer) { $ctx.Timer.Stop() } } catch {}
        try { if ($ctx -and $ctx.PS) { $ctx.PS.Stop() } } catch {}
        try { if ($ctx -and $ctx.PS) { $ctx.PS.Dispose() } } catch {}
        try { if ($ctx -and $ctx.Runspace) { $ctx.Runspace.Close(); $ctx.Runspace.Dispose() } } catch {}

        $kmsTestBtn.Tag = $null
        $script:kmsTestContext = $null
        $kmsTestBtn.Text = "Tester la connexion"
        $kmsTestBtn.Enabled = $chkActivate.Checked
        [System.Windows.Forms.MessageBox]::Show("Test de connexion KMS interrompu manuellement.", "Test interrompu", "OK", "Information") | Out-Null
        return
    }

    $raw = $kmsTextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($raw) -or $raw -eq "KMS.VOTRE-SERVEUR.FR") {
        [System.Windows.Forms.MessageBox]::Show("Renseignez l'adresse KMS (ex: kms.grandest.fr ou kms.grandest.fr:1688).", "Adresse manquante", "OK", "Warning") | Out-Null
        return
    }
    $addr = ConvertTo-KmsAddress $raw
    $kmsTestBtn.Tag = "running"
    $kmsTestBtn.Text = "Arreter le test" ; $kmsTestBtn.Enabled = $true ; $form.Refresh()

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.Open()
    $rs.SessionStateProxy.SetVariable("srv", $addr.Host)
    $rs.SessionStateProxy.SetVariable("prt", [int]$addr.Port)
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        $tcp = $null
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $r   = $tcp.BeginConnect($srv, $prt, $null, $null)
            if (-not $r.AsyncWaitHandle.WaitOne(3000, $false)) { return $false }
            $tcp.EndConnect($r)
            return $tcp.Connected
        } catch { return $false }
        finally { if ($tcp) { try { $tcp.Close() } catch {} } }
    })
    $handle = $ps.BeginInvoke()

    $kt = New-Object System.Windows.Forms.Timer ; $kt.Interval = 100
    $script:kmsTestContext = @{
        Timer          = $kt
        PS             = $ps
        Runspace       = $rs
        Handle         = $handle
        StartedAt      = Get-Date
        MaxDurationSec = 12
        Host           = $addr.Host
        Port           = $addr.Port
    }
    $kt.Add_Tick({
        $ctx = $script:kmsTestContext
        if (-not $ctx) { return }

        $timedOut = (-not $ctx.Handle.IsCompleted) -and (((Get-Date) - $ctx.StartedAt).TotalSeconds -ge $ctx.MaxDurationSec)
        if ((-not $ctx.Handle.IsCompleted) -and (-not $timedOut)) { return }

        $ctx.Timer.Stop()
        $ok = $null
        try {
            if ($timedOut) {
                try { $ctx.PS.Stop() } catch {}
                $ok = $false
            } else {
                $ok = $ctx.PS.EndInvoke($ctx.Handle)[0]
            }
        } catch { $ok = $false }
        finally {
            try { $ctx.PS.Dispose() } catch {}
            try { $ctx.Runspace.Close(); $ctx.Runspace.Dispose() } catch {}
            $kmsTestBtn.Tag = $null
            $script:kmsTestContext = $null
            $kmsTestBtn.Text = "Tester la connexion"
            $kmsTestBtn.Enabled = $chkActivate.Checked
        }

        if ($timedOut) {
            [System.Windows.Forms.MessageBox]::Show(
                "Le test KMS a depasse le delai ($($ctx.MaxDurationSec) s) et a ete arrete.`n`nReessayez ou verifiez le reseau/DNS.",
                "Timeout du test KMS", "OK", "Warning") | Out-Null
        } elseif ($ok) {
            [System.Windows.Forms.MessageBox]::Show(
                "Connexion KMS reussie !`n`nServeur : $($ctx.Host)`nPort    : $($ctx.Port)`n`nL'activation devrait fonctionner sur ce reseau.",
                "Connexion KMS OK", "OK", "Information") | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Connexion KMS echouee.`n`nServeur : $($ctx.Host)`nPort    : $($ctx.Port)`n`nSolutions :`n- PC sur le reseau lycee (cable ou VPN)`n- Verifiez l'adresse avec l'IT Grand Est`n- Port $($ctx.Port) doit etre ouvert",
                "Connexion KMS echouee", "OK", "Warning") | Out-Null
        }
    })
    $kt.Start()
})

$chkActivate.Add_CheckedChanged({ $kmsTextBox.Enabled=$chkActivate.Checked ; $kmsTestBtn.Enabled=$chkActivate.Checked })
$chkInstall.Add_CheckedChanged({
    if (-not $chkInstall.Checked) {
        $chkActivate.Checked=$false ; $chkActivate.Enabled=$false
        $kmsTextBox.Enabled=$false  ; $kmsTestBtn.Enabled=$false
    } else {
        $chkActivate.Enabled=$true
        $kmsTextBox.Enabled=$chkActivate.Checked
        $kmsTestBtn.Enabled=$chkActivate.Checked
    }
})

# ---- Ligne 5b : depot de sources (mode sans telechargement) ----
$chkUseDepot            = New-Object System.Windows.Forms.CheckBox
$chkUseDepot.Location   = New-Object System.Drawing.Point(10, 168)
$chkUseDepot.Size       = New-Object System.Drawing.Size(216, 22)
$chkUseDepot.Text       = "Depot de sources (sans telechargement)"
$chkUseDepot.Font       = $fontNormal
$optPanel.Controls.Add($chkUseDepot)

$depotPathBox           = New-Object System.Windows.Forms.TextBox
$depotPathBox.Location  = New-Object System.Drawing.Point(228, 167)
$depotPathBox.Size      = New-Object System.Drawing.Size(158, 22)
$depotPathBox.Font      = $fontSmall
$depotPathBox.Text      = ""
$depotPathBox.Enabled   = $false
$optPanel.Controls.Add($depotPathBox)

$depotBrowseBtn         = New-Object System.Windows.Forms.Button
$depotBrowseBtn.Location = New-Object System.Drawing.Point(390, 166)
$depotBrowseBtn.Size    = New-Object System.Drawing.Size(88, 24)
$depotBrowseBtn.Text    = "Parcourir..."
$depotBrowseBtn.Font    = $fontSmall
$depotBrowseBtn.FlatStyle = "Flat"
$depotBrowseBtn.FlatAppearance.BorderColor = $clrGray
$depotBrowseBtn.Enabled = $false
$optPanel.Controls.Add($depotBrowseBtn)

$depotBrowseBtn.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Dossier contenant le sous-dossier 'Office' produit par setup.exe /download"
    $fbd.ShowNewFolderButton = $false
    $current = $depotPathBox.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($current) -and (Test-Path $current)) { $fbd.SelectedPath = $current }
    if ($fbd.ShowDialog() -eq "OK") { $depotPathBox.Text = $fbd.SelectedPath }
})

$chkUseDepot.Add_CheckedChanged({
    $depotPathBox.Enabled   = $chkUseDepot.Checked
    $depotBrowseBtn.Enabled = $chkUseDepot.Checked
    $chkUseProxy.Enabled    = -not $chkUseDepot.Checked
    if ($chkUseDepot.Checked -and -not $chkActivate.Checked -and -not $chkInstall.Checked) { $chkActivate.Checked = $true }
    Update-ProxyControlState
})

# ---- Separateur XML ----
$sep2           = New-Object System.Windows.Forms.Label
$sep2.Location  = New-Object System.Drawing.Point(10, 196)
$sep2.Size      = New-Object System.Drawing.Size(468, 1)
$sep2.BorderStyle = "Fixed3D"
$optPanel.Controls.Add($sep2)

$xmlSectionLbl           = New-Object System.Windows.Forms.Label
$xmlSectionLbl.Location  = New-Object System.Drawing.Point(10, 204)
$xmlSectionLbl.Size      = New-Object System.Drawing.Size(200, 16)
$xmlSectionLbl.Text      = "Configuration XML :"
$xmlSectionLbl.Font      = $fontSection
$xmlSectionLbl.ForeColor = $clrBlueDark
$optPanel.Controls.Add($xmlSectionLbl)

# ---- Ligne 6 : XML auto + personnaliser ----
$radioXmlAuto           = New-Object System.Windows.Forms.RadioButton
$radioXmlAuto.Location  = New-Object System.Drawing.Point(10, 224)
$radioXmlAuto.Size      = New-Object System.Drawing.Size(170, 20)
$radioXmlAuto.Text      = "Generer automatiquement"
$radioXmlAuto.Font      = $fontNormal
$radioXmlAuto.Checked   = $true
$optPanel.Controls.Add($radioXmlAuto)

$btnCustomizeApps           = New-Object System.Windows.Forms.Button
$btnCustomizeApps.Location  = New-Object System.Drawing.Point(184, 222)
$btnCustomizeApps.Size      = New-Object System.Drawing.Size(156, 24)
$btnCustomizeApps.Text      = "Personnaliser les apps..."
$btnCustomizeApps.Font      = $fontSmall
$btnCustomizeApps.FlatStyle = "Flat"
$btnCustomizeApps.FlatAppearance.BorderColor = $clrBlueLight
$optPanel.Controls.Add($btnCustomizeApps)

# Label montrant les apps actuellement exclues
$xmlAppsLabel           = New-Object System.Windows.Forms.Label
$xmlAppsLabel.Location  = New-Object System.Drawing.Point(10, 248)
$xmlAppsLabel.Size      = New-Object System.Drawing.Size(468, 14)
$xmlAppsLabel.Text      = Get-AppsLabel
$xmlAppsLabel.Font      = $fontSmall
$xmlAppsLabel.ForeColor = $clrGray
$optPanel.Controls.Add($xmlAppsLabel)

$btnCustomizeApps.Add_Click({
    $result = Show-AppCustomizationDialog -CurrentLanguage $langCombo.SelectedItem
    if ($null -ne $result) {
        $script:xmlExcludedApps = $result
        $xmlAppsLabel.Text = Get-AppsLabel
    }
})

# ---- Ligne 7 : XML fichier existant ----
$radioXmlFile           = New-Object System.Windows.Forms.RadioButton
$radioXmlFile.Location  = New-Object System.Drawing.Point(10, 264)
$radioXmlFile.Size      = New-Object System.Drawing.Size(130, 20)
$radioXmlFile.Text      = "Utiliser un fichier :"
$radioXmlFile.Font      = $fontNormal
$optPanel.Controls.Add($radioXmlFile)

$xmlFilePathBox           = New-Object System.Windows.Forms.TextBox
$xmlFilePathBox.Location  = New-Object System.Drawing.Point(144, 263)
$xmlFilePathBox.Size      = New-Object System.Drawing.Size(200, 22)
$xmlFilePathBox.Text      = ""
$xmlFilePathBox.Font      = $fontSmall
$xmlFilePathBox.ForeColor = $clrGray
$xmlFilePathBox.Enabled   = $false
$optPanel.Controls.Add($xmlFilePathBox)

$xmlBrowseBtn           = New-Object System.Windows.Forms.Button
$xmlBrowseBtn.Location  = New-Object System.Drawing.Point(348, 262)
$xmlBrowseBtn.Size      = New-Object System.Drawing.Size(130, 24)
$xmlBrowseBtn.Text      = "Parcourir..."
$xmlBrowseBtn.Font      = $fontSmall
$xmlBrowseBtn.FlatStyle = "Flat"
$xmlBrowseBtn.FlatAppearance.BorderColor = $clrGray
$xmlBrowseBtn.Enabled   = $false
$optPanel.Controls.Add($xmlBrowseBtn)

$xmlBrowseBtn.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title  = "Selectionner un fichier de configuration XML (ODT)"
    $ofd.Filter = "Fichiers XML (*.xml)|*.xml|Tous les fichiers (*.*)|*.*"
    if ($ofd.ShowDialog() -eq "OK") {
        $xmlFilePathBox.Text = $ofd.FileName
        $script:xmlCustomFilePath = $ofd.FileName
    }
})

$radioXmlAuto.Add_CheckedChanged({
    $script:xmlUseCustomFile   = -not $radioXmlAuto.Checked
    $btnCustomizeApps.Enabled  = $radioXmlAuto.Checked
    $xmlAppsLabel.Visible      = $radioXmlAuto.Checked
    $xmlFilePathBox.Enabled    = -not $radioXmlAuto.Checked
    $xmlBrowseBtn.Enabled      = -not $radioXmlAuto.Checked
})

# ---- Separateur Proxy ----
$sep3           = New-Object System.Windows.Forms.Label
$sep3.Location  = New-Object System.Drawing.Point(10, 292)
$sep3.Size      = New-Object System.Drawing.Size(468, 1)
$sep3.BorderStyle = "Fixed3D"
$optPanel.Controls.Add($sep3)

$proxySectionLbl           = New-Object System.Windows.Forms.Label
$proxySectionLbl.Location  = New-Object System.Drawing.Point(10, 300)
$proxySectionLbl.Size      = New-Object System.Drawing.Size(200, 16)
$proxySectionLbl.Text      = "Configuration Proxy :"
$proxySectionLbl.Font      = $fontSection
$proxySectionLbl.ForeColor = $clrBlueDark
$optPanel.Controls.Add($proxySectionLbl)

$chkUseProxy          = New-Object System.Windows.Forms.CheckBox
$chkUseProxy.Location = New-Object System.Drawing.Point(10, 318)
$chkUseProxy.Size     = New-Object System.Drawing.Size(210, 22)
$chkUseProxy.Text     = "Utiliser un proxy pour le telechargement"
$chkUseProxy.Font     = $fontNormal
$optPanel.Controls.Add($chkUseProxy)

$proxyHostLabel          = New-Object System.Windows.Forms.Label
$proxyHostLabel.Location = New-Object System.Drawing.Point(10, 342)
$proxyHostLabel.Size     = New-Object System.Drawing.Size(94, 20)
$proxyHostLabel.Text     = "Adresse proxy :"
$proxyHostLabel.Font     = $fontSmall
$optPanel.Controls.Add($proxyHostLabel)

$proxyHostBox           = New-Object System.Windows.Forms.TextBox
$proxyHostBox.Location  = New-Object System.Drawing.Point(106, 340)
$proxyHostBox.Size      = New-Object System.Drawing.Size(180, 22)
$proxyHostBox.Font      = $fontSmall
$proxyHostBox.Text      = ""
$optPanel.Controls.Add($proxyHostBox)

$proxyPortLabel          = New-Object System.Windows.Forms.Label
$proxyPortLabel.Location = New-Object System.Drawing.Point(292, 342)
$proxyPortLabel.Size     = New-Object System.Drawing.Size(36, 20)
$proxyPortLabel.Text     = "Port :"
$proxyPortLabel.Font     = $fontSmall
$optPanel.Controls.Add($proxyPortLabel)

$proxyPortBox           = New-Object System.Windows.Forms.TextBox
$proxyPortBox.Location  = New-Object System.Drawing.Point(332, 340)
$proxyPortBox.Size      = New-Object System.Drawing.Size(64, 22)
$proxyPortBox.Font      = $fontSmall
$proxyPortBox.Text      = "8080"
$optPanel.Controls.Add($proxyPortBox)

# Le service Click-to-Run (LocalSystem) ne peut pas repondre a un 407 : les
# identifiants sont transmis par un relais local (Start-ProxyRelay).
$proxyAuthUserLabel          = New-Object System.Windows.Forms.Label
$proxyAuthUserLabel.Location = New-Object System.Drawing.Point(10, 364)
$proxyAuthUserLabel.Size     = New-Object System.Drawing.Size(66, 18)
$proxyAuthUserLabel.Text     = "Utilisateur :"
$proxyAuthUserLabel.Font     = $fontSmall
$optPanel.Controls.Add($proxyAuthUserLabel)

$proxyAuthUserBox           = New-Object System.Windows.Forms.TextBox
$proxyAuthUserBox.Location  = New-Object System.Drawing.Point(78, 362)
$proxyAuthUserBox.Size      = New-Object System.Drawing.Size(148, 22)
$proxyAuthUserBox.Font      = $fontSmall
$proxyAuthUserBox.Text      = ""
$optPanel.Controls.Add($proxyAuthUserBox)

$proxyAuthPassLabel          = New-Object System.Windows.Forms.Label
$proxyAuthPassLabel.Location = New-Object System.Drawing.Point(232, 364)
$proxyAuthPassLabel.Size     = New-Object System.Drawing.Size(78, 18)
$proxyAuthPassLabel.Text     = "Mot de passe :"
$proxyAuthPassLabel.Font     = $fontSmall
$optPanel.Controls.Add($proxyAuthPassLabel)

$proxyAuthPassBox                       = New-Object System.Windows.Forms.TextBox
$proxyAuthPassBox.Location              = New-Object System.Drawing.Point(314, 362)
$proxyAuthPassBox.Size                  = New-Object System.Drawing.Size(156, 22)
$proxyAuthPassBox.Font                  = $fontSmall
$proxyAuthPassBox.Text                  = ""
$proxyAuthPassBox.UseSystemPasswordChar = $true
$optPanel.Controls.Add($proxyAuthPassBox)

$proxyAuthTip = New-Object System.Windows.Forms.ToolTip
$proxyAuthTip.SetToolTip($proxyAuthUserBox, "Identifiants du proxy, utiles si le test repond 407. Le relais local les transmet au proxy ; le mot de passe n'est jamais enregistre.")
$proxyAuthTip.SetToolTip($proxyAuthPassBox, "Mot de passe du proxy, garde en memoire pendant le deploiement seulement : a ressaisir a chaque lancement.")

$proxyBypassLabel          = New-Object System.Windows.Forms.Label
$proxyBypassLabel.Location = New-Object System.Drawing.Point(10, 390)
$proxyBypassLabel.Size     = New-Object System.Drawing.Size(94, 20)
$proxyBypassLabel.Text     = "Exceptions :"
$proxyBypassLabel.Font     = $fontSmall
$optPanel.Controls.Add($proxyBypassLabel)

$proxyBypassBox           = New-Object System.Windows.Forms.TextBox
$proxyBypassBox.Location  = New-Object System.Drawing.Point(106, 388)
$proxyBypassBox.Size      = New-Object System.Drawing.Size(180, 22)
$proxyBypassBox.Font      = $fontSmall
$proxyBypassBox.Text      = $PROXY_BYPASS_DEF
$optPanel.Controls.Add($proxyBypassBox)

$proxyTestBtn           = New-Object System.Windows.Forms.Button
$proxyTestBtn.Location  = New-Object System.Drawing.Point(292, 387)
$proxyTestBtn.Size      = New-Object System.Drawing.Size(186, 24)
$proxyTestBtn.Text      = "Tester le proxy"
$proxyTestBtn.Font      = $fontSmall
$proxyTestBtn.FlatStyle = "Flat"
$proxyTestBtn.FlatAppearance.BorderColor = $clrGray
$optPanel.Controls.Add($proxyTestBtn)

function Update-ProxyControlState {
    $enabled = $chkUseProxy.Checked
    $proxyHostLabel.Enabled = $enabled
    $proxyHostBox.Enabled = $enabled
    $proxyPortLabel.Enabled = $enabled
    $proxyPortBox.Enabled = $enabled
    $proxyBypassLabel.Enabled = $enabled
    $proxyBypassBox.Enabled = $enabled
    $proxyAuthUserLabel.Enabled = $enabled
    $proxyAuthUserBox.Enabled = $enabled
    $proxyAuthPassLabel.Enabled = $enabled
    $proxyAuthPassBox.Enabled = $enabled
    $proxyTestBtn.Enabled = $enabled
    if ($null -ne $chkUseDepot) { $chkUseProxy.Enabled = -not $chkUseDepot.Checked }
}

# Test du proxy : verifie le proxy ET la destination Microsoft (CONNECT HTTP)
$proxyTestBtn.Add_Click({
    $h = $proxyHostBox.Text.Trim()
    $p = 0
    if ([string]::IsNullOrWhiteSpace($h) -or -not [int]::TryParse($proxyPortBox.Text.Trim(), [ref]$p) -or $p -le 0 -or $p -gt 65535) {
        [System.Windows.Forms.MessageBox]::Show("Renseignez une adresse et un port de proxy valides avant le test.", "Test proxy", "OK", "Warning") | Out-Null
        return
    }
    $proxyTestBtn.Enabled = $false
    $proxyTestBtn.Text    = "Test en cours..."
    $form.UseWaitCursor   = $true
    $statusLabel.Text     = "Test du proxy ${h}:${p} (quelques secondes)..."
    [System.Windows.Forms.Application]::DoEvents()

    $lines = @()
    $allOk = $true
    try {
        # Identifiants saisis : tester a travers le meme relais que le deploiement.
        $relaisTest = $null ; $hTest = $h ; $pTest = $p
        if (-not [string]::IsNullOrWhiteSpace($proxyAuthUserBox.Text.Trim())) {
            try {
                $relaisTest = Start-ProxyRelay -UpstreamHost $h -UpstreamPort $p -User $proxyAuthUserBox.Text.Trim() -Password $proxyAuthPassBox.Text
                $hTest = "127.0.0.1" ; $pTest = $relaisTest.Port
                $lines += "Relais local 127.0.0.1:$($relaisTest.Port) -> ${h}:${p} (identifiants transmis au proxy)."
            } catch {
                $allOk = $false
                $lines += "Relais local impossible : $($_.Exception.Message)"
            }
        }
        foreach ($r in (Test-ProxyReadiness -ProxyHost $hTest -ProxyPort $pTest)) {
            if (-not $r.OK) { $allOk = $false }
            $mark = if ($r.OK) { "OK " } else { "ERR" }
            $lines += "[$mark] $($r.Name)  -  $($r.Detail)"
            if ($r.Code -eq 407) {
                if ($null -ne $relaisTest) {
                    $lines += "       -> 407 : identifiants refuses par le proxy. Verifiez l'utilisateur et le mot de passe ci-dessus."
                } else {
                    $lines += "       -> 407 : le proxy exige une authentification. Renseignez l'utilisateur et le mot de passe ci-dessus : un relais local les transmet au proxy (le service Click-to-Run ne peut pas s'authentifier lui-meme)."
                }
            }
        }
    } catch {
        $allOk = $false
        $lines += "Erreur du test : $($_.Exception.Message)"
    }

    if ($null -ne $relaisTest) { Stop-ProxyRelay -Relais $relaisTest }
    $form.UseWaitCursor = $false
    $proxyTestBtn.Enabled = $chkUseProxy.Checked
    $proxyTestBtn.Text = "Tester le proxy"
    $statusLabel.Text = "Pret."
    $head = if ($allOk) { "Proxy operationnel : ODT peut joindre le CDN Microsoft." } else { "Test en echec - detail ci-dessous." }
    $icon = if ($allOk) { "Information" } else { "Warning" }
    [System.Windows.Forms.MessageBox]::Show("$head`n`n" + ($lines -join "`n"), "Test du proxy", "OK", $icon) | Out-Null
})

Update-ProxyControlState

$chkUseProxy.Add_CheckedChanged({ Update-ProxyControlState })
Update-ProxyControlState

# ---- Separateur KMS avance ----
$sep4           = New-Object System.Windows.Forms.Label
$sep4.Location  = New-Object System.Drawing.Point(10, 412)
$sep4.Size      = New-Object System.Drawing.Size(468, 1)
$sep4.BorderStyle = "Fixed3D"
$optPanel.Controls.Add($sep4)

$kmsAdvLbl           = New-Object System.Windows.Forms.Label
$kmsAdvLbl.Location  = New-Object System.Drawing.Point(10, 420)
$kmsAdvLbl.Size      = New-Object System.Drawing.Size(220, 16)
$kmsAdvLbl.Text      = "Activation KMS avancee :"
$kmsAdvLbl.Font      = $fontSection
$kmsAdvLbl.ForeColor = $clrBlueDark
$optPanel.Controls.Add($kmsAdvLbl)

$kmsSecondaryLbl           = New-Object System.Windows.Forms.Label
$kmsSecondaryLbl.Location  = New-Object System.Drawing.Point(10, 442)
$kmsSecondaryLbl.Size      = New-Object System.Drawing.Size(130, 20)
$kmsSecondaryLbl.Text      = "Serveur KMS secours :"
$kmsSecondaryLbl.Font      = $fontSmall
$optPanel.Controls.Add($kmsSecondaryLbl)

$kmsSecondaryBox           = New-Object System.Windows.Forms.TextBox
$kmsSecondaryBox.Location  = New-Object System.Drawing.Point(142, 440)
$kmsSecondaryBox.Size      = New-Object System.Drawing.Size(336, 22)
$kmsSecondaryBox.Text      = ""
$kmsSecondaryBox.Font      = $fontSmall
$optPanel.Controls.Add($kmsSecondaryBox)

$chkApplyProductKey          = New-Object System.Windows.Forms.CheckBox
$chkApplyProductKey.Location = New-Object System.Drawing.Point(10, 466)
$chkApplyProductKey.Size     = New-Object System.Drawing.Size(236, 22)
$chkApplyProductKey.Text     = "Appliquer une cle produit avant activation"
$chkApplyProductKey.Font     = $fontNormal
$chkApplyProductKey.Checked  = $false
$optPanel.Controls.Add($chkApplyProductKey)

$productKeyBox           = New-Object System.Windows.Forms.TextBox
$productKeyBox.Location  = New-Object System.Drawing.Point(250, 466)
$productKeyBox.Size      = New-Object System.Drawing.Size(228, 22)
$productKeyBox.Font      = $fontSmall
$productKeyBox.Enabled   = $false
$productKeyBox.Text      = ""
$optPanel.Controls.Add($productKeyBox)

$chkApplyProductKey.Add_CheckedChanged({
    $productKeyBox.Enabled = $chkApplyProductKey.Checked
    if (-not $chkApplyProductKey.Checked) { $productKeyBox.Text = "" }
})

# ----------------------------------------------------------------
#  JOURNAL
# ----------------------------------------------------------------
$logPanel             = New-Object System.Windows.Forms.Panel
$logPanel.Location    = New-Object System.Drawing.Point(12, 770)
$logPanel.Size        = New-Object System.Drawing.Size(492, 150)
$logPanel.BackColor   = $clrLogBg
$logPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($logPanel)

$logHeaderBg           = New-Object System.Windows.Forms.Panel
$logHeaderBg.Location  = New-Object System.Drawing.Point(0, 0)
$logHeaderBg.Size      = New-Object System.Drawing.Size(492, 22)
$logHeaderBg.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$logPanel.Controls.Add($logHeaderBg)

$logHeaderLbl           = New-Object System.Windows.Forms.Label
$logHeaderLbl.Location  = New-Object System.Drawing.Point(4, 3)
$logHeaderLbl.Size      = New-Object System.Drawing.Size(168, 16)
$logHeaderLbl.Text      = "Journal d'execution"
$logHeaderLbl.Font      = $fontSection
$logHeaderLbl.ForeColor = [System.Drawing.Color]::FromArgb(180, 200, 220)
$logHeaderLbl.BackColor = [System.Drawing.Color]::Transparent
$logHeaderBg.Controls.Add($logHeaderLbl)

# Bouton Copier
$copyLogBtn           = New-Object System.Windows.Forms.Button
$copyLogBtn.Location  = New-Object System.Drawing.Point(176, 1)
$copyLogBtn.Size      = New-Object System.Drawing.Size(84, 20)
$copyLogBtn.Text      = "Copier"
$copyLogBtn.Font      = $fontSmall
$copyLogBtn.FlatStyle = "Flat" ; $copyLogBtn.FlatAppearance.BorderSize=0
$copyLogBtn.BackColor = [System.Drawing.Color]::FromArgb(40,40,40)
$copyLogBtn.ForeColor = [System.Drawing.Color]::FromArgb(140,180,220)
$logHeaderBg.Controls.Add($copyLogBtn)
$copyLogBtn.Add_Click({ if ($logBox.Text.Trim()) { [System.Windows.Forms.Clipboard]::SetText($logBox.Text) ; [System.Windows.Forms.MessageBox]::Show("Journal copie dans le presse-papier.","Copie","OK","Information")|Out-Null } })

# Bouton Exporter
$exportLogBtn           = New-Object System.Windows.Forms.Button
$exportLogBtn.Location  = New-Object System.Drawing.Point(262, 1)
$exportLogBtn.Size      = New-Object System.Drawing.Size(114, 20)
$exportLogBtn.Text      = "Exporter..."
$exportLogBtn.Font      = $fontSmall
$exportLogBtn.FlatStyle = "Flat" ; $exportLogBtn.FlatAppearance.BorderSize=0
$exportLogBtn.BackColor = [System.Drawing.Color]::FromArgb(40,40,40)
$exportLogBtn.ForeColor = [System.Drawing.Color]::FromArgb(140,180,220)
$logHeaderBg.Controls.Add($exportLogBtn)
$exportLogBtn.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "Texte (*.txt)|*.txt|Tous (*.*)|*.*"
    $dlg.FileName = "Office2024_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if ($dlg.ShowDialog() -eq "OK") { $logBox.Text | Out-File $dlg.FileName -Encoding UTF8 ; [System.Windows.Forms.MessageBox]::Show("Exporte : $($dlg.FileName)","Export","OK","Information")|Out-Null }
})

# Bouton Ouvrir dossier
$openLogsBtn           = New-Object System.Windows.Forms.Button
$openLogsBtn.Location  = New-Object System.Drawing.Point(378, 1)
$openLogsBtn.Size      = New-Object System.Drawing.Size(112, 20)
$openLogsBtn.Text      = "Ouvrir dossier"
$openLogsBtn.Font      = $fontSmall
$openLogsBtn.FlatStyle = "Flat" ; $openLogsBtn.FlatAppearance.BorderSize=0
$openLogsBtn.BackColor = [System.Drawing.Color]::FromArgb(40,40,40)
$openLogsBtn.ForeColor = [System.Drawing.Color]::FromArgb(140,180,220)
$logHeaderBg.Controls.Add($openLogsBtn)
$openLogsBtn.Add_Click({ if (Test-Path $LOG_FOLDER) { Start-Process "explorer.exe" $LOG_FOLDER } else { [System.Windows.Forms.MessageBox]::Show("Dossier cree lors du premier deploiement.`n$LOG_FOLDER","Info","OK","Information")|Out-Null } })

$logBox               = New-Object System.Windows.Forms.RichTextBox
$logBox.Location      = New-Object System.Drawing.Point(0, 22)
$logBox.Size          = New-Object System.Drawing.Size(490, 124)
$logBox.ReadOnly      = $true
$logBox.BackColor     = $clrLogBg
$logBox.ForeColor     = [System.Drawing.Color]::FromArgb(200, 200, 200)
$logBox.Font          = $fontMono
$logBox.BorderStyle   = "None"
$logBox.ScrollBars    = "Vertical"
$logPanel.Controls.Add($logBox)

# ----------------------------------------------------------------
#  BARRE DE STATUT
# ----------------------------------------------------------------
$statusBar            = New-Object System.Windows.Forms.Panel
$statusBar.Location   = New-Object System.Drawing.Point(0, 930)
$statusBar.Size       = New-Object System.Drawing.Size(520, 52)
$statusBar.BackColor  = $clrWhite
$form.Controls.Add($statusBar)

$spinnerLabel         = New-Object System.Windows.Forms.Label
$spinnerLabel.Location = New-Object System.Drawing.Point(12, 8)
$spinnerLabel.Size    = New-Object System.Drawing.Size(18, 18)
$spinnerLabel.Font    = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
$spinnerLabel.ForeColor = $clrBlueLight
$statusBar.Controls.Add($spinnerLabel)

$statusLabel          = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(34, 6)
$statusLabel.Size     = New-Object System.Drawing.Size(200, 18)
$statusLabel.Text     = "Pret."
$statusLabel.Font     = $fontNormal
$statusLabel.ForeColor = $clrGray
$statusBar.Controls.Add($statusLabel)

$chronoLabel          = New-Object System.Windows.Forms.Label
$chronoLabel.Location = New-Object System.Drawing.Point(34, 26)
$chronoLabel.Size     = New-Object System.Drawing.Size(238, 16)
$chronoLabel.Visible  = $false
$chronoLabel.Font     = New-Object System.Drawing.Font("Consolas", 8)
$chronoLabel.ForeColor = $clrGray
$statusBar.Controls.Add($chronoLabel)

$progressBar          = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(34, 26)
$progressBar.Size     = New-Object System.Drawing.Size(238, 14)
$progressBar.Style    = "Marquee"
$progressBar.MarqueeAnimationSpeed = 25
$progressBar.Visible  = $false
$statusBar.Controls.Add($progressBar)

$startButton          = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(272, 9)
$startButton.Size     = New-Object System.Drawing.Size(122, 32)
$startButton.Text     = "Lancer deploiement"
$startButton.Font     = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$startButton.BackColor = $clrGray
$startButton.ForeColor = $clrWhite
$startButton.FlatStyle = "Flat"
$startButton.FlatAppearance.BorderSize = 0
$statusBar.Controls.Add($startButton)

$activateOnlyBtn          = New-Object System.Windows.Forms.Button
$activateOnlyBtn.Location = New-Object System.Drawing.Point(398, 9)
$activateOnlyBtn.Size     = New-Object System.Drawing.Size(110, 32)
$activateOnlyBtn.Text     = "Activer seulement"
$activateOnlyBtn.Font     = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$activateOnlyBtn.BackColor = $clrGreen
$activateOnlyBtn.ForeColor = $clrWhite
$activateOnlyBtn.FlatStyle = "Flat"
$activateOnlyBtn.FlatAppearance.BorderSize = 0
$statusBar.Controls.Add($activateOnlyBtn)

# ================================================================
#  TOOLTIPS
# ================================================================
$tooltip              = New-Object System.Windows.Forms.ToolTip
$tooltip.AutoPopDelay = 9000
$tooltip.InitialDelay = 600
$tooltip.ReshowDelay  = 300

$tooltip.SetToolTip($setupPathBox,    "Chemin complet vers setup.exe (outil ODT Microsoft).`nLe nom du fichier peut etre different de 'setup.exe'.")
$tooltip.SetToolTip($setupBrowseBtn,  "Ouvrir un selecteur de fichier pour localiser setup.exe.")
$tooltip.SetToolTip($refreshBtn,      "Relance les 4 verifications et met a jour les indicateurs.")
$tooltip.SetToolTip($chkRemoveOld,    "Supprime Office 2016/2019/365 avant installation.`nPeut prendre plusieurs minutes.")
$tooltip.SetToolTip($chkInstall,      "Lance l'installation des que le telechargement est termine.`nDe-cocher pour preparer les sources sans installer (multi-postes).")
$tooltip.SetToolTip($chkActivate,     "Active Office via KMS apres installation.`nNecessite le reseau lycee ou le VPN.")
$tooltip.SetToolTip($kmsTextBox,      "Adresse du serveur KMS.`nFormats acceptes :`n  kms.grandest.fr`n  kms.grandest.fr:1688`n  192.168.1.10:1688")
$tooltip.SetToolTip($kmsTestBtn,      "Test de connexion TCP asynchrone au serveur KMS.`nNe bloque pas l'interface.")
$tooltip.SetToolTip($btnCustomizeApps,"Choisir quelles applications Office installer ou exclure.`nApercu XML mis a jour en temps reel.")
$tooltip.SetToolTip($radioXmlAuto,    "Genere automatiquement le fichier XML selon les options selectionnees.")
$tooltip.SetToolTip($radioXmlFile,    "Utilise un fichier XML deja configure (ex : configuration IT Grand Est).")
$tooltip.SetToolTip($xmlBrowseBtn,    "Selectionner un fichier .xml genere avec l'outil ODT Configuration Tool.")
$tooltip.SetToolTip($chkUseProxy,     "Active un proxy HTTP(S) pour le telechargement Office via setup.exe.")
$tooltip.SetToolTip($proxyHostBox,    "Nom DNS ou IP du proxy (ex: proxy.grandest.fr).")
$tooltip.SetToolTip($proxyPortBox,    "Port du proxy (ex: 8080 ou 3128).")
$tooltip.SetToolTip($proxyBypassBox,  "Hotes a joindre SANS passer par le proxy (separees par ;).`nExemple : <local>;*.grandest.fr;10.0.0.0/8")
$tooltip.SetToolTip($proxyTestBtn,    "Teste le proxy ET la destination officecdn.microsoft.com (CONNECT HTTP).`nDetecte : proxy injoignable, destination bloquee, authentification 407.")
$tooltip.SetToolTip($chkUseDepot,     "Installe depuis des sources deja telechargees (dossier local ou partage).`nAucun acces Internet necessaire : installation + activation KMS sur le reseau pedagogique.")
$tooltip.SetToolTip($depotPathBox,    "Dossier contenant le sous-dossier 'Office' produit par setup.exe /download.`nExemple : \\\\srv-fichiers\\Logiciels\\Office2024")
$tooltip.SetToolTip($depotBrowseBtn,  "Choisir le dossier de depot des sources Office 2024.")
$tooltip.SetToolTip($kmsSecondaryBox, "Serveur KMS de secours (optionnel), format host ou host:port.")
$tooltip.SetToolTip($chkApplyProductKey, "Applique /inpkey avant /act. A utiliser selon les consignes de la region.")
$tooltip.SetToolTip($productKeyBox,   "Cle produit (format XXXXX-XXXXX-XXXXX-XXXXX-XXXXX).")
$tooltip.SetToolTip($copyLogBtn,      "Copie le journal dans le presse-papier (pour coller dans un ticket support).")
$tooltip.SetToolTip($activateOnlyBtn, "Pour les postes ou Office est installe mais pas encore active.")
$tooltip.SetToolTip($langCombo,       "Langue de l'interface Office installee. fr-fr = francais, en-us = anglais.")

# ================================================================
#  NOTIFY ICON
# ================================================================
$notifyIcon             = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon        = [System.Drawing.SystemIcons]::Information
$notifyIcon.Visible     = $true
$notifyIcon.Text        = "Office 2024 Deploiement - Lycee"
$notifyIcon.Add_Click({ $form.WindowState = "Normal" ; $form.Activate() })

# ================================================================
#  Update-PreChecks  (declaree apres tooltip + setupPathBox)
# ================================================================
function Update-PreChecks {
    $c = Get-PreChecks -SetupPath $setupPathBox.Text.Trim()
    $resolvedSetupPath = if ($c.ResolvedSetup) { $c.ResolvedSetup } else { Resolve-DeploymentPath $setupPathBox.Text.Trim() }

    if ($c.HasSetup) {
        $checkIcons["setup"].Text="OK" ; $checkIcons["setup"].ForeColor=$clrGreen
        $checkLabels["setup"].Text="setup.exe OK" ; $checkLabels["setup"].ForeColor=$clrGreen
    } else {
        $checkIcons["setup"].Text="X" ; $checkIcons["setup"].ForeColor=$clrRed
        $checkLabels["setup"].Text="setup.exe absent" ; $checkLabels["setup"].ForeColor=$clrRed
    }
    if ($c.HasEnoughSpace) {
        $checkIcons["disk"].Text="OK" ; $checkIcons["disk"].ForeColor=$clrGreen
        $checkLabels["disk"].Text="C:\ $($c.FreeGB) Go libres" ; $checkLabels["disk"].ForeColor=$clrGreen
    } else {
        $checkIcons["disk"].Text="X" ; $checkIcons["disk"].ForeColor=$clrRed
        $checkLabels["disk"].Text="C:\ $($c.FreeGB)Go (min $DISK_MIN_GB)" ; $checkLabels["disk"].ForeColor=$clrRed
    }
    if ($c.HasNet) {
        $checkIcons["net"].Text="OK" ; $checkIcons["net"].ForeColor=$clrGreen
        $checkLabels["net"].Text="Liaison locale" ; $checkLabels["net"].ForeColor=$clrGreen
    } else {
        $checkIcons["net"].Text="!" ; $checkIcons["net"].ForeColor=$clrOrange
        $checkLabels["net"].Text="Aucune liaison" ; $checkLabels["net"].ForeColor=$clrOrange
    }
    if ($c.IsWin10Plus) {
        $checkIcons["os"].Text="OK" ; $checkIcons["os"].ForeColor=$clrGreen
        $checkLabels["os"].Text="Windows OK" ; $checkLabels["os"].ForeColor=$clrGreen
    } else {
        $checkIcons["os"].Text="X" ; $checkIcons["os"].ForeColor=$clrRed
        $checkLabels["os"].Text="Win 10 requis" ; $checkLabels["os"].ForeColor=$clrRed
    }

    $officeInfo = if ($c.ExistingOffice) { "Detecte : $($c.ExistingOffice)" } else { "Aucun Office detecte" }
    $adminInfo = if ($c.IsAdmin) { "Admin : OUI" } else { "Admin : NON (relancer en Admin !)" }
    $adminColor = if ($c.IsAdmin) { $clrGray } else { $clrOrange }
    $sysInfoLabel.ForeColor = $adminColor
    $sysInfoLabel.Text = "PC : $($c.Hostname)  |  $($c.Username)  |  $officeInfo  |  $adminInfo"
    $office2024WarnLabel.Text = if ($c.HasOffice2024) { "Office 2024 deja detecte. Cocher 'Desinstaller' ou utiliser 'Activer seulement'." } else { "" }

    $canStart = $c.HasSetup -and $c.HasEnoughSpace -and $c.IsWin10Plus
    # Ne jamais reactiver les boutons pendant un deploiement en cours.
    $busy = ($null -ne $script:deployWorkerContext)
    $startButton.Enabled   = ($canStart -and -not $busy)
    $activateOnlyBtn.Enabled = -not $busy
    $startButton.BackColor = if ($canStart -and -not $busy) { $clrBlueLight } else { $clrGray }
    if (-not $canStart) {
        $issues = @()
        if (-not $c.HasSetup)       { $issues += "setup.exe absent (cliquez Parcourir pour le localiser)" }
        if (-not $c.HasEnoughSpace) { $issues += "espace insuffisant : $($c.FreeGB) Go libres, min $DISK_MIN_GB Go" }
        if (-not $c.IsWin10Plus)    { $issues += "Windows 10+ requis" }
        $tooltip.SetToolTip($startButton, "IMPOSSIBLE DE LANCER :`n" + ($issues -join "`n") + "`n`nCliquez Rafraichir apres correction.")
    } else {
        $tooltip.SetToolTip($startButton, "Lance le deploiement complet selon les options selectionnees.")
    }
    return $c
}

$refreshBtn.Add_Click({ Update-PreChecks | Out-Null })
$setupPathBox.Add_Leave({ Update-PreChecks | Out-Null })

# ================================================================
#  TIMER D'ANIMATION
# ================================================================
$animTimer          = New-Object System.Windows.Forms.Timer
$animTimer.Interval = 350
$script:animFrame   = 0
$script:logLines    = 0
$script:lastState   = [string[]]@("","","","","")
$spinFrames         = @("|", "/", "-", "\")

$animTimer.Add_Tick({
    $script:animFrame = ($script:animFrame + 1) % 4

    $entry = $null ; $count = 0
    while ($syncHash.LogQueue.TryDequeue([ref]$entry) -and $count -lt 30) {
        $logBox.SelectionColor = $entry.Color
        $logBox.AppendText("$($entry.Text)`n")
        $count++
    }
    if ($count -gt 0) {
        # Borner le journal : sur un telechargement long, le RichTextBox ralentit l'interface.
        $script:logLines += $count
        if ($script:logLines -gt 2000) {
            $cut = $logBox.GetFirstCharIndexFromLine(500)
            if ($cut -gt 0) { $logBox.Select(0, $cut) ; $logBox.SelectedText = "" ; $script:logLines -= 500 }
        }
        $logBox.ScrollToCaret()
    }

    $dw = $script:deployWorkerContext
    if ($dw -and $dw.Handle.IsCompleted) {
        try {
            $null = $dw.PS.EndInvoke($dw.Handle)
        } catch {
            $errMsg = $_.Exception.ToString()
            try {
                $syncHash.LogQueue.Enqueue(@{Text="[$(Get-Date -Format 'HH:mm:ss')] [ERR]  Erreur runspace worker : $errMsg";Color=[System.Drawing.Color]::FromArgb(240,80,80)})
            } catch {}
            $syncHash.ActiveStep = -1
            $syncHash.HasError = $true
            $syncHash.FinalMsg = "Erreur runspace worker : $errMsg"
            $syncHash.StatusText = "Erreur runspace - voir journal."
            $syncHash.Done = $true
        } finally {
            try { $dw.PS.Dispose() } catch {}
            try { $dw.Runspace.Close(); $dw.Runspace.Dispose() } catch {}
            $script:deployWorkerContext = $null
            Stop-ProxyRelay
        }
    }

    for ($i = 0; $i -lt 5; $i++) {
        $ns = $syncHash.StepState[$i]
        if ($ns -ne $script:lastState[$i]) {
            $script:lastState[$i] = $ns
            if ($ns -ne "running") { Set-StepState $i $ns }
        }
    }
    $active = $syncHash.ActiveStep
    if ($active -ge 0 -and $syncHash.StepState[$active] -eq "running") {
        $stepBubbles[$active].BackColor = $clrBlueLight ; $stepBubbles[$active].ForeColor = $clrWhite
        $stepBubbles[$active].Text = $spinFrames[$script:animFrame]
        $stepLabels[$active].ForeColor = $clrBlueLight
    }

    $spinnerLabel.Text = if ($syncHash.ActiveStep -ge 0 -and -not $syncHash.Done) { $spinFrames[$script:animFrame] } else { "" }
    $statusLabel.Text  = $syncHash.StatusText

    if ($syncHash.ActiveStep -ge 0 -and -not $syncHash.Done -and $null -ne $syncHash.StartTime) {
        $elapsed = (Get-Date) - $syncHash.StartTime
        $form.Text = "Deploiement en cours...  {0:mm}:{0:ss}  [$VERSION]" -f $elapsed
        $remMin = 0
        for ($s = $syncHash.ActiveStep; $s -lt 5; $s++) {
            if ($syncHash.StepState[$s] -in @("waiting","running")) {
                $mins = $STEP_ESTIMATES[$s]
                if ($s -eq $syncHash.ActiveStep -and $null -ne $syncHash.StepStartTime) {
                    $stepElapsed = ((Get-Date) - $syncHash.StepStartTime).TotalMinutes
                    $mins = [Math]::Max(0, $mins - [int]$stepElapsed)
                }
                $remMin += $mins
            }
        }
        $chronoLabel.Text = "{0:mm}m {0:ss}s ecoule  |  ~$remMin min restantes" -f $elapsed
        $chronoLabel.Visible = $true ; $progressBar.Visible = $true
    } elseif ($syncHash.Done) {
        $form.Text = if ($syncHash.HasError) { "Deploiement - Erreurs  [$VERSION]" } else { "Deploiement termine !  [$VERSION]" }
    } else { $form.Text = "Deploiement Office 2024 - Lycee  [$VERSION]" }

    if ($syncHash.ShowPopup) {
        $syncHash.ShowPopup = $false
        $icon = switch ($syncHash.PopupIcon) { "Error"{[System.Windows.Forms.MessageBoxIcon]::Error} "Warning"{[System.Windows.Forms.MessageBoxIcon]::Warning} default{[System.Windows.Forms.MessageBoxIcon]::Information} }
        [System.Windows.Forms.MessageBox]::Show($syncHash.PopupMsg, $syncHash.PopupTitle, "OK", $icon) | Out-Null
    }

    if ($syncHash.Done) {
        $animTimer.Stop()
        $progressBar.Visible = $false ; $chronoLabel.Visible = $false ; $spinnerLabel.Text = ""
        $startButton.Enabled = $true ; $activateOnlyBtn.Enabled = $true
        $startButton.Text = "Lancer deploiement"
        $activateOnlyBtn.Text = "Activer seulement"
        $statusLabel.ForeColor = $clrGray
        $form.UseWaitCursor = $false

        try {
            if (!(Test-Path $LOG_FOLDER)) { New-Item -ItemType Directory -Path $LOG_FOLDER -Force | Out-Null }
            $logBox.Text | Out-File -FilePath "$LOG_FOLDER\deploy_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" -Encoding UTF8
        } catch {}

        if ($syncHash.HasError) {
            try { [System.Media.SystemSounds]::Hand.Play() } catch {}
            $notifyIcon.ShowBalloonTip(8000, "Deploiement - Erreurs", "Des erreurs ont ete detectees. Consultez le journal.", [System.Windows.Forms.ToolTipIcon]::Error)
        } else {
            try { [System.Media.SystemSounds]::Asterisk.Play() } catch {}
            $notifyIcon.ShowBalloonTip(8000, "Office 2024 deploye !", "PC : $env:COMPUTERNAME - Succes.", [System.Windows.Forms.ToolTipIcon]::Info)
        }

        $icon  = if ($syncHash.HasError) { "Warning" } else { "Information" }
        $title = if ($syncHash.HasError) { "Deploiement - Erreurs" } else { "Deploiement termine !" }
        $extra = if (-not $syncHash.HasError) { "`n`nJournal sauvegarde dans : $LOG_FOLDER" } else { "" }
        [System.Windows.Forms.MessageBox]::Show("$($syncHash.FinalMsg)$extra", $title, "OK", [System.Windows.Forms.MessageBoxIcon]::$icon) | Out-Null
        Update-PreChecks | Out-Null
    }
})

# ================================================================
#  WORKER ARRIERE-PLAN
# ================================================================

$script:BgDoWorkHandler = {
    param($eventSender, $e)
    $cfg  = $e.Argument
    $sync = $cfg.SyncHash

    $logQueue = $null
    try { $logQueue = $sync['LogQueue'] } catch {}
    if ($null -eq $logQueue) {
        $logQueue = New-Object System.Collections.Concurrent.ConcurrentQueue[object]
        try { $sync['LogQueue'] = $logQueue } catch {}
    }

    $CLR = @{
        INFO  = [System.Drawing.Color]::FromArgb(180, 180, 180)
        OK    = [System.Drawing.Color]::FromArgb(80,  200, 120)
        WARN  = [System.Drawing.Color]::FromArgb(240, 180,  60)
        ERROR = [System.Drawing.Color]::FromArgb(240,  80,  80)
        STEP  = [System.Drawing.Color]::FromArgb(80,  180, 255)
    }
    $PFX = @{ INFO="       "; OK=" [OK]  "; WARN=" [!]   "; ERROR=" [ERR] "; STEP="---    " }

    function bg_Log { param([string]$M,[string]$L="INFO") $logQueue.Enqueue(@{Text="[$(Get-Date -Format 'HH:mm:ss')]$($PFX[$L])$M";Color=$CLR[$L]}) ; if($L -ne "STEP"){$sync.StatusText=$M} }
    function bg_Step { param([int]$i,[string]$S) $sync.StepState[$i]=$S ; if($S -eq "running"){$sync.ActiveStep=$i; $sync.StepStartTime=(Get-Date)} }
    function bg_Popup { param([string]$T,[string]$M,[string]$I="Error") $sync.PopupTitle=$T;$sync.PopupMsg=$M;$sync.PopupIcon=$I;$sync.ShowPopup=$true ; Start-Sleep -Milliseconds 200 }
    function bg_Abort { param([string]$M) $sync.HasError=$true;$sync.FinalMsg=$M;$sync.Done=$true }
    function bg_InvokeUninstall {
        param([string]$CommandLine)
        if ([string]::IsNullOrWhiteSpace($CommandLine)) { return 0 }

        $line = $CommandLine.Trim()
        if ($line -match '^"([^\"]+)"\s*(.*)$') {
            $exe  = $matches[1]
            $args = $matches[2]
        } else {
            $parts = $line.Split(' ', 2)
            $exe   = $parts[0]
            $args  = if ($parts.Count -gt 1) { $parts[1] } else { "" }
        }

        if ($exe -match '(?i)msiexec(\.exe)?$' -and $args -notmatch '(?i)(/quiet|/qn)') {
            $args = "$args /qn /norestart".Trim()
        }

        $p = Start-Process -FilePath $exe -ArgumentList $args -Wait -NoNewWindow -PassThru -ErrorAction Stop
        return $p.ExitCode
    }
    function bg_ConvertToKmsAddress {
        param([string]$Raw)
        $s = ([string]$Raw).Trim()
        if ($s -match '^(.+):(\d{1,5})$') {
            return @{ Host = $matches[1].Trim(); Port = $matches[2].Trim() }
        }
        return @{ Host = $s; Port = "1688" }
    }
    function bg_NormalizeProductKey {
        param([string]$RawKey)
        if ($null -eq $RawKey) { return "" }
        $k = $RawKey.Trim().ToUpper()
        if ([string]::IsNullOrWhiteSpace($k)) { return "" }
        $k = $k -replace '[\u2010\u2011\u2012\u2013\u2014\u2212\uFE58\uFE63\uFF0D]', '-'
        $k = $k -replace '\s', ''
        $compact = $k -replace '[^A-Z0-9]', ''
        if ($compact.Length -eq 25) {
            return (($compact -split '(.{5})' | Where-Object { $_ }) -join '-')
        }
        return $k
    }

    # --- Le depot est-il sur un disque local fixe (pas de copie necessaire) ? ---
    function bg_TestDepotLocal {
        param([string]$Path)
        try {
            if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
            if ($Path -match '^\\\\') { return $false }
            $root = [System.IO.Path]::GetPathRoot($Path)
            if ([string]::IsNullOrWhiteSpace($root)) { return $false }
            $di = New-Object System.IO.DriveInfo($root)
            return ($di.DriveType -eq [System.IO.DriveType]::Fixed)
        } catch { return $false }
    }

    # --- Impose SourcePath dans le XML avant tout telechargement ---
    #     Sans cet attribut, ODT ecrit les sources (~3 Go) dans le repertoire courant
    #     du processus, donc potentiellement dans le dossier de l'outil.
    function bg_EnsureSourcePath {
        param([string]$XmlPath, [string]$SourcePath, [string]$WorkDir)
        if ([string]::IsNullOrWhiteSpace($XmlPath)) { throw "XML de deploiement introuvable (chemin vide)." }
        if (-not (Test-Path -LiteralPath $XmlPath)) { throw "XML de deploiement introuvable : $XmlPath" }
        if ([string]::IsNullOrWhiteSpace($SourcePath)) { throw "Destination des sources inconnue : impossible d'imposer SourcePath." }

        $doc = New-Object System.Xml.XmlDocument
        try { $doc.Load($XmlPath) } catch { throw "XML illisible ($XmlPath) : $($_.Exception.Message)" }

        $adds = @($doc.SelectNodes("//*[local-name()='Add']"))
        if ($adds.Count -eq 0) { throw "XML invalide : aucun element <Add> dans $XmlPath." }

        $missing = 0
        foreach ($a in $adds) {
            if ([string]::IsNullOrWhiteSpace([string]$a.GetAttribute("SourcePath"))) {
                $att = $doc.CreateAttribute("SourcePath")
                $att.Value = $SourcePath
                [void]$a.Attributes.Append($att)
                $missing++
            }
        }
        # SourcePath deja present : on remonte celui du XML, c'est la vraie destination.
        if ($missing -eq 0) { return @{ Path = $XmlPath ; Injected = $false ; Source = [string]$adds[0].GetAttribute("SourcePath") } }

        # On ne modifie jamais le fichier fourni : copie dediee (telechargement comme installation).
        if (-not (Test-Path -LiteralPath $WorkDir)) { New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null }
        $out = Join-Path $WorkDir "OfficeConfig2024.sourcepath.xml"
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $settings.Indent = $true
        $writer = [System.Xml.XmlWriter]::Create($out, $settings)
        try { $doc.Save($writer) } finally { $writer.Close() }
        return @{ Path = $out ; Injected = $true ; Source = $SourcePath }
    }

    # --- Prepare le XML pour ODT : destination des sources verifiee ou imposee ---
    function bg_PrepareXmlSource {
        param($Cfg, [string]$Etape)
        $target = if (-not [string]::IsNullOrWhiteSpace([string]$Cfg.SourcePath)) { [string]$Cfg.SourcePath }
                  elseif ($Cfg.DepotEnabled -and -not [string]::IsNullOrWhiteSpace([string]$Cfg.DepotPath)) { [string]$Cfg.DepotPath }
                  else { [string]$Cfg.LocalDataPath }
        $info = bg_EnsureSourcePath -XmlPath ([string]$Cfg.ActualXMLPath) -SourcePath $target -WorkDir ([string]$Cfg.LocalDataPath)
        if ($info.Injected) {
            $Cfg.ActualXMLPath = $info.Path
            bg_Log ("XML sans SourcePath : $($info.Source) impose via une copie ($(Split-Path $info.Path -Leaf)) pour $Etape.") "WARN"
        }
        return $info
    }

    # --- Verdict d'espace pour la destination des sources (pur, donc testable) ---
    function bg_SpaceVerdict {
        param($FreeGB, [int]$NeedGB, [string]$Root)
        if ($null -eq $FreeGB) {
            return @{ Level = "WARN" ; Msg = "Espace libre non verifiable a cet emplacement (chemin reseau ou inaccessible) : prevoyez environ $NeedGB Go pour les sources." }
        }
        if ($FreeGB -lt $NeedGB) {
            return @{ Level = "WARN" ; Msg = "Espace insuffisant sur ${Root} : $FreeGB Go libres, environ $NeedGB Go necessaires pour les sources. Le telechargement va echouer : liberez de l'espace ou choisissez un autre dossier de travail." }
        }
        return @{ Level = "OK" ; Msg = "Espace disque OK a cet emplacement : $FreeGB Go libres sur ${Root} (environ $NeedGB Go necessaires)." }
    }

    # --- Acces systeme du proxy (un seul point, remplace par les controles automatises) ---
    #     Defaut : netsh (WinHTTP machine), registre HKCU (WinINet session), environnement.
    function bg_ProxySink {
        param($Cfg)
        if ($Cfg -and $Cfg.ProxySink) { return $Cfg.ProxySink }
        $connKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
        $inetKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
        return @{
            Read = {
                $whValue = $null ; $whExists = $false
                try { $whValue = Get-ItemPropertyValue -Path $connKey -Name "WinHttpSettings" -ErrorAction Stop ; $whExists = $true } catch {}
                $whDirect = -not $whExists
                if ($whExists) {
                    try { $whDirect = (((& netsh winhttp show proxy 2>&1) -join " ") -match '(?i)Direct access|No proxy server|acc.s direct|aucun serveur proxy') } catch {}
                }
                $ini = @{}
                foreach ($p in @("ProxyEnable","ProxyServer","ProxyOverride")) {
                    $ini[$p] = $null
                    try { $ini[$p] = Get-ItemPropertyValue -Path $inetKey -Name $p -ErrorAction Stop } catch {}
                }
                # dictionnaire sensible a la casse : HTTP_PROXY et http_proxy sont deux variables
                $envSt = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)
                foreach ($n in @("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy")) { $envSt[$n] = [System.Environment]::GetEnvironmentVariable($n, "Process") }
                @{ WinHttpExists = $whExists ; WinHttpValue = $whValue ; WinHttpDirect = $whDirect ; WinInet = $ini ; Env = $envSt }
            }
            Apply = {
                param($server,$bypass)
                try {
                    & netsh winhttp set proxy "$server" "$bypass" 2>&1 | Out-Null
                    $after = ((& netsh winhttp show proxy 2>&1) -join " ").Trim()
                    @{ Confirmed = ($after -match [regex]::Escape(($server -split ":")[0])) ; Detail = $after ; Error = "" }
                } catch { @{ Confirmed = $false ; Detail = "" ; Error = $_.Exception.Message } }
            }
            RestoreWinHttp = {
                param($State)
                if ($State.WinHttpDirect -or -not $State.WinHttpExists) { & netsh winhttp reset proxy 2>&1 | Out-Null }
                else { Set-ItemProperty -Path $connKey -Name "WinHttpSettings" -Value $State.WinHttpValue -Type Binary -Force }
            }
            WriteWinInet = {
                param($Values)
                $ok = $true
                foreach ($p in @("ProxyEnable","ProxyServer","ProxyOverride")) {
                    $t = if ($p -eq "ProxyEnable") { "DWord" } else { "String" }
                    try {
                        if ($null -eq $Values[$p]) { Remove-ItemProperty -Path $inetKey -Name $p -ErrorAction SilentlyContinue }
                        else { New-ItemProperty -Path $inetKey -Name $p -Value $Values[$p] -PropertyType $t -Force | Out-Null }
                    } catch { $ok = $false }
                }
                $ok
            }
            WriteEnv = {
                param($Values)
                foreach ($n in @("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy")) {
                    try { [System.Environment]::SetEnvironmentVariable($n, $Values[$n], "Process") } catch {}
                }
            }
        }
    }

    # --- Application du proxy : WinHTTP machine + WinINet session + environnement ---
    function bg_ApplyProxy {
        param($Cfg, [string]$ProxyHost, [string]$ProxyPort, [string]$BypassList)
        $sink = bg_ProxySink $Cfg
        $server = "${ProxyHost}:${ProxyPort}"
        if ([string]::IsNullOrWhiteSpace($BypassList)) { $BypassList = "<local>" }
        $st = & $sink.Read
        $st.Server = $server

        $res = & $sink.Apply $server $BypassList
        if ($res.Error) { bg_Log "Echec netsh winhttp set proxy : $($res.Error)" "WARN" }
        elseif ($res.Confirmed) { bg_Log "Proxy WinHTTP (machine) applique : $server  |  exceptions : $BypassList" "OK" }
        else { bg_Log "netsh winhttp n'a pas confirme le proxy : $($res.Detail)" "WARN" }

        if (& $sink.WriteWinInet @{ ProxyEnable = 1 ; ProxyServer = $server ; ProxyOverride = $BypassList }) { bg_Log "Proxy WinINet (session) applique." "OK" }
        else { bg_Log "Impossible d'ecrire le proxy de session (HKCU)." "WARN" }

        $envNew = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)
        foreach ($n in @("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy")) { $envNew[$n] = "http://$server" }
        & $sink.WriteEnv $envNew
        return $st
    }

    # --- Restauration de la configuration proxy d'origine ---
    function bg_RestoreProxy {
        param($Cfg, $State)
        if ($null -eq $State) { return }
        $sink = bg_ProxySink $Cfg
        & $sink.WriteEnv $State.Env
        [void](& $sink.WriteWinInet $State.WinInet)
        try {
            & $sink.RestoreWinHttp $State
            bg_Log "Configuration proxy d'origine restauree." "OK"
        } catch { bg_Log "Restauration proxy partielle : $($_.Exception.Message)" "WARN" }
    }

    # --- Test CONNECT du proxy vers une destination Microsoft ---
    #     Copie volontaire de Test-ProxyConnect (voir la remarque cote interface).
    function bg_TestProxyConnect {
        param([string]$ProxyHost,[int]$ProxyPort,[string]$TargetHost,[int]$TargetPort,[int]$TimeoutMs = 8000)
        $res = @{ Name = "${TargetHost}:${TargetPort}" ; OK = $false ; Code = 0 ; Detail = "" }
        $tcp = $null
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $iar = $tcp.BeginConnect($ProxyHost, $ProxyPort, $null, $null)
            if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { $res.Detail = "timeout : proxy injoignable" ; return $res }
            $tcp.EndConnect($iar)
            $tcp.ReceiveTimeout = $TimeoutMs
            $s = $tcp.GetStream()
            $sw = New-Object System.IO.StreamWriter($s, [System.Text.Encoding]::ASCII)
            $sw.NewLine = "`r`n"
            $sw.AutoFlush = $true
            $sw.WriteLine("CONNECT ${TargetHost}:${TargetPort} HTTP/1.1")
            $sw.WriteLine("Host: ${TargetHost}:${TargetPort}")
            $sw.WriteLine("")
            $sr = New-Object System.IO.StreamReader($s, [System.Text.Encoding]::ASCII)
            $line = $sr.ReadLine()
            if ($line -match 'HTTP/\d(\.\d)?\s+(\d{3})') {
                $res.Code = [int]$matches[2]
                $res.OK = ($res.Code -ge 200 -and $res.Code -lt 300)
                $res.Detail = $line.Trim()
            } else { $res.Detail = "reponse inattendue : $line" }
        } catch { $res.Detail = $_.Exception.Message }
        finally { if ($tcp) { try { $tcp.Close() } catch {} } }
        return $res
    }

    $ao = $cfg.ActivateOnly

    try {
    bg_Log "Worker demarre (ActivateOnly=$ao, RemoveOld=$($cfg.RemoveOld), DoInstall=$($cfg.DoInstall), DoActivate=$($cfg.DoActivate))" "INFO"
    bg_Log "SetupExePath=$($cfg.SetupExePath)" "INFO"
    bg_Log "XML cible=$($cfg.ActualXMLPath) | XML custom=$($cfg.UseCustomXML)" "INFO"

    # ======================================================
    # ETAPE 1 - Desinstallation
    # ======================================================
    bg_Log "ETAPE 1 - Desinstallation ancienne version" "STEP"
    if (-not $ao -and $cfg.RemoveOld) {
        bg_Step 0 "running" ; $sync.StatusText = "Etape 1/5 - Recherche d'anciennes installations..."
        try {
            bg_Log "Recherche des entrees de desinstallation dans le registre..." "INFO"
            $uninstallEntries = Get-ItemProperty `
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" `
                -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.DisplayName -and (
                        $_.DisplayName -like "Microsoft Office*" -or
                        $_.DisplayName -like "Microsoft 365*" -or
                        $_.DisplayName -like "*Office 1[69]*" -or
                        $_.DisplayName -like "*Office LTSC*"
                    )
                }

            $count = if ($uninstallEntries) { @($uninstallEntries).Count } else { 0 }
            bg_Log "Entrees Office candidates detectees : $count" "INFO"

            if (-not $uninstallEntries) {
                bg_Log "Aucune version Office ciblee pour desinstallation." "INFO"
            } else {
                foreach ($app in $uninstallEntries) {
                    $name = $app.DisplayName
                    $cmd = if (-not [string]::IsNullOrWhiteSpace($app.QuietUninstallString)) { [string]$app.QuietUninstallString } else { [string]$app.UninstallString }

                    if ([string]::IsNullOrWhiteSpace($cmd)) {
                        bg_Log "Pas de commande de desinstallation pour : $name" "WARN"
                        continue
                    }

                    bg_Log "Desinstallation : $name" "INFO"
                    bg_Log "Commande : $($cmd.Substring(0, [Math]::Min($cmd.Length, 180)))" "INFO"
                    $exitCode = bg_InvokeUninstall -CommandLine $cmd
                    if ($exitCode -in @(0, 1641, 3010)) {
                        bg_Log "Desinstallation terminee : $name (code $exitCode)" "OK"
                    } else {
                        bg_Log "Code de sortie desinstallation pour $name : $exitCode" "WARN"
                    }
                }
                # Click-to-Run rend souvent la main avant la fin reelle : attendre la fin des processus
                # evite les echecs d'installation 17002/17029 par coexistence.
                $waitStart = Get-Date
                $lastWaitLog = -30
                while (((Get-Date) - $waitStart).TotalSeconds -lt 300) {
                    $busyProc = @(Get-Process -Name "OfficeClickToRun", "setup" -ErrorAction SilentlyContinue)
                    if ($busyProc.Count -eq 0) { break }
                    $waited = ((Get-Date) - $waitStart).TotalSeconds
                    if (($waited - $lastWaitLog) -ge 20) {
                        $lastWaitLog = $waited
                        bg_Log ("Desinstallation en cours ($($busyProc[0].ProcessName)) - {0:mm}m{0:ss}s..." -f [timespan]::FromSeconds([math]::Floor($waited))) "INFO"
                    }
                    Start-Sleep -Seconds 5
                }
                bg_Log "Etape de desinstallation terminee." "OK"
            }
            bg_Step 0 "done"
        } catch {
            $solDesinst = "Fermez Word, Excel, Outlook et PowerPoint, attendez la fin des processus 'OfficeClickToRun' / 'setup' dans le Gestionnaire des taches, et redemarrez le poste si une mise a jour Office est en attente, puis relancez le deploiement. En dernier recours seulement : Parametres > Applications > Microsoft Office > Desinstaller, ou l'outil SaRA de Microsoft, puis relancez."
            bg_Step 0 "error" ; bg_Log "Echec : $($_.Exception.Message)" "ERROR" ; bg_Log "Solution : $solDesinst" "WARN"
            bg_Popup "Erreur desinstallation" "Echec.`n`nSolution :`n$solDesinst" "Error"
            bg_Abort "Deploiement interrompu." ; return
        }
    } else { bg_Step 0 "skipped" ; bg_Log $(if($ao){"Mode activation seule."}else{"Option non cochee."}) }

    # ======================================================
    # ETAPE 2 - Preparation (verification + XML)
    # ======================================================
    bg_Log "ETAPE 2 - Verification et preparation" "STEP"
    if (-not $ao) {
        bg_Step 1 "running" ; $sync.StatusText = "Etape 2/5 - Verification des prerequis..."

        # Verifier setup.exe (chemin libre)
        # Test-Path refuse une chaine vide : on teste le vide d'abord, sinon l'erreur
        # de liaison interrompt silencieusement le worker sans aucun message.
        if ([string]::IsNullOrWhiteSpace([string]$cfg.SetupExePath) -or -not (Test-Path -LiteralPath ([string]$cfg.SetupExePath))) {
            $solSetup = if ([string]::IsNullOrWhiteSpace([string]$cfg.SetupExePath)) {
                "Aucun chemin memorise : cliquez sur 'Parcourir...' et selectionnez setup.exe (Office Deployment Tool)."
            } else {
                "Ce fichier n'existe plus a cet emplacement : il a pu etre deplace, supprime par l'antivirus, ou le chemin enregistre vient d'un autre poste. Cliquez sur 'Parcourir...' pour le reselectionner, ou reinstallez l'ODT (archive Microsoft, lien dans le README) et extrayez setup.exe a cote de cet outil."
            }
            bg_Step 1 "error" ; bg_Log "setup.exe introuvable : $($cfg.SetupExePath)" "ERROR"
            bg_Log "Solution : $solSetup" "WARN"
            bg_Popup "setup.exe introuvable" "Chemin : $($cfg.SetupExePath)`n`nSolution :`n$solSetup" "Error"
            bg_Abort "Deploiement interrompu." ; return
        }
        bg_Log "setup.exe OK : $($cfg.SetupExePath)" "OK"

        # Verifier espace disque
        $freeGB = [math]::Round((Get-PSDrive C -ErrorAction SilentlyContinue).Free / 1GB, 1)
        if ($freeGB -lt $cfg.DiskMinGB) {
            bg_Step 1 "error" ; bg_Log "Espace insuffisant : $freeGB Go (min $($cfg.DiskMinGB) Go)" "ERROR"
            bg_Popup "Espace disque insuffisant" "Espace libre C:\ : $freeGB Go`nRequis : $($cfg.DiskMinGB) Go`n`nSolution : Liberez de l'espace sur C:\" "Error"
            bg_Abort "Deploiement interrompu." ; return
        }
        bg_Log "Espace disque OK : $freeGB Go libres." "OK"

        # Creer le dossier si absent
        if (!(Test-Path $cfg.InstallFolder)) { New-Item -ItemType Directory -Path $cfg.InstallFolder -Force | Out-Null }

        # === XML : genere ou fichier existant ===
        try {
            if ($cfg.UseCustomXML) {
                # Verifier le fichier XML fourni
                if ([string]::IsNullOrWhiteSpace([string]$cfg.CustomXMLPath) -or -not (Test-Path -LiteralPath ([string]$cfg.CustomXMLPath))) { throw "Fichier XML introuvable : $($cfg.CustomXMLPath)" }
                $cfg.ActualXMLPath = $cfg.CustomXMLPath
                bg_Log "XML personnalise utilise : $(Split-Path $cfg.CustomXMLPath -Leaf)" "OK"
            } else {
                # Generer le XML avec les exclusions d'applications
                # Le XML et les fichiers Office sont ecrits LOCALEMENT (ODT ne supporte pas les chemins UNC)
                # Emplacement des sources : telechargement local, ou depot fourni
                $cfg.SourcePath = $cfg.LocalDataPath
                $cfg.DepotCopy  = $false
                if ($cfg.DepotEnabled) {
                    $cfg.DepotCopy = -not (bg_TestDepotLocal ([string]$cfg.DepotPath))
                    if (-not $cfg.DepotCopy) { $cfg.SourcePath = [string]$cfg.DepotPath }
                    bg_Log "Mode depot : sources = $($cfg.SourcePath) (copie locale = $($cfg.DepotCopy))" "INFO"
                }

                $localXmlDir = $cfg.LocalDataPath
                if (!(Test-Path $localXmlDir)) {
                    New-Item -ItemType Directory -Path $localXmlDir -Force | Out-Null
                }
                $localXmlPath = Join-Path $localXmlDir "OfficeConfig2024.xml"
                $excl  = $cfg.ExcludedApps
                $lines = $excl | ForEach-Object { "      <ExcludeApp ID=`"$_`" />" }
                $exclBlock = if ($lines) { "`n" + ($lines -join "`n") } else { "" }
                $xml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024" SourcePath="$($cfg.SourcePath)">
    <Product ID="ProPlus2024Volume">
      <Language ID="$($cfg.LanguageID)" />$exclBlock
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
</Configuration>
"@
                # Ecriture sans BOM (UTF-8 pur) car ODT peut rejeter le BOM
                [System.IO.File]::WriteAllText($localXmlPath, $xml, [System.Text.UTF8Encoding]::new($false))
                $cfg.ActualXMLPath = $localXmlPath
                $exclMsg = if ($excl.Count -gt 0) { "Exclusions : " + ($excl -join ", ") } else { "Toutes les apps incluses" }
                bg_Log "XML genere ($($cfg.LanguageID)) - $exclMsg" "OK"
                bg_Log "Repertoire de telechargement local : $localXmlDir" "INFO"
            }
            bg_Step 1 "done"
        } catch {
            bg_Step 1 "error" ; bg_Log "Echec preparation XML : $($_.Exception.Message)" "ERROR"
            bg_Log "Solution : Verifiez le chemin du fichier XML ou les droits d'ecriture." "WARN"
            bg_Popup "Erreur preparation XML" "$($_.Exception.Message)`n`nSolution : Verifiez le chemin et les droits d'acces." "Error"
            bg_Abort "Deploiement interrompu." ; return
        }
    } else { bg_Step 1 "skipped" ; bg_Log "Mode activation seule." }

    # ======================================================
    # ETAPE 3 - Telechargement
    # ======================================================
    bg_Log "ETAPE 3 - Acquisition des sources Office 2024" "STEP"
    if (-not $ao) {
        bg_Step 2 "running" ; $sync.StatusText = "Etape 3/5 - Acquisition des sources..."

        # ---------- MODE DEPOT : aucun acces Internet necessaire ----------
        if ($cfg.DepotEnabled) {
            try {
                $depot = [string]$cfg.DepotPath
                bg_Log "Mode depot : $depot (aucun telechargement)" "INFO"
                if (-not (Test-Path $depot)) { throw "Depot de sources introuvable : $depot" }

                $cfg.DepotCopy  = -not (bg_TestDepotLocal $depot)
                $cfg.SourcePath = if ($cfg.DepotCopy) { [string]$cfg.LocalDataPath } else { $depot }
                if ($cfg.UseCustomXML) {
                    bg_Log "XML personnalise : SourcePath sera impose/verifie avant l'installation ($($cfg.SourcePath))." "INFO"
                }

                if ($cfg.DepotCopy) {
                    $dest = [string]$cfg.LocalDataPath
                    bg_Log "Copie du depot vers $dest (le service Click-to-Run n'accede pas aux partages sous LocalSystem)..." "INFO"
                    if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Path $dest -Force | Out-Null }
                    $rcArgs = @("`"$depot`"", "`"$dest`"", "/MIR", "/NFL", "/NDL", "/NJH", "/NJS", "/NP", "/R:2", "/W:2")
                    $rc = Start-Process -FilePath "robocopy.exe" -ArgumentList $rcArgs -Wait -PassThru -WindowStyle Hidden
                    if ($rc.ExitCode -ge 8) { throw "Copie impossible depuis le depot (robocopy code $($rc.ExitCode)). Verifiez les droits sur $depot." }
                    bg_Log "Copie du depot terminee." "OK"
                } else {
                    bg_Log "Depot sur disque local : utilisation directe, sans copie." "INFO"
                }

                $srcData = Join-Path ([string]$cfg.SourcePath) "Office\Data"
                if (-not (Test-Path $srcData)) {
                    throw "Sources Office introuvables : $srcData`nLe depot doit contenir le dossier 'Office' produit par 'setup.exe /download'."
                }
                bg_Log "Sources validees : $srcData" "OK"
                bg_Step 2 "done"
            } catch {
                bg_Step 2 "error" ; bg_Log "Echec depot : $($_.Exception.Message)" "ERROR"
                bg_Log "Solution : corriger le chemin du depot, ou decocher 'Depot de sources' pour telecharger." "WARN"
                bg_Popup "Depot de sources" "$($_.Exception.Message)`n`nSolution : le depot doit contenir le dossier 'Office\Data'." "Error"
                bg_Abort "Deploiement interrompu." ; return
            }
        } else {
            # ---------- MODE TELECHARGEMENT : proxy applique au niveau SYSTEME ----------
            bg_Step 2 "running" ; $sync.StatusText = "Etape 3/5 - Telechargement (10 a 20 min)..."
            bg_Log "Connexion aux serveurs Microsoft - patience, 10 a 20 minutes." "INFO"
            bg_Log "Vous pouvez minimiser cette fenetre - une notification vous alertera." "INFO"
            # La sauvegarde/restauration des variables d'environnement est faite par
            # bg_ApplyProxy / bg_RestoreProxy, comme celle de WinHTTP et WinINet.
            $proxyState = $null
            try {
                if ($cfg.ProxyEnabled) {
                    $proxyHost = $cfg.ProxyHost
                    $proxyPort = $cfg.ProxyPort
                    $proxyBypass = [string]$cfg.ProxyBypass
                    if ([string]::IsNullOrWhiteSpace($proxyBypass)) { $proxyBypass = "<local>" }
                    $proxyPortN = 0
                    if (-not [int]::TryParse([string]$proxyPort, [ref]$proxyPortN) -or $proxyPortN -le 0) { throw "Port proxy invalide : $proxyPort" }

                    # 1) Verifier le proxy AVANT de lancer ~20 minutes de telechargement
                    $checks  = @()
                    $checks += bg_TestProxyConnect -ProxyHost $proxyHost -ProxyPort $proxyPortN -TargetHost "officecdn.microsoft.com" -TargetPort 443
                    $checks += bg_TestProxyConnect -ProxyHost $proxyHost -ProxyPort $proxyPortN -TargetHost "config.office.com" -TargetPort 443
                    $okCount = 0 ; $authReq = $false
                    foreach ($c in $checks) {
                        if ($c.OK) { $okCount++ ; bg_Log "Proxy : $($c.Name) joignable ($($c.Detail))" "OK" }
                        else {
                            if ($c.Code -eq 407) { $authReq = $true }
                            bg_Log "Proxy : $($c.Name) ECHEC - $($c.Detail)" "ERROR"
                        }
                    }
                    if ($okCount -eq 0) {
                        if ($authReq) {
                            if ($cfg.ProxyRelay) {
                                throw "Le proxy $($cfg.ProxyRelayTarget) refuse les identifiants fournis (407), malgre le relais local. Verifiez l'utilisateur et le mot de passe du proxy dans Options : le mot de passe n'est jamais enregistre, il faut le ressaisir a chaque lancement. Si le proxy exige NTLM ou Kerberos (et non Basic), le relais ne suffit pas -> utilisez le mode 'Depot de sources'."
                            }
                            throw "Le proxy ${proxyHost}:${proxyPort} exige une authentification (407). Le service Click-to-Run tourne en LocalSystem et ne peut pas s'authentifier -> utilisez le mode 'Depot de sources' (telechargement une fois, puis installation sur le reseau pedagogique)."
                        }
                        throw "Aucun acces au CDN Microsoft via le proxy ${proxyHost}:${proxyPort}. Verifiez l'adresse, le port, les exceptions, et les regles de filtrage (destination officecdn.microsoft.com:443)."
                    }

                    # 2) Appliquer le proxy a WINDOWS (WinHTTP machine + WinINet session), puis restaurer
                    $proxyState = bg_ApplyProxy $cfg -ProxyHost $proxyHost -ProxyPort $proxyPort -BypassList $proxyBypass
                    bg_Log "Proxy systeme applique : WinHTTP (machine) + WinINet (session). Les variables HTTP_PROXY seules ne sont pas lues par Office." "INFO"
                } else {
                    bg_Log "Aucun proxy : utilisation de la configuration systeme existante." "INFO"
                    # Verification informative : sans proxy, le filtrage reseau peut bloquer le CDN.
                    $cdnOk = $false
                    $tcpProbe = $null
                    try {
                        $tcpProbe = New-Object System.Net.Sockets.TcpClient
                        $iarProbe = $tcpProbe.BeginConnect("officecdn.microsoft.com", 443, $null, $null)
                        if ($iarProbe.AsyncWaitHandle.WaitOne(5000, $false)) { $tcpProbe.EndConnect($iarProbe) ; $cdnOk = $tcpProbe.Connected }
                    } catch {} finally { if ($tcpProbe) { try { $tcpProbe.Close() } catch {} } }
                    if ($cdnOk) { bg_Log "CDN Microsoft joignable directement (officecdn.microsoft.com:443)." "OK" }
                    else { bg_Log "CDN Microsoft NON joignable en direct : le telechargement va echouer. Renseignez le proxy, ou utilisez le mode 'Depot de sources'." "WARN" }
                }

                $setupDir      = Split-Path $cfg.SetupExePath -Parent
                $downloadStart = Get-Date

                # Sans SourcePath, ODT ecrit les sources dans le repertoire courant : la source
                # reelle (imposee ou deja presente dans le XML) sert de destination et de suivi.
                $srcInfo   = bg_PrepareXmlSource $cfg "le telechargement"
                $officeDir = Join-Path $srcInfo.Source "Office"
                bg_Log "Destination des sources : $officeDir" "INFO"

                # Espace disque a CET emplacement (pas seulement sur C:)
                # Sources Office : ~3 Go en pratique, 4 Go avec marge.
                $srcNeedGB = 4
                $srcFreeGB   = $null
                $srcRootPath = $null
                try {
                    $srcRootPath = [System.IO.Path]::GetPathRoot($srcInfo.Source)
                    if (-not [string]::IsNullOrWhiteSpace($srcRootPath)) {
                        $srcDrive  = New-Object System.IO.DriveInfo($srcRootPath)
                        $srcFreeGB = [math]::Round($srcDrive.AvailableFreeSpace / 1GB, 1)
                    }
                } catch {}
                $spaceCheck = bg_SpaceVerdict -FreeGB $srcFreeGB -NeedGB $srcNeedGB -Root $srcRootPath
                bg_Log $spaceCheck.Msg $spaceCheck.Level

                # Lancer setup.exe sans -Wait pour pouvoir surveiller la progression
                $proc = Start-Process -FilePath $cfg.SetupExePath `
                    -ArgumentList "/download `"$($cfg.ActualXMLPath)`"" `
                    -PassThru -WindowStyle Hidden -WorkingDirectory $setupDir

                # Boucle de surveillance : taille du dossier Office\ + vitesse
                $lastSizeMB   = 0
                $lastCheck    = Get-Date
                $reportedStart = $false
                $prevLogSec   = -30   # force un log immediat

                while (-not $proc.HasExited) {
                    Start-Sleep -Seconds 8

                    $nowSec = ((Get-Date) - $downloadStart).TotalSeconds

                    # Calcul taille telechargee
                    $sizeMB = 0
                    if (Test-Path $officeDir) {
                        try {
                            $sizeMB = [math]::Round(
                                (Get-ChildItem $officeDir -Recurse -File -ErrorAction SilentlyContinue |
                                 Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum / 1MB, 1)
                        } catch {}
                    }

                    # Vitesse en Ko/s sur les 4 dernieres secondes
                    $elapsed4 = ((Get-Date) - $lastCheck).TotalSeconds
                    $speedKBs = if ($elapsed4 -gt 0) { [math]::Round(($sizeMB - $lastSizeMB) * 1024 / $elapsed4, 0) } else { 0 }
                    $lastSizeMB = $sizeMB
                    $lastCheck  = Get-Date

                    # Log toutes les 15 secondes ou des le premier octet detecte
                    $doLog = ($nowSec - $prevLogSec) -ge 15
                    if ($sizeMB -gt 0 -and -not $reportedStart) { $doLog = $true ; $reportedStart = $true }

                    if ($doLog) {
                        $prevLogSec = $nowSec
                        $elapsed = [timespan]::FromSeconds([math]::Floor($nowSec))
                        if ($sizeMB -gt 0) {
                            $speedTxt = if ($speedKBs -gt 0) { "  |  ~$speedKBs Ko/s" } else { "" }
                            bg_Log ("Telechargement en cours : $sizeMB Mo recu  |  {0:mm}m{0:ss}s ecoule$speedTxt" -f $elapsed) "INFO"
                        } else {
                            bg_Log ("Connexion aux serveurs Microsoft en cours... ({0:mm}m{0:ss}s)" -f $elapsed) "INFO"
                        }
                    }

                    # Mettre a jour le statut visible
                    $sync.StatusText = if ($sizeMB -gt 0) {
                        "Telechargement : $sizeMB Mo $(if ($speedKBs -gt 0){"($speedKBs Ko/s)"})"
                    } else {
                        "Connexion aux serveurs Microsoft..."
                    }
                }

                $exitCode = $proc.ExitCode
                $totalMB  = 0
                if (Test-Path $officeDir) {
                    try {
                        $totalMB = [math]::Round(
                            (Get-ChildItem $officeDir -Recurse -File -ErrorAction SilentlyContinue |
                             Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum / 1MB, 0)
                    } catch {}
                }
                $duration = [math]::Round(((Get-Date) - $downloadStart).TotalSeconds, 0)
                if ($exitCode -ne 0) { throw "setup.exe code de sortie : $exitCode" }
                $downloadedData = Join-Path $officeDir "Data"
                if (-not (Test-Path $downloadedData)) {
                    throw "Telechargement incomplet : '$downloadedData' est absent (setup.exe a rendu 0 sans produire les sources)."
                }
                bg_Log "Sources telechargees avec succes. ($totalMB Mo en ${duration}s)" "OK"
                bg_Step 2 "done"
            } catch {
                bg_Step 2 "error" ; bg_Log "Echec : $($_.Exception.Message)" "ERROR"
                # Lire le log ODT pour avoir la vraie raison
                $odtLogDetail = ""
                try {
                    $odtLogDir = "$env:TEMP\Microsoft Office"
                    $odtLogFile = Get-ChildItem $odtLogDir -Filter "*.log" -ErrorAction SilentlyContinue |
                        Where-Object { $_.LastWriteTime -gt $downloadStart } |
                        Sort-Object LastWriteTime -Descending | Select-Object -First 1
                    if ($odtLogFile) {
                        bg_Log "Log ODT : $($odtLogFile.Name)" "INFO"
                        $logLines = Get-Content $odtLogFile.FullName -Tail 20 -ErrorAction SilentlyContinue
                        $errorLines = $logLines | Where-Object { $_ -match '(?i)error|fail|0x[89A-F]' }
                        if ($errorLines) {
                            $errorLines | Select-Object -Last 5 | ForEach-Object { bg_Log "  ODT: $_" "ERROR" }
                            $odtLogDetail = "`n`nDetail ODT :`n" + (($errorLines | Select-Object -Last 3) -join "`n")
                        } else {
                            bg_Log "Aucune erreur detaillee trouvee dans le log ODT." "INFO"
                        }
                    }
                } catch {}
                # Message contextuel selon le code de sortie
                $ec = if ($_.Exception.Message -match 'code de sortie\s*:\s*(\d+)') { [int]$matches[1] } else { -1 }
                $sol = switch ($ec) {
                    32    { "Code 32 = fichier verrouille. Causes : antivirus qui scanne le dossier C:\InstallOffice2024, service Click-to-Run actif, ou droits insuffisants. Essayez : 1) Desactiver Defender temps reel brievement  2) Relancer en Administrateur  3) Redemarrer le poste" }
                    17002 { "Code 17002 = installation incomplete. Redemarrez le poste et reessayez." }
                    30066 { "Code 30066 = fichiers introuvables. Verifiez le chemin setup.exe." }
                    default {
                        if ($cfg.ProxyEnabled -and $cfg.ProxyRelay) {
                            "Le proxy $($cfg.ProxyRelayTarget) n'a pas accepte la connexion malgre le relais local (identifiants refuses ou proxy indisponible). Verifiez l'utilisateur et le mot de passe du proxy dans Options : le mot de passe n'est jamais enregistre. Si le proxy exige NTLM ou Kerberos, le relais (Basic) ne suffit pas : utilisez le mode 'Depot de sources'."
                        } elseif ($cfg.ProxyEnabled) {
                            "Aucun acces au CDN Microsoft via le proxy $($cfg.ProxyHost):$($cfg.ProxyPort). Verifiez dans Options : adresse et port du proxy, et les exceptions (officecdn.microsoft.com et config.office.com doivent etre autorises en 443). Si le proxy exige une authentification (erreur 407), le telechargement est impossible : le service Click-to-Run tourne en LocalSystem et ne peut pas s'authentifier. Utilisez le mode 'Depot de sources' : telecharger les sources une fois depuis un poste autorise, puis cocher ce mode ici pour installer sans aucun acces Internet."
                        } else {
                            "Aucun acces au CDN Microsoft depuis ce poste. Verifiez le filtrage reseau, ou renseignez le proxy dans Options. Alternative hors ligne : mode 'Depot de sources' (sources telechargees une fois, puis installation et activation KMS sans Internet)."
                        }
                    }
                }
                bg_Log "Solution : $sol" "WARN"
                bg_Popup "Erreur telechargement" "$($_.Exception.Message)$odtLogDetail`n`nSolution :`n$sol" "Error"
                bg_Abort "Deploiement interrompu." ; return
            } finally {
                bg_RestoreProxy $cfg -State $proxyState
            }
        }
    } else { bg_Step 2 "skipped" ; bg_Log "Mode activation seule." }

    # ======================================================
    # ETAPE 4 - Installation
    # ======================================================
    bg_Log "ETAPE 4 - Installation Microsoft Office 2024" "STEP"
    if (-not $ao -and $cfg.DoInstall) {
        bg_Step 3 "running" ; $sync.StatusText = "Etape 4/5 - Installation (5 a 15 min)..."
        bg_Log "Installation silencieuse. Ne pas fermer ce programme." "INFO"

        # Verifier si une version Office est deja installee (ODT refuse d'installer en coexistence)
        $existingForInstall = Get-ItemProperty `
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" `
            -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and ($_.DisplayName -like "Microsoft Office*" -or $_.DisplayName -like "Microsoft 365*") } |
            Select-Object -First 1
        if ($existingForInstall) {
            bg_Log "ATTENTION : '$($existingForInstall.DisplayName)' detecte sur ce poste." "WARN"
            bg_Log "ODT peut refuser l'installation. Si echec, relancez avec la case 'Desinstaller ancienne version' cochee." "WARN"
        }

        try {
            # Sans SourcePath, /configure retombe sur le CDN : l'installation hors ligne
            # (mode depot) echouerait. On impose la source reellement disponible.
            $null = bg_PrepareXmlSource $cfg "l'installation"

            $setupDir2 = Split-Path $cfg.SetupExePath -Parent
            $proc = Start-Process -FilePath $cfg.SetupExePath -ArgumentList "/configure `"$($cfg.ActualXMLPath)`"" -Wait -PassThru -WindowStyle Hidden -WorkingDirectory $setupDir2
            $installCode = $proc.ExitCode
            if ($installCode -ne 0) { throw "setup.exe code de sortie : $installCode" }
            bg_Log "Office 2024 installe avec succes dans Program Files." "OK" ; bg_Step 3 "done"
            # Nettoyage des fichiers sources temporaires (ne pas garder ~3 Go inutiles)
            $officeDataDir = Join-Path $cfg.LocalDataPath "Office"
            if ($cfg.DepotEnabled -and -not $cfg.DepotCopy) {
                bg_Log "Nettoyage ignore : les sources sont le depot lui-meme ($officeDataDir)." "INFO"
            } elseif (Test-Path $officeDataDir) {
                bg_Log "Nettoyage des fichiers sources : $officeDataDir" "INFO"
                try {
                    Remove-Item $officeDataDir -Recurse -Force -ErrorAction Stop
                    bg_Log "Fichiers sources supprimes (espace libere)." "OK"
                } catch {
                    bg_Log "Nettoyage non bloquant : $($_.Exception.Message)" "WARN"
                }
            }
        } catch {
            bg_Step 3 "error" ; bg_Log "Echec : $($_.Exception.Message)" "ERROR"
            $ec4 = if ($_.Exception.Message -match 'code de sortie\s*:\s*(-?\d+)') { [int]$matches[1] } else { -999 }
            $sol4 = switch ($ec4) {
                -999   { "setup.exe n'a pas pu demarrer : le fichier est peut-etre bloque (clic droit > Proprietes > Debloquer) ou mis en quarantaine par l'antivirus. Verifiez le chemin setup.exe puis relancez." }
                32     { "Fichier verrouille pendant l'installation : fermez Word/Excel/Outlook, empechez l'antivirus de scanner le dossier des sources (C:\InstallOffice2024), ou redemarrez le poste, puis relancez." }
                30066  { "Une autre installation est en cours (Windows Installer ou Click-to-Run occupe). Fermez les applications Office, attendez la fin des processus 'OfficeClickToRun' / 'setup' (redemarrez si besoin), puis relancez." }
                17029  { "Une version d'Office est deja installee : cochez 'Desinstaller l'ancienne version d'Office' puis relancez le deploiement." }
                17002  { "Installation inachevee, souvent une coexistence avec une installation precedente : redemarrez le poste, verifiez qu'aucune application Office ne tourne, puis relancez." }
                default { "Causes usuelles : desinstallation precedente incomplete (cochez 'Desinstaller l'ancienne version'), antivirus qui verrouille les fichiers sources, application Office restee ouverte, ou redemarrage en attente. Le journal ODT detaille est ecrit dans %TEMP%. L'espace disque a deja ete verifie a l'etape 2." }
            }
            bg_Log "Solution : $sol4" "WARN"
            bg_Popup "Erreur installation" "$($_.Exception.Message)`n`nSolution :`n$sol4" "Error"
            bg_Abort "Deploiement interrompu." ; return
        }
    } else {
        bg_Step 3 "skipped"
        bg_Log $(if($ao){"Activation seule."}elseif(-not $cfg.DoInstall){"Installation differee."}else{"Ignoree."})
    }

    # ======================================================
    # ETAPE 5 - Activation KMS (host:port gere)
    # ======================================================
    bg_Log "ETAPE 5 - Activation KMS" "STEP"
    $doKms = $ao -or ($cfg.DoInstall -and $cfg.DoActivate)
    if ($doKms) {
        bg_Step 4 "running" ; $sync.StatusText = "Etape 5/5 - Activation KMS..."

        if ([string]::IsNullOrWhiteSpace($cfg.KmsHost) -or $cfg.KmsHost -eq "KMS.VOTRE-SERVEUR.FR") {
            bg_Step 4 "error" ; bg_Log "Adresse KMS non renseignee." "ERROR"
            bg_Log "Solution : Renseignez l'adresse KMS dans le champ (format : host ou host:port)." "WARN"
            bg_Popup "KMS - Adresse manquante" "Office installe mais PAS active.`n`nSolution : Renseignez l'adresse KMS (ex: kms.grandest.fr:1688) puis relancez 'Activer seulement'." "Warning"
        } else {
            $kmsCandidates = @("$($cfg.KmsHost):$($cfg.KmsPort)")
            if (-not [string]::IsNullOrWhiteSpace($cfg.KmsSecondary)) {
                $kmsCandidates += $cfg.KmsSecondary
            }
            $kmsCandidates = $kmsCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
            bg_Log "Serveur(s) KMS : $($kmsCandidates -join ' | ')" "INFO"

            $OsppPath = @("C:\Program Files\Microsoft Office\Office16\ospp.vbs","C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs") | Where-Object { Test-Path $_ } | Select-Object -First 1
            if (-not $OsppPath) {
                bg_Step 4 "error" ; bg_Log "ospp.vbs introuvable - installation incomplete ?" "ERROR"
                bg_Popup "Fichier activation manquant" "ospp.vbs introuvable.`n`nSolution : Relancez le deploiement complet." "Error"
            } else {
                bg_Log "ospp.vbs : $OsppPath" "OK"
                try {
                    if ($cfg.ApplyProductKey -and -not [string]::IsNullOrWhiteSpace($cfg.ProductKey)) {
                        $normalizedPk = bg_NormalizeProductKey $cfg.ProductKey
                        bg_Log "Application de la cle produit avant activation..." "INFO"
                        if ([string]::IsNullOrWhiteSpace($normalizedPk)) {
                            throw "Cle produit vide apres normalisation."
                        }
                        if ($normalizedPk -notmatch '^([A-Z0-9]{5}-){4}[A-Z0-9]{5}$') {
                            bg_Log "Cle produit en format non standard : tentative /inpkey tout de meme." "WARN"
                        }
                        $pkOut = (cscript //nologo "$OsppPath" /inpkey:$normalizedPk 2>&1) -join " "
                        if ($pkOut -notmatch "successful|installed product key") {
                            throw "Echec /inpkey : $pkOut"
                        }
                        bg_Log "Cle produit appliquee." "OK"
                    }

                    $activated = $false
                    $lastErr = ""
                    foreach ($kmsRaw in $kmsCandidates) {
                        $srv = bg_ConvertToKmsAddress $kmsRaw
                        bg_Log "Tentative KMS : $($srv.Host):$($srv.Port)" "INFO"
                        cscript //nologo "$OsppPath" /sethst:$($srv.Host)  2>&1 | Out-Null
                        cscript //nologo "$OsppPath" /setprt:$($srv.Port)  2>&1 | Out-Null
                        $out = (cscript //nologo "$OsppPath" /act 2>&1) -join " "
                        if ($out -match "successful") {
                            $activated = $true
                            bg_Log "Activation KMS reussie sur $($srv.Host):$($srv.Port) !" "OK"
                            break
                        }
                        if     ($out -match "0xC004F074") { $lastErr = "Serveur KMS injoignable (0xC004F074)" }
                        elseif ($out -match "0x80070005") { $lastErr = "Acces refuse (0x80070005)" }
                        elseif ($out -match "0xC004F038") { $lastErr = "Seuil KMS insuffisant (0xC004F038)" }
                        else                              { $lastErr = "Reponse inattendue : $out" }
                        bg_Log "Echec tentative KMS sur $($srv.Host):$($srv.Port) - $lastErr" "WARN"
                    }

                    if ($activated) { bg_Step 4 "done" }
                    else { throw $lastErr }
                } catch {
                    bg_Step 4 "error"
                    $err = $_.Exception.Message
                    $sol = switch -Wildcard ($err) {
                        "*injoignable*" { "Le serveur KMS ne repond pas : verifiez que le poste est sur le reseau du lycee (cable ou VPN), que l'adresse KMS et le port $($cfg.KmsPort) sont corrects (resolution du nom avec 'nslookup'), et que le pare-feu autorise le port $($cfg.KmsPort) en sortie vers ce serveur." }
                        "*Acces*"       { "Acces refuse : relancez l'outil en Administrateur, verifiez que le service 'Office Software Protection Platform' (osppsvc) est demarre, et que l'antivirus ne bloque pas cscript/ospp.vbs." }
                        "*insuffisant*" { "Le serveur KMS n'a pas encore atteint son seuil d'activation (25 clients Windows / 5 clients Office vus). Ce n'est pas un probleme de ce poste : reessayez plus tard ou contactez l'IT Grand Est." }
                        default         { "Erreur d'activation non identifiee : notez le code affiche, verifiez l'adresse KMS dans Options, puis contactez l'IT Grand Est avec ce message." }
                    }
                    bg_Log "Echec KMS : $err" "ERROR" ; bg_Log "Solution : $sol" "WARN"
                    bg_Popup "Erreur KMS" "Echec.`nDetail : $err`n`nSolution : $sol" "Error"
                }
            }
        }
    } else { bg_Step 4 "skipped" ; bg_Log "Activation ignoree (option non cochee)." }

    # ======================================================
    # BILAN STRUCTURE
    # ======================================================
    $elapsed = if ($sync.StartTime) { (Get-Date) - $sync.StartTime } else { [timespan]::Zero }
    $sym  = @{ "done"="[OK] "; "skipped"="[ - ]"; "error"="[ERR]"; "waiting"="[ ? ]"; "running"="[...]" }
    $nms  = @("Desinstallation "; "Preparation    "; "Telechargement  "; "Installation   "; "Activation KMS ")
    $col  = @{ "done"="OK"; "skipped"="INFO"; "error"="ERROR"; "waiting"="INFO"; "running"="INFO" }

    bg_Log "================================================" "STEP"
    bg_Log "  BILAN DU DEPLOIEMENT" "STEP"
    bg_Log "================================================" "STEP"
    bg_Log "  Duree totale  : $("{0:mm}m {0:ss}s" -f $elapsed)" "INFO"
    bg_Log "  PC            : $env:COMPUTERNAME" "INFO"
    bg_Log "  Utilisateur   : $env:USERNAME" "INFO"
    bg_Log "  Date          : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" "INFO"
    bg_Log "  ------------------------------------------------" "STEP"
    for ($i = 0; $i -lt 5; $i++) { $st = $sync.StepState[$i] ; bg_Log "  $($sym[$st])  $($nms[$i])" $col[$st] }
    bg_Log "================================================" "STEP"

    $sync.ActiveStep = -1
    # StepState est une hashtable : '-contains' testerait les CLES, pas les valeurs.
    # Sans ce test sur les valeurs, un echec KMS finissait sur un message de succes.
    $stepValues = if ($sync.StepState -is [System.Collections.IDictionary]) { @($sync.StepState.Values) } else { @($sync.StepState) }
    $sync.HasError = ($stepValues -contains "error")
    $sync.FinalMsg = if (-not $sync.HasError) {
        "Microsoft Office 2024 deploye avec succes !"
    } elseif ($sync.StepState[4] -eq "error") {
        if ($sync.StepState[3] -eq "done") {
            "Office 2024 est installe, mais l'activation KMS a echoue.`nConsultez le journal, puis relancez 'Activer seulement' une fois le serveur KMS joignable."
        } else {
            "Activation KMS echouee.`nConsultez le journal pour la cause."
        }
    } else {
        "Deploiement termine avec des erreurs.`nConsultez le journal."
    }
    $sync.StatusText = if ($sync.HasError) { "Termine avec erreurs - voir journal." } else { "Succes !" }
    $sync.Done       = $true
    } catch {
        $errMsg = $_.Exception.Message
        $sync.ActiveStep = -1
        $sync.HasError   = $true
        $sync.FinalMsg   = "Erreur inattendue pendant le deploiement : $errMsg"
        $sync.StatusText = "Erreur inattendue - voir journal."
        $logQueue.Enqueue(@{Text="[$(Get-Date -Format 'HH:mm:ss')] [ERR]  Erreur fatale worker : $errMsg";Color=[System.Drawing.Color]::FromArgb(240,80,80)})
        $sync.Done       = $true
    }
}

# ================================================================
#  LANCEMENT
# ================================================================
function Start-Deployment {
    param([bool]$ActivateOnly = $false)

    if ($script:deployWorkerContext -and -not $script:deployWorkerContext.Handle.IsCompleted) {
        [System.Windows.Forms.MessageBox]::Show("Un deploiement est deja en cours.", "Actif", "OK", "Warning") | Out-Null ; return
    }

    # Valider le fichier XML personnalise si selectionne
    if ($radioXmlFile.Checked -and [string]::IsNullOrWhiteSpace($xmlFilePathBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Vous avez selectionne 'Utiliser un fichier XML' mais aucun fichier n'est renseigne.`n`nCliquez sur 'Parcourir...' pour selectionner votre fichier XML.", "Fichier XML manquant", "OK", "Warning") | Out-Null ; return
    }

    if ($chkUseProxy.Checked -and -not $chkUseDepot.Checked -and [string]::IsNullOrWhiteSpace($proxyHostBox.Text.Trim())) {
        [System.Windows.Forms.MessageBox]::Show("Proxy active mais adresse vide.`n`nRenseignez l'adresse du proxy (ex: proxy.grandest.fr).", "Proxy incomplet", "OK", "Warning") | Out-Null ; return
    }

    if ($chkUseProxy.Checked -and -not $chkUseDepot.Checked) {
        $portText = $proxyPortBox.Text.Trim()
        $portNum = 0
        if (-not [int]::TryParse($portText, [ref]$portNum) -or $portNum -le 0 -or $portNum -gt 65535) {
            [System.Windows.Forms.MessageBox]::Show("Port proxy invalide.`n`nValeurs autorisees : 1 a 65535.", "Proxy incomplet", "OK", "Warning") | Out-Null ; return
        }
    }

    if ($chkUseDepot.Checked) {
        $depotCheck = $depotPathBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($depotCheck) -or -not (Test-Path $depotCheck)) {
            [System.Windows.Forms.MessageBox]::Show("Depot de sources introuvable :`n$depotCheck`n`nIndiquez le dossier contenant 'Office\Data' (resultat de 'setup.exe /download').", "Depot de sources", "OK", "Warning") | Out-Null ; return
        }
        if (-not (Test-Path (Join-Path $depotCheck "Office\Data"))) {
            [System.Windows.Forms.MessageBox]::Show("Le dossier choisi ne contient pas 'Office\Data'.`n`n$depotCheck`n`nVerifiez qu'il s'agit bien du dossier de telechargement ODT.", "Depot de sources", "OK", "Warning") | Out-Null ; return
        }
    }

    if ($chkActivate.Checked -and ([string]::IsNullOrWhiteSpace($kmsTextBox.Text.Trim()) -or $kmsTextBox.Text.Trim() -eq $KMS_SERVEUR)) {
        [System.Windows.Forms.MessageBox]::Show("L'activation KMS est cochee mais l'adresse du serveur n'est pas renseignee.`n`nRenseignez-la (ex: kms.grandest.fr:1688), decochez 'Activer via KMS', ou utilisez 'Activer seulement' plus tard.", "Adresse KMS manquante", "OK", "Warning") | Out-Null ; return
    }

    $normalizedProductKey = ""
    if ($chkApplyProductKey.Checked) {
        $normalizedProductKey = Normalize-ProductKey $productKeyBox.Text
        if ([string]::IsNullOrWhiteSpace($normalizedProductKey)) {
            [System.Windows.Forms.MessageBox]::Show("Option '/inpkey' activee mais la cle produit est vide.", "Cle produit manquante", "OK", "Warning") | Out-Null ; return
        }

        # Appliquer la version normalisee dans l'UI pour eviter les erreurs de saisie (espaces/tirets exotiques).
        $productKeyBox.Text = $normalizedProductKey

        if ($normalizedProductKey -notmatch '^([A-Z0-9]{5}-){4}[A-Z0-9]{5}$') {
            $r = [System.Windows.Forms.MessageBox]::Show(
                "Format de cle non standard detecte.`n`nLe programme tentera quand meme '/inpkey'.`nContinuer ?",
                "Cle produit - format non standard", "YesNo", [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($r -ne "Yes") { return }
        }
    }

    # Parser l'adresse KMS (host:port ou host)
    $kmsAddr = ConvertTo-KmsAddress $kmsTextBox.Text.Trim()
    $kmsSecondaryRaw = $kmsSecondaryBox.Text.Trim()

    # Synchroniser l'etat XML script/UI avant confirmation et lancement.
    $script:xmlUseCustomFile = [bool]$radioXmlFile.Checked
    $script:xmlCustomFilePath = $xmlFilePathBox.Text.Trim()

    # Sauvegarder la config
    Save-Config @{
        SetupExePath  = $setupPathBox.Text.Trim()
        KmsServer     = $kmsTextBox.Text.Trim()
        Language      = $langCombo.SelectedItem
        RemoveOld     = $chkRemoveOld.Checked
        AutoInstall   = $chkInstall.Checked
        AutoActivate  = $chkActivate.Checked
        ExcludedApps  = $script:xmlExcludedApps
        UseCustomXML  = $radioXmlFile.Checked
        CustomXMLPath = $xmlFilePathBox.Text.Trim()
        ProxyEnabled  = $chkUseProxy.Checked
        ProxyHost     = $proxyHostBox.Text.Trim()
        ProxyPort     = $proxyPortBox.Text.Trim()
        ProxyBypass   = $proxyBypassBox.Text.Trim()
        ProxyUser     = $proxyAuthUserBox.Text.Trim()
        DepotEnabled  = $chkUseDepot.Checked
        DepotPath     = $depotPathBox.Text.Trim()
        KmsSecondaryServer = $kmsSecondaryRaw
        ApplyProductKey = $chkApplyProductKey.Checked
        ProductKeyEnc = (Protect-SecretText $normalizedProductKey)
        SavedAt       = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }

    # Confirmation
    $opts = @{
        RemoveOld=$chkRemoveOld.Checked
        AutoInstall=$chkInstall.Checked
        AutoActivate=$chkActivate.Checked
        Language=$langCombo.SelectedItem
        ApplyProductKey=$chkApplyProductKey.Checked
        UseDepot=$chkUseDepot.Checked
        DepotPath=$depotPathBox.Text.Trim()
        UseProxy=$chkUseProxy.Checked
        ProxyHost=$proxyHostBox.Text.Trim()
        ProxyPort=$proxyPortBox.Text.Trim()
    }
    if (-not (Show-LaunchConfirmation -ActivateOnly $ActivateOnly -Opts $opts -KmsRaw $kmsTextBox.Text.Trim() -KmsSecondaryRaw $kmsSecondaryRaw)) { return }

    $startButton.Enabled = $false ; $activateOnlyBtn.Enabled = $false
    $startButton.Text = if ($ActivateOnly) { "Patientez..." } else { "En cours..." }
    $activateOnlyBtn.Text = if ($ActivateOnly) { "En cours..." } else { "Patientez..." }
    $logBox.Clear() ; $script:logLines = 0 ; $chronoLabel.Visible = $true ; $progressBar.Visible = $true
    $chronoLabel.Text = "Initialisation du deploiement..."
    $statusLabel.Text = if ($ActivateOnly) { "Demarrage de l'activation..." } else { "Demarrage du deploiement..." }
    $statusLabel.ForeColor = $clrBlueLight
    $spinnerLabel.Text = "|"
    $form.Text = "Preparation du deploiement...  [$VERSION]"
    $form.UseWaitCursor = $false

    $logBox.SelectionColor = [System.Drawing.Color]::FromArgb(130, 170, 220)
    $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')]       Initialisation en cours...`n")
    $logBox.SelectionColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
    $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')]       Lancement du worker arriere-plan...`n")
    $logBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    for ($i = 0; $i -lt 5; $i++) { $script:lastState[$i] = "" ; Set-StepState $i "waiting" }

    $syncHash.Done=$false ; $syncHash.HasError=$false ; $syncHash.ActiveStep=-1
    $syncHash.StatusText="Demarrage..." ; $syncHash.StartTime=Get-Date ; $syncHash.ShowPopup=$false
    $syncHash.StepStartTime=$null ; $syncHash.StepState=[string[]]@("waiting","waiting","waiting","waiting","waiting")
    if ($null -eq $syncHash['LogQueue']) {
        $syncHash['LogQueue'] = New-Object System.Collections.Concurrent.ConcurrentQueue[object]
    }
    $entry=$null ; while ($syncHash.LogQueue.TryDequeue([ref]$entry)) {}

    $rawSetupPath = $setupPathBox.Text.Trim()
    $setupExePath = Resolve-DeploymentPath $rawSetupPath
    if ([string]::IsNullOrWhiteSpace($setupExePath)) { $setupExePath = $rawSetupPath }

    # Choisir le dossier de telechargement ODT selon les droits
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    if ($isAdmin) {
        $script:ResolvedDataPath = $LOG_FOLDER   # C:\InstallOffice2024
    } else {
        $script:ResolvedDataPath = Join-Path $env:LOCALAPPDATA "Office2024Deploy"
    }
    $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')]       Dossier ODT  : $script:ResolvedDataPath$(if (-not $isAdmin){' (non-admin, dossier utilisateur)'})`n")

    $installFolder = $script:ScriptRoot
    if (-not [string]::IsNullOrWhiteSpace($setupExePath)) {
        try {
            $tmpFolder = Split-Path -Path $setupExePath -Parent -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($tmpFolder)) { $installFolder = $tmpFolder }
        } catch {}
    }
    $xmlPathPreview = Join-Path -Path $installFolder -ChildPath "OfficeConfig2024.xml"

    $logBox.SelectionColor = [System.Drawing.Color]::FromArgb(130, 170, 220)
    $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')]       Parametres : ActivateOnly=$ActivateOnly, RemoveOld=$($chkRemoveOld.Checked), AutoInstall=$($chkInstall.Checked), AutoActivate=$($chkActivate.Checked)`n")
    $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')]       Source setup : $($setupPathBox.Text.Trim())`n")
    $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')]       Setup resolu : $setupExePath`n")
    $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')]       XML cible    : $xmlPathPreview`n")
    $logBox.ScrollToCaret()

    try {
        if ([string]::IsNullOrWhiteSpace($setupExePath)) { throw "Chemin setup.exe vide." }
        if ([string]::IsNullOrWhiteSpace($installFolder)) { $installFolder = $LOG_FOLDER }
        $xmlPath = $xmlPathPreview

        $logBox.SelectionColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
        $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')]       RunspaceAsync: envoi de la configuration au worker...`n")
        $logBox.ScrollToCaret()

        # Proxy authentifie : relais loopback, seul moyen d'ajouter l'authentification
        # pour un service LocalSystem. Le mot de passe ne quitte pas la memoire du processus.
        $proxyHostCfg = $proxyHostBox.Text.Trim()
        $proxyPortCfg = $proxyPortBox.Text.Trim()
        $proxyRelayCfg = $false
        $proxyTargetCfg = "$($proxyHostBox.Text.Trim()):$($proxyPortBox.Text.Trim())"
        Stop-ProxyRelay
        if ($chkUseProxy.Checked -and (-not $chkUseDepot.Checked) -and (-not [string]::IsNullOrWhiteSpace($proxyAuthUserBox.Text.Trim()))) {
            $pRelais = 0
            if (-not [int]::TryParse($proxyPortCfg, [ref]$pRelais) -or $pRelais -le 0 -or $pRelais -gt 65535) { throw "Port proxy invalide : $proxyPortCfg" }
            $relais = Start-ProxyRelay -UpstreamHost $proxyHostCfg -UpstreamPort $pRelais -User $proxyAuthUserBox.Text.Trim() -Password $proxyAuthPassBox.Text
            $proxyHostCfg = "127.0.0.1"
            $proxyPortCfg = [string]$relais.Port
            $proxyRelayCfg = $true
            $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')]       Proxy authentifie : relais local 127.0.0.1:$($relais.Port) -> $proxyTargetCfg (mot de passe non enregistre)`n")
            $logBox.ScrollToCaret()
        }

        $script:LastDeployConfig = [pscustomobject]@{
            SyncHash      = $syncHash
            ActivateOnly  = $ActivateOnly
            SetupExePath  = $setupExePath
            InstallFolder = $installFolder
            XMLPath       = $xmlPath
            ActualXMLPath = $xmlPath   # sera mis a jour dans DoWork
            UseCustomXML  = $radioXmlFile.Checked
            CustomXMLPath = $xmlFilePathBox.Text.Trim()
            ExcludedApps  = $script:xmlExcludedApps
            LanguageID    = $langCombo.SelectedItem
            RemoveOld     = $chkRemoveOld.Checked
            DoInstall     = $chkInstall.Checked
            DoActivate    = $chkActivate.Checked
            KmsHost       = $kmsAddr.Host
            KmsPort       = $kmsAddr.Port
            ProxyEnabled  = $chkUseProxy.Checked
            ProxyHost     = $proxyHostCfg
            ProxyPort     = $proxyPortCfg
            ProxyRelay    = $proxyRelayCfg
            ProxyRelayTarget = $proxyTargetCfg
            ProxyBypass   = $proxyBypassBox.Text.Trim()
            DepotEnabled  = $chkUseDepot.Checked
            DepotPath     = $depotPathBox.Text.Trim()
            SourcePath    = $script:ResolvedDataPath
            DepotCopy     = $false
            KmsSecondary  = $kmsSecondaryRaw
            ApplyProductKey = $chkApplyProductKey.Checked
            ProductKey    = $normalizedProductKey
            DiskMinGB     = $DISK_MIN_GB
            LocalDataPath = $script:ResolvedDataPath
        }

        $workerText = $script:BgDoWorkHandler.ToString()
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.Open()
        $rs.SessionStateProxy.SetVariable("cfg", $script:LastDeployConfig)
        $rs.SessionStateProxy.SetVariable("workerText", $workerText)

        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript({
            $worker = [scriptblock]::Create($workerText)
            $e = [System.ComponentModel.DoWorkEventArgs]::new($cfg)
            & $worker $null $e
        })

        $handle = $ps.BeginInvoke()
        $script:deployWorkerContext = @{ PS = $ps; Runspace = $rs; Handle = $handle }
        $animTimer.Start()
    } catch {
        try {
            if ($script:deployWorkerContext) {
                try { $script:deployWorkerContext.PS.Dispose() } catch {}
                try { $script:deployWorkerContext.Runspace.Close(); $script:deployWorkerContext.Runspace.Dispose() } catch {}
                $script:deployWorkerContext = $null
            }
        } catch {}
        $form.UseWaitCursor = $false
        $progressBar.Visible = $false ; $chronoLabel.Visible = $false ; $spinnerLabel.Text = ""
        $startButton.Enabled = $true ; $activateOnlyBtn.Enabled = $true
        $startButton.Text = "Lancer deploiement"
        $activateOnlyBtn.Text = "Activer seulement"
        $statusLabel.Text = "Echec au demarrage du deploiement."
        $statusLabel.ForeColor = $clrRed
        $logBox.SelectionColor = [System.Drawing.Color]::FromArgb(240,80,80)
        $logBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] [ERR]  Echec lancement worker : $($_.Exception.Message)`n")
        $logBox.ScrollToCaret()
        [System.Windows.Forms.MessageBox]::Show("Impossible de lancer le deploiement.`n`n$($_.Exception.Message)", "Erreur de lancement", "OK", "Error") | Out-Null
        return
    }
}

$startButton.Add_Click({     Start-Deployment $false })
$activateOnlyBtn.Add_Click({ Start-Deployment $true  })

# ================================================================
#  FERMETURE + CLEANUP
# ================================================================
$form.Add_FormClosing({
    if ($script:deployWorkerContext -and -not $script:deployWorkerContext.Handle.IsCompleted) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Un deploiement est en cours.`nQuitter maintenant pourrait laisser Office dans un etat instable.`n`nVoulez-vous vraiment quitter ?",
            "Deploiement en cours", "YesNo", [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($r -eq "No") { $_.Cancel = $true ; return }
    }
    Stop-ProxyRelay
    try {
        if ($script:deployWorkerContext) {
            try { $script:deployWorkerContext.PS.Stop() } catch {}
            try { $script:deployWorkerContext.PS.Dispose() } catch {}
            try { $script:deployWorkerContext.Runspace.Close(); $script:deployWorkerContext.Runspace.Dispose() } catch {}
            $script:deployWorkerContext = $null
        }
    } catch {}
    $notifyIcon.Visible = $false ; $notifyIcon.Dispose() ; $animTimer.Stop()
})

# ================================================================
#  INITIALISATION POST-AFFICHAGE
# ================================================================
$form.Add_Shown({
    Update-PreChecks | Out-Null

    # Restaurer la config sauvegardee
    $saved = Get-Config
    if ($saved) {
        if ($saved.SetupExePath) { $setupPathBox.Text = $saved.SetupExePath }
        if ($saved.KmsServer -and $saved.KmsServer -ne "KMS.VOTRE-SERVEUR.FR") {
            $kmsTextBox.Text = $saved.KmsServer ; $kmsTextBox.ForeColor = [System.Drawing.Color]::Black
        }
        if ($saved.Language -and $langCombo.Items.Contains($saved.Language)) { $langCombo.SelectedItem = $saved.Language }
        if ($null -ne $saved.RemoveOld)    { $chkRemoveOld.Checked = [bool]$saved.RemoveOld }
        if ($null -ne $saved.AutoInstall)  { $chkInstall.Checked   = [bool]$saved.AutoInstall }
        if ($null -ne $saved.AutoActivate) { $chkActivate.Checked  = [bool]$saved.AutoActivate }
        if ($saved.ExcludedApps)           { $script:xmlExcludedApps = @($saved.ExcludedApps) ; $xmlAppsLabel.Text = Get-AppsLabel }
        if ($null -ne $saved.UseCustomXML -and [bool]$saved.UseCustomXML) {
            $radioXmlFile.Checked = $true
            if ($saved.CustomXMLPath) { $xmlFilePathBox.Text = $saved.CustomXMLPath }
        }
        $script:xmlUseCustomFile = [bool]$radioXmlFile.Checked
        $script:xmlCustomFilePath = $xmlFilePathBox.Text.Trim()
        if ($null -ne $saved.ProxyEnabled) { $chkUseProxy.Checked = [bool]$saved.ProxyEnabled }
        if ($saved.ProxyHost)              { $proxyHostBox.Text = [string]$saved.ProxyHost }
        if ($saved.ProxyPort)              { $proxyPortBox.Text = [string]$saved.ProxyPort }
        if ($saved.ProxyBypass)            { $proxyBypassBox.Text = [string]$saved.ProxyBypass }
        if ($saved.ProxyUser)              { $proxyAuthUserBox.Text = [string]$saved.ProxyUser }
        if ($null -ne $saved.DepotEnabled) { $chkUseDepot.Checked = [bool]$saved.DepotEnabled }
        if ($saved.DepotPath)              { $depotPathBox.Text = [string]$saved.DepotPath }
        if ($saved.KmsSecondaryServer)     { $kmsSecondaryBox.Text = [string]$saved.KmsSecondaryServer }
        if ($null -ne $saved.ApplyProductKey) { $chkApplyProductKey.Checked = [bool]$saved.ApplyProductKey }
        if ($saved.ProductKeyEnc)          { $productKeyBox.Text = Unprotect-SecretText ([string]$saved.ProductKeyEnc) }
        $sysInfoLabel.Text += "  [Config restauree]"
    }

    $depotPathBox.Enabled   = $chkUseDepot.Checked
    $depotBrowseBtn.Enabled = $chkUseDepot.Checked
    Update-ProxyControlState

    # Relancer pre-check avec le chemin restaure
    Update-PreChecks | Out-Null
})

# ================================================================
#  LANCEMENT
# ================================================================
# Ecrans de faible hauteur (1024x768) : autoriser le defilement plutot que de couper la fenetre.
$workArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
if ($workArea.Height -lt ($form.Height + 20)) {
    $form.AutoScroll = $true
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox = $true
    $form.ClientSize = New-Object System.Drawing.Size(($form.ClientSize.Width + 18), [Math]::Max(360, ($workArea.Height - 90)))
}

[void]$form.ShowDialog()
