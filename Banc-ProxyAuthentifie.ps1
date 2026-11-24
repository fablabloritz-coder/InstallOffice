#Requires -Version 5.1
<#
    Banc de preuve : proxy authentifie de bout en bout.

    Ce que le banc montre, dans l'ordre :
      1. un VRAI proxy amont local (processus a lui, journal par requete) qui exige
         un Basic et repond 407 sans identifiants ;
      2. le relais de l'outil (Start-AuthenticatedProxy, extrait du moteur) devant ce
         proxy : le client parle au relais sans identifiants, le proxy voit l'en-tete ;
      3. une cible HTTP locale atteinte A TRAVERS le tunnel, ce qui prouve que la
         chaine complete transporte des octets et pas seulement un en-tete ;
      4. le meme essai avec un mauvais mot de passe : 407 remonte au client ;
      5. un proxy qui n'offre que NTLM : le 407 et son en-tete Proxy-Authenticate
         remontent au client, ce qui permet de nommer le schema ;
      6. une connexion que le client voudrait reutiliser : le relais annonce la
         fermeture (Connection: close) et la requete suivante, sur une connexion
         neuve, est authentifiee elle aussi ;
      7. une reponse dont l'en-tete est tronque : le relais ferme, il ne relaie pas
         une reponse partielle ;
      5. en option (-Odt), un vrai `setup.exe /download` derriere le relais, borne
         dans le temps, pour voir si Click-to-Run passe par le relais. Le journal du
         proxy tranche : une ligne CONNECT vers officecdn.microsoft.com vaut preuve.

    Le banc n'installe rien et ne touche au poste que si -Odt est demande : il regle
    alors les variables d'environnement (et WinINet avec -WinInet), puis les restaure.
    WinHTTP demande les droits administrateur : sans eux, ODT peut ignorer le relais,
    et le rapport le dit au lieu de le masquer.

    Usage :
      powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Banc-ProxyAuthentifie.ps1
      ... -Rapport .\banc.json
      ... -Odt -SecondesOdt 60
      ... -Odt -WinInet
#>
param(
    [string]$Utilisateur = "banc",
    [string]$MotDePasse = "secret-banc",
    [string]$Rapport = "",
    [switch]$SansConsole,
    [switch]$Odt,
    [int]$SecondesOdt = 60,
    [switch]$WinInet,
    [string]$SetupExe = ""
)

$ErrorActionPreference = "Stop"
$racine = $PSScriptRoot
$moteur = Join-Path $racine "Install-Office2024-KMS.ps1"
if ([string]::IsNullOrWhiteSpace($SetupExe)) { $SetupExe = Join-Path $racine "setup.exe" }

function Write-Banc {
    param([string]$Texte)
    if (-not $SansConsole) { Write-Host $Texte }
}

# --- Le relais de l'outil, extrait du moteur (meme code que l'ecran) --------
function Import-FonctionMoteur {
    param([string]$Chemin, [string[]]$Noms)
    if (-not (Test-Path -LiteralPath $Chemin)) { throw ("Moteur introuvable : " + $Chemin) }
    $jetons = $null ; $erreurs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Chemin, [ref]$jetons, [ref]$erreurs)
    $defs = @{}
    foreach ($d in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        if (($Noms -contains $d.Name) -and -not $defs.ContainsKey($d.Name)) { $defs[$d.Name] = $d.Extent.Text }
    }
    $absents = @($Noms | Where-Object { -not $defs.ContainsKey($_) })
    if ($absents.Count -gt 0) { throw ("Moteur : fonction(s) introuvable(s) : " + ($absents -join ", ")) }
    return (@($Noms | ForEach-Object { $defs[$_] }) -join "`r`n`r`n")
}

. ([scriptblock]::Create((Import-FonctionMoteur $moteur @("Start-ProxyRelay", "Stop-ProxyRelay", "Start-AuthenticatedProxy"))))

# --- Proxy amont reel et cible locale --------------------------------------
$code = @'
using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

public static class BancReseau
{
    // Entete HTTP lu octet par octet jusqu'a la ligne vide : la meme lecture que le
    // relais, pour que les deux bouts se comprennent sans negociation.
    public static string LireEntete(NetworkStream s)
    {
        StringBuilder sb = new StringBuilder();
        byte[] un = new byte[1];
        while (sb.Length < 16384)
        {
            if (s.Read(un, 0, 1) <= 0) { break; }
            sb.Append((char)un[0]);
            if (sb.Length >= 4 && sb[sb.Length - 1] == '\n' && sb[sb.Length - 3] == '\n') { break; }
        }
        return sb.ToString();
    }

    public static bool Decouper(string cible, out string hote, out int port)
    {
        hote = ""; port = 0;
        if (string.IsNullOrEmpty(cible)) { return false; }
        int i = cible.LastIndexOf(':');
        if (i <= 0 || i == cible.Length - 1) { return false; }
        hote = cible.Substring(0, i).Trim('[', ']');
        return int.TryParse(cible.Substring(i + 1), out port);
    }
}

// Un vrai proxy direct : il valide le Basic, ouvre le tunnel vers la cible demandee
// et journalise chaque requete. C'est ce journal qui sert de preuve.
public class BancProxyAmont
{
    private readonly string _attendu;
    private readonly string _journal;
    private readonly object _verrou = new object();
    private TcpListener _l;
    private int _port;
    private int _tunnels;
    private int _refus;
    private int _requetes;

    public BancProxyAmont(string user, string password, string journal)
    {
        _attendu = "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(user + ":" + password));
        _journal = journal;
    }

    // Arme, ce proxy ne propose que NTLM : le cas d'un proxy d'etablissement qui
    // refuse Basic, celui que le diagnostic doit savoir nommer.
    public bool NegocieSeul;

    public int Port { get { return _port; } }
    public int Tunnels { get { return _tunnels; } }
    public int Refus { get { return _refus; } }
    public int Requetes { get { return _requetes; } }

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

    private void Log(string ligne)
    {
        lock (_verrou)
        {
            try { File.AppendAllText(_journal, DateTime.Now.ToString("HH:mm:ss.fff") + " " + ligne + Environment.NewLine, new UTF8Encoding(false)); } catch { }
        }
    }

    private static void Repondre(NetworkStream s, string texte)
    {
        try
        {
            byte[] b = Encoding.ASCII.GetBytes(texte);
            s.Write(b, 0, b.Length);
            s.Flush();
        }
        catch { }
    }

    private void Serve(TcpClient c)
    {
        try
        {
            NetworkStream s = c.GetStream();
            string entete = BancReseau.LireEntete(s);
            string[] lignes = entete.Split('\n');
            string premiere = lignes.Length > 0 ? lignes[0].Trim() : "";
            string[] parties = premiere.Split(' ');
            string methode = parties.Length > 0 ? parties[0] : "";
            string cible = parties.Length > 1 ? parties[1] : "";

            string auth = "";
            foreach (string l in lignes)
            {
                string t = l.Trim();
                if (t.StartsWith("Proxy-Authorization:", StringComparison.OrdinalIgnoreCase))
                {
                    auth = t.Substring("Proxy-Authorization:".Length).Trim();
                }
            }
            bool ok = (auth == _attendu);
            string utilisateur = "";
            if (auth.StartsWith("Basic ", StringComparison.OrdinalIgnoreCase))
            {
                try { utilisateur = Encoding.UTF8.GetString(Convert.FromBase64String(auth.Substring(6))).Split(':')[0]; } catch { }
            }
            Interlocked.Increment(ref _requetes);
            Log(methode + " " + cible + " auth=" + (ok ? "oui" : "non") + (utilisateur != "" ? " user=" + utilisateur : ""));

            if (!ok || NegocieSeul)
            {
                Interlocked.Increment(ref _refus);
                Log("407 " + cible + " (identifiants absents ou faux)");
                string defi = NegocieSeul ? "NTLM" : "Basic realm=\"banc\"";
                Repondre(s, "HTTP/1.1 407 Proxy Authentication Required\r\nProxy-Authenticate: " + defi + "\r\nContent-Length: 0\r\nConnection: close\r\n\r\n");
                return;
            }
            if (methode != "CONNECT")
            {
                // Requete en clair a URL absolue : le CDN Office sert les .cab du
                // telechargement comme cela. Le banc se comporte donc en vrai proxy
                // direct, sinon il refuserait le trafic reel et ne prouverait rien.
                Uri uri;
                if (!Uri.TryCreate(cible, UriKind.Absolute, out uri) || string.IsNullOrEmpty(uri.Host))
                {
                    Log("400 " + cible + " (URL absolue attendue)");
                    Repondre(s, "HTTP/1.1 400 Bad Request\r\nContent-Length: 0\r\nConnection: close\r\n\r\n");
                    return;
                }
                TcpClient origine = new TcpClient();
                origine.Connect(uri.Host, uri.Port);
                NetworkStream os = origine.GetStream();
                StringBuilder req = new StringBuilder();
                req.Append(methode).Append(' ').Append(uri.PathAndQuery).Append(" HTTP/1.1\r\n");
                bool hostVu = false;
                foreach (string l in lignes)
                {
                    string t = l.TrimEnd('\r');
                    if (t.Length == 0 || t.Equals(premiere, StringComparison.Ordinal)) { continue; }
                    if (t.StartsWith("Proxy-Authorization:", StringComparison.OrdinalIgnoreCase)) { continue; }
                    if (t.StartsWith("Proxy-Connection:", StringComparison.OrdinalIgnoreCase)) { continue; }
                    if (t.StartsWith("Host:", StringComparison.OrdinalIgnoreCase))
                    {
                        req.Append("Host: ").Append(uri.Host).Append("\r\n");
                        hostVu = true;
                        continue;
                    }
                    req.Append(t).Append("\r\n");
                }
                if (!hostVu) { req.Append("Host: ").Append(uri.Host).Append("\r\n"); }
                req.Append("\r\n");
                byte[] rb2 = Encoding.ASCII.GetBytes(req.ToString());
                os.Write(rb2, 0, rb2.Length);
                os.Flush();
                Interlocked.Increment(ref _tunnels);
                Log("200 " + cible + " flux ouvert vers " + uri.Host + ":" + uri.Port);
                // L'en-tete de reponse est lu ici pour y prendre Content-Length : c'est
                // lui qui dira, dans le rapport, si le fichier recu est entier.
                string rep = BancReseau.LireEntete(os);
                long longueur = -1;
                long total = -1;
                foreach (string l in rep.Split('\n'))
                {
                    string t = l.TrimEnd('\r');
                    if (t.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
                    {
                        long v;
                        if (long.TryParse(t.Substring(15).Trim(), out v)) { longueur = v; }
                    }
                    else if (t.StartsWith("Content-Range:", StringComparison.OrdinalIgnoreCase))
                    {
                        // "bytes 0-1048575/4050066" : le total, c'est la taille du fichier.
                        // ODT telecharge les gros .cab par plages : comparer le fichier au
                        // Content-Length d'une plage le declarerait tronque a tort.
                        string[] morceaux = t.Substring(14).Trim().Split('/');
                        long v;
                        if (morceaux.Length == 2 && long.TryParse(morceaux[1].Trim(), out v)) { total = v; }
                    }
                }
                if (total < 0) { total = longueur; }
                Log("LEN " + total + " " + cible);
                byte[] rb3 = Encoding.ASCII.GetBytes(rep);
                s.Write(rb3, 0, rb3.Length);
                s.Flush();
                Task ta = s.CopyToAsync(os);
                Task tb = os.CopyToAsync(s);
                Task.WaitAny(ta, tb);
                try { origine.Close(); } catch { }
                Log("flux ferme " + cible);
                return;
            }

            string hote; int port;
            if (!BancReseau.Decouper(cible, out hote, out port))
            {
                Log("400 " + cible);
                Repondre(s, "HTTP/1.1 400 Bad Request\r\nContent-Length: 0\r\nConnection: close\r\n\r\n");
                return;
            }

            TcpClient amont = new TcpClient();
            amont.Connect(hote, port);
            Interlocked.Increment(ref _tunnels);
            Log("200 " + cible + " tunnel ouvert");
            Repondre(s, "HTTP/1.1 200 Connection established\r\n\r\n");

            NetworkStream us = amont.GetStream();
            Task a = s.CopyToAsync(us);
            Task b = us.CopyToAsync(s);
            Task.WaitAny(a, b);
            try { amont.Close(); } catch { }
            Log("tunnel ferme " + cible);
        }
        catch (Exception ex)
        {
            Log("erreur " + ex.Message);
        }
        finally
        {
            try { c.Close(); } catch { }
        }
    }
}

// Cible locale : repond un corps connu, ce qui prouve que le tunnel transporte
// vraiment des octets dans les deux sens.
public class BancCibleHttp
{
    private TcpListener _l;
    private int _port;
    private int _hits;

    // Arme, la cible coupe son en-tete en pleine ligne : le relais doit fermer, pas
    // relayer une reponse partielle.
    public bool Tronque;

    public int Port { get { return _port; } }
    public int Hits { get { return _hits; } }

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
            try
            {
                NetworkStream s = c.GetStream();
                BancReseau.LireEntete(s);
                Interlocked.Increment(ref _hits);
                if (Tronque)
                {
                    byte[] partiel = Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Length: 5\r\n");
                    s.Write(partiel, 0, partiel.Length);
                    s.Flush();
                    return;
                }
                string corps = "BANC-OK";
                byte[] b = Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\nContent-Length: " + corps.Length + "\r\nConnection: close\r\n\r\n" + corps);
                s.Write(b, 0, b.Length);
                s.Flush();
            }
            catch { }
            finally { try { c.Close(); } catch { } }
        }
    }
}
'@
if (-not ("BancProxyAmont" -as [type])) { Add-Type -TypeDefinition $code -Language CSharp -ErrorAction Stop }

# --- Client : CONNECT au relais, puis une requete HTTP dans le tunnel ------
function Invoke-TunnelBanc {
    param([int]$PortRelais, [string]$Cible)
    $r = @{ Ok = $false ; Statut = "" ; Corps = "" ; Detail = "" ; Entete = "" }
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $client.Connect("127.0.0.1", $PortRelais)
        $flux = $client.GetStream()
        $demande = "CONNECT $Cible HTTP/1.1`r`nHost: $Cible`r`n`r`n"
        $b = [System.Text.Encoding]::ASCII.GetBytes($demande)
        $flux.Write($b, 0, $b.Length) ; $flux.Flush()
        $entete = [BancReseau]::LireEntete($flux)
        $r.Entete = $entete
        $r.Statut = ($entete -split "`r`n")[0].Trim()
        if ($r.Statut -notmatch " 200 ") { return $r }
        $demande2 = "GET /banc HTTP/1.1`r`nHost: $Cible`r`nConnection: close`r`n`r`n"
        $b2 = [System.Text.Encoding]::ASCII.GetBytes($demande2)
        $flux.Write($b2, 0, $b2.Length) ; $flux.Flush()
        $sb = New-Object System.Text.StringBuilder
        $tampon = New-Object byte[] 4096
        while ($true) {
            $n = $flux.Read($tampon, 0, $tampon.Length)
            if ($n -le 0) { break }
            [void]$sb.Append([System.Text.Encoding]::ASCII.GetString($tampon, 0, $n))
        }
        $r.Corps = $sb.ToString()
        $r.Ok = ($r.Corps -match "BANC-OK")
    } catch {
        $r.Detail = $_.Exception.Message
    } finally {
        try { $client.Close() } catch { }
    }
    return $r
}

# --- Client simple : une requete, un en-tete, un corps lu jusqu'a la fermeture ---
# Le tunnel (test 1) garde son client a deux temps : CONNECT, puis requete dans le
# tunnel. Tout le reste passe par ici : une seule facon d'ouvrir et de lire.
function Invoke-EchangeBanc {
    param([int]$PortRelais, [string]$Demande)
    $r = @{ Statut = "" ; Entete = "" ; Corps = "" ; Detail = "" }
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $client.Connect("127.0.0.1", $PortRelais)
        $flux = $client.GetStream()
        $b = [System.Text.Encoding]::ASCII.GetBytes($Demande)
        $flux.Write($b, 0, $b.Length) ; $flux.Flush()
        $entete = [BancReseau]::LireEntete($flux)
        $r.Entete = $entete
        $r.Statut = ($entete -split "`r`n")[0].Trim()
        $sb = New-Object System.Text.StringBuilder
        $tampon = New-Object byte[] 4096
        while ($true) {
            $n = $flux.Read($tampon, 0, $tampon.Length)
            if ($n -le 0) { break }
            [void]$sb.Append([System.Text.Encoding]::ASCII.GetString($tampon, 0, $n))
        }
        $r.Corps = $sb.ToString()
    } catch {
        $r.Detail = $_.Exception.Message
    } finally {
        try { $client.Close() } catch { }
    }
    return $r
}

# --- Deroulement ------------------------------------------------------------
$dossier = Join-Path $env:TEMP ("banc_proxy_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $dossier -Force | Out-Null
$journalProxy = Join-Path $dossier "proxy-amont.log"
$tests = @()
$odtResultat = @{ Tente = $false }
$proxy = $null ; $cible = $null ; $relaisOk = $null ; $relaisKo = $null
$envAvant = @{}
$inetAvant = $null
$sourcesOdt = ""

try {
    $proxy = New-Object BancProxyAmont -ArgumentList @($Utilisateur, $MotDePasse, $journalProxy)
    $proxy.Start()
    $cible = New-Object BancCibleHttp
    $cible.Start()
    $cibleAdresse = "127.0.0.1:" + $cible.Port

    Write-Banc ("Proxy amont : 127.0.0.1:" + $proxy.Port + " (Basic, 407 sans identifiants)")
    Write-Banc ("Cible locale : " + $cibleAdresse)
    Write-Banc ""

    # 1. Relais de l'outil + tunnel + Basic ajoute par le relais.
    $relaisOk = Start-AuthenticatedProxy -ProxyHost "127.0.0.1" -ProxyPort ([string]$proxy.Port) -User $Utilisateur -Password $MotDePasse
    Write-Banc ("Relais de l'outil : " + $relaisOk.Host + ":" + $relaisOk.Port + " -> " + $relaisOk.Target)
    $a = Invoke-TunnelBanc -PortRelais ([int]$relaisOk.Port) -Cible $cibleAdresse
    $tests += [ordered]@{ Nom = "tunnel + Basic ajoute par le relais" ; Ok = ($a.Ok -and $a.Statut -match " 200 ") ; Statut = $a.Statut ; Detail = $(if ($a.Ok) { "corps recu : BANC-OK" } else { $a.Detail }) }

    # 2. Meme relais, mauvais mot de passe : le 407 du proxy remonte au client.
    $relaisKo = Start-AuthenticatedProxy -ProxyHost "127.0.0.1" -ProxyPort ([string]$proxy.Port) -User $Utilisateur -Password "faux-mot-de-passe"
    $b = Invoke-TunnelBanc -PortRelais ([int]$relaisKo.Port) -Cible $cibleAdresse
    $tests += [ordered]@{ Nom = "mauvais mot de passe -> 407" ; Ok = ($b.Statut -match " 407 ") ; Statut = $b.Statut ; Detail = "le client voit le refus du proxy, pas un tunnel muet" }

    # 3. Requete en clair a URL absolue : c'est ainsi que le CDN Office sert les .cab
    #    du telechargement. Le relais doit la transmettre telle quelle, avec l'en-tete,
    #    et non la convertir en CONNECT (defaut vu sur un vrai setup.exe /download).
    $urlBanc = "http://127.0.0.1:" + $cible.Port + "/banc"
    $autoriteBanc = "127.0.0.1:" + $cible.Port
    $c = Invoke-EchangeBanc -PortRelais ([int]$relaisOk.Port) -Demande ("GET $urlBanc HTTP/1.1`r`nHost: $autoriteBanc`r`nConnection: close`r`n`r`n")
    $tests += [ordered]@{ Nom = "requete HTTP en clair (URL absolue) + Basic" ; Ok = ($c.Corps -match "BANC-OK") ; Statut = $c.Statut ; Detail = "le CDN Office sert les .cab ainsi : transmise, pas convertie en CONNECT" }

    # 4. Proxy qui n'offre que NTLM : le relais doit laisser remonter le 407 et son
    #    en-tete Proxy-Authenticate. C'est ce que lit le diagnostic du moteur pour
    #    nommer le schema au lieu de faire chercher un mot de passe faux.
    $ntlm = New-Object BancProxyAmont -ArgumentList @($Utilisateur, $MotDePasse, $journalProxy)
    $ntlm.NegocieSeul = $true
    $ntlm.Start()
    $relaisNtlm = Start-AuthenticatedProxy -ProxyHost "127.0.0.1" -ProxyPort ([string]$ntlm.Port) -User $Utilisateur -Password $MotDePasse
    $d = Invoke-TunnelBanc -PortRelais ([int]$relaisNtlm.Port) -Cible $cibleAdresse
    $tests += [ordered]@{ Nom = "proxy NTLM seul : le 407 et son schema remontent" ; Ok = (($d.Statut -match " 407 ") -and ($d.Entete -match '(?i)Proxy-Authenticate:\s*NTLM')) ; Statut = $d.Statut ; Detail = "le diagnostic peut nommer NTLM au lieu d'accuser le mot de passe" }
    Stop-ProxyRelay -Relais $relaisNtlm
    $ntlm.Stop()

    # 5. Connexion reutilisee : le relais ferme proprement (Connection: close) au lieu de
    #    laisser repartir une requete sans en-tete ; la requete suivante, sur une
    #    connexion neuve, est authentifiee elle aussi et ne prend aucun 407.
    $refusAvant = $proxy.Refus
    $k1 = Invoke-EchangeBanc -PortRelais ([int]$relaisOk.Port) -Demande ("GET $urlBanc HTTP/1.1`r`nHost: $autoriteBanc`r`nConnection: keep-alive`r`n`r`n")
    $k2 = Invoke-EchangeBanc -PortRelais ([int]$relaisOk.Port) -Demande ("GET $urlBanc HTTP/1.1`r`nHost: $autoriteBanc`r`nConnection: close`r`n`r`n")
    $refusApres = $proxy.Refus
    $ferme = ($k1.Entete -match '(?i)Connection:\s*close')
    $tests += [ordered]@{ Nom = "connexion reutilisee : fermeture annoncee, requete suivante authentifiee" ; Ok = (($k1.Corps -match "BANC-OK") -and ($k2.Corps -match "BANC-OK") -and $ferme -and (($refusApres - $refusAvant) -eq 0)) ; Statut = $k1.Statut + " / " + $k2.Statut ; Detail = "Connection: close annonce=" + $ferme + " refus en plus=" + ($refusApres - $refusAvant) }

    # 6. En-tete tronque : le relais ferme au lieu de relayer une reponse coupee.
    $tronque = New-Object BancCibleHttp
    $tronque.Tronque = $true
    $tronque.Start()
    $urlTronque = "http://127.0.0.1:" + $tronque.Port + "/banc"
    $x = Invoke-EchangeBanc -PortRelais ([int]$relaisOk.Port) -Demande ("GET $urlTronque HTTP/1.1`r`nHost: 127.0.0.1:" + $tronque.Port + "`r`nConnection: close`r`n`r`n")
    $tronque.Stop()
    $tests += [ordered]@{ Nom = "en-tete tronque : rien de partiel n'est relaye" ; Ok = ($x.Entete -eq "") ; Statut = $(if ($x.Statut -eq "") { "connexion fermee sans en-tete" } else { $x.Statut }) ; Detail = "le relais ferme au lieu de donner une reponse coupee au client" }

    # 3. ODT reel, borne dans le temps : le journal du proxy tranche.
    if ($Odt) {
        $odtResultat = @{ Tente = $true ; Ok = $false ; Secondes = $SecondesOdt ; Lignes = @() ; Cabs = @() ; Complet = $false ; Note = "" }
        if (-not (Test-Path -LiteralPath $SetupExe)) {
            $odtResultat.Note = "setup.exe absent : " + $SetupExe
        } else {
            $sourcesOdt = Join-Path $dossier "sources"
            New-Item -ItemType Directory -Path $sourcesOdt -Force | Out-Null
            $xml = Join-Path $dossier "banc-odt.xml"
            $contenu = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024" SourcePath="$sourcesOdt">
    <Product ID="ProPlus2024Volume">
      <Language ID="fr-fr" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="Lync" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
            [System.IO.File]::WriteAllText($xml, $contenu, (New-Object System.Text.UTF8Encoding($false)))

            # Le relais pour ODT : sans identifiants dans la configuration, c'est le
            # relais qui les ajoute.
            foreach ($n in @("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy")) {
                $envAvant[$n] = [System.Environment]::GetEnvironmentVariable($n, "Process")
                [System.Environment]::SetEnvironmentVariable($n, ("http://127.0.0.1:" + $relaisOk.Port), "Process")
            }
            $cleInet = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
            if ($WinInet) {
                $inetAvant = @{}
                foreach ($p in @("ProxyEnable", "ProxyServer", "ProxyOverride")) {
                    $inetAvant[$p] = $null
                    try { $inetAvant[$p] = Get-ItemPropertyValue -Path $cleInet -Name $p -ErrorAction Stop } catch { }
                }
                New-ItemProperty -Path $cleInet -Name "ProxyEnable" -Value 1 -PropertyType DWord -Force | Out-Null
                New-ItemProperty -Path $cleInet -Name "ProxyServer" -Value ("127.0.0.1:" + $relaisOk.Port) -PropertyType String -Force | Out-Null
                New-ItemProperty -Path $cleInet -Name "ProxyOverride" -Value "<local>" -PropertyType String -Force | Out-Null
            }

            Write-Banc ("ODT : setup.exe /download derriere le relais, " + $SecondesOdt + " s au maximum...")
            $p = Start-Process -FilePath $SetupExe -ArgumentList @("/download", $xml) -PassThru -WindowStyle Hidden
            if (-not $p.WaitForExit($SecondesOdt * 1000)) {
                try { Stop-Process -Id $p.Id -Force } catch { }
                $odtResultat.Note = ("arrete apres " + $SecondesOdt + " s (le telechargement complet n'est pas le sujet)")
            } else {
                $odtResultat.Note = ("setup.exe termine avec le code " + $p.ExitCode)
            }
            Start-Sleep -Milliseconds 500
            $lignes = @()
            if (Test-Path -LiteralPath $journalProxy) { $lignes = @(Get-Content -LiteralPath $journalProxy) }
            $versCdn = @($lignes | Where-Object { $_ -match "CONNECT .*officecdn|CONNECT .*microsoft" })
            $odtResultat.Lignes = @($versCdn | ForEach-Object { [string]$_ })

            # Un cab est complet quand sa taille sur disque egale le Content-Length annonce
            # par le CDN : c'est la difference entre "le flux passe" et "le fichier arrive
            # entier". Un .cab servi en HTTPS (tunnel) n'annonce pas sa taille : inconnue.
            $attentes = @{}
            foreach ($l in $lignes) {
                if ($l -match '^[0-9:.]+\s+LEN\s+(-?\d+)\s+(\S+)$') {
                    $n = [long]$matches[1]
                    $nomCab = [System.IO.Path]::GetFileName(([string]$matches[2] -split '\?')[0])
                    if (($n -ge 0) -and ($nomCab -ne "")) { $attentes[$nomCab.ToLower()] = $n }
                }
            }
            $cabs = @()
            if (Test-Path -LiteralPath $sourcesOdt) {
                foreach ($f in @(Get-ChildItem -LiteralPath $sourcesOdt -Recurse -Filter *.cab -File -ErrorAction SilentlyContinue)) {
                    $attendu = -1
                    if ($attentes.ContainsKey($f.Name.ToLower())) { $attendu = [long]$attentes[$f.Name.ToLower()] }
                    $complet = $null
                    if ($attendu -ge 0) { $complet = ($f.Length -eq $attendu) }
                    $cabs += [ordered]@{ Nom = $f.Name ; Octets = [long]$f.Length ; Attendu = $attendu ; Complet = $complet }
                }
            }
            $odtResultat.Cabs = @($cabs)
            $odtResultat.Complet = (@($cabs | Where-Object { $_.Complet -eq $true }).Count -gt 0)
            $odtResultat.Ok = ($versCdn.Count -gt 0) -and $odtResultat.Complet
            if (-not $odtResultat.Ok) {
                if ($versCdn.Count -eq 0) {
                    $odtResultat.Note += " ; aucune ligne CONNECT vers le CDN : Click-to-Run n'a pas utilise le relais (WinHTTP n'est reglable qu'en administrateur)"
                } elseif (-not $odtResultat.Complet) {
                    $odtResultat.Note += " ; le flux passe mais aucun .cab complet : augmenter -SecondesOdt"
                }
            }
        }
    }
} finally {
    if ($relaisOk) { try { $relaisOk.Relay.Stop() } catch { } }
    if ($relaisKo) { try { $relaisKo.Relay.Stop() } catch { } }
    if ($null -ne $proxy) { try { $proxy.Stop() } catch { } }
    if ($null -ne $cible) { try { $cible.Stop() } catch { } }
    foreach ($n in @("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy")) {
        if ($envAvant.ContainsKey($n)) { [System.Environment]::SetEnvironmentVariable($n, $envAvant[$n], "Process") }
    }
    if ($null -ne $inetAvant) {
        $cleInet = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
        foreach ($p in @("ProxyEnable", "ProxyServer", "ProxyOverride")) {
            if ($null -eq $inetAvant[$p]) { Remove-ItemProperty -Path $cleInet -Name $p -ErrorAction SilentlyContinue }
            else {
                $type = if ($p -eq "ProxyEnable") { "DWord" } else { "String" }
                New-ItemProperty -Path $cleInet -Name $p -Value $inetAvant[$p] -PropertyType $type -Force | Out-Null
            }
        }
    }
}

# --- Journal du proxy en regard --------------------------------------------
$lignesProxy = @()
if (Test-Path -LiteralPath $journalProxy) { $lignesProxy = @(Get-Content -LiteralPath $journalProxy) }

Write-Banc ""
Write-Banc "=== Journal du proxy amont ==="
foreach ($l in $lignesProxy) { Write-Banc ("  " + $l) }
Write-Banc ""
Write-Banc "=== Resultats ==="
foreach ($t in $tests) {
    Write-Banc ("  " + $(if ($t.Ok) { "OK   " } else { "ECHEC" }) + "  " + $t.Nom + "  [" + $t.Statut + "] " + $t.Detail)
}
if ($odtResultat.Tente) {
    Write-Banc ("  " + $(if ($odtResultat.Ok) { "OK   " } else { "ECHEC" }) + "  ODT derriere le relais : " + $odtResultat.Note)
    foreach ($l in $odtResultat.Lignes) { Write-Banc ("         " + $l) }
    foreach ($c in @($odtResultat.Cabs)) {
        $vue = if ($null -eq $c.Complet) { "taille inconnue (HTTPS)" } elseif ($c.Complet) { "complet" } else { "TRONQUE" }
        Write-Banc ("         " + $c.Nom + "  " + $c.Octets + " / " + $c.Attendu + " octets  " + $vue)
    }
}

$resume = [ordered]@{
    Banc        = "proxy authentifie"
    Date        = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Poste       = $env:COMPUTERNAME
    Administrateur = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    Proxy       = [ordered]@{ Port = $(if ($null -ne $proxy) { $proxy.Port } else { 0 }) ; Utilisateur = $Utilisateur ; Requetes = $(if ($null -ne $proxy) { $proxy.Requetes } else { 0 }) ; Tunnels = $(if ($null -ne $proxy) { $proxy.Tunnels } else { 0 }) ; Refus = $(if ($null -ne $proxy) { $proxy.Refus } else { 0 }) }
    Relais      = [ordered]@{ Port = $(if ($null -ne $relaisOk) { [int]$relaisOk.Port } else { 0 }) ; Cible = $(if ($null -ne $relaisOk) { $relaisOk.Target } else { "" }) }
    Cible       = [ordered]@{ Port = $(if ($null -ne $cible) { $cible.Port } else { 0 }) ; Requetes = $(if ($null -ne $cible) { $cible.Hits } else { 0 }) }
    Tests       = @($tests)
    Odt         = $odtResultat
    # Chaines explicites : un tableau de PSObject se serialiserait en objets {value:...}.
    JournalProxy = @($lignesProxy | ForEach-Object { [string]$_ })
}
$ok = (@($tests | Where-Object { -not $_.Ok }).Count -eq 0)
$resume["Ok"] = $ok

if (-not [string]::IsNullOrWhiteSpace($Rapport)) {
    $json = ConvertTo-Json -InputObject $resume -Depth 6
    $dossierRapport = Split-Path -Parent $Rapport
    if (-not [string]::IsNullOrWhiteSpace($dossierRapport) -and -not (Test-Path -LiteralPath $dossierRapport)) {
        New-Item -ItemType Directory -Path $dossierRapport -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Rapport, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Banc ("Rapport : " + $Rapport)
}

# Le dossier de travail ne survit pas au banc : sources ODT comprises.
Remove-Item -LiteralPath $dossier -Recurse -Force -ErrorAction SilentlyContinue

if ($ok) { exit 0 }
exit 1
