# InstallOffice — déploiement d'Office LTSC 2024 + activation KMS

Outil graphique (PowerShell / WinForms) pour installer Microsoft Office LTSC 2024
(`ProPlus2024Volume`, canal `PerpetualVL2024`) puis l'activer par KMS, sur un poste
d'un réseau pédagogique.

## Contenu du dépôt

| Fichier | Rôle |
| --- | --- |
| `Install-Office2024-KMS.ps1` | Source de l'interface (v8.1) |
| `Install-Office2024-KMS.exe` | Version compilée (PS2EXE, manifeste `requireAdmin`) |
| `OfficeConfig2024.xml` | Modèle de configuration ODT (référence / XML personnalisé) |
| `Test-InstallOffice.ps1` | Contrôles automatisés hors ligne (90 cas, `-Fast` pour la section 1 seule) |
| `.githooks/pre-commit` | Contrôle rapide de l'index avant chaque commit, refuse l'EXE désynchronisé |
| `.github/workflows/controle-exe.yml` | Suite complète + contrat de fins de ligne (checkout sans conversion) sur GitHub Actions, à chaque push sur `main` et à chaque pull request |
| `setup.exe` | Office Deployment Tool — **non versionné**, à récupérer (voir ci-dessous) |

## Prérequis

- Windows 10/11, 64 bits, PowerShell 5.1
- Droits administrateur (l'EXE les demande via son manifeste)
- `setup.exe` (Office Deployment Tool) : <https://www.microsoft.com/download/details.aspx?id=49117>
  — le placer à côté du script/EXE (il y est détecté automatiquement)
- Serveur KMS joignable sur le port 1688 pour l'activation

## Les deux modes

**1. Téléchargement → installation → activation** (poste avec accès au CDN Microsoft)
Le script applique le proxy au niveau système, télécharge les sources (`setup.exe /download`),
installe en silencieux, puis active par KMS.

**2. Dépôt de sources** (recommandé sur le réseau pédagogique)
Télécharger les sources **une fois** dans un dossier (local ou partagé), puis cocher
« Dépôt de sources » sur chaque poste : aucune connexion Internet n'est alors nécessaire,
l'installation et l'activation KMS se font en une seule passe. Si le dépôt est un partage
UNC, les sources sont recopiées localement (`robocopy`), car le service Click-to-Run ne
lit pas les partages sous le compte `LocalSystem`.

## Pourquoi un réglage proxy dédié ?

`setup.exe /download` ne lit **pas** `HTTP_PROXY`. Le téléchargement est effectué par le
service **Microsoft Office Click-to-Run**, qui tourne en `LocalSystem` et utilise la
configuration **WinHTTP de la machine**. Le script :

1. teste réellement le proxy (`CONNECT officecdn.microsoft.com:443`) avant de lancer
   les ~20 minutes de téléchargement (détecte proxy injoignable, destination filtrée, 407) ;
2. applique le proxy à WinHTTP (machine) **et** à WinINet (session), plus les variables
   d'environnement ;
3. **restaure** la configuration d'origine en fin de déploiement.

Un proxy **avec authentification** n'est pas exploitable par un service `LocalSystem` :
Un proxy **avec authentification** (407) ne peut pas être satisfait par le service
Click-to-Run, qui tourne en `LocalSystem` et n'a aucun moyen de saisir un mot de passe.
C'est pourquoi les champs « Utilisateur » et « Mot de passe » du proxy existent : le
programme démarre alors un **relais local** (sur `127.0.0.1`, port libre) qui ajoute
l'en-tête `Proxy-Authorization` vers le vrai proxy et laisse passer le tunnel ; Windows est
pointé sur ce relais, qui n'est pour Office qu'un proxy ordinaire. Le relais ne vit que le
temps du déploiement et le mot de passe n'est **jamais** écrit sur disque : il faut le
ressaisir à chaque lancement. Seule l'authentification **Basic** est prise en charge ; un
proxy qui n'accepte que NTLM ou Kerberos reste hors de portée, d'où le mode « Dépôt de
sources ». Si le programme est tué pendant le téléchargement, avant la restauration, le poste
garderait un proxy WinHTTP pointant sur un port mort. Le port du relais est donc
**enregistré au lancement** (`relay-port.txt`, à côté des journaux) : au démarrage
suivant, la remise en accès direct (`netsh winhttp reset proxy`) n'est proposée que si
WinHTTP pointe sur **ce** port précis et qu'il ne répond plus — un autre proxy local
(agent, antivirus, cache) n'est jamais touché, même en loopback. La décision est écrite
dans le journal, le marqueur est effacé dès que WinHTTP ne pointe plus dessus, et une
fermeture normale restaure tout toute seule.

## Compilation de l'EXE

```powershell
Invoke-PS2EXE -InputFile "Install-Office2024-KMS.ps1" `
              -OutputFile "Install-Office2024-KMS.exe" `
              -NoConsole -RequireAdmin -x64 `
              -Title "Deploiement Office 2024" `
              -Product "InstallOffice" `
              -Description "Installation Office LTSC 2024 et activation KMS" `
              -Company "Lycee Loritz" `
              -Version "8.2.0.0"
```

(`Install-Module ps2exe` si le module est absent.)

Compiler **directement** vers `Install-Office2024-KMS.exe` : PS2EXE reprend le nom du
fichier de sortie dans `OriginalFilename`/`InternalName`, or le contrôle vérifie que ces
champs valent bien `Install-Office2024-KMS.exe` (compiler vers un fichier temporaire puis le
renommer laisse `_b.exe` dans les propriétés).

## Journaux et configuration

- Journal : `C:\InstallOffice2024\deploy_AAAAMMJJ_HHMMSS.log` (bouton « Exporter » dans l'interface)
- Configuration du poste : `%APPDATA%\Office2024Deploy\config.json`
  (mots de passe et clé produit chiffrés en DPAPI, donc liés au compte utilisateur)

## Contrôles automatisés

```powershell
powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1
powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1 -Fast
```

`-Fast` ne joue que la section 1 — syntaxe du script, empreinte du script embarqué et
propriétés de l'EXE — soit quelques secondes : c'est exactement ce que lance le hook avant
chaque commit. Sans commutateur, la suite complète (90 cas) tourne, comme sur GitHub.

Le **dernier contrôle** de la suite relit ce README, y cherche toutes les mentions
« N cas » et exige qu'elles valent le total réellement exécuté : un cas ajouté ou retiré
sans mettre la documentation à jour fait donc échouer le contrôle, localement comme sur
GitHub. C'est ce qui empêche le chiffre annoncé de dériver à nouveau (60, 67, 69…).

Les fins de ligne font partie du contrat : l'EXE embarque le script en CRLF et le contrôle
compare des octets, donc `.gitattributes` impose `*.ps1 text eol=crlf` (et `*.exe binary`).
Un clone avec `core.autocrlf=false` ou un runner Linux obtient malgré tout du CRLF ; si un
script arrive quand même en LF, la suite le dit explicitement — « contenu identique mais fins
de ligne différentes » — au lieu de réclamer une recompilation inutile, et un contrôle dédié
signale l'absence de la déclaration dans `.gitattributes`.

Vérifie d'abord que l'EXE distribué correspond au script : l'empreinte SHA-256 du script
embarqué dans le binaire doit être celle du `.ps1`, sinon le contrôle s'arrête avec un
message **RECOMPILER** (section « Compilation de l'EXE »). Si le binaire n'expose pas le
script en clair (format compressé ou construction différente), l'échec dit que l'empreinte
n'est **pas vérifiable** — sans conseil de recompilation, qui serait alors trompeur.

Les autres propriétés du binaire livré sont contrôlées aussi : version `8.2.0.0`, titre
`Deploiement Office 2024`, produit `InstallOffice`, société `Lycee Loritz` et manifeste
`requireAdministrator`, plus `OriginalFilename` et `InternalName`, qui doivent porter le nom
du fichier livré. Chaque écart nomme la valeur trouvée, celle attendue et l'option
`Invoke-PS2EXE` correspondante.

Rejoue ensuite les cas clés sans rien installer : `SourcePath` imposé au téléchargement et à
l'installation, destination des sources et verdict d'espace disque, codes de sortie ODT
(32, 30066, 17002, 17029), échecs KMS (0xC004F074, 0x80070005, 0xC004F038, réponse
inconnue, `ospp.vbs` absent), dépôt invalide, proxy injoignable et bilan final.

Le proxy est aussi vérifié dans l'autre sens : le script applique les trois couches
(WinHTTP machine, WinINet session, quatre variables d'environnement) **puis** les restaure
à l'identique. Ces contrôles passent par `bg_ProxySink`, le point d'accès système unique des
fonctions proxy : le script de contrôle le remplace par une machine simulée, si bien que
le poste n'est jamais modifié (vérifié en comparant HKCU, l'environnement et `netsh`
avant/après). L'état sauvegardé est aussi vérifié comme une **copie** : une mutation de la
machine après `bg_ApplyProxy` ne doit pas l'altérer, et un échec de `netsh` doit tout de
même aboutir à une restauration complète.

Le relais d'authentification est vérifié de la même façon, contre un proxy amont simulé
sur la loopback : un `CONNECT` muni des bons identifiants doit rendre `200` et laisser
passer les données dans les deux sens, un mauvais mot de passe doit rendre `407`, et
`Stop-ProxyRelay` doit libérer le port d'écoute. Le câblage côté déploiement est
couvert de la même façon : face à un proxy qui répond `407`, le worker produit le
message des identifiants quand la configuration porte `ProxyRelay`, le message
historique (`LocalSystem`) sinon, et dans les deux cas sans appliquer de proxy à la
machine. Le reliquat d'un arrêt brutal est couvert de la même façon : WinHTTP pointé sur
un port loopback fermé est détecté, un relais encore vivant ou un accès direct ne sont
pas signalés, et la remise en accès direct n'est annoncée réussie que si `netsh`
confirme (un échec est remonté tel quel). Le port enregistré fait foi : un port
différent n'est jamais signalé, et sans port enregistré aucune proposition n'est faite.

Comment : un faux `setup.exe` est compilé dans `%TEMP%` (il enregistre ses arguments et
renvoie le code voulu) et un `ospp.vbs` factice remplace celui d'Office ; aucun
téléchargement, aucune modification du poste. Code de sortie `0` si tout passe, `1` sinon.

### Exécution automatique

Le contrôle ne dépend pas de la mémoire de celui qui commite :

- **avant chaque commit**, le hook `.githooks/pre-commit` lance le contrôle **rapide**
  (`-Fast`, section 1 seule, quelques secondes) dès que
  `Install-Office2024-KMS.ps1`, `Install-Office2024-KMS.exe`, `Test-InstallOffice.ps1`
  ou `.gitattributes` est dans le commit, et refuse le commit s'il échoue (les échecs
  sont affichés, pas avalés). `.gitattributes` est surveillé parce que retirer
  `*.ps1 text eol=crlf` change les octets livrés : c'était le trou par lequel la
  régression de fins de ligne passait sans bruit. Un commit ne coûte donc pas la suite entière : rejouer 150 s à chaque fois
  poussait à `--no-verify`, ce qui revenait à ne rien contrôler du tout. Il ne lit pas l'arbre de travail : l'index est exporté par
  `git checkout-index` dans un dossier temporaire et la suite y tourne sur le contenu
  **tel qu'il sera commité** — une modification non indexée ne peut donc ni masquer une
  dérive, ni en créer une. L'export impose `core.autocrlf=false`, comme le job CI : le
  hook voit les octets du dépôt, pas ceux que le réglage local de Git fabriquerait. La section 10 de la suite le prouve en faisant diverger index
  et arbre de travail dans un dépôt jetable, et vérifie que le hook reste bien en mode
  rapide. Il s'active une fois par clone :

  ```powershell
  git config core.hooksPath .githooks
  ```

  Un commit qui ne touche aucun de ces trois fichiers (README, XML, workflow) n'est pas
  ralenti ; `git commit --no-verify` reste possible pour un cas exceptionnel, et se voit
  dans l'historique.
- **sur GitHub**, le workflow `.github/workflows/controle-exe.yml` rejoue la suite
  **complète** (90 cas) sur `windows-latest` à chaque push sur `main` et à chaque pull
  request : un EXE oublié ou compilé autrement fait échouer la vérification côté serveur,
  même si le hook local n'était pas installé — et sans rien imposer à la boucle locale.
- **le contrat de fins de ligne y est vérifié aussi** : un second job refait un checkout
  avec `core.autocrlf=false` et `core.eol=lf`, puis lance le contrôle rapide. Sans la
  déclaration `*.ps1 text eol=crlf`, il recevrait des `.ps1` en LF et échouerait ; le run
  publie les octets réellement reçus, donc un faux « RECOMPILER » ne peut plus passer
  inaperçu sur un poste donné pour la seule raison qu'il a le bon réglage local.

## Limites connues

- Windows uniquement, architecture 64 bits
- Proxy authentifié NTLM/Kerberos non pris en charge : seul Basic passe par le relais local (voir ci-dessus)
- Binaire non signé : SmartScreen peut afficher un avertissement
- L'activation KMS suppose que le poste est sur le réseau du lycée (ou en VPN)
