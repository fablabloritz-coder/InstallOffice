# InstallOffice — déploiement d'Office LTSC 2024 + activation KMS

Outil graphique (PowerShell / WPF, interface XAML) pour installer Microsoft Office LTSC
2024 (`ProPlus2024Volume`, canal `PerpetualVL2024`) puis l'activer par KMS, sur un poste
d'un réseau pédagogique.

L'interface livrée est `Apercu-Maquette.ps1` (fenêtres dans `Maquette/`), compilée dans
`Install-Office2024-KMS.exe`. Le déroulé d'installation n'est pas réécrit : l'EXE extrait
du moteur `Install-Office2024-KMS.ps1`, livré à côté, le scriptblock de déploiement et
l'exécute tel quel. Les deux fichiers voyagent donc ensemble.

## Démarrage rapide

1. Télécharger le ZIP du dépôt (**Code → Download ZIP**) et l'extraire où vous voulez.
2. Lancer `Install-Office2024-KMS.exe` et accepter l'élévation demandée.
3. Renseigner le serveur KMS, choisir le mode, puis « Installer Office 2024 » — ou
   « Activer seulement » sur un poste déjà équipé.

Le dossier extrait contient tout : le programme, ses fenêtres (`Maquette\`), le moteur et
`setup.exe`. Rien d'autre à télécharger.

## Contenu du dépôt

| Fichier | Rôle |
| --- | --- |
| `Apercu-Maquette.ps1` | Interface graphique livrée (WPF/XAML), **script compilé dans l'EXE** |
| `Maquette/*.xaml` | Fenêtres et styles de l'interface, lus au démarrage — à livrer avec l'EXE |
| `Install-Office2024-KMS.exe` | Version compilée (PS2EXE, manifeste `requireAdmin`) |
| `Install-Office2024-KMS.ps1` | Moteur : déroulé de déploiement extrait par l'EXE ; son interface WinForms historique reste exécutable telle quelle |
| `OfficeConfig2024.xml` | Modèle de configuration ODT (référence / XML personnalisé) |
| `Test-InstallOffice.ps1` | Contrôles automatisés hors ligne (163 cas, `-Fast` pour la section 1 seule) |
| `version.txt` | Version de l'outil, source unique : le badge, le README, le script et `Compare-ExeRecompilation.ps1` s'y comparent |
| `Set-Version.ps1` | Vérifie ou écrit cette version partout, et refuse si un emplacement diverge |
| `New-Livraison.ps1` | Construit `Livraison.zip` en liste blanche, avec `SHA256SUMS` |
| `Verifier-Livraison.ps1` | Vérificateur livré dans le ZIP : à lancer sur un poste, sans git |
| `.githooks/pre-commit` | Contrôle rapide de l'index avant chaque commit, refuse l'EXE désynchronisé |
| `.github/workflows/controle-exe.yml` | Suite complète + contrat de fins de ligne (checkout sans conversion) sur GitHub Actions, à chaque push sur `main` et à chaque pull request |
| `setup.exe` | Office Deployment Tool (16.0.18129.20030, signé Microsoft) — versionné pour que le ZIP soit utilisable tel quel |

## Ce qu'on copie sur un poste

Quatre éléments, dans un même dossier :

| Élément | Rôle |
| --- | --- |
| `Install-Office2024-KMS.exe` | le programme (il demande l'élévation par son manifeste) |
| `Maquette\` | les fenêtres XAML : sans ce dossier, l'EXE ne peut pas s'ouvrir |
| `Install-Office2024-KMS.ps1` | le moteur, dont le worker est extrait à chaque déploiement |
| `setup.exe` | Office Deployment Tool (détecté à côté du programme) |

Les quatre sont déjà réunis dans le ZIP du dépôt. `OfficeConfig2024.xml` n'est utile que
pour un XML de référence personnalisé. L'EXE
retrouve ces fichiers **relativement à lui-même** : un binaire PS2EXE n'a pas de
`$PSScriptRoot`, le programme retombe donc sur le dossier du binaire. Déplacer l'EXE seul
sans `Maquette\` ni le `.ps1` le rend inutilisable.

## Prérequis

- Windows 10/11, 64 bits, PowerShell 5.1
- Droits administrateur (l'EXE les demande via son manifeste)
- `setup.exe` (Office Deployment Tool) : fourni dans le dépôt (16.0.18129.20030). Pour le
  mettre à jour : <https://www.microsoft.com/download/details.aspx?id=49117>, remplacer le
  fichier à côté de l'EXE, puis committer le nouveau binaire
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

## Diagnostic sans installation

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Install-Office2024-KMS.ps1 -Diagnostic
```

Lecture seule, sans élévation et sans fenêtre : six points sont mesurés, puis un verdict.
`-ConfigFile <chemin>` relit un autre `config.json` que celui du profil (configuration de
test isolée) ; le mode ne l'écrit jamais.

| Point | Ce qui est mesuré |
| --- | --- |
| Espace disque | Go libres sur `C:`, comparés au minimum de 4 Go |
| Office déjà présent | entrées d'installation Office dans le registre, dont Office LTSC 2024 |
| Droits administrateur | session courante, élevée ou non |
| Proxy joignable | proxy de la configuration : port, puis `CONNECT officecdn.microsoft.com:443` ; sans proxy, accès direct annoncé |
| DNS KMS | résolution du serveur KMS configuré (port 1688) ; l'adresse d'exemple est signalée |
| Horloge | date, fuseau, source de temps Windows et état du service `W32Time` |

Chaque ligne porte `[OK]`, `[ATTENTION]` ou `[BLOQUANT]` ; le code de sortie vaut 1 s'il
reste un point bloquant (espace disque insuffisant), 0 sinon. Rien n'est écrit : ni journal,
ni `config.json`, ni réglage proxy. Lancé sans droits administrateur, il le dit au lieu de
demander une élévation : c'est justement l'un des points mesurés.

## Pourquoi un réglage proxy dédié ?

`setup.exe /download` ne lit **pas** `HTTP_PROXY`. Le téléchargement est effectué par le
service **Microsoft Office Click-to-Run**, qui tourne en `LocalSystem` et utilise la
configuration **WinHTTP de la machine**. Le script :

1. teste réellement le proxy (`CONNECT officecdn.microsoft.com:443`) avant de lancer
   les ~20 minutes de téléchargement (détecte proxy injoignable, destination filtrée, 407) ;
2. applique le proxy à WinHTTP (machine) **et** à WinINet (session), plus les variables
   d'environnement ;
3. **restaure** la configuration d'origine en fin de déploiement.

Un proxy **avec authentification** (407) ne peut pas être satisfait par le service
Click-to-Run, qui tourne en `LocalSystem` et n'a aucun moyen de saisir un mot de passe.
C'est pourquoi les champs « Utilisateur » et « Mot de passe » du proxy existent : le
programme démarre alors un **relais local** (sur `127.0.0.1`, port libre) qui ajoute
l'en-tête `Proxy-Authorization` vers le vrai proxy et laisse passer le tunnel ; Windows est
pointé sur ce relais, qui n'est pour Office qu'un proxy ordinaire. Le relais ne vit que le
temps du déploiement et le mot de passe n'est **jamais** écrit sur disque : il faut le
ressaisir à chaque lancement. Deux modes : **Basic** quand un compte est saisi, et **Negotiate** (NTLM/Kerberos)
quand les champs identifiants restent vides — le relais demande alors `Negotiate` et
laisse Windows SSPI signer avec le compte de la session (voir *Proxy d'établissement :
Basic, ou identifiants Windows*). Si le programme est tué pendant le téléchargement, avant la restauration, le poste
garderait le relais loopback dans **deux** configurations : WinHTTP (machine, ce que lit
Office) et le proxy de session (`HKCU\…\Internet Settings`, `ProxyEnable=1`, ce que lit
le navigateur). Le port du relais est donc **enregistré au lancement**
(`relay-port.txt`, à côté des journaux) : au démarrage suivant, la remise en accès direct
n'est proposée que pour les couches qui pointent sur **ce** port précis et qui ne
répondent plus — `netsh winhttp reset proxy` pour la machine, `ProxyEnable=0` pour la
session. Un autre proxy local (agent, antivirus, cache) n'est jamais touché, même en
loopback, et une couche qui n'est pas en cause n'est pas modifiée. La décision est écrite
dans le journal, le marqueur est effacé dès que plus aucune couche ne pointe dessus, et
une fermeture normale restaure tout toute seule.

## Banc de preuve : proxy authentifié de bout en bout

`Banc-ProxyAuthentifie.ps1` monte la chaîne complète sur le poste, sans rien installer :
un **vrai proxy amont** local (processus à lui, Basic exigé, `407` sans identifiants,
journal par requête), le relais de l'outil devant, et une cible locale atteinte à travers
le tunnel.

```powershell
powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Banc-ProxyAuthentifie.ps1
```

Huit preuves, hors ligne, en une seconde :

- le client parle au relais **sans identifiants**, et le proxy voit `auth=oui user=…` ;
- un mauvais mot de passe remonte un **407** au client, pas un tunnel muet ;
- une requête **HTTP en clair à URL absolue** traverse la chaîne — c'est la forme que le
  CDN Office utilise pour les `.cab` du téléchargement ;
- un proxy qui n'offre que **NTLM** laisse remonter son `407` et son `Proxy-Authenticate`,
  ce qui permet de nommer le schéma au lieu d'accuser le mot de passe ;
- un proxy qui exige **Negotiate** authentifie le relais avec le **compte Windows** de la
  session : le proxy journalise `AUTH DOMAINE\compte`, et le corps `BANC-OK` arrive ;
- le même proxy, sans relais, refuse la requête (`407` + `Negotiate`) : l'authentification
  est bien exigée, pas contournée ;
- une connexion que le client voudrait **réutiliser** est fermée proprement
  (`Connection: close`) : la requête suivante repart sur une connexion neuve, authentifiée
  elle aussi, au lieu de partir sans en-tête et de prendre un 407 ;
- une réponse dont l'en-tête est **tronqué** fait fermer la connexion : le relais ne
  relaie jamais une réponse partielle.

Le journal du proxy est imprimé en regard des résultats, et le rapport JSON (`-Rapport`)
reprend les mêmes lignes.

### Le téléchargement ODT réel

```powershell
... -Odt -WinInet
```

`-SecondesOdt` vaut 300 s par défaut : c'est la borne qui laisse le temps d'un fichier
entier (`.cab` ou `.dat`), puisque la preuve exige un fichier complet ; sur un lien plus
lent, l'augmenter.

Le banc règle alors le proxy de session et les variables d'environnement sur le relais,
lance un vrai `setup.exe /download` borné dans le temps, puis relit le journal : des lignes
`CONNECT officeclient.microsoft.com:443` et `GET http://…/…cab auth=oui` montrent que
Click-to-Run passe par le relais. Sans `-WinInet`, seul l'environnement est réglé. WinHTTP
demande les droits administrateur : sans eux, ODT peut ignorer le relais, et le banc le dit
au lieu de le masquer.

Le rapport va plus loin que « le flux passe ». Pour chaque `.cab` **et** chaque `.dat`
écrit dans les sources — les `stream.x64…dat` portent l'essentiel du téléchargement —, il
compare la taille du fichier au `Content-Length` annoncé par le CDN — ou au total du
`Content-Range` quand ODT télécharge par plages, ce qui est le cas des gros fichiers. Un
fichier n'est déclaré `Complet` que si les deux coïncident, et `Odt.Ok` exige au moins un
fichier complet : un flux qui traverse le relais sans qu'aucun fichier n'arrive entier ne
passe plus pour une réussite. Un fichier servi en HTTPS (donc dans un tunnel) n'annonce pas
sa taille : il est marqué « inconnue », jamais « complet ».

Ce que ce banc a appris, et qui est corrigé : le relais ne traitait que `CONNECT`. Or ODT
télécharge les `.cab` en **HTTP clair**, sous forme de requête à URL absolue ; le relais les
transformait en `CONNECT` et le proxy amont répondait `400`. Il transmet désormais ces
requêtes telles quelles, en ajoutant l'en-tête d'authentification.

Second défaut trouvé de la même façon : l'en-tête n'était ajouté qu'à la **première**
requête d'une connexion. Si le client réutilisait la connexion, la deuxième partait sans
`Proxy-Authorization` et prenait un `407`. Le relais traite maintenant **une requête par
connexion amont** et annonce `Connection: close` aux deux bouts : le client rouvre, et la
connexion neuve est authentifiée. Un parseur HTTP pour réinjecter l'en-tête à chaque
message aurait coûté bien plus cher, pour un trafic fait de `GET` de fichiers.

## Proxy d'établissement : Basic, ou identifiants Windows

| Ce que le proxy exige | Ce que l'outil fait | Statut |
| --- | --- | --- |
| Aucune authentification | WinHTTP et WinINet pointés sur le proxy | pris en charge |
| **Basic** | relais local qui ajoute `Proxy-Authorization` | pris en charge, prouvé par le banc |
| **NTLM / Kerberos** (`Negotiate`) | relais local qui delegate les jetons a SSPI | pris en charge, prouvé par le banc |

Le second mode ne calcule aucun jeton : le relais envoie la requete sans en-tete, lit le
`407` et son `Proxy-Authenticate: Negotiate`, puis confie le Type1 et le Type2 a
`NegotiateStream` en memoire — Windows signe avec le compte de la session et rend le Type3.
Le proxy voit donc le compte ouvre, sans qu'aucun mot de passe ne soit saisi ni enregistre.
Le handshake tient sur la connexion amont, en trois requetes successives, comme un proxy
NTLM l'exige.

Le banc le prouve : `Banc-ProxyAuthentifie.ps1` monte un proxy qui refuse une requete
anonyme (`407` + `Negotiate`) et n'accepte que le Type3 ; il verifie que l'identite
annoncee par le proxy est celle du compte qui execute le banc, et que le corps recu est
`BANC-OK`.

Ce qui reste hors de portee : un proxy qui n'annonce **que** `NTLM` sans `Negotiate`, et
les requetes a corps (le relais ne traite que `GET`, `HEAD` et `CONNECT`, ce qui couvre le
telechargement). Dans ces cas, le mode « Depot de sources » reste l'issue recommandee.

## Déploiement sans surveillance (GPO, PsExec)

L'outil sait aussi tourner **sans fenêtre**, pour équiper un parc depuis un script. Le
déroulement est celui de l'écran — le même worker du moteur — mais les réglages viennent
d'un fichier JSON et le résultat part dans un rapport.

```powershell
Install-Office2024-KMS.exe -Silent -Config .\poste.json -Rapport .\rapport_PC042.json
```

`-Rapport` est facultatif : sans lui, le rapport s'écrit dans le dossier des journaux,
sous `rapport_<poste>_<date>.json`, avec le journal du déploiement à côté (même nom, en
`.log`). Rien n'est écrit dans `config.json` : un déploiement scripté ne modifie pas les
réglages du poste.

### Le fichier de configuration

Mêmes clés que `config.json`, plus `ActivateOnly` pour rejouer une activation seule :

| Clé | Rôle |
| --- | --- |
| `SetupExePath` | `setup.exe` (par défaut, à côté de l'EXE) |
| `KmsServer` | `hote` ou `hote:port` |
| `KmsSecondaryServer` | serveur de secours |
| `Language` | langue ODT (`fr-fr` par défaut) |
| `RemoveOld` | désinstaller l'ancienne version |
| `AutoInstall` / `AutoActivate` | installer / activer |
| `ActivateOnly` | activation seule, sans installation |
| `ExcludedApps` | identifiants ODT exclus |
| `UseCustomXML` / `CustomXMLPath` | XML fourni au lieu du XML généré |
| `DepotEnabled` / `DepotPath` | sources locales au lieu du téléchargement |
| `ProxyEnabled`, `ProxyHost`, `ProxyPort`, `ProxyBypass`, `ProxyUser` | proxy |
| `ApplyProductKey` | appliquer une clé avant l'activation |

### Les secrets ne sont jamais dans le fichier

Un mot de passe ou une clé en clair est **refusé** (code 10) : ce fichier part souvent sur
un partage. Deux formes sont acceptées, une par secret :

- `ProxyPasswordEnv` / `ProductKeyEnv` : nom d'une variable d'environnement, que la GPO
  renseigne avant l'appel ;
- `ProxyPasswordEnc` / `ProductKeyEnc` : valeur chiffrée en DPAPI sur **ce** poste.

### Codes de sortie

| Code | Sens |
| --- | --- |
| 0 | déploiement réussi |
| 10 | configuration refusée (fichier, secret, pré-vol) |
| 20 | sources (téléchargement ou copie du dépôt) |
| 30 | installation, y compris la désinstallation de l'ancienne version |
| 40 | activation KMS |
| 50 | relais proxy impossible |
| 60 | erreur hors étapes |

Un refus écrit quand même son rapport : un poste qui n'a pas démarré laisse une trace
lisible, avec ses cinq étapes en attente.

## Compilation de l'EXE

```powershell
Invoke-PS2EXE -InputFile "Apercu-Maquette.ps1" `
              -OutputFile "Install-Office2024-KMS.exe" `
              -NoConsole -RequireAdmin -x64 `
              -Title "Deploiement Office 2024" `
              -Product "InstallOffice" `
              -Description "Installation Office LTSC 2024 et activation KMS" `
              -Company "Lycee Loritz" `
              -Version "8.5.1.0"
```

(`Install-Module ps2exe` si le module est absent.)

### Version : `version.txt` est la source unique

Le badge de la fenêtre, le README, `Compare-ExeRecompilation.ps1` et la constante du
script ne peuvent pas lire un fichier à l'exécution (XAML statique, texte, commande
PS2EXE). `version.txt` porte donc la version, et `Set-Version.ps1` tient les copies
alignées :

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Set-Version.ps1            # vérifie
powershell -NoProfile -ExecutionPolicy Bypass -File .\Set-Version.ps1 -Version <N.N.N.N>
```

Sans argument, le script vérifie que tous les emplacements valent `version.txt` et sort en
erreur en nommant le fichier, la ligne et la valeur trouvée. Avec `-Version`, il refait
cette vérification **avant** d'écrire : si un seul emplacement diverge, rien n'est écrit.
Le test lit `version.txt` — il ne doit contenir aucune version en dur, et le script refuse
s'il en trouve une. Après un bump, l'EXE est à recompiler et les captures à régénérer (le
script rappelle les deux commandes).

La CI recompile ce script et compare le résultat au binaire committé : le script embarqué
doit être identique, et les propriétés du fichier (titre, produit, société, version,
manifeste `requireAdministrator`) doivent être celles de la commande ci-dessus. Les octets
des deux EXE ne sont jamais comparés — PS2EXE n'est pas reproductible. Un binaire compilé
depuis un autre script, ou avec d'autres options, échoue donc côté serveur même si son
empreinte interne est cohérente.

Le binaire garde le nom `Install-Office2024-KMS.exe` : les raccourcis et la documentation
restent valables, et l'interface WinForms historique peut être retirée sans changer les
postes. Le script embarqué est `Apercu-Maquette.ps1` — c'est son empreinte qui est
contrôlée — tandis que le moteur n'est **pas** embarqué : il est lu à l'exécution, à côté
de l'EXE.

Compiler **directement** vers `Install-Office2024-KMS.exe` : PS2EXE reprend le nom du
fichier de sortie dans `OriginalFilename`/`InternalName`, or le contrôle vérifie que ces
champs valent bien `Install-Office2024-KMS.exe` (compiler vers un fichier temporaire puis le
renommer laisse `_b.exe` dans les propriétés).

## Livraison : le ZIP, son empreinte, son vérificateur

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\New-Livraison.ps1
```

Construit `Livraison.zip` à partir d'une **liste blanche** : les quatre éléments livrés,
`README.md`, `Verifier-Livraison.ps1` et un `SHA256SUMS` calculé au passage. Rien d'autre
n'entre — ni `config.json` (il livrerait le serveur KMS et les chemins du poste), ni
journal, ni sources `Office\`, ni la suite de tests, qui a besoin du dépôt. La suite
construit le ZIP dans un dossier de travail et le vérifie : éléments présents, aucun
fichier local, `SHA256SUMS` exact, puis le vérificateur rejoué sur l'archive extraite.

Sur un poste sans git, après extraction :

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Verifier-Livraison.ps1
```

Il recompte les empreintes de `SHA256SUMS`, vérifie les cinq XAML, la signature Microsoft
de `setup.exe`, le nom et le produit de l'EXE, et l'absence de `config.json` ou de journal
dans le dossier. Rien n'est installé ni modifié ; code de sortie 0 si la livraison est
conforme.

Empreinte du binaire livré, publiée ici pour être comparée sans git :

```
SHA256 Install-Office2024-KMS.exe = F3366F64DCB4E16353EDD6D44615D0E3127872DC6A361513ABB0778D8946473E
```

La suite compare cette ligne au fichier livré : une recompilation sans mise à jour du
README fait échouer le contrôle, et `SHA256SUMS` la reprend dans le ZIP.

### Publication automatique (tag `v*`)

Pousser un tag `v*` déclenche `.github/workflows/publication-livraison.yml` :

1. le tag doit correspondre à `version.txt` (`v8.5.1` pour `8.5.1.0`), sinon le job
   s'arrête avant de construire quoi que ce soit ;
2. avant toute construction, `Test-InstallOffice.ps1 -Fast` vérifie que l'EXE du tag
   embarque bien le script tagué (empreinte et version), puis le workflow compare
   l'empreinte du binaire à la ligne `SHA256 Install-Office2024-KMS.exe` du README :
   un binaire désynchronisé refuse la publication ;
3. `New-Livraison.ps1` construit `Livraison.zip` depuis le commit tagué ;
4. l'archive est extraite dans un dossier de travail et le `Verifier-Livraison.ps1`
   qu'elle contient est rejoué dessus, comme sur un poste sans git ;
5. la taille et le SHA-256 du ZIP sont calculés, puis `Livraison.zip` est attaché à la
   release GitHub du tag, avec ces valeurs dans les notes.

Chaque étape s'arrête sur un code non nul : tag incohérent, EXE désynchronisé, ZIP
incomplet ou refusé, empreinte absente. Aucune release n'est créée dans ces cas. Un
re-run du même tag remplace l'asset (`--clobber`) au lieu d'échouer sur une release
existante. Le jeton du dépôt suffit : le workflow demande `contents: write`.

**Répétition avant un tag** : lancer le workflow manuellement (Actions → *Publication
de la livraison* → *Run workflow*) rejoue la construction et toutes les vérifications
sur la branche choisie — `Test-InstallOffice.ps1 -Fast`, empreinte du README,
`Livraison.zip`, extraction et `Verifier-Livraison.ps1`, taille et SHA-256 — mais
saute le contrôle de tag et la publication. Aucune release n'est créée : le journal
donne la taille et le SHA-256 du ZIP construit, ce qui valide le workflow avant de
poser le vrai tag.

## Journaux et configuration

- Journal : `C:\InstallOffice2024\deploy_AAAAMMJJ_HHMMSS.log` (bouton « Exporter » dans l'interface)
- Configuration du poste : `%APPDATA%\Office2024Deploy\config.json`
  (mots de passe et clé produit chiffrés en DPAPI, donc liés au compte utilisateur)

## Contrôles automatisés

```powershell
powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1
powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1 -Fast
```

`-Fast` ne joue que la section 1 — syntaxe de l'interface et du moteur, empreinte du script
embarqué, propriétés de l'EXE, présence des fichiers livrés avec lui et interface ouverte
hors écran avec chaque nom résolu — soit quelques secondes : c'est exactement ce que lance
le hook avant chaque commit. Sans commutateur, la
suite complète (163 cas) tourne, comme sur GitHub.

Le **dernier contrôle** de la suite relit ce README, y cherche toutes les mentions
« N cas » et exige qu'elles valent le total réellement exécuté : un cas ajouté ou retiré
sans mettre la documentation à jour fait donc échouer le contrôle, localement comme sur
GitHub. C'est ce qui empêche le chiffre annoncé de dériver à nouveau (60, 67, 69…).

Les quatre captures de `Maquette\` sont régénérées par la suite (hors mode `-Fast`) et
comparées au pixel aux PNG livrés. Cette comparaison dépendant des polices du poste,
`Maquette\captures.txt` enregistre en plus l'empreinte des cinq XAML et du script : une
référence périmée est détectée même là où le rendu diffère, et le pixel ne tranche que sur
une pile de polices identique à celle des références. Après régénération des PNG :
`Test-InstallOffice.ps1 -EnregistrerCaptures`.

Et l'interface est vérifiée **à l'exécution** : l'EXE livré (ou, sans droits administrateur,
le script qu'il embarque, le binaire portant un manifeste `requireAdministrator`) est lancé
avec un commutateur caché, `-VerifierInterface`, qui charge les cinq XAML, ouvre les
fenêtres hors écran et résout chaque nom que le script demande à la fenêtre. Un nom présent
dans le XAML mais placé dans un modèle ou un dictionnaire de ressources ne s'y résout pas :
ni l'empreinte des sources ni le contrôle statique des noms ne peuvent le voir.

Avant livraison, la suite exporte le dépôt (`git archive HEAD`, l'archive que sert le
bouton **Code → Download ZIP**), vérifie qu'elle contient `Maquette\`, le moteur et
`setup.exe`, puis y rejoue la section 1 : un ZIP incomplet est signalé avant d'être
distribué. Dans une copie déjà extraite du ZIP (pas de `.git`), l'export n'est pas possible :
la suite vérifie alors la présence des fichiers livrés et le dit, au lieu d'échouer sur
l'absence de dépôt.

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

Les autres propriétés du binaire livré sont contrôlées aussi : version `8.5.1.0`, titre
`Deploiement Office 2024`, produit `InstallOffice`, société `Lycee Loritz` et manifeste
`requireAdministrator`, plus `OriginalFilename` et `InternalName`, qui doivent porter le nom
du fichier livré. Chaque écart nomme la valeur trouvée, celle attendue et l'option
`Invoke-PS2EXE` correspondante. La version de l'Office Deployment Tool annoncée dans ce
README est confrontée de la même façon à celle du `setup.exe` livré. Et tout autre nombre à
quatre parties du README doit être l'une de ces deux versions ou une adresse de bouclage
(`127.0.0.0/8`) : rien de non confronté à un fichier livré ne reste dans la prose.

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
`Stop-ProxyRelay` doit libérer le port d'écoute. Les tunnels ouverts chez le proxy amont
doivent aussi égaler le nombre de `CONNECT` que le harnais a vus aboutir — un tunnel
ouvert sans réponse au client, ou un tunnel doublé, se voit. Cette égalité est prouvée
à chaque exécution par une **variante fautive du relais**, compilée par la suite elle-même à partir du code réel plus une seconde connexion amont authentifiée par
`CONNECT` client : le contrôle exige que l'égalité échoue, sans réseau réel ni total en
dur. Le câblage côté déploiement est
couvert de la même façon : face à un proxy qui répond `407`, le worker produit le
message des identifiants quand la configuration porte `ProxyRelay`, le message
historique (`LocalSystem`) sinon, et dans les deux cas sans appliquer de proxy à la
machine. Le reliquat d'un arrêt brutal est couvert de la même façon : WinHTTP pointé sur
un port loopback fermé est détecté, un relais encore vivant ou un accès direct ne sont
pas signalés, et la remise en accès direct n'est annoncée réussie que si `netsh`
confirme (un échec est remonté tel quel). Le port enregistré fait foi : un port
différent n'est jamais signalé, et sans port enregistré aucune proposition n'est faite.
Le proxy de session est couvert séparément : détecté seul ou avec WinHTTP,
`ProxyEnable=0` avec ce port ou un port différent ignoré, et la remise en direct ne
touche que les couches effectivement en cause. Un dernier cas interroge le poste **réel**,
en lecture seule (`netsh winhttp show proxy` et la clé `HKCU`), et confronte le résultat à
une lecture indépendante : une erreur dans la lecture réelle ne peut donc pas passer
inaperçue derrière les sinks simulés, et le contrôle vérifie qu'aucune écriture n'a eu
lieu. Le bloc `Add_Shown` lui-même est exercé : extrait tel quel, il est rejoué avec le
test, la question et la remise en direct remplacés, pour vérifier que le démarrage teste
le reliquat puis propose la remise en direct, ne touche à rien si l'utilisateur refuse,
et efface le marqueur quand il est périmé.

Comment : un faux `setup.exe` est compilé dans `%TEMP%` (il enregistre ses arguments et
renvoie le code voulu) et un `ospp.vbs` factice remplace celui d'Office ; aucun
téléchargement, aucune modification du poste. Code de sortie `0` si tout passe, `1` sinon.

### Exécution automatique

Le contrôle ne dépend pas de la mémoire de celui qui commite :

- **avant chaque commit**, le hook `.githooks/pre-commit` lance le contrôle **rapide**
  (`-Fast`, section 1 seule, quelques secondes) dès que
  `Apercu-Maquette.ps1`, `Install-Office2024-KMS.ps1`, `Install-Office2024-KMS.exe`,
  `Test-InstallOffice.ps1` ou `.gitattributes` est dans le commit, et refuse le commit
  s'il échoue (les échecs
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

  Un commit qui ne touche aucun de ces fichiers (README, XML, workflow) n'est pas
  ralenti ; `git commit --no-verify` reste possible pour un cas exceptionnel, et se voit
  dans l'historique.
- **sur GitHub**, le workflow `.github/workflows/controle-exe.yml` rejoue la suite
  **complète** (163 cas) sur `windows-latest` à chaque push sur `main` et à chaque pull
  request : un EXE oublié ou compilé autrement fait échouer la vérification côté serveur,
  même si le hook local n'était pas installé — et sans rien imposer à la boucle locale.
- **et la recompilation y est vérifiée aussi** : un job à part installe `ps2exe`,
  recompile `Apercu-Maquette.ps1` avec la commande documentée et compare au binaire
  committé le script embarqué (SHA-256) et les propriétés du fichier — jamais les octets,
  PS2EXE n'est pas reproductible. Un EXE compilé depuis un autre script, ou avec d'autres
  options, échoue donc côté serveur même quand son empreinte interne est cohérente.
- **le contrat de fins de ligne y est vérifié aussi** : un second job refait un checkout
  avec `core.autocrlf=false` et `core.eol=lf`, puis lance le contrôle rapide. Sans la
  déclaration `*.ps1 text eol=crlf`, il recevrait des `.ps1` en LF et échouerait ; le run
  publie les octets réellement reçus, donc un faux « RECOMPILER » ne peut plus passer
  inaperçu sur un poste donné pour la seule raison qu'il a le bon réglage local.

## Travailler depuis un autre poste

Le dépôt contient tout ce qu'il faut pour reprendre le travail sur un autre PC : sources,
fenêtres XAML, captures et leurs empreintes, EXE, moteur, `setup.exe` et la suite de
contrôles. Après un clone (ou l'extraction du ZIP), deux réglages sont propres à la machine
et ne voyagent pas avec le dépôt :

```powershell
git config core.hooksPath .githooks        # hook de pré-commit, sinon jamais lancé
Install-Module ps2exe -Scope CurrentUser   # seulement pour recompiler l'EXE
powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1 -Fast
```

`config.json` (réglages du poste) et `Office/` (sources Office téléchargées) restent hors du
dépôt : leur absence est normale. Sur un poste dont les polices diffèrent de celles qui ont
produit les captures, la suite signale les PNG différents sans échouer ; l'empreinte des
sources, elle, reste décisive.

## Limites connues

- Windows uniquement, architecture 64 bits
- Proxy NTLM/Kerberos : pris en charge par le schéma `Negotiate` ; un proxy qui n'annonce que `NTLM` brut reste hors de portée (voir ci-dessus)
- Binaire non signé : SmartScreen peut afficher un avertissement
- L'EXE n'est pas autonome : `Maquette\` et `Install-Office2024-KMS.ps1` doivent rester
  à côté de lui (voir « Ce qu'on copie sur un poste »)
- L'activation KMS suppose que le poste est sur le réseau du lycée (ou en VPN)
