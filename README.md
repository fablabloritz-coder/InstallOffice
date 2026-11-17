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
| `Test-InstallOffice.ps1` | Contrôles automatisés hors ligne (53 cas) |
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
c'est précisément le cas d'usage du mode « Dépôt de sources ».

## Compilation de l'EXE

```powershell
Invoke-PS2EXE -InputFile "Install-Office2024-KMS.ps1" `
              -OutputFile "Install-Office2024-KMS.exe" `
              -NoConsole -RequireAdmin -x64 `
              -Title "Deploiement Office 2024" `
              -Product "InstallOffice" `
              -Description "Installation Office LTSC 2024 et activation KMS" `
              -Company "Lycee Loritz" `
              -Version "8.1.0.0"
```

(`Install-Module ps2exe` si le module est absent.)

## Journaux et configuration

- Journal : `C:\InstallOffice2024\deploy_AAAAMMJJ_HHMMSS.log` (bouton « Exporter » dans l'interface)
- Configuration du poste : `%APPDATA%\Office2024Deploy\config.json`
  (mots de passe et clé produit chiffrés en DPAPI, donc liés au compte utilisateur)

## Contrôles automatisés

```powershell
powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1
```

Rejoue les cas clés sans rien installer : `SourcePath` imposé au téléchargement et à
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

Comment : un faux `setup.exe` est compilé dans `%TEMP%` (il enregistre ses arguments et
renvoie le code voulu) et un `ospp.vbs` factice remplace celui d'Office ; aucun
téléchargement, aucune modification du poste. Code de sortie `0` si tout passe, `1` sinon.

## Limites connues

- Windows uniquement, architecture 64 bits
- Pas de proxy authentifié (voir ci-dessus)
- Binaire non signé : SmartScreen peut afficher un avertissement
- L'activation KMS suppose que le poste est sur le réseau du lycée (ou en VPN)
