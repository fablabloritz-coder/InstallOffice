# InstallOffice — repères pour travailler sur ce dépôt

Outil Windows (PowerShell 5.1 + WPF) qui installe Office LTSC 2024 puis l'active par KMS.
`README.md` décrit le produit ; ce fichier ne dit que ce qu'il faut savoir pour modifier le
dépôt sans casser la livraison.

## Mise en route d'un clone (une fois par poste)

```powershell
git clone https://github.com/fablabloritz-coder/InstallOffice.git
cd InstallOffice
git config core.hooksPath .githooks        # le hook n'est pas porté par le clone
Install-Module ps2exe -Scope CurrentUser   # seulement pour recompiler l'EXE
powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1 -Fast
```

`config.json` et `Office/` sont volontairement hors du dépôt (réglages du poste, sources
Office de plusieurs Go) : leur absence est normale.

## Règles que les contrôles font respecter

- **Fins de ligne** : `.gitattributes` impose CRLF aux `.ps1` et LF au hook et au workflow.
  L'EXE embarque `Apercu-Maquette.ps1` et le contrôle compare des octets : ne jamais
  committer un `.ps1` converti.
- **EXE aligné** : après toute modification de `Apercu-Maquette.ps1`, recompiler
  `Install-Office2024-KMS.exe` (commande dans le README, sortie écrite directement sous le
  nom livré). Sinon la suite dit RECOMPILER et le hook refuse le commit. Le moteur
  `Install-Office2024-KMS.ps1` n'est pas embarqué : le modifier ne demande rien de tel.
- **Total annoncé** : la suite compte ses cas et relit le README ; tout « N cas » du README
  doit valoir le total réellement exécuté. Un cas ajouté ou retiré sans mettre à jour les
  mentions du README fait échouer le dernier contrôle.
- **Captures** : les quatre PNG de `Maquette\` sont comparés au pixel, mais seulement là où
  les polices du poste sont celles des références ; l'empreinte des cinq XAML et du script
  (`Maquette\captures.txt`) tranche partout. Après régénération des quatre captures :
  `Test-InstallOffice.ps1 -EnregistrerCaptures`.
- **Une preuve par comportement** : un correctif se prouve en faisant échouer la garde
  concernée, puis en restaurant les fichiers à l'identique (SHA-256 comparé) ; les messages
  d'échec nomment le fichier, la ligne et la valeur trouvée.

## Vérifier avant de pousser

```powershell
powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\Test-InstallOffice.ps1
```

113 cas, environ 80 s ; le code de sortie vaut 0 seulement si tout passe. Sur un poste dont
les polices diffèrent de celles des références, la suite signale les captures différentes
sans échouer ; c'est attendu, pas une régression.
