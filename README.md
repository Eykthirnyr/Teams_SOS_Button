# Alerte Salariés

Application Windows legere pour declencher une alerte discrete depuis la barre des taches, notifier Teams via webhook admin, puis enregistrer le micro local pendant 15 minutes.

## Resume

- L'utilisateur clique sur l'icone `Alerte Salaries`.
- L'application envoie une alerte Teams vers le webhook defini par l'administration.
- L'application enregistre le micro local pendant 15 minutes.
- Le fichier audio est depose sur le Bureau du poste.
- Un second clic pendant l'enregistrement est ignore.

<img width="816" height="638" alt="image" src="https://github.com/user-attachments/assets/52d52725-cc3a-40ce-a374-67470dee1b56" />


## Fonctionnement utilisateur

- Au premier lancement, une fenetre de configuration s'ouvre.
- L'utilisateur renseigne son nom et son bureau habituel.
- Les parametres utilisateur sont sauvegardes dans `%APPDATA%\AlerteSalaries\user.config.json`.
- Le lien Teams est fourni par l'administration dans `%ProgramData%\AlerteSalaries\AlerteSalaries.Admin.ini`.
- Aux lancements suivants, aucun ecran ne s'ouvre.
- Un clic sur l'icone:
  - envoie une alerte dans Teams avec l'identite, le poste, le bureau et l'heure,
  - lance un enregistrement audio local de 15 minutes,
  - sauvegarde le fichier sur le Bureau.
- Pendant l'enregistrement, tout nouveau clic est ignore pour eviter les doubles alertes et les doubles captures.

## Configuration admin Teams

Le lien Teams n'est pas saisi par l'utilisateur final. Il doit etre fourni par l'administration dans un fichier ini separe:

L'application lit ce lien en priorite dans `%ProgramData%\AlerteSalaries\AlerteSalaries.Admin.ini`.

Exemple:

```ini
[Teams]
WebhookUrl=https://votre-url-workflow-ou-power-automate
```

## Installation et build

1. Executer `Build-AlerteSalaries.ps1` pour compiler l'application.
2. Fournir un fichier `AlerteSalaries.Admin.ini` dans le dossier de deploiement.
3. Executer `Install-AlerteSalaries.ps1` pour l'installation machine-wide.
4. Ou executer `Build-Msi.ps1` si WiX v4 est disponible pour generer un MSI.
5. Un dossier d'installation sera cree dans `%ProgramFiles%\AlerteSalaries`.
6. La configuration admin sera copiee dans `%ProgramData%\AlerteSalaries`.
7. Des raccourcis communs seront crees pour tous les utilisateurs:
   - `Alerte Salaries`
8. Epingler le raccourci `Alerte Salaries` dans la barre des taches, ou appliquer `deploy\TaskbarLayoutModification.xml` via GPO/Intune.

## Commandes de build

### Recompiler l'application

```powershell
powershell.exe -ExecutionPolicy Bypass -File Build-AlerteSalaries.ps1
```

### Regenerer le MSI

```powershell
powershell.exe -ExecutionPolicy Bypass -File Build-Msi.ps1
```

## Deploiement silencieux

Le chemin recommande  sans PowerShell:

- utilisez `Deploy-AlerteSalaries-Silent.cmd`

Ce script batch natif Windows:

- copie l'application dans `%ProgramFiles%\AlerteSalaries`
- copie le fichier admin ini dans `%ProgramData%\AlerteSalaries`
- cree les raccourcis communs via `wscript.exe`

Le script PowerShell de deploiement reste disponible seulement en secours.

## MSI et parametres de deploiement

Le MSI supporte des proprietes publiques pour le deploiement silencieux.

### Parametres MSI disponibles

- `TEAMSWEBHOOKURL`
  - URL du webhook Teams ou Power Automate a utiliser pour toutes les alertes.
  - Le MSI ecrit cette valeur dans `AlerteSalaries.Admin.ini` installe avec l'application.
- `INSTALLDESKTOPSHORTCUT`
  - `1` par defaut.
  - Mettre `0` pour ne pas creer le raccourci Bureau.

### Exemple d'installation MSI manuelle

```powershell
msiexec /i AlerteSalaries.msi TEAMSWEBHOOKURL="https://votre-url" INSTALLDESKTOPSHORTCUT=1 /qn /norestart
```

### Exemple GPO

Si vous publiez le MSI depuis un partage reseau, la ligne equivalente est de type:

```cmd
msiexec /i "\\serveur\share\AlerteSalaries.msi" TEAMSWEBHOOKURL="https://votre-url" INSTALLDESKTOPSHORTCUT=1 /qn /norestart
```

### Commande exacte avec le webhook actuel

Commande a lancer depuis PowerShell:

```powershell
msiexec /i AlerteSalaries.msi TEAMSWEBHOOKURL="[webhookurl](https://votre-url)" INSTALLDESKTOPSHORTCUT=1 /qn /norestart /L*V "$env:TEMP\AlerteSalaries-MSI.log"
```

Note:

- pour un vrai deploiement GPO ordinateur, il est souvent preferable d'utiliser l'affectation de package MSI standard;
- si vous avez besoin de passer des proprietes MSI personnalisees, un script de demarrage ordinateur reste souvent plus souple.
- l'application lit d'abord `%ProgramData%\AlerteSalaries\AlerteSalaries.Admin.ini`, puis en secours `AlerteSalaries.Admin.ini` a cote de l'executable.

## Emplacements utilises

- Application: `%ProgramFiles%\AlerteSalaries`
- Config admin: `%ProgramData%\AlerteSalaries\AlerteSalaries.Admin.ini`
- Config utilisateur: `%APPDATA%\AlerteSalaries\user.config.json`
- Logs utilisateur: `%APPDATA%\AlerteSalaries\app.log`

## Verifications apres installation

### Verifier le fichier admin installe

```powershell
Get-Content "C:\Program Files (x86)\AlerteSalaries\AlerteSalaries.Admin.ini"
```

Si l'application est installee en 64 bits, verifier aussi:

```powershell
Get-Content "C:\Program Files\AlerteSalaries\AlerteSalaries.Admin.ini"
```

### Verifier le raccourci menu Demarrer

```powershell
Get-ChildItem "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Alerte Salaries"
```

### Verifier le raccourci Bureau de l'utilisateur courant

```powershell
Get-ChildItem "$HOME\Desktop"
```

### Verifier le log MSI

```powershell
Get-Content "$env:TEMP\AlerteSalaries-MSI.log" | Select-String "TEAMSWEBHOOKURL","AlerteSalaries.Admin.ini","DesktopShortcut","ProgramMenuShortcut"
```

### Verifier le log applicatif utilisateur

```powershell
Get-Content "C:\Users\c.ghanem\AppData\Roaming\AlerteSalaries\app.log" -Tail 50
```

## Deploiement taskbar

L'epinglage automatique supporte par Windows doit passer par une strategie de type XML/GPO/Intune.

Fichier fourni:

- `deploy\TaskbarLayoutModification.xml`

Ce fichier epingle le raccourci commun du menu Demarrer:

- `%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Alerte Salaries\Alerte Salaries.lnk`

## Audio

- L'application enregistre par defaut en `wav`.
- Si `ffmpeg.exe` est disponible dans le `PATH` Windows ou dans le dossier d'installation, le fichier sera converti automatiquement en `mp3`.

## Limites connues

- Le premier clic ouvrira l'ecran de configuration si l'application n'a pas encore ete parametree.
- L'epinglage automatique dans la barre des taches doit passer par la methode supportee par Windows, typiquement via `TaskbarLayoutModification.xml` applique par GPO ou Intune.
- Le premier lancement sert a configurer l'application. L'alerte sera envoyee a partir du clic suivant.

## Technique

### Stack

- Application: C# WinForms compilee en `.exe`
- Packaging MSI: WiX
- Deploiement silencieux alternatif: `.cmd` + `wscript.exe`

### Comportement technique

- Mutex local pour bloquer les doubles clics pendant l'alerte
- Envoi Teams en Adaptive Card
- Lecture de la config admin en ini
- Lecture et ecriture de la config utilisateur en JSON
- Enregistrement audio local via `winmm.dll`
