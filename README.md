# Office Unattended Installation and Activation

Scripts pour installer et activer Microsoft Office de manière automatique (unattended).

Basé sur [ohook par asdcorp](https://github.com/asdcorp/ohook) pour l'activation.

## Scripts disponibles

| Script | Description |
|--------|-------------|
| `Office-Install-Activate.cmd` | Installation automatique complète (configurable) |
| `Office-Menu-Install.cmd` | Version interactive avec menu |
| `Ohook-Activate.cmd` | Activation uniquement (version détaillée) |
| `Ohook-Activate-Silent.cmd` | Activation uniquement (version silencieuse) |
| `Disable-Telemetry.cmd` | Désactive télémétrie Windows/Office/Edge + recommandations |
| `config-office365.xml` | Configuration pour Office 365 |
| `config-office2021.xml` | Configuration pour Office 2021 LTSC |

## Utilisation rapide

### Installation automatique complète

1. **Clic droit** sur `Office-Install-Activate.cmd`
2. Sélectionner **"Exécuter en tant qu'administrateur"**
3. Attendre la fin de l'installation et activation

### Version avec menu

1. **Clic droit** sur `Office-Menu-Install.cmd`
2. Sélectionner **"Exécuter en tant qu'administrateur"**
3. Choisir l'option désirée dans le menu

### Activation uniquement

Si Office est déjà installé :

```batch
:: Version détaillée avec logs
Ohook-Activate.cmd

:: Version silencieuse
Ohook-Activate-Silent.cmd

:: Version silencieuse avec logs
Ohook-Activate-Silent.cmd /log
```

## Désactivation de la télémétrie

Le script `Disable-Telemetry.cmd` désactive :

**Télémétrie :**
- Windows (DiagTrack, données de diagnostic, publicité)
- Microsoft Office (feedback, données client, LinkedIn)
- Microsoft Edge (métriques, expériences, SmartScreen)

**Recommandations et widgets :**
- Widgets Windows 11
- Suggestions dans le menu Démarrer
- Recommandations dans l'Explorateur
- Recherche Bing dans la barre des tâches
- Copilot
- Chat Teams dans la barre des tâches

### URLs à configurer

Dans les scripts d'installation, configurez ces variables :

```batch
set "OHOOK_SCRIPT_URL=https://raw.githubusercontent.com/VOTRE_USER/VOTRE_REPO/main/Ohook-Activate.cmd"
set "TELEMETRY_SCRIPT_URL=https://raw.githubusercontent.com/VOTRE_USER/VOTRE_REPO/main/Disable-Telemetry.cmd"
```

## Comment fonctionne Ohook

Ohook fonctionne en plaçant un fichier `sppc.dll` personnalisé dans le dossier Office. Ce fichier intercepte les appels de vérification d'activation et retourne toujours que Office est activé.

**Avantages :**
- Ne modifie pas les fichiers système Windows
- Survit aux mises à jour Office
- Ne nécessite pas de serveur KMS

**Fichiers utilisés :**
- `sppc64.dll` (64-bit) - SHA256: `393a1fa26deb3663854e41f2b687c188a9eacd87b23f17ea09422c4715cb5a9f`
- `sppc32.dll` (32-bit) - SHA256: `09865ea5993215965e8f27a74b8a41d15fd0f60f5f404cb7a8b3c7757acdab02`

## Configuration

### Modifier la version d'Office

Éditer `Office-Install-Activate.cmd` et modifier ces variables :

```batch
:: Office Edition
set "OFFICE_PRODUCT=O365ProPlusRetail"

:: Autres options :
:: O365ProPlusRetail    - Office 365 ProPlus (recommandé)
:: ProPlus2021Volume    - Office 2021 LTSC Professional Plus
:: ProPlus2019Volume    - Office 2019 Professional Plus
:: Standard2021Volume   - Office 2021 Standard
```

### Modifier la langue

```batch
set "OFFICE_LANG=fr-fr"

:: Autres langues :
:: en-us (Anglais)
:: de-de (Allemand)
:: es-es (Espagnol)
:: it-it (Italien)
```

### Modifier l'architecture

```batch
set "OFFICE_ARCH=64"

:: Options : 64 ou 32
```

### Exclure des applications

```batch
set "EXCLUDE_APPS=Publisher,Access,OneDrive,Teams"

:: Applications disponibles :
:: Access, Excel, OneDrive, OneNote, Outlook
:: PowerPoint, Publisher, Word, Teams
```

### Modifier l'URL des DLL Ohook

Dans `Ohook-Activate.cmd` ou `Ohook-Activate-Silent.cmd` :

```batch
set "OHOOK_VERSION=0.5"
set "DLL64_URL=https://github.com/asdcorp/ohook/releases/download/%OHOOK_VERSION%/sppc64.dll"
set "DLL32_URL=https://github.com/asdcorp/ohook/releases/download/%OHOOK_VERSION%/sppc32.dll"
```

## Prérequis

- Windows 10/11 (ou Windows Server 2016+)
- Connexion Internet
- Droits Administrateur
- PowerShell 5.0+

## Produits supportés

| Produit | ID |
|---------|-----|
| Office 365 ProPlus | `O365ProPlusRetail` |
| Office 2021 Pro Plus | `ProPlus2021Volume` |
| Office 2021 Standard | `Standard2021Volume` |
| Office 2019 Pro Plus | `ProPlus2019Volume` |
| Office 2019 Standard | `Standard2019Volume` |
| Visio 2021 | `VisioPro2021Volume` |
| Project 2021 | `ProjectPro2021Volume` |

## Dépannage

### L'activation échoue

1. Fermez toutes les applications Office
2. Relancez le script d'activation en administrateur
3. Redémarrez l'ordinateur et réessayez

### Vérifier le statut d'activation

```cmd
cscript "%ProgramFiles%\Microsoft Office\root\Office16\OSPP.VBS" /dstatus
```

### Réinstaller proprement

1. Désinstaller Office via les paramètres Windows
2. Utiliser [Office Removal Tool](https://aka.ms/SaRA-OfficeUninstallFromPC)
3. Relancer le script d'installation

### Les DLL ne se téléchargent pas

Vérifiez que GitHub n'est pas bloqué sur votre réseau. Vous pouvez télécharger manuellement les DLL depuis :
- https://github.com/asdcorp/ohook/releases

## Sources

- [ohook par asdcorp](https://github.com/asdcorp/ohook) - Code source des DLL d'activation
- [Office Deployment Tool](https://docs.microsoft.com/deployoffice/overview-office-deployment-tool) - Outil d'installation Office
- [MAS Documentation](https://massgrave.dev/ohook) - Documentation sur Ohook

## Avertissement

Ce script est fourni à des fins éducatives. Utilisez-le de manière responsable et conformément aux lois en vigueur dans votre pays.
