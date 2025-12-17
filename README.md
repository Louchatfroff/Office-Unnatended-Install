# Office Unattended Installation and Activation

Scripts pour installer et activer Microsoft Office de manière automatique (unattended).

Basé sur [Microsoft Activation Scripts (MAS)](https://massgrave.dev) - méthode Ohook.

## Scripts disponibles

| Script | Description |
|--------|-------------|
| `Office-Install-Activate.cmd` | Installation automatique complète (configurable) |
| `Office-Menu-Install.cmd` | Version interactive avec menu |
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

## Activation uniquement

Si Office est déjà installé, vous pouvez l'activer avec :

### Méthode 1 : Via le menu
Lancez `Office-Menu-Install.cmd` et choisissez l'option 5.

### Méthode 2 : Commande PowerShell directe
```powershell
irm https://get.activated.win | iex
```
Puis sélectionner "Ohook" dans le menu.

### Méthode 3 : Activation silencieuse
```powershell
& ([ScriptBlock]::Create((irm https://get.activated.win))) /Ohook /S
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
Visitez : https://massgrave.dev/troubleshoot

### Vérifier le statut d'activation
```cmd
cscript "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /dstatus
```

### Réinstaller proprement
1. Désinstaller Office via les paramètres Windows
2. Utiliser [Office Removal Tool](https://aka.ms/SaRA-OfficeUninstallFromPC)
3. Relancer le script d'installation

## Crédits

- [Microsoft Activation Scripts (MAS)](https://github.com/massgravel/Microsoft-Activation-Scripts)
- [Office Deployment Tool](https://docs.microsoft.com/deployoffice/overview-office-deployment-tool)

## Avertissement

Ce script est fourni à des fins éducatives. Utilisez-le de manière responsable et conformément aux lois en vigueur dans votre pays.
