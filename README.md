# 🛠️ Administration Active Directory via Powershell

## 🇫🇷 Description

`Administration.ps1` est un script PowerShell conçu pour automatiser diverses tâches d'administration système sur les postes Windows. Il centralise plusieurs fonctions utiles pour la gestion, l'analyse, le nettoyage ou la configuration des machines dans un environnement d'entreprise ou d'administration locale.

## 🇬🇧 Description

`Administration.ps1` is a PowerShell script designed to automate various system administration tasks on Windows machines. It consolidates several useful functions for managing, analyzing, cleaning, or configuring machines in a business or local administration environment.

---

## 📦 🇫🇷 Fonctions incluses / 🇬🇧 Included Features

- 🔍 Vérification de la configuration système / System configuration check
- 📂 Nettoyage des répertoires temporaires / Temporary folder cleanup
- 🔐 Gestion des utilisateurs et des droits / User and permission management
- 📝 Analyse et journalisation des événements / Event log analysis
- 🖥️ Inventaire matériel et logiciel / Hardware and software inventory
- 🔄 Automatisation de tâches répétitives / Repetitive task automation
- 💾 Sauvegardes/restauration de configurations / Configuration backup and restore

---

## 🖥️ 🇫🇷 Prérequis / 🇬🇧 Requirements

- PowerShell 5.1 or later
- Lancement en tant qu’administrateur pour certaines commandes / Run as administrator for some commands
- Politique d’exécution permettant l’exécution de scripts (`Set-ExecutionPolicy`)  
  / Execution policy allowing script execution (`Set-ExecutionPolicy`)

---

## 🚀 🇫🇷 Utilisation / 🇬🇧 Usage

### 1. 🇫🇷 Cloner le dépôt / 🇬🇧 Clone the repository

```bash
git clone https://github.com/votre-utilisateur/nom-du-repo.git
cd nom-du-repo
```

### 2. 🇫🇷 Exécuter le script / 🇬🇧 Run the script

```powershell
# 🇫🇷 Depuis PowerShell (en mode administrateur de préférence)
# 🇬🇧 From PowerShell (preferably in administrator mode)
.\Administration.ps1
```

💡 **Astuce / Tip** :  
🇫🇷 Utilisez le planificateur de tâches Windows pour l’exécuter automatiquement  
🇬🇧 Use Windows Task Scheduler for automated execution

---

## 🧪 🇫🇷 Conseils de sécurité / 🇬🇧 Security Tips

- 🇫🇷 Lisez et comprenez le script avant exécution  
  🇬🇧 Read and understand the script before running it
- 🇫🇷 Exécutez d’abord sur une machine de test  
  🇬🇧 Test it first on a development machine
- 🇫🇷 Sauvegardez vos données avant modifications massives  
  🇬🇧 Backup your data before making major changes

---

## 📄 Licence / License

Ce projet est sous licence [MIT](LICENSE).  
This project is licensed under the [MIT License](LICENSE).

---

## ✅ 🇫🇷 Qualité & Analyse / 🇬🇧 Quality & Analysis

Ce script a été analysé avec [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) pour garantir la qualité du code et le respect des bonnes pratiques PowerShell.  
This script was tested using [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) to ensure code quality and adherence to PowerShell best practices.

## 🙋 🇫🇷 Support & Contribution / 🇬🇧 Support & Contribution

Les contributions sont les bienvenues !  
Contributions are welcome!

Merci de soumettre vos idées via des issues ou des pull requests.  
Please submit your ideas via issues or pull requests.