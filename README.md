# 🛠️ Administration Active Directory via Powershell

## Description

`Administration.ps1` est un script PowerShell conçu pour automatiser diverses tâches d'administration système sur les postes Windows. Il centralise plusieurs fonctions utiles pour la gestion, l'analyse, le nettoyage ou la configuration des machines dans un environnement d'entreprise ou d'administration locale.

## 📦 Fonctions incluses

Voici une liste non exhaustive des fonctionnalités prises en charge par le script :

- 🔍 Vérification de la configuration système
- 📂 Nettoyage des répertoires temporaires
- 🔐 Gestion des utilisateurs et des droits
- 📝 Analyse et journalisation des événements
- 🖥️ Inventaire matériel et logiciel
- 🔄 Automatisation de tâches répétitives
- 💾 Sauvegardes/restauration de configurations

## 🖥️ Prérequis

- PowerShell 5.1 ou supérieur
- Lancement en tant qu’administrateur pour certaines commandes
- Politique d’exécution permettant l’exécution de scripts (`Set-ExecutionPolicy`)

## 🚀 Utilisation

### 1. Cloner ce dépôt

```bash
git clone https://github.com/votre-utilisateur/nom-du-repo.git
cd nom-du-repo
```

### 2. Exécuter le script PowerShell

```powershell
# Depuis PowerShell (en mode administrateur de préférence)
.\Administration.ps1
```

> 💡 Astuce : Pour automatiser l’exécution au démarrage ou à une heure définie, utilisez le planificateur de tâches Windows.

## 🧪 Conseils de sécurité

- Lisez et comprenez le script avant exécution.
- Exécutez uniquement sur des machines de test si vous n’êtes pas certain de son comportement.
- Sauvegardez votre système ou créez un point de restauration avant toute modification massive.

## 📄 Licence

Ce projet est sous licence [MIT](LICENSE). Vous pouvez l’utiliser, le modifier et le redistribuer librement.

## 🙋 Support & Contribution

Les contributions sont les bienvenues ! Merci de soumettre vos idées via des issues ou des pull requests.