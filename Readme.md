<div align="center">

# NeuRail

Client de bureau Electron et Python pour centraliser la gestion des tâches, des emails, des rappels, des leads et de l'agenda dans une interface unique.

[![Version](https://img.shields.io/badge/version-1.0.0-blue?style=flat-square)](package.json)
[![Electron](https://img.shields.io/badge/Electron-33-47848f?style=flat-square&logo=electron&logoColor=white)](https://www.electronjs.org/)
[![Licence](https://img.shields.io/badge/licence-MIT-green?style=flat-square)](LICENSE)
[![Build](https://img.shields.io/badge/build-electron--builder-orange?style=flat-square)](https://www.electron.build/)

</div>

---

## Table des matières

- [ISENAPP](#isenapp)
  - [Table des matières](#table-des-matières)
  - [Fonctionnalités](#fonctionnalités)
  - [L'objectif du projet est de réduire les changements d'outils en regroupant dans un seul client les opérations de communication, de suivi et d'organisation.](#lobjectif-du-projet-est-de-réduire-les-changements-doutils-en-regroupant-dans-un-seul-client-les-opérations-de-communication-de-suivi-et-dorganisation)
  - [Installation (Développement)](#installation-développement)
    - [Prérequis](#prérequis)
    - [Mise en place](#mise-en-place)
  - [Variables d'environnement](#variables-denvironnement)
  - [Build \& Distribution](#build--distribution)
    - [Générer les icônes (optionnel, nécessite `rsvg-convert`)](#générer-les-icônes-optionnel-nécessite-rsvg-convert)
    - [Packager par OS](#packager-par-os)
  - [Structure du projet](#structure-du-projet)
  - [Communication IPC](#communication-ipc)
    - [1. Preload — Exposition de l'API sécurisée](#1-preload--exposition-de-lapi-sécurisée)
    - [2. Main — Gestion des événements IPC](#2-main--gestion-des-événements-ipc)
    - [3. Renderer — Utilisation côté interface](#3-renderer--utilisation-côté-interface)
  - [Contribution](#contribution)
    - [Conventions](#conventions)
  - [Licence](#licence)

---

## Fonctionnalités

ISENAPP est conçu comme un poste de travail unifié pour les équipes qui gèrent à la fois leurs échanges email, leurs relances, leur organisation quotidienne et leur suivi commercial.

- Gestion des tâches dans une interface intégrée au reste du flux de travail.
- Messagerie multi-comptes avec consultation de la boîte de réception, rédaction, réponse et envoi depuis l'application.
- Configuration email facilitée grâce à l'auto-détection IMAP/SMTP et à la gestion centralisée des comptes.
- Suivi des échanges sortants avec rappels pour relancer un destinataire lorsqu'une réponse est attendue.
- Assistance IA pour reformuler un message, corriger la rédaction et générer des réponses à partir d'un contexte ou d'un prompt.
- Sauvegarde des emails au format `.eml` et archivage en Markdown pour faciliter l'indexation, l'analyse et l'exploitation par des outils IA.
- Visualisation des archives et des connaissances associées sous forme de graphe.
- Intégration de Google Agenda pour consulter les calendriers et piloter les événements depuis l'application.
- Gestion des leads, des projets et de l'organisation d'équipe dans le même environnement de travail.

L'objectif du projet est de réduire les changements d'outils en regroupant dans un seul client les opérations de communication, de suivi et d'organisation.
---

## Installation (Développement)

### Prérequis

| Outil | Version requise |
|---|---|
| **Node.js** | `>= 18 LTS` |
| **npm** | `>= 9` |
| **Python** | `>= 3.8` |

### Mise en place

```bash
# 1. Cloner le dépôt
git clone https://github.com/naikenisen/ISENAPP.git
cd ISENAPP

# 2. Installer les dépendances Node.js
npm install

# 3. Créer un environnement virtuel Python et installer les dépendances
python3 -m venv venv
source venv/bin/activate        # Linux/macOS
# .\venv\Scripts\activate       # Windows (PowerShell)
pip install -r requirements.txt

# 4. Lancer l'application en mode développement
npm start
```

> **Note :** Le serveur Python backend (`server.py`) est automatiquement démarré par le processus Main d'Electron au lancement. Assurez-vous que le port **8080** est disponible.

---

## Variables d'environnement

L'application utilise des variables d'environnement optionnelles pour personnaliser le stockage des données :

```bash
# .env.example

# Répertoire de stockage des données runtime (comptes, tâches, emails indexés).
# Par défaut : ~/.local/share/isenapp (ou $XDG_DATA_HOME/isenapp)
ISENAPP_DATA_DIR=

# Clé API Google Gemini pour les fonctionnalités d'assistance IA
# (configurée dans l'interface de l'application)
```

---

## Build & Distribution

ISENAPP utilise [electron-builder](https://www.electron.build/) pour le packaging.

### Générer les icônes (optionnel, nécessite `rsvg-convert`)

```bash
npm run icons:generate
```

### Packager par OS

```bash
# Linux (AppImage + .deb)
npm run build

# macOS (DMG + ZIP)
npm run build:mac

# Windows (NSIS installer + portable)
npm run build:win

# Toutes les plateformes
npm run build:all
```

Les artefacts sont générés dans le dossier `dist/`.

| Plateforme | Formats | Catégorie |
|---|---|---|
| Linux | AppImage, `.deb` | Office |
| macOS | `.dmg`, `.zip` | Productivity |
| Windows | NSIS installer, Portable | — |

---

## Structure du projet

```
ISENAPP/
├── assets/                  # Ressources statiques (logo SVG, icônes)
├── build/                   # Icônes générées (icon.png, icon.icns, icon.ico)
├── data/                    # Données par défaut embarquées (bootstrap)
│   ├── data.json            #   État initial de l'application
│   └── contacts_*.csv       #   Carnet de contacts par défaut
├── dist/                    # Artefacts de build (généré)
├── src/
│   ├── main/
│   │   ├── main.js          # Processus principal Electron
│   │   └── preload.js       # Script preload (contextBridge)
│   ├── renderer/
│   │   ├── index.html       # Interface utilisateur complète
│   │   ├── styles.css       # Feuille de styles
│   │   └── renderer.js      #    Point d'entrée renderer
│   └── backend/
│       ├── server.py            # Serveur HTTP Python (API locale)
│       ├── app_config.py        # Configuration et chemins
│       ├── json_store.py        # Lecture/écriture atomique JSON
│       ├── account_store.py     # CRUD comptes email
│       ├── mail_utils.py        # Parsing email, .eml I/O, index
│       ├── mail_service.py      # Protocoles POP3/IMAP/SMTP
│       ├── mail_to_md.py        # Conversion emails → Markdown Graph
│       ├── google_calendar_service.py  # OAuth2 Google + Calendar API
│       ├── calendar_routes.py   # Routes HTTP calendrier
│       ├── ai_service.py        # Appels IA Google Gemini
│       ├── graph_service.py     # Graphe de connaissances + export email
│       └── autoconfig_service.py # Auto-détection IMAP/SMTP
├── package.json             # Configuration npm & electron-builder
├── requirements.txt         # Dépendances Python
└── README.md                # Ce fichier
```

---

## Communication IPC

Toute communication entre le Renderer et le Main process transite par le **preload script** via `contextBridge`, garantissant une isolation complète.

### 1. Preload — Exposition de l'API sécurisée

```js
// src/main/preload.js
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  // Contrôles fenêtre
  minimize:    () => ipcRenderer.send('window:minimize'),
  maximize:    () => ipcRenderer.send('window:maximize'),
  close:       () => ipcRenderer.send('window:close'),
  isMaximized: () => ipcRenderer.invoke('window:isMaximized'),

  // Fichiers
  readFile:  (path)          => ipcRenderer.invoke('fs:readFile', path),
  writeFile: (path, content) => ipcRenderer.invoke('fs:writeFile', path, content),

  // Coffre-fort mots de passe
  passwordVaultList:   () => ipcRenderer.invoke('passwordVault:list'),
  passwordVaultUpsert: (payload) => ipcRenderer.invoke('passwordVault:upsert', payload),
  // ... autres canaux
});
```

### 2. Main — Gestion des événements IPC

```js
// src/main/main.js
const { ipcMain } = require('electron');

// Événement unidirectionnel (send → on)
ipcMain.on('window:minimize', () => {
  if (mainWindow) mainWindow.minimize();
});

// Événement bidirectionnel (invoke → handle)
ipcMain.handle('window:isMaximized', () => {
  return mainWindow ? mainWindow.isMaximized() : false;
});
```

### 3. Renderer — Utilisation côté interface

```js
// Dans le code du renderer (index.html)
// Appel simple (fire-and-forget)
window.electronAPI.minimize();

// Appel avec réponse (async)
const isMax = await window.electronAPI.isMaximized();
```

---

## Contribution

Les contributions sont les bienvenues ! Merci de suivre ces étapes :

1. **Forkez** le dépôt
2. Créez une branche pour votre fonctionnalité (`git checkout -b feat/ma-fonctionnalite`)
3. Committez vos changements (`git commit -m "feat: description"`)
4. Poussez vers la branche (`git push origin feat/ma-fonctionnalite`)
5. Ouvrez une **Pull Request**

### Conventions

- Commits : suivez la convention [Conventional Commits](https://www.conventionalcommits.org/)
- Code JavaScript : pas de TypeScript — vanilla JS, `const`/`let`, template literals
- Sécurité : toute nouvelle API IPC doit passer par le preload avec des canaux explicitement validés

---

## Licence

Ce projet est sous licence **MIT**. Voir le fichier [LICENSE](LICENSE) pour plus de détails.

---

<div align="center">

par [Isen Naiken](https://github.com/naikenisen) 

</div>
