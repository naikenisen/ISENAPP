# NexoMail

Client email Electron avec gestion de tâches, visualisation de coffre Obsidian et export Markdown.

---

## Prérequis

- **Node.js** ≥ 18
- **Python 3** (accessible via `python3`)
- **pip** (gestionnaire de paquets Python)

---

## Installation des dépendances

### Node.js

```bash
npm install
```

### Python

```bash
pip install -r requirements-v3.txt
```

Ou manuellement :

```bash
pip install html2text
```

---

## Lancement en mode développement

```bash
npm start
```

Cela lance Electron et démarre automatiquement le serveur Python (`server.py`) sur le port 8080.

---

## Compilation

### Générer les exécutables Linux (AppImage + .deb)

```bash
npm run build
```

Les fichiers générés se trouvent dans le dossier `dist/` :

| Fichier | Description |
|---|---|
| `NexoMail-1.0.0.AppImage` | Exécutable portable, aucune installation requise |
| `nexomail_1.0.0_amd64.deb` | Paquet Debian installable |

### Générer uniquement le dossier décompressé (sans empaquetage)

```bash
npm run build:dir
```

Le résultat se trouve dans `dist/linux-unpacked/`.

---

## Installation

### AppImage

```bash
chmod +x dist/NexoMail-1.0.0.AppImage
./dist/NexoMail-1.0.0.AppImage
```

> **Note :** si l'AppImage ne se lance pas, essayer avec l'option `--no-sandbox` :
>
> ```bash
> ./dist/NexoMail-1.0.0.AppImage --no-sandbox
> ```

### Paquet .deb

```bash
sudo dpkg -i dist/nexomail_1.0.0_amd64.deb
```

Puis lancer depuis le menu d'applications ou via :

```bash
nexomail
```

---

## Structure du projet

```
main.js          → Processus principal Electron
preload.js       → Bridge IPC (contextIsolation)
index.html       → Interface complète (CSS + HTML + JS)
server.py        → Serveur HTTP Python (API todo, email, vault)
v3.py            → Utilitaire de traitement email
data.json        → Données persistantes (tâches, emails)
icon.png         → Icône de l'application
package.json     → Configuration npm et electron-builder
```

---

## Remarques

- **Python 3 doit être installé** sur la machine cible, même pour l'AppImage ou le .deb, car le serveur backend est lancé en tant que processus enfant.
- Le serveur Python écoute sur `localhost:8080`. Assurez-vous que ce port est disponible.
- Les données sont stockées dans `data.json` à côté du serveur (dans `resources/` en mode packagé).
