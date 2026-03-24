# Architecture Technique

## Arborescence

- `src/main` : processus principal Electron (`main.js`, `preload.js`)
- `src/renderer` : interface utilisateur (`index.html`, `renderer.js`, `styles.css`)
- `src/backend` : logique Python (`server.py`, `v3.py`)
- `assets` : ressources statiques (logos, icones)
- `data` : jeux de donnees et fichiers de bootstrap (`.json`, `.csv`)
- `docs` : documentation technique

## Contrainte de compatibilite

La logique applicative est inchangée :
- mêmes routes API backend
- mêmes interactions Electron <-> renderer
- mêmes mecanismes d'execution Python

Seuls les chemins, l'organisation des fichiers et les points d'entree de structure ont ete realignes.
