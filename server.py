#!/usr/bin/env python3
"""Serveur local pour l'app Todo — sauvegarde dans data.json."""

import http.server
import json
import os
import urllib.request

PORT = 8080
DIR = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(DIR, "data.json")


def load():
    try:
        with open(DATA, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {"sections": [], "settings": {}}


def save(data):
    with open(DATA, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def ai_organize(payload):
    token = payload.get("token", "")
    sections = payload.get("sections", [])

    lines = []
    for s in sections:
        lines.append(f"\n## {s.get('emoji', '')} {s.get('title', '')}")
        if s.get("description"):
            lines.append(f"   {s['description']}")
        for t in s.get("tasks", []):
            mark = "x" if t.get("done") else " "
            indent = "  " * t.get("indent", 0)
            line = f"{indent}- [{mark}] {t.get('label', '')}"
            if t.get("note"):
                line += f"  (Note: {t['note']})"
            lines.append(line)

    prompt = (
        "Tu es un assistant d'organisation. Voici une todo-list.\n"
        "Réorganise, reformule clairement et trie les tâches en sections logiques.\n"
        "Conserve le statut fait/pas fait de chaque tâche.\n"
        "Propose un emoji approprié pour chaque section.\n"
        "Réponds UNIQUEMENT en JSON valide (sans balises markdown) avec cette structure :\n"
        '{"sections":[{"emoji":"📞","title":"...","badge":"...","color":"blue|orange|green|purple|pink|slate",'
        '"description":"...","tasks":[{"label":"...","note":"","done":false,"indent":0}]}]}\n\n'
        + "\n".join(lines)
    )

    body = json.dumps({
        "model": "gpt-4o",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3,
    }).encode()

    req = urllib.request.Request(
        "https://models.inference.ai.azure.com/chat/completions",
        data=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
    )

    with urllib.request.urlopen(req, timeout=60) as r:
        result = json.loads(r.read())

    content = result["choices"][0]["message"]["content"]
    if "```" in content:
        content = content.split("```json")[-1] if "```json" in content else content.split("```")[1]
        content = content.split("```")[0]
    return json.loads(content.strip())


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=DIR, **kw)

    def do_GET(self):
        if self.path == "/api/state":
            return self._json(load())
        super().do_GET()

    def do_POST(self):
        try:
            raw = self.rfile.read(int(self.headers.get("Content-Length", 0)))
            data = json.loads(raw) if raw else {}
        except (json.JSONDecodeError, ValueError):
            return self._json({"error": "JSON invalide"}, 400)

        if self.path == "/api/state":
            save(data)
            return self._json({"ok": True})

        if self.path == "/api/organize":
            try:
                result = ai_organize(data)
                return self._json(result)
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        self.send_error(404)

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass  # silencieux


if __name__ == "__main__":
    print(f"🚀 Todo → http://localhost:{PORT}")
    http.server.HTTPServer(("", PORT), Handler).serve_forever()
