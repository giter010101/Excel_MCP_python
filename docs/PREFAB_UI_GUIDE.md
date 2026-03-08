# Guide Prefab UI — Utilisation, structure et intégration FastMCP

> Documentation générée à partir de [prefab.prefect.io](https://prefab.prefect.io) via Firecrawl.

---

## 1. Introduction à Prefab

**Prefab** est un framework frontend avec un DSL Python qui compile en JSON. Il permet de décrire une interface — layouts, formulaires, graphiques, tableaux de données, interactivité complète — qu’un renderer React transforme en application autonome.

- **Python → JSON → React** : le flux principal
- Compatible **MCP Apps** : interfaces interactives dans Claude Desktop, ChatGPT, etc.
- **35+ composants** : layout, typographie, formulaires, affichage de données, éléments interactifs

> ⚠️ Prefab est en développement actif (pré-1.0). L’API peut évoluer rapidement.

---

## 2. Installation

```bash
pip install prefab-ui
```

Ou avec uv :

```bash
uv add prefab-ui
```

**Prérequis** : Python 3.10+

### Versioning

Pour un usage en production, **fixer la version** :

```
prefab-ui==0.8.0   # Recommandé
prefab-ui>=0.5,<0.6 # OK pour recevoir les patches
```

---

## 3. Utilisation rapide

### 3.1 Premier exemple (Quickstart)

```python
from prefab_ui.app import PrefabApp
from prefab_ui.components import Button, Card, CardContent, CardFooter, Column, H3, Input, Muted
from prefab_ui.actions import ShowToast

with Column(css_class="max-w-md mx-auto") as view:
    with Card():
        with CardContent():
            with Column(gap=3):
                name_input = Input(name="name", placeholder="Your name...")
                H3(f"Hello, {name_input.rx}!")
                Muted("Type your name and watch the heading update in real time.")
        with CardFooter():
            Button("Say hi", on_click=ShowToast(f"Hey there, {name_input.rx}!"))

app = PrefabApp(view=view, state={"name": "world"})
```

### 3.2 Lancer en local

```bash
prefab serve app.py --reload
```

L’interface est disponible sur `http://127.0.0.1:5175`.

---

## 4. Concepts fondamentaux

Prefab repose sur **quatre concepts** :

| Concept | Rôle |
|--------|------|
| **Components** | Ce que l’utilisateur voit (Text, Button, Card, etc.) |
| **State** | Données côté client (clé-valeur) |
| **Expressions** | Lecture du state via `{{ key }}` pour garder l’affichage à jour |
| **Actions** | Réponse aux interactions (SetState, CallTool, ShowToast, etc.) |

### Cycle de fonctionnement

1. Python construit l’arbre de composants → le renderer l’affiche
2. L’utilisateur interagit (clic, saisie, etc.)
3. Une action est déclenchée (mise à jour du state ou appel serveur)
4. Les expressions se réévaluent → l’affichage se met à jour

---

## 5. Protocole JSON (wire format)

Chaque réponse Prefab est un objet JSON :

```json
{
  "version": "0.2",
  "view": { ... },
  "defs": { ... },
  "state": {
    "count": 42,
    "name": "Alice"
  }
}
```

| Clé | Type | Description |
|-----|------|-------------|
| `version` | string | Version du protocole (`"0.2"`) |
| `view` | Component | Arbre de composants racine |
| `defs` | object | Templates réutilisables (Define/Use) |
| `state` | object | State initial, accessible via `{{ key }}` |

### Interpolation

Les chaînes supportent `{{ key }}` pour interpoler le state :

```json
{
  "type": "P",
  "content": "Hello, {{ name }}! You have {{ count }} items."
}
```

`{{ $event }}` contient la valeur de l’événement déclencheur (position du slider, texte saisi, etc.).

---

## 6. Structure du package `prefab_ui`

```
prefab_ui/
├── __init__.py          # Exports principaux
├── app.py               # PrefabApp, set_initial_state
├── cli/                 # Commande `prefab` (serve, version, etc.)
├── components/          # Composants UI (layout, forms, charts, etc.)
│   ├── charts/          # BarChart, LineChart, PieChart, etc.
│   └── control_flow/    # If, Else, ForEach, etc.
├── actions/             # Actions (SetState, ShowToast, CallTool, etc.)
│   └── mcp.py           # CallTool, SendMessage, UpdateContext
├── rx/                  # Expressions réactives (Rx)
├── define.py            # Define (templates)
├── use.py               # Use (référence aux templates)
├── css.py               # Helpers CSS (Tailwind)
├── themes.py            # Thèmes
└── renderer/             # Bundle React embarqué
```

### Imports typiques

```python
from prefab_ui.app import PrefabApp, set_initial_state
from prefab_ui.components import Card, Column, Row, Button, Input, Text
from prefab_ui.components.charts import BarChart, ChartSeries
from prefab_ui.components.control_flow import If, Else, ForEach
from prefab_ui.actions import SetState, ShowToast
from prefab_ui.actions.mcp import CallTool, SendMessage, UpdateContext
from prefab_ui.rx import Rx
```

---

## 7. Intégration FastMCP (MCP Apps)

Prefab s’intègre à **FastMCP** pour créer des **MCP Apps** : UIs interactives affichées directement dans la conversation (Claude Desktop, ChatGPT, etc.).

### Dépendances

```toml
# pyproject.toml
dependencies = [
    "fastmcp[apps]>=3.1.0",
    "prefab-ui==0.8.0",
]
```

### Exemple minimal

```python
from fastmcp import FastMCP
from prefab_ui.app import PrefabApp
from prefab_ui.components import Column, Heading, Text
from prefab_ui.components.control_flow import ForEach

mcp = FastMCP("My Server")

ITEMS = [{"name": "Widget"}, {"name": "Gadget"}, {"name": "Gizmo"}]

@mcp.tool(app=True)
def browse() -> PrefabApp:
    """Show all items."""
    with Column(gap=4) as view:
        Heading("Items")
        with ForEach("items"):
            Text("{{ name }}")
    return PrefabApp(view=view, state={"items": ITEMS})
```

### Flux FastMCP + Prefab

1. **Appel outil** — Le host MCP appelle le tool
2. **PrefabApp → structuredContent** — FastMCP sérialise la valeur de retour en JSON Prefab
3. **Rendu** — Le host charge le renderer Prefab (`ui://prefab/renderer.html`) et affiche l’UI

### Retourner une UI

Deux possibilités :

```python
# Option 1 : PrefabApp complet
@mcp.tool(app=True)
def dashboard() -> PrefabApp:
    with Column(gap=4) as view:
        Heading("Dashboard")
        Text(f"Welcome, {Rx('name')}")
    return PrefabApp(title="Dashboard", view=view, state={"name": "Alice"})

# Option 2 : Composant seul (enveloppé automatiquement)
@mcp.tool(app=True)
def dashboard() -> Column:
    with Column(gap=4) as view:
        Heading("Dashboard")
        Text("Hello!")
    return view
```

### Appeler le serveur depuis l’UI : `CallTool`

```python
from prefab_ui.actions.mcp import CallTool
from prefab_ui.components import Slot

@mcp.tool(app=True)
def browse() -> PrefabApp:
    with Column(gap=4) as view:
        Input(
            name="q",
            placeholder="Search...",
            on_change=[
                SetState("q", "{{ $event }}"),
                CallTool("search", arguments={"q": "{{ $event }}"}, result_key="results"),
            ],
        )
        Slot("results")
    return PrefabApp(view=view, state={"q": "", "results": None})

@mcp.tool
def search(q: str = "") -> PrefabApp:
    matches = [i for i in ITEMS if q.lower() in i["name"].lower()] if q else ITEMS
    with ForEach("items") as view:
        Text("{{ name }}")
    return PrefabApp(view=view, state={"items": matches})
```

`Slot("results")` affiche le contenu renvoyé par `CallTool` avec `result_key="results"`.

### Visibilité des tools (AppConfig)

Pour masquer certains tools du modèle tout en les gardant accessibles via `CallTool` :

```python
from fastmcp.server.apps import AppConfig

@mcp.tool(app=True)
def browse() -> PrefabApp:
    """Point d'entrée — visible par le modèle."""
    ...

@mcp.tool(app=AppConfig(visibility=["app"]))
def search(q: str = "") -> PrefabApp:
    """Appelé uniquement depuis l'UI — caché du modèle."""
    ...
```

### Actions MCP spécifiques

| Action | Description |
|--------|-------------|
| `CallTool` | Appel d’un tool MCP côté serveur |
| `SendMessage` | Envoi d’un message dans la conversation |
| `UpdateContext` | Mise à jour silencieuse du contexte du modèle |

---

## 8. Comparaison FastMCP vs API Server

| | FastMCP | API Server |
|---|---------|------------|
| **Transport** | Protocole MCP | HTTP (fetch) |
| **Action serveur** | `CallTool` | `Fetch` |
| **Actions host** | `SendMessage`, `UpdateContext` | — |
| **Hébergement** | Dans Claude Desktop, ChatGPT, etc. | Page web standalone |
| **Renderer** | Fourni par le host MCP | Inclus dans `PrefabApp.html()` |

---

## 9. Ressources

- **Documentation** : [prefab.prefect.io](https://prefab.prefect.io)
- **Index complet** : [prefab.prefect.io/docs/llms.txt](https://prefab.prefect.io/docs/llms.txt)
- **Playground** : [prefab.prefect.io/docs/playground](https://prefab.prefect.io/docs/playground)
- **Dépôt** : [github.com/PrefectHQ/prefab](https://github.com/PrefectHQ/prefab)
- **Exemple MCP** : [examples/hitchhikers-guide](https://github.com/PrefectHQ/prefab/tree/main/examples/hitchhikers-guide)

---

## 10. Méthodologie de documentation (Firecrawl)

Cette documentation a été produite en utilisant **Firecrawl** pour extraire le contenu des pages Prefab.

### Utilisation de Firecrawl pour documenter Prefab

```json
// Scraper une page de documentation
{
  "url": "https://prefab.prefect.io/docs/running/fastmcp.md",
  "formats": ["markdown"]
}
```

**Index complet** : [https://prefab.prefect.io/docs/llms.txt](https://prefab.prefect.io/docs/llms.txt) — liste toutes les pages disponibles.

### Pages scrapées pour ce guide

- `getting-started/quickstart.md`
- `getting-started/installation.md`
- `running/fastmcp.md`
- `protocol/overview.md`
- `concepts/core-concepts.md`
