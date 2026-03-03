# Excel MCP Server — Python

Version Python du serveur MCP Excel, portée depuis [negokaz/excel-mcp-server](https://github.com/negokaz/excel-mcp-server) (Go) vers **Python + FastMCP + openpyxl**.

## Outils disponibles (29)

### 📖 Lecture / Métadonnées

| Outil | Description |
|---|---|
| `excel_describe_sheets` | Liste les feuilles, plages utilisées, tables et plages paginées |
| `excel_read_sheet` | Lecture paginée d'une feuille (HTML) |
| `excel_validate_formula` | Valider la syntaxe d'une formule sans l'écrire |
| `excel_get_merged_cells` | Lister les cellules fusionnées d'une feuille |
| `excel_get_validation_info` | Lire les règles de validation de données existantes |

### ✏️ Écriture / Modification

| Outil | Description |
|---|---|
| `excel_write_to_sheet` | Écriture de valeurs / formules (range, startCell, append) |
| `excel_create_workbook` | Créer un nouveau classeur |
| `excel_create_table` | Créer un tableau structuré Excel |
| `excel_create_chart` | Créer un graphique (line, bar, pie, scatter, area) |
| `excel_create_pivot_table` | Tableau croisé dynamique via pandas (sum, mean, count, min, max, median) |

### 📄 Gestion des feuilles

| Outil | Description |
|---|---|
| `excel_copy_sheet` | Copier une feuille |
| `excel_delete_sheet` | Supprimer une feuille |
| `excel_rename_sheet` | Renommer une feuille |
| `excel_manage_named_ranges` | Gérer les plages nommées (list / create / delete) |

### 📐 Plages / Cellules

| Outil | Description |
|---|---|
| `excel_copy_range` | Copier une plage (valeurs + styles + traduction des formules) |
| `excel_delete_range` | Supprimer une plage avec décalage (haut / gauche) |
| `excel_move_range` | Déplacer une plage avec traduction optionnelle des formules |
| `excel_merge_cells` | Fusionner des cellules |
| `excel_unmerge_cells` | Défusionner des cellules |

### ↕️ Lignes / Colonnes

| Outil | Description |
|---|---|
| `excel_insert_rows` | Insérer des lignes |
| `excel_delete_rows` | Supprimer des lignes |
| `excel_insert_columns` | Insérer des colonnes |
| `excel_delete_columns` | Supprimer des colonnes |

### 🎨 Mise en forme

| Outil | Description |
|---|---|
| `excel_format_range` | Formater des cellules (police, nom de police, fond, bordure, alignement, protection, gradient…) |
| `excel_conditional_formatting` | Mise en forme conditionnelle (cellIs, colorScale, dataBar, iconSet, formula) |
| `excel_set_dimensions` | Hauteur de ligne, largeur de colonne, masquer/afficher lignes et colonnes |

### ⚙️ Fonctionnalités avancées

| Outil | Description |
|---|---|
| `excel_auto_filter` | Ajouter/supprimer des filtres automatiques |
| `excel_add_comment` | Ajouter, modifier ou supprimer des commentaires |
| `excel_data_validation` | Validation de données (listes déroulantes, contraintes numériques, formules…) |

## Prérequis

- Python 3.11+

## Installation

Après un `git clone` :

```bash
cd Excel_MCP_Python
pip install -e .
```

> **Note** : l'installation inclut `openpyxl`, `fastmcp` et `pandas` (utilisé pour les tableaux croisés dynamiques).

## Lancement

```bash
python server.py
# ou via FastMCP CLI :
fastmcp run server.py
```

## Configuration Claude Desktop / Cursor

```json
{
  "mcpServers": {
    "excel_python": {
      "command": "python",
      "args": ["C:/chemin/vers/Excel_MCP_Python/server.py"]
    }
  }
}
```

Ou avec le script installé :

```json
{
  "mcpServers": {
    "excel_python": {
      "command": "excel-mcp-server"
    }
  }
}
```

### Variable d'environnement

| Variable | Défaut | Description |
|---|---|---|
| `EXCEL_MCP_PAGING_CELLS_LIMIT` | `2000` | Nombre max de cellules par page de lecture |

## Notes

- **Chemins** : toujours utiliser des chemins absolus.
- **Serveur destiné à Excel français** : les formules sont automatiquement affichées en français par Excel.
- **Formules en ANGLAIS obligatoire** : openpyxl utilise le format interne `.xlsx` qui stocke les noms de fonctions en anglais. Excel français traduit automatiquement à l'affichage :
  - `=SUM(A1:A10)` → affiché `=SOMME(A1:A10)` dans Excel français ✅
  - `=SOMME(A1:A10)` → stocké tel quel → **erreur #NOM?** ❌
- **Séparateurs** : utiliser la virgule (`,`), pas le point-virgule (`;`). Excel français traduit automatiquement.
- **Fonctions avec point** : éviter les fonctions contenant un `.` (ex: `NORM.DIST`). Utiliser les versions legacy (`NORMDIST`).
- **Tableaux croisés** : calculés via pandas — le résultat est un tableau plat, pas un PivotTable natif Excel.

### Correspondances formules courantes (Anglais → Français dans Excel)

| Écrire (anglais) | Affiché dans Excel FR |
|---|---|
| `SUM` | `SOMME` |
| `AVERAGE` | `MOYENNE` |
| `IF` | `SI` |
| `VLOOKUP` | `RECHERCHEV` |
| `COUNTIF` | `NB.SI` |
| `SUMIF` | `SOMME.SI` |
| `INDEX` | `INDEX` |
| `MATCH` | `EQUIV` |
| `LEFT` / `RIGHT` / `MID` | `GAUCHE` / `DROITE` / `STXT` |
| `TODAY` / `NOW` | `AUJOURDHUI` / `MAINTENANT` |
| `TRUE` / `FALSE` | `VRAI` / `FAUX` |
