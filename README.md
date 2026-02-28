# Excel MCP Server — Python

Version Python du serveur MCP Excel, portée depuis [negokaz/excel-mcp-server](https://github.com/negokaz/excel-mcp-server) (Go) vers **Python + FastMCP + openpyxl**.

## Outils disponibles

| Outil | Description |
|---|---|
| `excel_describe_sheets` | Liste les feuilles et leurs plages |
| `excel_read_sheet` | Lecture paginée d'une feuille |
| `excel_write_to_sheet` | Écriture de valeurs / formules |
| `excel_create_workbook` | Créer un nouveau classeur |
| `excel_rename_sheet` | Renommer une feuille |
| `excel_delete_sheet` | Supprimer une feuille |
| `excel_copy_sheet` | Copier une feuille |
| `excel_insert_rows` | Insérer des lignes |
| `excel_delete_rows` | Supprimer des lignes |
| `excel_insert_columns` | Insérer des colonnes |
| `excel_delete_columns` | Supprimer des colonnes |
| `excel_format_range` | Formater des cellules (police, fond, bordure…) |
| `excel_create_table` | Créer un tableau structuré |
| `excel_create_chart` | Créer un graphique |
| `excel_create_pivot_table` | Tableau croisé (limité, voir note) |
| `excel_merge_cells` | Fusionner des cellules |
| `excel_unmerge_cells` | Défusionner des cellules |
| `excel_manage_named_ranges` | Gérer les plages nommées (list/create/delete) |

## Installation

```bash
pip install fastmcp openpyxl
# ou avec uv :
uv pip install fastmcp openpyxl
```

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
    "excel": {
      "command": "python",
      "args": ["C:/chemin/vers/Excel_MCP_Python/server.py"]
    }
  }
}
```

### Variable d'environnement

| Variable | Défaut | Description |
|---|---|---|
| `EXCEL_MCP_PAGING_CELLS_LIMIT` | `2000` | Nombre max de cellules par page |

## Notes

- **Pivot tables** : openpyxl ne supporte pas la création de tableaux croisés dynamiques. Utilisez Excel ou xlwings pour cette fonctionnalité.
- **VBA** : les macros VBA ne sont pas supportées dans cette version Python.
- **Chemins** : toujours utiliser des chemins absolus.
