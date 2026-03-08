"""Tool: excel_interactive_range"""

from fastmcp import FastMCP
from fastmcp.server.apps import AppConfig
from prefab_ui.app import PrefabApp
from prefab_ui.components import Column, Row, Input, Button, Heading, Text, Badge, Separator
from prefab_ui.actions.mcp import CallTool

from excel_engine import open_workbook, parse_range, cell_name


def register_interactive_range(mcp: FastMCP):
    @mcp.tool(
        name="excel_interactive_range",
        description="View and interactively edit an Excel range using a UI form.",
        app=True
    )
    def excel_interactive_range(
        fileAbsolutePath: str,
        sheetName: str,
        range_str: str,
    ) -> PrefabApp:
        """
        Creates an interactive UI to edit cell values in a given range.
        
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
            range_str: Range of cells to read (e.g. "A1:C5")
        """
        wb = open_workbook(fileAbsolutePath, read_only=False)
        if sheetName not in wb.sheetnames:
            wb.close()
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]
        
        min_col, min_row, max_col, max_row = parse_range(range_str)
        
        state = {}
        cell_values_args = {}

        # Collect state and build arguments for save tool
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                c_name = cell_name(c, r)
                val = ws.cell(row=r, column=c).value
                state[c_name] = "" if val is None else str(val)
                cell_values_args[c_name] = f"{{{{ {c_name} }}}}"

        wb.close()

        from prefab_ui.components import Slot
        import string

        def col_letter(col_idx: int) -> str:
            """Convert 1-based column index to Excel letter(s)."""
            result = ""
            while col_idx > 0:
                col_idx, rem = divmod(col_idx - 1, 26)
                result = string.ascii_uppercase[rem] + result
            return result

        n_cols = max_col - min_col + 1
        filename = fileAbsolutePath.replace("\\", "/").split("/")[-1]

        with Column(gap=4, css_class="p-4 rounded-xl") as view:

            # ── Header ────────────────────────────────────────────────
            with Row(css_class="items-center justify-between"):
                Heading(f"Editing Range {range_str}")
                Badge(sheetName, variant="secondary")

            Text(
                filename,
                css_class="text-xs text-muted-foreground -mt-2",
            )

            Separator()

            # ── Column letters row ────────────────────────────────────
            with Column(gap=1, css_class="overflow-x-auto"):
                with Row(gap=1, css_class="mb-1 items-center"):
                    # offset cell aligned with row-number gutter
                    Text("", css_class="w-8 shrink-0")
                    for c in range(min_col, max_col + 1):
                        Badge(
                            col_letter(c),
                            variant="secondary",
                            css_class="w-28 min-w-16 justify-center font-mono text-xs",
                        )

                # ── Data rows ─────────────────────────────────────────
                for r in range(min_row, max_row + 1):
                    is_header = (r == min_row)
                    row_bg = "bg-primary/10" if is_header else ""
                    with Row(gap=1, css_class=f"items-center {row_bg} rounded"):
                        # Row-number gutter
                        Text(
                            str(r),
                            css_class=(
                                "w-8 shrink-0 text-center text-xs font-mono "
                                "text-muted-foreground select-none"
                            ),
                        )
                        for c in range(min_col, max_col + 1):
                            c_name = cell_name(c, r)
                            Input(
                                name=c_name,
                                placeholder=c_name,
                                css_class=(
                                    "w-28 min-w-16 font-mono text-sm "
                                    + ("font-bold" if is_header else "")
                                ),
                            )

            Separator()

            # ── Save button ───────────────────────────────────────────
            Button(
                "Save Changes",
                on_click=CallTool(
                    "excel_save_range_changes",
                    arguments={
                        "fileAbsolutePath": fileAbsolutePath,
                        "sheetName": sheetName,
                        "cell_values": cell_values_args,
                    },
                    result_key="save_result",
                ),
            )
            Slot("save_result")

        return PrefabApp(view=view, state=state, title=f"Editor {range_str}")

    @mcp.tool(app=AppConfig(visibility=["app"]))
    def excel_save_range_changes(
        fileAbsolutePath: str,
        sheetName: str,
        cell_values: dict[str, str]
    ) -> Column:
        """
        Internal tool called by the UI to save changes to the Excel file.
        Values are coerced to their natural Python type before writing.
        """
        from excel_engine import open_workbook, save_workbook
        from datetime import datetime

        def coerce(raw: str):
            """Convert a string value to the most appropriate Python type."""
            if raw == "":
                return None
            # Formulas — keep as-is
            if raw.startswith("="):
                return raw
            # Integer
            try:
                i = int(raw)
                # avoid treating "2024" date-like strings as int when they contain separators
                return i
            except ValueError:
                pass
            # Float (handle both '.' and ',' as decimal separator)
            try:
                return float(raw.replace(",", "."))
            except ValueError:
                pass
            # Date / datetime — common formats
            for fmt in (
                "%Y-%m-%d",
                "%d/%m/%Y",
                "%m/%d/%Y",
                "%d-%m-%Y",
                "%Y-%m-%d %H:%M:%S",
                "%d/%m/%Y %H:%M:%S",
            ):
                try:
                    return datetime.strptime(raw, fmt)
                except ValueError:
                    pass
            # Fallback: plain string
            return raw

        wb = open_workbook(fileAbsolutePath, read_only=False)
        if sheetName not in wb.sheetnames:
            wb.close()
            with Column() as result:
                Text("Erreur : feuille introuvable.", css_class="text-red-500")
            return result

        ws = wb[sheetName]
        try:
            type_summary: dict[str, int] = {
                "formule": 0, "entier": 0, "décimal": 0,
                "date": 0, "texte": 0, "vide": 0,
            }
            for c_name, raw in cell_values.items():
                coerced = coerce(raw)
                ws[c_name].value = coerced
                if coerced is None:
                    type_summary["vide"] += 1
                elif isinstance(coerced, str) and coerced.startswith("="):
                    type_summary["formule"] += 1
                elif isinstance(coerced, int):
                    type_summary["entier"] += 1
                elif isinstance(coerced, float):
                    type_summary["décimal"] += 1
                elif isinstance(coerced, datetime):
                    type_summary["date"] += 1
                else:
                    type_summary["texte"] += 1

            save_workbook(wb, fileAbsolutePath)

            summary_parts = [
                f"{n} {t}" for t, n in type_summary.items() if n > 0
            ]
            with Column(gap=2) as result:
                Text(
                    f"✓ {len(cell_values)} cellules sauvegardées",
                    css_class="text-green-500 font-bold",
                )
                Text(
                    "Types détectés : " + " · ".join(summary_parts),
                    css_class="text-xs text-muted-foreground",
                )
            return result
        except Exception as e:
            wb.close()
            with Column() as result:
                Text(f"Erreur : {e}", css_class="text-red-500")
            return result
