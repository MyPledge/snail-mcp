from pathlib import Path
from mcp.server.fastmcp import FastMCP
from mcp.excel.read import read_excel as _read_excel

mcp = FastMCP("snail-mcp", description="General-purpose MCP server with multiple tools. Extensible for more data/file operations.")

READ_EXCEL_DESC = """Read cell data from a .xlsx workbook with configurable output fields and format.

**When to use**: Call when you need Excel table content, formulas, or styles (bold, font, color).

**Output format** (controlled by fmt):
- fmt=j (default): JSON. Structure: {"SheetName": {"cells": [{"row", "col", "value", "bold", "formula"?}, ...]}}. Cell fields depend on out.
- fmt=t: Plain text. One line ">SheetName" per sheet, then one line per cell "c\\trow\\tcol\\tvalue\\t...", tab-separated, columns follow out order. Tab/newline in values are escaped as \\\\t/\\\\n.

**out** (combine letters e.g. "vb" for value and bold): v=value, b=bold, f=formula, n=font_name, z=font_size, c=font_color, g=fill_color.

**Range**: row_start/row_end/col_start/col_end=0 means use sheet's used range; otherwise the given row/col range.

**sparse**: True (default) output only non-blank cells; False output every cell in range."""

@mcp.tool(description=READ_EXCEL_DESC)
def read_excel(
    filepath: str,
    out: str = "vbf",
    fmt: str = "j",
    sheets: str = "",
    row_start: int = 0,
    row_end: int = 0,
    col_start: int = 0,
    col_end: int = 0,
    sparse: bool = True,
) -> str:
    """Read xlsx. filepath: path to .xlsx. out: v=value b=bold f=formula n=font_name z=font_size c=font_color g=fill. fmt: j=json t=text. sheets: comma names or empty=all. row/col_start/end: 0=used range. sparse: True=non-blank only."""
    p = Path(filepath)
    if not p.exists():
        return f"error: file not found: {filepath}"
    return _read_excel(str(p), out=out, fmt=fmt, sheets=sheets, row_start=row_start, row_end=row_end, col_start=col_start, col_end=col_end, sparse=sparse)

def main():
    mcp.run(transport="stdio")

if __name__ == "__main__":
    main()
