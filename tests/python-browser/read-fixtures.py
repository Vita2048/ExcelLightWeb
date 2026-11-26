"""Project the two trusted, frozen demo files into xl() inputs for the probe."""
from datetime import datetime, timedelta
from pathlib import Path
import json
import sys
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
EXCEL_EPOCH = datetime(1899, 12, 30)


def iso(value):
    return (EXCEL_EPOCH + timedelta(days=value)).isoformat(timespec="milliseconds") + "Z"


def load_fixture(path):
    with ZipFile(ROOT / path) as archive:
        book = json.loads(archive.read("workbook.json"))
        sheets = {}
        scripts = {}
        for name in archive.namelist():
            if not name.startswith("sheets/"):
                continue
            sheet = json.loads(archive.read(name))
            cells = {}
            for cell in sheet["cells"]:
                coord = (cell["row"], cell["column"])
                content = cell["content"]
                if content["kind"] == "python":
                    scripts[f'{sheet["name"]}!{coord[0]}:{coord[1]}'] = content["block"]["Source"]
                elif content["kind"] == "formula":
                    cells[coord] = cell.get("cache", {}).get("value")
                else:
                    cells[coord] = content.get("value")
            sheets[Path(name).stem] = (sheet["name"], cells)

    snapshots = {}
    for table in book["names"]:
        if not table["IsTable"]:
            continue
        area, design = table["Area"], table["Design"]
        if not design["HeaderRow"]:
            raise ValueError("This probe expects visible fixture headers")
        _, cells = sheets[area["SheetId"]]
        row, col = area["Row"], area["Column"]
        columns = range(col, col + area["Columns"])
        headers = [cells[row, c] for c in columns]
        rows = []
        for r in range(row + 1, row + area["Rows"] - int(design["TotalRow"])):
            values = []
            for header, c in zip(headers, columns):
                value = cells.get((r, c))
                if header in {"OrderDate", "Opened UTC"} and isinstance(value, (int, float)):
                    value = iso(value)
                values.append(value)
            rows.append(values)
        snapshots[table["Name"] + "[#Data]"] = {"headers": headers, "rows": rows}

    if "Support-SLA" in path:
        by_name = dict(sheets.values())
        snapshots["Dashboard!P4"] = {"value": iso(by_name["Dashboard"][4, 16])}
        analysis = by_name["Analysis"]
        # Saved spill rows are an independent check on the new Python output.
        pivots = [[
            [analysis.get((r, c)) for c in range(3, 9)]
            for r in range(start, start + 5)
        ] for start in (27, 36)]
    else:
        pivots = []
    return {"scripts": scripts, "snapshots": snapshots, "baselinePivots": pivots}


json.dump(load_fixture(sys.argv[1]), sys.stdout, ensure_ascii=True, allow_nan=False)
