"""Transport adapter for the browser feasibility probe."""
import ast
import base64
import io
import json
import math
from contextlib import redirect_stderr, redirect_stdout

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import numpy as np
import pandas as pd


def scalar(value, *, in_table=False):
    if isinstance(value, np.generic):
        value = value.item()
    if in_table and (value is pd.NA or value is pd.NaT or
                     isinstance(value, float) and math.isnan(value)):
        return None
    if value is None or isinstance(value, (bool, str, int)):
        return value
    if isinstance(value, float) and math.isfinite(value):
        return value
    raise ValueError(f"Unsupported result value: {type(value).__name__}")


def table(rows, *, headers=None):
    if not rows or not rows[0] or any(len(row) != len(rows[0]) for row in rows):
        raise ValueError("Expected a nonempty rectangular result")
    if len(rows) * len(rows[0]) > 100_000:
        raise ValueError("Table exceeds the result cell limit")
    return {"kind": "table", "headers": headers,
            "rows": [[scalar(value, in_table=True) for value in row] for row in rows]}


def encode(value):
    if isinstance(value, Figure):
        png = io.BytesIO()
        value.savefig(png, format="png")
        return {"kind": "png", "base64": base64.b64encode(png.getvalue()).decode("ascii")}
    if isinstance(value, pd.Series):
        value = value.to_frame(name=value.name if value.name is not None else "Value")
    if isinstance(value, pd.DataFrame):
        headers = [str(column) for column in value.columns]
        return table(list(value.itertuples(index=False, name=None)), headers=headers) if len(value) else {
            "kind": "table", "headers": headers, "rows": []
        }
    if isinstance(value, (list, tuple, np.ndarray)):
        if isinstance(value, np.ndarray) and value.ndim != 2:
            raise ValueError("Expected a two-dimensional array")
        rows = list(value)
        if any(not isinstance(row, (list, tuple, np.ndarray)) for row in rows):
            raise ValueError("Expected a rectangular result")
        return table([list(row) for row in rows])
    return {"kind": "scalar", "value": scalar(value)}


def run(source, snapshots_json):
    snapshots = json.loads(snapshots_json)

    def xl(reference, *, headers=False):
        item = snapshots[reference]
        if "rows" not in item:
            return item["value"]
        if not headers:
            return pd.DataFrame(item["rows"])
        labels = item["headers"]
        if any(not isinstance(label, str) or not label.strip() for label in labels) or len(set(labels)) != len(labels):
            raise ValueError(f"Invalid header for {reference}")
        return pd.DataFrame(item["rows"], columns=labels)

    body = ast.parse(source, filename="<cell>")
    namespace = {"xl": xl}
    out, err = io.StringIO(), io.StringIO()
    try:
        with redirect_stdout(out), redirect_stderr(err):
            last = body.body[-1] if body.body else None
            if isinstance(last, ast.Expr):
                body.body.pop()
                exec(compile(body, "<cell>", "exec"), namespace)
                value = eval(compile(ast.Expression(last.value), "<cell>", "eval"), namespace)
            else:
                exec(compile(body, "<cell>", "exec"), namespace)
                if "result" not in namespace:
                    raise ValueError("No result or final expression")
                value = namespace["result"]
            result = encode(value)
        return json.dumps({"result": result, "stdout": out.getvalue(), "stderr": err.getvalue()}, allow_nan=False)
    finally:
        plt.close("all")
