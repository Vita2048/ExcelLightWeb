"""Bounded Pyodide adapter for explicitly requested workbook Python runs."""
import ast
import base64
import io
import json
import math
import struct
import traceback
from collections import Counter
from contextlib import redirect_stderr, redirect_stdout

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import numpy as np
import pandas as pd

MAX_SOURCE_BYTES = 1_048_576
MAX_LOG_BYTES = 1_048_576
MAX_TABLE_CELLS = 100_000
MAX_RESULT_BYTES = 16 * 1024 * 1024
MAX_FRAME_BYTES = 32 * 1024 * 1024


class BoundedText(io.StringIO):
    def __init__(self, limit):
        super().__init__()
        self.limit = limit
        self.bytes_written = 0
        self.truncated = False

    def write(self, value):
        encoded = str(value).encode("utf-8")
        remaining = self.limit - self.bytes_written
        if remaining <= 0:
            self.truncated = True
            return len(value)
        if len(encoded) > remaining:
            encoded = encoded[:remaining].decode("utf-8", errors="ignore").encode("utf-8")
            self.truncated = True
        self.bytes_written += len(encoded)
        super().write(encoded.decode("utf-8"))
        return len(value)


def _preflight(source, declared_calls):
    if not isinstance(source, str) or len(source.encode("utf-8")) > MAX_SOURCE_BYTES:
        raise ValueError("Source exceeds 1 MiB.")
    tree = ast.parse(source, filename="<python-cell>", mode="exec")
    parent = {}
    for node in ast.walk(tree):
        for child in ast.iter_child_nodes(node):
            parent[child] = node

    calls = []
    xl_functions = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "xl":
            if (len(node.args) != 1 or not isinstance(node.args[0], ast.Constant)
                    or type(node.args[0].value) is not str or len(node.keywords) > 1
                    or any(item.arg != "headers" or not isinstance(item.value, ast.Constant)
                           or type(item.value.value) is not bool for item in node.keywords)):
                raise ValueError('Use xl("A1:B3", headers=True/False) with literal arguments; dynamic references are unsupported.')
            cursor = node
            while cursor in parent:
                cursor = parent[cursor]
                if isinstance(cursor, ast.JoinedStr):
                    raise ValueError("Calls to xl inside interpolated strings are unsupported; assign the value first.")
            headers = node.keywords[0].value.value if node.keywords else False
            calls.append((node.args[0].value, headers))
            xl_functions.add(id(node.func))

    for node in ast.walk(tree):
        if (isinstance(node, ast.Name) and node.id == "xl" and id(node) not in xl_functions
                or isinstance(node, ast.arg) and node.arg == "xl"
                or isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)) and node.name == "xl"
                or isinstance(node, ast.alias) and (node.asname or node.name.split(".")[0]) == "xl"
                or isinstance(node, ast.Attribute) and node.attr == "xl"
                or isinstance(node, (ast.Global, ast.Nonlocal)) and "xl" in node.names):
            raise ValueError("xl is reserved for direct literal calls; aliases and shadowing are unsupported.")

    declared = Counter((item["reference"], item["headers"]) for item in declared_calls)
    if Counter(calls) != declared:
        raise ValueError("Python xl references changed before execution; run the current source again.")
    return tree


def _scalar(value, *, in_table=False, notices=None):
    if isinstance(value, np.generic):
        value = value.item()
    if in_table and (value is pd.NA or value is pd.NaT or
                     isinstance(value, (float, np.floating)) and math.isnan(float(value))):
        if notices is not None and "Missing table values were converted to blank cells." not in notices:
            notices.append("Missing table values were converted to blank cells.")
        return None
    if value is None or isinstance(value, (bool, str, int)):
        return value
    if isinstance(value, float) and math.isfinite(value):
        return value
    raise ValueError("Unsupported result value: " + type(value).__name__)


def _table(headers, rows):
    notices = []
    clean_headers = None if headers is None else [str(value) for value in headers]
    if clean_headers is not None and (not clean_headers or any(not item.strip() for item in clean_headers)
                                      or len(set(clean_headers)) != len(clean_headers)):
        raise ValueError("Table output needs nonblank unique column labels.")
    clean_rows = [[_scalar(value, in_table=True, notices=notices) for value in row] for row in rows]
    columns = len(clean_headers) if clean_headers is not None else (len(clean_rows[0]) if clean_rows else 0)
    if columns < 1 or any(len(row) != columns for row in clean_rows):
        raise ValueError("Expected a nonempty rectangular table result.")
    if len(clean_rows) * columns > MAX_TABLE_CELLS:
        raise ValueError("Table exceeds the 100,000-cell result limit.")
    result = {"kind": "table", "headers": clean_headers, "rows": clean_rows, "notices": notices}
    if len(json.dumps(result, ensure_ascii=False).encode("utf-8")) > MAX_RESULT_BYTES:
        raise ValueError("Table output exceeds 16 MiB.")
    return result


def _encode(value):
    if isinstance(value, Figure):
        width_in, height_in = value.get_size_inches()
        width = round(width_in * value.dpi)
        height = round(height_in * value.dpi)
        if width < 1 or height < 1 or width > 8192 or height > 8192 or width * height > 16_000_000:
            raise ValueError("Figure exceeds the 16-million-pixel output limit.")
        png = io.BytesIO()
        value.savefig(png, format="png")
        raw = png.getvalue()
        if len(raw) > MAX_RESULT_BYTES:
            raise ValueError("Figure PNG exceeds 16 MiB.")
        if len(raw) < 24 or raw[:8] != b"\x89PNG\r\n\x1a\n":
            raise ValueError("Figure did not produce a valid PNG.")
        actual_width, actual_height = struct.unpack(">II", raw[16:24])
        if actual_width != width or actual_height != height or actual_width * actual_height > 16_000_000:
            raise ValueError("Figure dimensions did not match the requested canvas.")
        return {"kind": "figure", "figure": {"pngBase64": base64.b64encode(raw).decode("ascii"),
                "width": actual_width, "height": actual_height}}

    if isinstance(value, pd.Series):
        name = value.name if value.name is not None else "Value"
        value = value.to_frame(name=str(name))
    if isinstance(value, pd.DataFrame):
        headers = [str(column) for column in value.columns]
        return _table(headers, list(value.itertuples(index=False, name=None)))
    if isinstance(value, np.ndarray):
        if value.ndim == 0:
            return {"kind": "scalar", "value": _scalar(value.item())}
        if value.ndim != 2:
            raise ValueError("Expected a two-dimensional array result.")
        value = value.tolist()
    if isinstance(value, (list, tuple)):
        rows = list(value)
        if not rows or any(not isinstance(row, (list, tuple)) for row in rows):
            raise ValueError("Expected a rectangular two-dimensional result.")
        return _table(None, rows)
    return {"kind": "scalar", "value": _scalar(value)}


def _as_dataframe(item, headers=False):
    if item.get("scalarCell") is True:
        return item["scalar"]
    rows = item.get("rows", [])
    if not headers and len(rows) == 1 and len(rows[0]) == 1:
        return rows[0][0]
    if headers:
        labels = item.get("headers")
        if labels is None:
            if not rows:
                raise ValueError("xl(headers=True) needs a header row.")
            labels, rows = rows[0], rows[1:]
        if any(not isinstance(label, str) or not label.strip() for label in labels) or len(set(labels)) != len(labels):
            raise ValueError("xl headers must be nonblank unique text; correct the first row or use headers=False.")
        return pd.DataFrame(rows, columns=labels)
    return pd.DataFrame(rows)


def run(source, input_json):
    stdout, stderr = BoundedText(MAX_LOG_BYTES), BoundedText(MAX_LOG_BYTES)
    try:
        payload = json.loads(input_json)
        inputs = payload["values"]
        calls = payload["calls"]
        bindings = payload["bindings"]
        tree = _preflight(source, calls)
        for call in calls:
            if call["key"] not in inputs:
                raise ValueError("Python input snapshot is missing for " + call["reference"] + ".")
            item = inputs[call["key"]]
            if call["headers"]:
                _as_dataframe(item, True)  # Validate all headers before user statements.
        by_call = {(item["reference"], item["headers"]): item for item in calls}

        def xl(reference, *, headers=False):
            metadata = by_call.get((reference, headers))
            if metadata is None:
                raise ValueError("xl reference was not supplied by the host; use direct literal calls.")
            return _as_dataframe(inputs[metadata["key"]], headers)

        namespace = {"xl": xl}
        for binding in bindings:
            if binding["key"] not in inputs:
                raise ValueError("Legacy Python binding snapshot is missing.")
            namespace[binding["name"]] = _as_dataframe(inputs[binding["key"]])

        with redirect_stdout(stdout), redirect_stderr(stderr):
            last = tree.body[-1] if tree.body else None
            if isinstance(last, ast.Expr):
                tree.body.pop()
                exec(compile(tree, "<python-cell>", "exec"), namespace, namespace)
                value = eval(compile(ast.Expression(last.value), "<python-cell>", "eval"), namespace, namespace)
            else:
                exec(compile(tree, "<python-cell>", "exec"), namespace, namespace)
                if "result" not in namespace:
                    raise ValueError("No result or final expression.")
                value = namespace["result"]
            result = _encode(value)
        output = {"status": "succeeded", "result": result, "stdout": stdout.getvalue(),
                  "stderr": stderr.getvalue(), "outputTruncated": stdout.truncated or stderr.truncated}
    except BaseException:
        message = traceback.format_exc()
        if len(message) > 32767:
            message = message[-32767:]
        output = {"status": "failed", "stdout": stdout.getvalue(), "stderr": stderr.getvalue(),
                  "error": message, "outputTruncated": stdout.truncated or stderr.truncated}
    finally:
        plt.close("all")

    encoded = json.dumps(output, ensure_ascii=False, allow_nan=False)
    if len(encoded.encode("utf-8")) > MAX_FRAME_BYTES:
        fallback = {"status": "failed", "stdout": "", "stderr": "", "error": "Python result exceeds 32 MiB.",
                    "outputTruncated": False}
        return json.dumps(fallback)
    return encoded
