"""Bounded Pyodide adapter for explicitly requested workbook Python runs."""
import ast
import base64
import copy
import hashlib
import io
import json
import itertools
import math
import re
import struct
import traceback
import types
from collections import Counter
from contextlib import redirect_stderr, redirect_stdout
from collections import OrderedDict

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import numpy as np
import pandas as pd
from matplotlib.backends.backend_agg import FigureCanvasAgg

MAX_SOURCE_BYTES = 1_048_576
MAX_LOG_BYTES = 1_048_576
MAX_TABLE_CELLS = 100_000
MAX_RESULT_BYTES = 16 * 1024 * 1024
MAX_FRAME_BYTES = 32 * 1024 * 1024
MAX_RETAINED_FIGURES = 8
_retained_figures = OrderedDict()


def _copy_figure(source):
    """Deep-copy a live Figure, recreating Matplotlib's opaque counters without sharing them."""
    counter_type = type(itertools.count())
    memo, visited = {}, set()

    def seed_counters(value):
        identity = id(value)
        if identity in visited:
            return
        visited.add(identity)
        if isinstance(value, counter_type):
            match = re.fullmatch(r"count\((-?\d+)(?:,\s*(-?\d+))?\)", repr(value))
            if not match:
                raise ValueError("Matplotlib Figure contains an unsupported live counter.")
            memo[identity] = itertools.count(*(int(item) for item in match.groups() if item is not None))
            return
        if isinstance(value, dict):
            for key, item in value.items():
                seed_counters(key)
                seed_counters(item)
            return
        if isinstance(value, (list, tuple, set, frozenset)):
            for item in value:
                seed_counters(item)
            return
        if isinstance(value, (str, bytes, bytearray, int, float, bool, type(None), type,
                              types.FunctionType, types.BuiltinFunctionType, types.ModuleType)):
            return
        attributes = getattr(value, "__dict__", None)
        if isinstance(attributes, dict):
            seed_counters(attributes)

    seed_counters(source)
    return copy.deepcopy(source, memo)


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


def _encode(value, *, retain_figure=True):
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
        if (actual_width < 1 or actual_height < 1 or actual_width > 8192 or actual_height > 8192 or
                actual_width * actual_height > 16_000_000):
            raise ValueError("Figure PNG exceeds the 16-million-pixel output limit.")
        figure_id = hashlib.sha256(raw).hexdigest().upper()
        if retain_figure:
            _retained_figures[figure_id] = value
            _retained_figures.move_to_end(figure_id)
            while len(_retained_figures) > MAX_RETAINED_FIGURES:
                _retained_figures.popitem(last=False)
        return {"kind": "figure", "figure": {"pngBase64": base64.b64encode(raw).decode("ascii"),
                "width": actual_width, "height": actual_height, "figureId": figure_id}}

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


def render_figure(figure_id, width, height, dpi):
    """Redraw one worker-retained Figure from a private copy at a bounded canvas size."""
    stdout, stderr = BoundedText(MAX_LOG_BYTES), BoundedText(MAX_LOG_BYTES)
    try:
        if (not isinstance(figure_id, str) or len(figure_id) != 64 or
                any(character not in "0123456789ABCDEF" for character in figure_id)):
            raise ValueError("Figure identity is invalid.")
        if (type(width) is not int or type(height) is not int or not 1 <= width <= 8192 or
                not 1 <= height <= 8192 or width * height > 16_000_000 or
                type(dpi) not in (int, float) or not math.isfinite(dpi) or not 0 < dpi <= 384):
            raise ValueError("Invalid figure render dimensions.")
        baseline = _retained_figures.get(figure_id)
        if baseline is None:
            raise ValueError("The live Figure is no longer available. Rebuild it from its Python source.")
        _retained_figures.move_to_end(figure_id)
        figure = _copy_figure(baseline)
        try:
            figure.set_dpi(float(dpi))
            size = [width / dpi, height / dpi]
            result = None
            for _ in range(4):
                figure.set_size_inches(size, forward=False)
                FigureCanvasAgg(figure)
                result = _encode(figure, retain_figure=False)
                actual_width = result["figure"]["width"]
                actual_height = result["figure"]["height"]
                if actual_width == width and actual_height == height:
                    break
                size[0] += (width - actual_width + (.001 if width > actual_width else -.001 if width < actual_width else 0)) / dpi
                size[1] += (height - actual_height + (.001 if height > actual_height else -.001 if height < actual_height else 0)) / dpi
            if result is None or result["figure"]["width"] != width or result["figure"]["height"] != height:
                raise ValueError("Figure redraw did not match the requested canvas dimensions.")
        finally:
            plt.close(figure)
        output = {"status": "succeeded", "result": result, "stdout": "", "stderr": "", "outputTruncated": False}
    except BaseException:
        error = traceback.format_exc()
        if len(error) > 32767:
            error = error[-32767:]
        output = {"status": "failed", "stdout": stdout.getvalue(), "stderr": stderr.getvalue(),
                  "error": error, "outputTruncated": stdout.truncated or stderr.truncated}

    if len(json.dumps(output, ensure_ascii=False, allow_nan=False).encode("utf-8")) > MAX_FRAME_BYTES:
        return json.dumps({"status": "failed", "stdout": "", "stderr": "", "error": "Figure output exceeds 32 MiB.",
                           "outputTruncated": False})
    return json.dumps(output, ensure_ascii=False, allow_nan=False)


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
