# m2py_core.py
# Pragmatic Power Query (M) -> pandas code generator
# Focused on a subset of common transforms and simple literals.

from __future__ import annotations
import re
from typing import Dict, List, Optional

import pandas as pd
import numpy as np


def _normalize_var(name: str) -> str:
    """
    Turn an M binding name into a Python identifier we can reuse safely.
    Handles #"Changed Type" → Changed_Type; strips surrounding quotes/hashes.
    """
    name = name.strip()
    if name.startswith('#"') and name.endswith('"'):
        name = name[2:-1]
    name = name.strip().strip('"').strip("'")
    # Replace disallowed chars with underscores
    name = re.sub(r"[^0-9A-Za-z_]", "_", name)
    # Avoid leading digits
    if re.match(r"^\d", name):
        name = "_" + name
    # Avoid empty
    return name or "step"


def _split_let_body(body: str) -> List[str]:
    """
    Split the 'let ... in' body into 'Name = Expr' steps, at commas that are not
    inside (), [], {}, or strings.
    """
    parts: List[str] = []
    buf = []
    depth_par = depth_br = depth_curly = 0
    in_str = False
    str_delim = ""
    esc = False

    for ch in body:
        buf.append(ch)
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif str_delim in ("'", '"') and ch == str_delim:
                in_str = False
            elif str_delim in ("'''", '"""') and "".join(buf[-3:]) == str_delim:
                in_str = False
        else:
            if ch == '"':
                in_str = True
                str_delim = '"'
            elif ch == "'":
                in_str = True
                str_delim = "'"
            elif ch == "(":
                depth_par += 1
            elif ch == ")":
                depth_par -= 1
            elif ch == "[":
                depth_br += 1
            elif ch == "]":
                depth_br -= 1
            elif ch == "{":
                depth_curly += 1
            elif ch == "}":
                depth_curly -= 1
            elif ch == "," and depth_par == depth_br == depth_curly == 0:
                # step boundary
                parts.append("".join(buf[:-1]).strip())
                buf = []

    tail = "".join(buf).strip()
    if tail:
        parts.append(tail)
    # Drop empties
    return [p for p in (p.strip().rstrip(",") for p in parts) if p]


def convert_m_to_python(m_code: str, query_name: str = "Result") -> str:
    """
    Convert a subset of Power Query (M) into executable pandas code.
    Pragmatic approach: step-by-step pattern matching; not a full M parser.
    """
    m_code = m_code or ""
    py: List[str] = [
        "import pandas as pd",
        "import numpy as np",
        "",
    ]

    # Find LET ... IN
    let_match = re.search(r"\blet\b(.*)\bin\b\s*([#\"A-Za-z0-9_\.]+)\s*$", m_code, flags=re.S | re.I)
    if let_match:
        let_body = let_match.group(1)
        out_name_raw = let_match.group(2)
    else:
        # Treat whole string as a single assignment if there's '='
        let_body = m_code.strip()
        out_name_raw = query_name

    steps = _split_let_body(let_body)
    env: Dict[str, str] = {}
    last_df: Optional[str] = None

    def add(line: str) -> None:
        py.append(line)

    # Emit unsupported as a comment but keep pipeline alive
    def unsupported(lhs: str, rhs: str):
        nonlocal last_df
        add(f"# Unsupported: {lhs} = {rhs}")
        if last_df:
            add(f"{lhs} = {last_df}.copy()")
        else:
            add(f"{lhs} = None  # unsupported start")
        last_df = lhs

    for raw in steps:
        if not raw or "=" not in raw:
            continue
        lhs_raw, rhs = raw.split("=", 1)
        lhs_raw = lhs_raw.strip()
        rhs = rhs.strip().rstrip(",")
        lhs = _normalize_var(lhs_raw)

        # --- Csv.Document ----------------------------------------------------
        # Csv.Document(File.Contents("file"), [Delimiter=";", Encoding=65001, QuoteStyle=QuoteStyle.Csv])
        m = re.search(r'Csv\.Document\(\s*File\.Contents\("([^"]+)"\)\s*(?:,\s*\[([^\]]*)\])?\s*\)\s*$', rhs)
        if m:
            csv_path = m.group(1)
            opts = m.group(2) or ""
            sep = None
            enc = None
            quote_none = False
            m_delim = re.search(r'Delimiter\s*=\s*"([^"]+)"', opts)
            if m_delim:
                sep = m_delim.group(1)
            m_enc = re.search(r'Encoding\s*=\s*(\d+)', opts)
            if m_enc:
                cp = m_enc.group(1)
                enc = "utf-8" if cp == "65001" else ("cp1252" if cp == "1252" else None)
            if re.search(r'QuoteStyle\s*=\s*QuoteStyle\.None', opts):
                quote_none = True
            args = [f"'{csv_path}'", "header=None"]
            if sep is not None:
                args.append(f"sep='{sep}'")
            if enc is not None:
                args.append(f"encoding='{enc}'")
            if quote_none:
                args.append("quoting=3")  # csv.QUOTE_NONE
            add(f"{lhs} = pd.read_csv({', '.join(args)})")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # --- Table.PromoteHeaders -------------------------------------------
        # Table.PromoteHeaders(Source, [PromoteAllScalars=true])
        m = re.search(r'Table\.PromoteHeaders\(\s*([^\),]+)\s*(?:,\s*\[[^\]]*\])?\s*\)\s*$', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            add(f"{lhs} = {src}.copy()")
            add(f"{lhs}.columns = {lhs}.iloc[0]")
            add(f"{lhs} = {lhs}.iloc[1:].reset_index(drop=True)")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # --- TransformColumnTypes -------------------------------------------
        # Table.TransformColumnTypes(Source, {{"col", type text}, {"num", type number}})
        m = re.search(r"Table\.TransformColumnTypes\(\s*([^,]+)\s*,\s*\{\{(.+?)\}\}\s*\)\s*$", rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            specs = m.group(2)
            add(f"{lhs} = {src}.copy()")
            pairs = re.findall(r'\{\s*"([^"]+)"\s*,\s*([A-Za-z0-9_\.]+)\s*\}', specs)
            for col, typ in pairs:
                t = typ.split(".")[-1].lower()
                if t in ("text",):
                    add(f"{lhs}['{col}'] = {lhs}['{col}'].astype('string')")
                elif t in ("number", "double", "single", "decimal"):
                    add(f"{lhs}['{col}'] = {lhs}['{col}'].astype('float')")
                elif t in ("int64", "int32", "int16", "int8"):
                    add(f"{lhs}['{col}'] = pd.to_numeric({lhs}['{col}'], errors='coerce').astype('Int64')")
                elif t in ("date", "datetime", "datetimezone"):
                    add(f"{lhs}['{col}'] = pd.to_datetime({lhs}['{col}'], errors='coerce')")
                elif t in ("logical",):
                    add(f"{lhs}['{col}'] = {lhs}['{col}'].astype('boolean')")
                else:
                    add(f"# {lhs}['{col}'] = {lhs}['{col}'].astype('object')  # unhandled type: {t}")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # --- Table.FromRecords (minimal literal support) --------------------
        # Table.FromRecords({ [A=1, B="x"], [A=2, Rec=[x=1,y=2]] })
        m = re.search(r'Table\.FromRecords\(\s*\{(.*)\}\s*\)\s*$', rhs, flags=re.S)
        if m:
            body = m.group(1)

            def _parse_record_literal(txt: str) -> dict:
                d: Dict[str, object] = {}
                # split on commas not in nested [..]
                parts: List[str] = []
                depth = 0
                tok = ""
                for ch in txt:
                    if ch == "[":
                        depth += 1
                    elif ch == "]":
                        depth -= 1
                    if ch == "," and depth == 0:
                        parts.append(tok)
                        tok = ""
                    else:
                        tok += ch
                if tok.strip():
                    parts.append(tok)
                for p in parts:
                    if "=" not in p:
                        continue
                    k, v = p.split("=", 1)
                    key = k.strip().strip('"')
                    v = v.strip()
                    if v.startswith("[") and v.endswith("]"):
                        inner = v[1:-1]
                        val = _parse_record_literal(inner)
                    elif v.startswith('"') and v.endswith('"'):
                        val = v[1:-1]
                    else:
                        try:
                            val = int(v)
                        except ValueError:
                            try:
                                val = float(v)
                            except ValueError:
                                val = v
                    d[key] = val
                return d

            # split outer { [..], [..], ... }
            recs: List[dict] = []
            buf = ""
            depth = 0
            i = 0
            while i < len(body):
                ch = body[i]
                if ch == "[":
                    depth += 1
                elif ch == "]":
                    depth -= 1
                buf += ch
                if depth == 0 and buf.strip():
                    if buf.strip().endswith("]"):
                        inner = buf.strip()
                        if inner.startswith("[") and inner.endswith("]"):
                            inner = inner[1:-1]
                        recs.append(_parse_record_literal(inner))
                        # skip following commas/space
                        j = i + 1
                        while j < len(body) and body[j] in " ,\n\t\r":
                            j += 1
                        i = j - 1
                        buf = ""
                i += 1

            add(f"{lhs} = pd.DataFrame({recs!r})")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # --- #table (minimal literal support) -------------------------------
        # #table(type table [A=number, B=text], {{1,"x"},{2,"y"}})
        # #table({"A","B"}, {{1,"x"},{2,"y"}})
        m = re.search(
            r'#table\(\s*(?:type\s+table\s+\[([^\]]+)\]|(\{.*?\}))\s*,\s*\{\{(.*?)\}\}\s*\)\s*$',
            rhs, flags=re.S
        )
        if m:
            cols_spec = m.group(1)    # e.g. 'A=number, B=text'
            cols_list = m.group(2)    # e.g. '{"A","B"}'
            rows_raw  = m.group(3)    # content between the outer {{ ... }}

            if cols_spec:
                cols = [part.split("=", 1)[0].strip().strip('"') for part in re.split(r",\s*", cols_spec)]
            else:
                cols = re.findall(r'"([^"]+)"', cols_list)

            # Extract row tokens robustly
            rows_tokens: List[str] = []
            depth = 0
            buf = ""
            for ch in rows_raw:
                if ch == "{":
                    depth += 1
                    if depth == 1:
                        buf = ""
                        continue
                if ch == "}":
                    depth -= 1
                    if depth == 0:
                        rows_tokens.append(buf)
                        continue
                if depth >= 1:
                    buf += ch

            # Fallback: if no braces found inside (edge cases), split on '},{'
            if not rows_tokens:
                s = rows_raw.strip()
                if s:
                    chunks = re.split(r'\}\s*,\s*\{', s)
                    if chunks:
                        chunks[0] = chunks[0].lstrip('{')
                        chunks[-1] = chunks[-1].rstrip('}')
                        rows_tokens = chunks

            # Parse each token into values (respect quoted commas)
            rows: List[dict] = []
            for t in rows_tokens:
                vals: List[str] = []
                q = False
                cur = ""
                i = 0
                while i < len(t):
                    ch = t[i]
                    if ch == '"' and (i == 0 or t[i-1] != "\\"):
                        q = not q
                        cur += ch
                    elif ch == "," and not q:
                        vals.append(cur.strip())
                        cur = ""
                    else:
                        cur += ch
                    i += 1
                if cur.strip():
                    vals.append(cur.strip())

                parsed: List[object] = []
                for v in vals:
                    v = v.strip()
                    if v.startswith('"') and v.endswith('"'):
                        parsed.append(v[1:-1])
                    else:
                        try:
                            parsed.append(int(v))
                        except ValueError:
                            try:
                                parsed.append(float(v))
                            except ValueError:
                                parsed.append(v)

                rows.append({cols[i]: (parsed[i] if i < len(parsed) else None) for i in range(len(cols))})

            add(f"{lhs} = pd.DataFrame({rows!r})")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # --- ExpandRecordColumn ---------------------------------------------
        # Table.ExpandRecordColumn(Source, "Rec", {"x","y"}, {"X","Y"})
        m = re.search(
            r'Table\.ExpandRecordColumn\(\s*([^,]+)\s*,\s*"([^"]+)"\s*,\s*\{([^\}]*)\}\s*,\s*\{([^\}]*)\}\s*\)\s*$',
            rhs
        )
        if m:
            src = _normalize_var(m.group(1).strip())
            col = m.group(2)
            fields = re.findall(r'"([^"]+)"', m.group(3))
            newnames = re.findall(r'"([^"]+)"', m.group(4))
            if not newnames or len(newnames) != len(fields):
                newnames = fields[:]
            add(f"{lhs} = {src}.drop(columns=['{col}'], errors='ignore').copy()")
            add(f"_exp = {src}['{col}'].apply(lambda x: pd.Series(x) if isinstance(x, dict) else pd.Series(dtype='object'))")
            if fields:
                add(f"_exp = _exp[{fields!r}]")
                add(f"_exp = _exp.rename(columns={{**dict(zip({fields!r}, {newnames!r}))}})")
            add(f"{lhs} = {lhs}.join(_exp)")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # --- ExpandTableColumn ----------------------------------------------
        # Table.ExpandTableColumn(Source, "Tbl", {"x","y"}, {"X","Y"})
        m = re.search(
            r'Table\.ExpandTableColumn\(\s*([^,]+)\s*,\s*"([^"]+)"\s*,\s*\{([^\}]*)\}\s*,\s*\{([^\}]*)\}\s*\)\s*$',
            rhs
        )
        if m:
            src = _normalize_var(m.group(1).strip())
            col = m.group(2)
            fields = re.findall(r'"([^"]+)"', m.group(3))
            newnames = re.findall(r'"([^"]+)"', m.group(4))
            if not newnames or len(newnames) != len(fields):
                newnames = fields[:]
            add(f"{lhs} = {src}.copy()")
            add(f"_tbl = {lhs}.pop('{col}') if '{col}' in {lhs}.columns else pd.Series(index={lhs}.index, dtype='object')")
            add(f"_tbl = _tbl.apply(lambda t: t if isinstance(t, (list, tuple)) else ([] if t is None else [t]))")
            add(f"_tbl = _tbl.explode()")
            add(f"_df = pd.DataFrame(_tbl.tolist()) if not _tbl.empty else pd.DataFrame()")
            if fields:
                add(f"_df = _df.reindex(columns={fields!r})")
            if newnames and fields:
                add(f"_df = _df.rename(columns={{**dict(zip({fields!r}, {newnames!r}))}})")
            add(f"{lhs} = {lhs}.join(_df.reset_index(drop=True))")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # --- Join (single or multi-key) -------------------------------------
        # Table.Join(A, "k", B, "k", JoinKind.Inner)
        # Table.Join(A, {"k1","k2"}, B, {"k1","k2"}, JoinKind.LeftOuter)
        m = re.search(
            r'Table\.Join\(\s*([^,]+)\s*,\s*("([^"]+)"|\{[^\}]+\})\s*,\s*([^,]+)\s*,\s*("([^"]+)"|\{[^\}]+\})\s*,\s*JoinKind\.([A-Za-z]+)\s*\)\s*$',
            rhs
        )
        if m:
            left = _normalize_var(m.group(1).strip())
            left_keys_raw = m.group(2)
            right = _normalize_var(m.group(4).strip())
            right_keys_raw = m.group(5)
            kind = m.group(7)
            def _parse_keys(tok: str) -> List[str]:
                tok = tok.strip()
                if tok.startswith("{"):
                    return re.findall(r'"([^"]+)"', tok)
                return [tok.strip('"')]
            lk = _parse_keys(left_keys_raw)
            rk = _parse_keys(right_keys_raw)
            how = {"Inner": "inner", "LeftOuter": "left", "RightOuter": "right", "FullOuter": "outer"}.get(kind, "inner")
            add(f"{lhs} = pd.merge({left}, {right}, how='{how}', left_on={lk!r}, right_on={rk!r})")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # --- Table.SelectRows ------------------------------------------------
        # Table.SelectRows(Source, each [B] = "X" and [A] >= 2)
        m = re.search(r'Table\.SelectRows\(\s*([^,]+)\s*,\s*each\s+(.+)\)\s*$', rhs, flags=re.S)
        if m:
            src = _normalize_var(m.group(1).strip())
            cond = m.group(2).strip()

            # Replace [Col] -> src['Col']
            cond = re.sub(r'\[([^\]]+)\]', lambda mm: f"{src}['{mm.group(1)}']", cond)

            # Operators / keywords
            cond = re.sub(r'<>', '!=', cond)
            cond = re.sub(r'(?<![<>=!])=(?!=)', '==', cond)  # bare '=' -> '=='
            cond = re.sub(r'\band\b', '&', cond, flags=re.I)
            cond = re.sub(r'\bor\b', '|', cond, flags=re.I)
            cond = re.sub(r'\bnot\b', '~', cond, flags=re.I)

            # Literals
            cond = re.sub(r'\bnull\b', 'None', cond, flags=re.I)
            cond = re.sub(r'\btrue\b', 'True', cond, flags=re.I)
            cond = re.sub(r'\bfalse\b', 'False', cond, flags=re.I)

            add(f"{lhs} = {src}[{cond}].copy()")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Table.Sort ------------------------------------------------------
        # Table.Sort(Filtered, {{"A", Order.Descending}, {"B", Order.Ascending}})
        m = re.search(r'Table\.Sort\(\s*([^,]+)\s*,\s*\{(.+)\}\s*\)\s*$', rhs, flags=re.S)
        if m:
            src = _normalize_var(m.group(1).strip())
            spec = m.group(2)
            pairs = re.findall(r'\{\s*"([^"]+)"\s*,\s*Order\.(Ascending|Descending)\s*\}', spec)
            cols = [c for c, _ in pairs] if pairs else []
            asc = [True if order == 'Ascending' else False for _, order in pairs] if pairs else True
            if cols:
                add(f"{lhs} = {src}.sort_values(by={cols!r}, ascending={asc!r}).reset_index(drop=True)")
            else:
                add(f"{lhs} = {src}.copy()")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Fallback --------------------------------------------------------
        unsupported(lhs_raw, rhs)

    # Final binding
    out_name = _normalize_var(query_name)
    # If the 'in ...' named a particular step, prefer it
    if let_match:
        out_ref = env.get(out_name_raw, last_df or out_name)
    else:
        out_ref = last_df or out_name
    add(f"{out_name} = {out_ref}")
    return "\n".join(py)
