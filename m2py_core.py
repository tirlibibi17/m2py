import re
import base64
import zlib
import json

# --- helpers ---------------------------------------------------------------

def _normalize_var(name: str) -> str:
    """Power Query step names like #\"Changed Type\" -> Changed_Type"""
    name = name.strip()
    if name.startswith('#"') and name.endswith('"'):
        name = name[2:-1]
    return re.sub(r"\s+", "_", name)

def _replace_record_refs(expr: str, df_name: str) -> str:
    """
    Convert M column refs like [Col A] into df['Col A'] (for the given df_name).
    Keeps string literals intact.
    """
    out, i, in_str, quote = [], 0, False, None
    while i < len(expr):
        ch = expr[i]
        if not in_str and ch in ('"', "'"):
            in_str, quote = True, ch
            out.append(ch); i += 1
        elif in_str:
            out.append(ch)
            if ch == quote:
                in_str, quote = False, None
            i += 1
        else:
            m = re.match(r"\[([^\]]+)\]", expr[i:])
            if m:
                col = m.group(1).strip()
                out.append(f"{df_name}['{col}']")
                i += m.end()
            else:
                out.append(ch); i += 1
    return "".join(out)

def _translate_condition(cond: str, df_name: str) -> str:
    """
    Translate M row filter lambda like:
      [Country] = "FR" and [Age] > 30 or not [Active]
    -> precedence-safe pandas boolean expression.
    """
    s = _replace_record_refs(cond, df_name)
    s = re.sub(r"(?<![=!<>])=(?!=)", "==", s)                # = -> == (where safe)
    s = re.sub(r"\band\b", "&", s, flags=re.I)
    s = re.sub(r"\bor\b", "|", s, flags=re.I)
    s = re.sub(r"\bnot\b", "~", s, flags=re.I)
    s = re.sub(r"([)\]\}A-Za-z0-9_'\"]\s*(?:==|!=|>=|<=|>|<)\s*[^&|]+)", r"(\1)", s)
    if "&" in s and "|" in s:
        s = f"({s})"
    return s

# --- main ------------------------------------------------------------------

def convert_m_to_python(m_code: str) -> str:
    """
    Best-effort M -> pandas translator (MVP). Supports:
      - Excel.CurrentWorkbook(){[Name="TableName"]}[Content]
      - Csv.Document(File.Contents("file.csv")) + Table.PromoteHeaders
      - Table.RenameColumns
      - Table.SelectColumns / Table.RemoveColumns
      - Table.Sort
      - Table.SelectRows (and/or/not, =, !=, >, <, >=, <=)
      - Table.AddColumn (arithmetic/column refs 'each ...', Record.Field(_, "Col"))
      - Table.TransformColumnTypes (commented hint)
      - Table.Group (Sum/Avg/Min/Max/Count, multiple)
      - Table.Join (Inner/Left/Right/Full)
      - Table.Distinct
      - Table.RemoveRowsWithErrors
      - Table.FromRows(Json.Document(Binary.Decompress(...)))  -> static DataFrame literal
      - Table.Unpivot / Table.UnpivotOtherColumns / Table.Pivot
    """
    lines = [ln.rstrip() for ln in m_code.splitlines()]
    py = ["import pandas as pd", "import numpy as np", ""]
    env = {}
    last_df = None

    def add(line: str):
        py.append(line)

    for raw in lines:
        line = raw.strip()
        if not line or line.startswith("//"):
            continue
        if line.endswith(","):
            line = line[:-1].strip()
        if line.lower() in ("let", "in"):
            continue
        if "=" not in line:
            continue

        lhs_raw, rhs = map(str.strip, line.split("=", 1))
        lhs = _normalize_var(lhs_raw)

        # --- Enter Data: Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText(...)))) ----
        m = re.search(
            r'Table\.FromRows\(\s*Json\.Document\(\s*Binary\.Decompress\(\s*Binary\.FromText\("([^"]+)"\s*,\s*BinaryEncoding\.Base64\)\s*,\s*Compression\.Deflate\)\s*\)\s*,\s*.*?type\s+table\s*\[([^\]]+)\]\s*\)',
            rhs, re.DOTALL
        )
        if m:
            b64 = m.group(1)
            cols_spec = m.group(2)
            col_names = [c.strip().split("=", 1)[0].strip() for c in cols_spec.split(",")]
            try:
                raw = base64.b64decode(b64)

                # Try zlib header first, then raw DEFLATE
                try:
                    data = zlib.decompress(raw)
                except zlib.error:
                    data = zlib.decompress(raw, -zlib.MAX_WBITS)

                rows = json.loads(data.decode("utf-8"))

                dict_rows = []
                for r in rows:
                    r_list = list(r) if isinstance(r, (list, tuple)) else [r]
                    if len(r_list) < len(col_names):
                        r_list += [None] * (len(col_names) - len(r_list))
                    if len(r_list) > len(col_names):
                        r_list = r_list[:len(col_names)]
                    dict_rows.append({k: v for k, v in zip(col_names, r_list)})

                add(f"{lhs} = pd.DataFrame({dict_rows})")

            except Exception:
                # runtime fallback, but also try raw DEFLATE
                add("import base64, zlib, json  # runtime fallback for Enter Data")
                add(f"_b64 = {b64!r}")
                add(f"_raw = base64.b64decode(_b64)")
                add("_tmp = None")
                add("try:\n    _tmp = zlib.decompress(_raw)\nexcept zlib.error:\n    _tmp = zlib.decompress(_raw, -zlib.MAX_WBITS)")
                add("_rows = json.loads(_tmp.decode('utf-8'))")
                add(f"{lhs} = pd.DataFrame(_rows, columns={[c for c in col_names]})")

            env[lhs_raw] = lhs
            last_df = lhs
            continue
        # --- Sources --------------------------------------------------------
        # Excel.CurrentWorkbook(){[Name="Table1"]}[Content]
        m = re.search(r'Excel\.CurrentWorkbook\(\)\{\[Name="(.*?)"\]\}\[Content\]', rhs)
        if m:
            tbl = m.group(1)
            add(f"{lhs} = pd.read_excel('workbook.xlsx', sheet_name=None)['{tbl}']")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # Csv.Document(File.Contents("file.csv"))
        m = re.search(r'Csv\.Document\(File\.Contents\("([^"]+)"\)\)', rhs)
        if m:
            csv = m.group(1)
            add(f"{lhs} = pd.read_csv('{csv}', header=None)")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # Table.PromoteHeaders(Source)
        m = re.search(r'Table\.PromoteHeaders\(([^)]+)\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            add(f"{lhs} = {src}.copy()")
            add(f"{lhs}.columns = {lhs}.iloc[0]")
            add(f"{lhs} = {lhs}.iloc[1:].reset_index(drop=True)")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- RenameColumns --------------------------------------------------
        m = re.search(r'Table\.RenameColumns\(([^,]+),\s*\{\{(.+?)\}\}\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            pairs = [p.strip() for p in m.group(2).split("},")]
            rename_map = {}
            for p in pairs:
                cols = re.findall(r'"([^"]+)"', p)
                if len(cols) >= 2:
                    rename_map[cols[0]] = cols[1]
            add(f"{lhs} = {src}.rename(columns={rename_map})")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Select/Remove Columns -----------------------------------------
        m = re.search(r'Table\.SelectColumns\(([^,]+),\s*\{(.+?)\}\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            cols = [c.strip().strip('"') for c in m.group(2).split(",")]
            add(f"{lhs} = {src}[[{', '.join(repr(c) for c in cols)}]]")
            env[lhs_raw] = lhs; last_df = lhs; continue

        m = re.search(r'Table\.RemoveColumns\(([^,]+),\s*\{(.+?)\}\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            cols = [c.strip().strip('"') for c in m.group(2).split(",")]
            add(f"{lhs} = {src}.drop(columns={cols})")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Sort -----------------------------------------------------------
        m = re.search(r'Table\.Sort\(([^,]+),\s*\{(.+?)\}\)', rhs, re.DOTALL)
        if m:
            src = _normalize_var(m.group(1).strip())
            items = re.findall(r'\{"([^"]+)"\s*,\s*Order\.(Ascending|Descending)\}', m.group(2))
            cols = [c for c,_ in items]
            asc  = [o == "Ascending" for _,o in items]
            add(f"{lhs} = {src}.sort_values(by={cols}, ascending={asc})")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- SelectRows (filters) ------------------------------------------
        m = re.search(r'Table\.SelectRows\(([^,]+),\s*each\s*(.+)\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            cond_raw = m.group(2).strip()
            cond = _translate_condition(cond_raw, src)
            add(f"{lhs} = {src}[{cond}]")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- AddColumn ------------------------------------------------------
        # 1) Record.Field(_, "Col")
        m = re.search(r'Table\.AddColumn\(\s*([^,]+),\s*"([^"]+)",\s*each\s*Record\.Field\(_, "([^"]+)"\)\s*(?:,\s*type\s*\w+)?\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip()); new_col = m.group(2); field_col = m.group(3)
            add(f"{lhs} = {src}.copy()")
            add(f"{lhs}['{new_col}'] = {src}['{field_col}']")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # 2) Generic each expr
        m = re.search(r'Table\.AddColumn\(\s*([^,]+),\s*"([^"]+)",\s*each\s*(.*?)(?:,\s*type\s*\w+)?\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip()); new_col = m.group(2); expr = m.group(3).strip()
            expr_py = _replace_record_refs(expr, src)
            add(f"{lhs} = {src}.copy()")
            add(f"{lhs}['{new_col}'] = {expr_py}")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- TransformColumnTypes (comment) --------------------------------
        m = re.search(r'Table\.TransformColumnTypes\(([^,]+),\s*\{\{(.+?)\}\}\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            col_specs = [s.strip() for s in m.group(2).split("},")]
            add(f"{lhs} = {src}.copy()")
            for spec in col_specs:
                cols = re.findall(r'"([^"]+)"', spec)
                if cols:
                    c = cols[0]
                    add(f"# {lhs}['{c}'] = {lhs}['{c}'].astype(...)  # adjust dtype")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Group ----------------------------------------------------------
        m = re.search(r'Table\.Group\(([^,]+),\s*\{([^}]+)\},\s*\{(.+)\}\)', rhs, re.DOTALL)
        if m:
            src = _normalize_var(m.group(1).strip())
            keys = [k.strip().strip('"') for k in m.group(2).split(",")]
            aggs_block = m.group(3)
            agg_defs = re.findall(r'\{\s*"([^"]+)"\s*,\s*each\s*(?:List\.(\w+)\(\[([^\]]+)\]\)|Table\.(RowCount)\(_\))', aggs_block)
            lines = []
            func_map = {"Sum": "sum", "Average": "mean", "Min": "min", "Max": "max", "Count": "count"}
            add(f"{lhs} = {src}.copy()")
            add(f"{lhs}['__grp__'] = 1")
            for alias, list_func, list_col, rowcount in agg_defs:
                if rowcount == "RowCount":
                    lines.append(f"{alias}=('__grp__', 'size')")
                else:
                    pyfunc = func_map.get(list_func, "sum")
                    lines.append(f"{alias}=('{list_col}', '{pyfunc}')")
            add(f"{lhs} = {lhs}.groupby({keys}).agg(\n    " + ",\n    ".join(lines) + "\n).reset_index()")
            add(f"{lhs} = {lhs}.drop(columns=['__grp__'])")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Join -----------------------------------------------------------
        m = re.search(r'Table\.Join\(([^,]+),\s*"([^"]+)",\s*([^,]+),\s*"([^"]+)",\s*JoinKind\.(\w+)\)', rhs)
        if m:
            left = _normalize_var(m.group(1).strip()); left_key = m.group(2)
            right = _normalize_var(m.group(3).strip()); right_key = m.group(4)
            kind = m.group(5)
            how = {"Inner":"inner","LeftOuter":"left","RightOuter":"right","FullOuter":"outer"}.get(kind, "inner")
            add(f"{lhs} = pd.merge({left}, {right}, left_on='{left_key}', right_on='{right_key}', how='{how}')")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Distinct -------------------------------------------------------
        m = re.search(r'Table\.Distinct\(([^)]+)\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            add(f"{lhs} = {src}.drop_duplicates()")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- RemoveRowsWithErrors (naive -> dropna) -------------------------
        m = re.search(r'Table\.RemoveRowsWithErrors\(([^)]+)\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            add(f"{lhs} = {src}.dropna()")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Unpivot --------------------------------------------------------
        m = re.search(r'Table\.Unpivot\(\s*([^,]+),\s*\{([^}]+)\},\s*"([^"]+)",\s*"([^"]+)"\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            value_vars = [c.strip().strip('"') for c in m.group(2).split(",")]
            attr, val = m.group(3), m.group(4)
            add(f"{lhs} = {src}.melt(")
            add(f"    id_vars=[c for c in {src}.columns if c not in {value_vars}],")
            add(f"    value_vars={value_vars}, var_name='{attr}', value_name='{val}')")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- UnpivotOtherColumns -------------------------------------------
        m = re.search(r'Table\.UnpivotOtherColumns\(\s*([^,]+),\s*\{([^}]+)\},\s*"([^"]+)",\s*"([^"]+)"\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            keep_cols = [c.strip().strip('"') for c in m.group(2).split(",")]
            attr, val = m.group(3), m.group(4)
            add(f"{lhs} = {src}.melt(")
            add(f"    id_vars={keep_cols},")
            add(f"    value_vars=[c for c in {src}.columns if c not in {keep_cols}],")
            add(f"    var_name='{attr}', value_name='{val}')")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Pivot ----------------------------------------------------------
        m = re.search(r'Table\.Pivot\(\s*([^,]+),\s*.+?,\s*"([^"]+)",\s*"([^"]+)"', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            pivot_col, value_col = m.group(2), m.group(3)
            add(f"_idx_cols = [c for c in {src}.columns if c not in ['{pivot_col}', '{value_col}']]")
            add(f"{lhs} = {src}.pivot_table(index=_idx_cols, columns='{pivot_col}', values='{value_col}', aggfunc='first').reset_index()")
            add(f"{lhs}.columns = [c if not isinstance(c, tuple) else c[-1] for c in {lhs}.columns]")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Fallback -------------------------------------------------------
        add(f"# Unsupported: {lhs_raw} = {rhs}")

    return "\n".join(py)
