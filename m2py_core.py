# m2py_core.py
import re
import base64
import zlib
import json

# --- helpers ---------------------------------------------------------------

def _normalize_var(name: str) -> str:
    """
    Convert Power Query step/query names to valid Python identifiers.

    Examples:
      # "Changed Type" -> Changed_Type
      "Changed Type"  -> Changed_Type
    """
    name = name.strip()
    if name.startswith('#"') and name.endswith('"'):
        name = name[2:-1]
    return re.sub(r"\s+", "_", name)


def _replace_record_refs(expr: str, df_name: str) -> str:
    """
    Convert M column refs like [Col A] into df['Col A'] (for the given df_name).
    Keeps string literals intact with a simple scanner.
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
    Translate an M filter lambda into a precedence-safe pandas boolean expression.

    Example:
      [Country] = "FR" and [Age] > 30 or not [Active]
      ->
      ((df['Country'] == "FR") & (df['Age'] > 30) | ~df['Active'])
    """
    s = _replace_record_refs(cond, df_name)
    # single '=' to '==', but avoid '=='/'!=', etc. already present
    s = re.sub(r"(?<![=!<>])=(?!=)", "==", s)
    s = re.sub(r"\band\b", "&", s, flags=re.I)
    s = re.sub(r"\bor\b", "|", s, flags=re.I)
    s = re.sub(r"\bnot\b", "~", s, flags=re.I)
    # parenthesize each simple comparison; crude but helps precedence
    s = re.sub(r"([)\]\}A-Za-z0-9_'\"]\s*(?:==|!=|>=|<=|>|<)\s*[^&|]+)", r"(\1)", s)
    # wrap if mixing & and |
    if "&" in s and "|" in s:
        s = f"({s})"
    return s


def _extract_final_symbol(m_code: str) -> str | None:
    """
    Return the symbol referenced by the final 'in' clause.
    Handles both:
      in  #"Step"
    and
      in
          #"Step"
    """
    lines = [ln.rstrip() for ln in m_code.splitlines()]
    n = len(lines)
    for i, ln in enumerate(lines):
        s = ln.strip()
        if s.lower() == "in" or s.lower().startswith("in "):
            rhs = s[2:].strip() if s.lower().startswith("in ") else ""
            if not rhs:
                # take the first non-empty line after 'in'
                k = i + 1
                while k < n and not lines[k].strip():
                    k += 1
                if k < n:
                    rhs = lines[k].strip()
            rhs = rhs.rstrip(",")

            m = re.fullmatch(r'#"(.*)"', rhs)
            if m:
                return _normalize_var(m.group(1))

            m = re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', rhs)
            if m:
                return _normalize_var(rhs)
            return None
    return None


# --- main ------------------------------------------------------------------

def convert_m_to_python(m_code: str, query_name: str | None = None) -> str:
    """
    Best-effort M -> pandas translator (regex-based).

    Supported patterns (subset of M):
      - Excel.CurrentWorkbook(){[Name="TableName"]}[Content]
      - Csv.Document(File.Contents("file.csv")) + Table.PromoteHeaders
      - Table.RenameColumns
      - Table.SelectColumns / Table.RemoveColumns
      - Table.Sort
      - Table.SelectRows (and/or/not, =, !=, >, <, >=, <=)
      - Table.AddColumn (arithmetic/column refs 'each ...', Record.Field(_, "Col"))
      - Table.TransformColumnTypes (emitted as dtype comment hints)
      - Table.Group (Sum/Avg/Min/Max/Count, multi-agg)
      - Table.Join (Inner/Left/Right/Full)
      - Table.Distinct
      - Table.RemoveRowsWithErrors
      - Table.FromRows(Json.Document(Binary.Decompress(...))) -> static DataFrame literal
      - Table.Unpivot / Table.UnpivotOtherColumns / Table.Pivot
      - Cross-query refs: Step = #"Other Step"  and  Step = Other_Step

    If `query_name` is provided, the output appends:
        <QueryName> = <FinalStepFromInClause>

    Notes:
    - This is a pragmatic converter for prototyping/tests, not a full M parser.
    - Unsupported lines are emitted as comments to keep context.
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

        # --- cross-query refs ------------------------------------------------
        # bare: This = That
        m = re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', rhs)
        if m:
            ref = _normalize_var(rhs)
            add(f"{lhs} = {ref}")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # quoted: This = #"That"
        m = re.fullmatch(r'#"([^"]+)"', rhs)
        if m:
            ref = _normalize_var(m.group(1))
            add(f"{lhs} = {ref}")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # --- Enter Data ------------------------------------------------------
        m = re.search(
            r'Table\.FromRows\(\s*Json\.Document\(\s*Binary\.Decompress\(\s*Binary\.FromText\("([^"]+)"\s*,\s*BinaryEncoding\.Base64\)\s*,\s*Compression\.Deflate\)\s*\)\s*,\s*.*?type\s+table\s*\[([^\]]+)\]\s*\)',
            rhs, re.DOTALL
        )
        if m:
            b64 = m.group(1)
            cols_spec = m.group(2)
            col_names = [c.strip().split("=", 1)[0].strip() for c in cols_spec.split(",")]
            try:
                raw_bytes = base64.b64decode(b64)
                try:
                    data = zlib.decompress(raw_bytes)
                except zlib.error:
                    # some PQ exports use raw-deflate (no zlib header)
                    data = zlib.decompress(raw_bytes, -zlib.MAX_WBITS)
                rows = json.loads(data.decode("utf-8"))
                dict_rows = []
                for r in rows:
                    r_list = list(r) if isinstance(r, (list, tuple)) else [r]
                    if len(r_list) < len(col_names):
                        r_list += [None] * (len(col_names) - len(r_list))
                    if len(r_list) > len(col_names):
                        r_list = r_list[:len(col_names)]
                    dict_rows.append({k: v for k, v in zip(col_names, r_list)})
                add(f"{lhs} = pd.DataFrame({dict_rows})")  # static literal good for tests
            except Exception as e:
                add(f"# Enter Data decode failed at conversion time: {e!r}")
                add(f"{lhs} = pd.DataFrame(columns={[c for c in col_names]})  # TODO: fill rows")
            env[lhs_raw] = lhs
            last_df = lhs
            continue

        # --- Sources ---------------------------------------------------------
        m = re.search(r'Excel\.CurrentWorkbook\(\)\{\[Name="(.*?)"\]\}\[Content\]', rhs)
        if m:
            tbl = m.group(1)
            add(f"{lhs} = pd.read_excel('workbook.xlsx', sheet_name=None)['{tbl}']")
            env[lhs_raw] = lhs; last_df = lhs; continue

        m = re.search(r'Csv\.Document\(File\.Contents\("([^"]+)"\)\s*(?:,\s*\[([^\]]*)\])?\)', rhs)
        if m:
            csv = m.group(1)
            opts = m.group(2) or ""
            sep = None
            enc = None
            quote_none = False
            m_delim = re.search(r'Delimiter\s*=\s*"([^"]+)"', opts)
            if m_delim: sep = m_delim.group(1)
            m_enc = re.search(r'Encoding\s*=\s*(\d+)', opts)
            if m_enc:
                codepage = m_enc.group(1)
                enc = "utf-8" if codepage == "65001" else ("cp1252" if codepage == "1252" else None)
            if re.search(r'QuoteStyle\s*=\s*QuoteStyle\.None', opts):
                quote_none = True
            args = [f"'{csv}'", "header=None"]
            if sep is not None:
                args.append(f"sep='{sep}'")
            if enc is not None:
                args.append(f"encoding='{enc}'")
            if quote_none:
                args.append("quoting=3")  # csv.QUOTE_NONE
            add(f"{lhs} = pd.read_csv({', '.join(args)})")
            env[lhs_raw] = lhs; last_df = lhs; continue

        m = re.search(r'Table\.PromoteHeaders\(([^)]+)\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            add(f"{lhs} = {src}.copy()")
            add(f"{lhs}.columns = {lhs}.iloc[0]")
            add(f"{lhs} = {lhs}.iloc[1:].reset_index(drop=True)")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- RenameColumns ---------------------------------------------------
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

        # --- Select/Remove Columns ------------------------------------------
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

        # --- Sort ------------------------------------------------------------
        m = re.search(r'Table\.Sort\(([^,]+),\s*\{(.+?)\}\)', rhs, re.DOTALL)
        if m:
            src = _normalize_var(m.group(1).strip())
            items = re.findall(r'\{"([^"]+)"\s*,\s*Order\.(Ascending|Descending)\}', m.group(2))
            cols = [c for c, _ in items]
            asc = [o == "Ascending" for _, o in items]
            add(f"{lhs} = {src}.sort_values(by={cols}, ascending={asc})")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- SelectRows (filters) -------------------------------------------
        m = re.search(r'Table\.SelectRows\(([^,]+),\s*each\s*(.+)\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            cond_raw = m.group(2).strip()
            cond = _translate_condition(cond_raw, src)
            add(f"{lhs} = {src}[{cond}]")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- AddColumn -------------------------------------------------------
        m = re.search(r'Table\.AddColumn\(\s*([^,]+),\s*"([^"]+)",\s*each\s*Record\.Field\(_, "([^"]+)"\)\s*(?:,\s*type\s*\w+)?\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip()); new_col = m.group(2); field_col = m.group(3)
            add(f"{lhs} = {src}.copy()")
            add(f"{lhs}['{new_col}'] = {src}['{field_col}']")
            env[lhs_raw] = lhs; last_df = lhs; continue

        m = re.search(r'Table\.AddColumn\(\s*([^,]+),\s*"([^"]+)",\s*each\s*(.*?)(?:,\s*type\s*\w+)?\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip()); new_col = m.group(2); expr = m.group(3).strip()
            expr_py = _replace_record_refs(expr, src)
            add(f"{lhs} = {src}.copy()")
            add(f"{lhs}['{new_col}'] = {expr_py}")
            env[lhs_raw] = lhs; last_df = lhs; continue

                # --- TransformColumnTypes -----------------------------------------
        m = re.search(r'Table\.TransformColumnTypes\(([^,]+),\s*\{\{(.+?)\}\}\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            specs = m.group(2)
            add(f"{lhs} = {src}.copy()")
            pairs = re.findall(r'\{\s*\"([^\"]+)\"\s*,\s*type\s*([A-Za-z0-9_\.]+)\s*\}', specs)
            for col, typ in pairs:
                t = typ.split('.')[-1].lower()
                if t in ('text',):
                    add(f"{lhs}['{col}'] = {lhs}['{col}'].astype('string')")
                elif t in ('number','double','single','decimal'):
                    add(f"{lhs}['{col}'] = {lhs}['{col}'].astype('float')")
                elif t in ('int64','int32','int16','int8'):
                    add(f"{lhs}['{col}'] = pd.to_numeric({lhs}['{col}'], errors='coerce').astype('Int64')")
                elif t in ('date','datetime','datetimezone'):
                    add(f"{lhs}['{col}'] = pd.to_datetime({lhs}['{col}'], errors='coerce')")
                elif t in ('logical',):
                    add(f"{lhs}['{col}'] = {lhs}['{col}'].astype('boolean')")
                else:
                    add(f"# {lhs}['{col}'] = {lhs}['{col}'].astype('object')  # unhandled type: {t}")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Group -----------------------------------------------------------
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
            env[lhs_raw] = lhs
            last_df = lhs
            continue

                # --- Join (single or multi-key) ----------------------------------
        m = re.search(r'Table\.Join\(\s*([^,]+),\s*("[^"]+"|\{[^\}]+\})\s*,\s*([^,]+),\s*("[^"]+"|\{[^\}]+\})\s*,\s*JoinKind\.(\w+)\s*\)', rhs)
        if m:
            left = _normalize_var(m.group(1).strip())
            left_keys_raw = m.group(2).strip()
            right = _normalize_var(m.group(3).strip())
            right_keys_raw = m.group(4).strip()
            kind = m.group(5)
            def _parse_keys(k):
                if k.startswith('{'):
                    return [x for x in re.findall(r'"([^"]+)"', k)]
                return [k.strip('"')]
            lk = _parse_keys(left_keys_raw)
            rk = _parse_keys(right_keys_raw)
            how = {'Inner':'inner','LeftOuter':'left','RightOuter':'right','FullOuter':'outer'}.get(kind, 'inner')
            add(f"{lhs} = pd.merge({left}, {right}, how='{how}', left_on={lk!r}, right_on={rk!r})")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Distinct --------------------------------------------------------
        m = re.search(r'Table\.Distinct\(([^)]+)\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            add(f"{lhs} = {src}.drop_duplicates()")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- RemoveRowsWithErrors (naive) -----------------------------------
        m = re.search(r'Table\.RemoveRowsWithErrors\(([^)]+)\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            add(f"{lhs} = {src}.dropna()")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- Unpivot ---------------------------------------------------------
        m = re.search(r'Table\.Unpivot\(\s*([^,]+),\s*\{([^}]+)\},\s*"([^"]+)",\s*"([^"]+)"\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            value_vars = [c.strip().strip('"') for c in m.group(2).split(",")]
            attr, val = m.group(3), m.group(4)
            add(f"{lhs} = {src}.melt(")
            add(f"    id_vars=[c for c in {src}.columns if c not in {value_vars}],")
            add(f"    value_vars={value_vars}, var_name='{attr}', value_name='{val}')")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- UnpivotOtherColumns --------------------------------------------
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

        # --- Pivot -----------------------------------------------------------
        m = re.search(r'Table\.Pivot\(\s*([^,]+),\s*.+?,\s*"([^"]+)",\s*"([^"]+)"', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            pivot_col, value_col = m.group(2), m.group(3)
            add(f"_idx_cols = [c for c in {src}.columns if c not in ['{pivot_col}', '{value_col}']]")
            add(f"{lhs} = {src}.pivot_table(index=_idx_cols, columns='{pivot_col}', values='{value_col}', aggfunc='first').reset_index()")
            add(f"{lhs}.columns = [c if not isinstance(c, tuple) else c[-1] for c in {lhs}.columns]")
            env[lhs_raw] = lhs; last_df = lhs; continue

                # --- ExpandRecordColumn --------------------------------------
        m = re.search(r'Table\.ExpandRecordColumn\(\s*([^,]+),\s*"([^"]+)"\s*,\s*\{([^\}]*)\}\s*,\s*\{([^\}]*)\}\s*\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            col = m.group(2)
            fields = [s for s in re.findall(r'"([^"]+)"', m.group(3))]
            newnames = [s for s in re.findall(r'"([^"]+)"', m.group(4))]
            if not newnames or len(newnames) != len(fields):
                newnames = fields[:]
            add(f"{lhs} = {src}.drop(columns=['{col}'], errors='ignore').copy()")
            add(f"_exp = {src}['{col}'].apply(lambda x: pd.Series(x) if isinstance(x, dict) else pd.Series(dtype='object'))")
            if fields:
                add(f"_exp = _exp[{fields!r}]")
            if newnames and fields:
                add(f"_exp = _exp.rename(columns={{**dict(zip({fields!r}, {newnames!r}))}})")
            add(f"{lhs} = {lhs}.join(_exp)")
            env[lhs_raw] = lhs; last_df = lhs; continue

        # --- ExpandTableColumn ---------------------------------------
        m = re.search(r'Table\.ExpandTableColumn\(\s*([^,]+),\s*"([^"]+)"\s*,\s*\{([^\}]*)\}\s*,\s*\{([^\}]*)\}\s*\)', rhs)
        if m:
            src = _normalize_var(m.group(1).strip())
            col = m.group(2)
            fields = [s for s in re.findall(r'"([^"]+)"', m.group(3))]
            newnames = [s for s in re.findall(r'"([^"]+)"', m.group(4))]
            if not newnames or len(newnames) != len(fields):
                newnames = fields[:]
            add(f"{lhs} = {src}.copy()")
            add(f"_tbl = {lhs}.pop('{col}') if '{col}' in {lhs}.columns else pd.Series(index={lhs}.index, dtype='object')")
            add(f"_tbl = _tbl.apply(lambda t: t if isinstance(t, (list, tuple)) else ([] if t is None else [t]))")
            add(f"_tbl = _tbl.explode()")
            add(f"_df = pd.DataFrame(_tbl.tolist()) if not _tbl.empty else pd.DataFrame(columns={fields!r})")
            if fields:
                add(f"_df = _df.reindex(columns={fields!r})")
            if newnames and fields:
                add(f"_df = _df.rename(columns={{**dict(zip({fields!r}, {newnames!r}))}})")
            add(f"{lhs} = {lhs}.join(_df.reset_index(drop=True))")
            env[lhs_raw] = lhs; last_df = lhs; continue
# --- Fallback --------------------------------------------------------
        add(f"# Unsupported: {lhs_raw} = {rhs}")

    # Append final alias if we know the query name and can find the final symbol
    if query_name:
        final_sym = _extract_final_symbol(m_code)
        if final_sym:
            py.append(f"{_normalize_var(query_name)} = {final_sym}")
        elif last_df:
            # fallback: alias the last defined step
            py.append(f"{_normalize_var(query_name)} = {last_df}")

    return "\n".join(py)
