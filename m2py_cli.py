
import argparse
import pathlib
import re

from m2py_core import convert_m_to_python

# --------- inline dependency resolver ---------

def _strip_line_comments(m_code: str) -> str:
    cleaned_lines = []
    for line in (m_code or "").splitlines():
        s = line
        out, i, in_str, quote = [], 0, False, ""
        while i < len(s):
            ch = s[i]
            if not in_str and ch in ('"', "'"):
                in_str, quote = True, ch
                out.append(ch); i += 1; continue
            if in_str:
                out.append(ch)
                if ch == quote:
                    in_str, quote = False, ""
                i += 1; continue
            if ch == "/" and i + 1 < len(s) and s[i + 1] == "/":
                break
            out.append(ch); i += 1
        cleaned_lines.append("".join(out))
    return "\n".join(cleaned_lines)

REF_QUOTED = re.compile(r'#"(.*?)"')
IDENT = re.compile(r'\b([A-Za-z_][A-Za-z0-9_]*)\b')
M_KEYWORDS = {"let","in","each","and","or","not","true","false","null","as","if","then","else","error","try","otherwise"}

def find_query_refs(m_code: str, known_names: set[str]) -> set[str]:
    src = _strip_line_comments(m_code or "")
    refs = set(re.findall(r'#"(.*?)"', src))
    cand = set(IDENT.findall(src))
    refs |= {n for n in cand if n in known_names and n.lower() not in M_KEYWORDS}
    return refs

def topo_order_queries(queries: dict[str,str]) -> list[str]:
    known = set(queries.keys())
    graph = {name: {r for r in find_query_refs(m, known) if r in known and r != name}
             for name, m in queries.items()}
    indeg = {n: 0 for n in queries}
    for n, deps in graph.items():
        for _ in deps: indeg[n] += 1
    from collections import deque
    q = deque([n for n,d in indeg.items() if d==0])
    order = []
    while q:
        u = q.popleft()
        order.append(u)
        for v,deps in graph.items():
            if u in deps:
                indeg[v] -= 1
                if indeg[v]==0 and v not in order and v not in q:
                    q.append(v)
    remaining = [n for n in queries if n not in order]
    return order + remaining

def dependency_chain_for(target: str, queries: dict[str,str]) -> list[str]:
    order = topo_order_queries(queries)
    known = set(queries.keys())
    from collections import defaultdict
    rev = defaultdict(set)
    for n,m in queries.items():
        for r in find_query_refs(m, known):
            if r in known and r != n: rev[n].add(r)
    seen = set(); stack=[target]
    while stack:
        cur = stack.pop()
        if cur in seen: continue
        seen.add(cur); stack.extend(rev[cur])
    return [n for n in order if n in seen]

# --------- inline Excel COM extractor ---------

def extract_queries_from_excel_via_com(path_xlsx: str) -> dict[str,str]:
    try:
        import gc, pythoncom
        import win32com.client as win32
        from contextlib import contextmanager
    except Exception:
        raise SystemExit("Requires Windows + Excel + 'pywin32'. Install: pip install pywin32")

    @contextmanager
    def _com_apartment():
        pythoncom.CoInitialize()
        try:
            yield
        finally:
            pythoncom.CoUninitialize()

    path_xlsx = str(pathlib.Path(path_xlsx).resolve())

    with _com_apartment():
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = None
        try:
            wb = excel.Workbooks.Open(path_xlsx, ReadOnly=True, UpdateLinks=0)
            result = {}
            for q in wb.Queries:
                result[str(q.Name)] = str(q.Formula)
            return result
        finally:
            try:
                if wb is not None: wb.Close(False)
            except Exception: pass
            try:
                excel.Quit()
            except Exception: pass
            del wb; del excel; gc.collect()

# --------- CLI ---------

def main():
    parser = argparse.ArgumentParser(description="Convert Power Query M to pandas Python.")
    g = parser.add_mutually_exclusive_group(required=True)
    g.add_argument("--file", help="Input .m file")
    g.add_argument("--excel", help="Input Excel workbook (.xlsx/.xlsm) — Windows only (COM)")

    parser.add_argument("--query", help="Query name to convert (with --excel)")
    parser.add_argument("--output", required=True,
                        help="Output .py file (with --file/--query) or output folder (with --excel for all)")
    parser.add_argument("--list", action="store_true", help="List queries in the Excel file and exit")

    args = parser.parse_args()

    if args.excel:
        queries = extract_queries_from_excel_via_com(args.excel)

        if args.list:
            for name in sorted(queries): print(name)
            return

        if args.query:
            if args.query not in queries:
                raise SystemExit(f"Query not found: {args.query}")
            chain = dependency_chain_for(args.query, queries)  # deps first, then target
            blocks: list[str] = []
            for name in chain:
                m_text = queries[name]
                py_text = convert_m_to_python(m_text, query_name=name)  # alias appended
                blocks.append(f"# === {name} ===\n{py_text}\n")
            pathlib.Path(args.output).write_text("\n".join(blocks), encoding="utf-8")
        else:
            out_dir = pathlib.Path(args.output)
            out_dir.mkdir(parents=True, exist_ok=True)
            order = topo_order_queries(queries)
            for name in order:
                chain = dependency_chain_for(name, queries)
                blocks: list[str] = []
                for n in chain:
                    py_text = convert_m_to_python(queries[n], query_name=n)
                    blocks.append(f"# === {n} ===\n{py_text}\n")
                safe = "".join(c if c.isalnum() or c in "._- " else "_" for c in name).strip()
                (out_dir / f"{safe}.py").write_text("\n".join(blocks), encoding="utf-8")
        return

    # Single .m file mode
    with open(args.file, "r", encoding="utf-8") as f:
        m_code = f.read()
    # Use filename stem as query name for alias
    query_name = pathlib.Path(args.file).stem
    py_code = convert_m_to_python(m_code, query_name=query_name)
    pathlib.Path(args.output).write_text(py_code, encoding="utf-8")

if __name__ == "__main__":
    main()
