
import argparse
import pathlib
import re

from m2py_core import convert_m_to_python

# Import dependency resolver (deduplicated)
from query_resolver import find_query_refs, topo_order_queries, dependency_chain_for

# Excel COM extractor (Windows-only) — lazy import with fallback
try:
    from excel_com_extractor import extract_queries_from_excel_via_com  # noqa: F401
except Exception:  # pragma: no cover
    def extract_queries_from_excel_via_com(*args, **kwargs):  # type: ignore
        raise SystemExit("Excel COM extraction requires Windows + Excel + pywin32.")

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
