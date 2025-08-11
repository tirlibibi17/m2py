# query_resolver.py
"""
Resolve dependencies between Power Query (M) queries so conversion can emit
dependencies first (topological order).

Features:
- Detects references of the form  #"Some Query"
- Detects bare identifier refs like  Some_Query  (when it matches another query name)
- Topological sort of the whole query set
- Dependency chain for a chosen target (deps first, then target)
"""

from __future__ import annotations

import re
from collections import defaultdict, deque
from typing import Dict, List, Set

# Matches quoted references: #"Some Query"
REF_QUOTED = re.compile(r'#"([^"]+)"')

# Matches bare identifiers: Foo, Bar123, _tmp
IDENT = re.compile(r'\b([A-Za-z_][A-Za-z0-9_]*)\b')

# Common M keywords to exclude when scanning bare identifiers
M_KEYWORDS = {
    "let", "in", "each", "and", "or", "not", "true", "false", "null",
    # operators sometimes tokenized as idents by the regex:
    "as", "if", "then", "else", "error", "try", "otherwise"
}


def _strip_line_comments(m_code: str) -> str:
    """
    Remove single-line '//' comments (Power Query style) for cleaner ref scanning.
    Keeps string literals intact (simple heuristic: remove from // to end of line).
    """
    cleaned_lines = []
    for line in (m_code or "").splitlines():
        # crude but practical: split on // if not inside quotes
        s = line
        out, i, in_str, quote = [], 0, False, ""
        while i < len(s):
            ch = s[i]
            if not in_str and ch in ('"', "'"):
                in_str, quote = True, ch
                out.append(ch); i += 1
                continue
            if in_str:
                out.append(ch)
                if ch == quote:
                    in_str, quote = False, ""
                i += 1
                continue
            # not in string: comment start?
            if ch == "/" and i + 1 < len(s) and s[i + 1] == "/":
                break  # drop rest of line
            out.append(ch); i += 1
        cleaned_lines.append("".join(out))
    return "\n".join(cleaned_lines)


def find_query_refs(m_code: str, known_names: Set[str]) -> Set[str]:
    """
    Return the set of query names referenced in an M script.

    We detect:
      - Quoted refs:   #"Some Query"
      - Bare idents:   Some_Query   (only kept if it matches a known query name)
    """
    # Remove line comments to avoid false positives inside comments
    src = _strip_line_comments(m_code or "")

    refs: Set[str] = set(REF_QUOTED.findall(src))

    # Bare identifiers that might be query names (no spaces).
    # We only keep those that exist in known_names and are not obvious M keywords.
    candidates = set(IDENT.findall(src))
    refs |= {n for n in candidates if n in known_names and n.lower() not in M_KEYWORDS}
    return refs


def topo_order_queries(queries: Dict[str, str]) -> List[str]:
    """
    Topologically sort queries so dependencies appear before dependents.

    If a cycle exists, remaining nodes are appended at the end in stable order.
    """
    known = set(queries.keys())

    # Build dependency graph: name -> set(dependencies)
    graph: Dict[str, Set[str]] = {
        name: {r for r in find_query_refs(m, known) if r in known and r != name}
        for name, m in queries.items()
    }

    indeg = {n: 0 for n in queries}
    for n, deps in graph.items():
        for _ in deps:
            indeg[n] += 1

    q = deque([n for n, d in indeg.items() if d == 0])
    order: List[str] = []

    while q:
        u = q.popleft()
        order.append(u)
        for v, deps in graph.items():
            if u in deps:
                indeg[v] -= 1
                if indeg[v] == 0 and v not in order and v not in q:
                    q.append(v)

    # Anything left (cycles, unresolved) gets appended in original dict order
    remaining = [n for n in queries if n not in order]
    return order + remaining


def dependency_chain_for(target: str, queries: Dict[str, str]) -> List[str]:
    """
    Return a topologically ordered list containing the target's transitive
    dependencies **and** the target (deps first, then target).

    If the target doesn't exist, returns [].
    """
    if target not in queries:
        return []

    order = topo_order_queries(queries)
    known = set(queries.keys())

    # Build reverse adjacency: node -> its direct dependencies
    rev: Dict[str, Set[str]] = defaultdict(set)
    for n, m in queries.items():
        for r in find_query_refs(m, known):
            if r in known and r != n:
                rev[n].add(r)

    # DFS from target to collect all (transitive) dependencies + target
    seen: Set[str] = set()
    stack = [target]
    while stack:
        cur = stack.pop()
        if cur in seen:
            continue
        seen.add(cur)
        stack.extend(rev[cur])

    # Keep only the reachable set, respecting the global topo order
    chain = [n for n in order if n in seen]
    return chain
