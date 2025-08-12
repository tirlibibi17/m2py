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

from typing import Dict, List, Set, Tuple
import re
from collections import defaultdict, deque


def _normalize_query_name(name: str) -> str:
    """Keep the original name, but normalize how we compare."""
    return name.strip()


def find_query_refs(m_text: str, known_names: Set[str]) -> Set[str]:
    """
    Return the set of query names referenced by this M text.

    We consider:
      - #"Some Query"
      - Bare identifiers that are an exact name of another query (after basic cleaning)
    """
    refs: Set[str] = set()
    text = m_text or ""

    # 1) Explicit #"Name" references
    for m in re.finditer(r'#"([^"]+)"', text):
        refs.add(_normalize_query_name(m.group(1)))

    # 2) Bare identifiers: tokens that match a known query name exactly.
    #    We only look for tokens that look like identifiers and then match known names.
    #    (Keeps false positives low.)
    tokens = set(re.findall(r'\b[A-Za-z_][A-Za-z0-9_\.]*\b', text))
    for t in tokens:
        if t in known_names:
            refs.add(t)

    # A query does not depend on itself
    return refs - {t for t in known_names if t in refs and len(refs) == 1 and next(iter(refs)) == t}


def topo_order_queries(queries: Dict[str, str]) -> List[str]:
    """
    Return a global topological order for all queries in the workbook.
    Independent components are ordered deterministically (by name).
    """
    names = set(queries.keys())
    edges: Dict[str, Set[str]] = {n: set() for n in names}   # n -> deps
    rev: Dict[str, Set[str]] = {n: set() for n in names}     # dep -> users

    for name, text in queries.items():
        deps = find_query_refs(text, names)
        edges[name] = set(d for d in deps if d in names and d != name)
        for d in edges[name]:
            rev[d].add(name)

    # Kahn's algorithm
    indeg = {n: 0 for n in names}
    for n in names:
        for d in edges[n]:
            indeg[n] += 1

    q = deque(sorted([n for n in names if indeg[n] == 0]))
    order: List[str] = []
    while q:
        n = q.popleft()
        order.append(n)
        for user in rev[n]:
            indeg[user] -= 1
            if indeg[user] == 0:
                q.append(user)
        # keep deterministic
        q = deque(sorted(q))

    # If there was a cycle, fall back to a stable name order with a warning order
    if len(order) != len(names):
        # put any remaining nodes deterministically at the end
        remaining = sorted(n for n in names if n not in order)
        order.extend(remaining)

    return order


def dependency_chain_for(target: str, queries: Dict[str, str]) -> List[str]:
    """
    Return the dependency chain for 'target': all transitive deps (in topo order) + target at the end.
    """
    if target not in queries:
        return [target]  # let caller handle "missing" gracefully

    names = set(queries.keys())

    # Build graph (name -> deps), and reverse graph
    edges: Dict[str, Set[str]] = {n: set() for n in names}
    rev: Dict[str, Set[str]] = {n: set() for n in names}

    for name, text in queries.items():
        deps = find_query_refs(text, names)
        edges[name] = set(d for d in deps if d in names and d != name)
        for d in edges[name]:
            rev[d].add(name)

    # Global order to keep output stable
    order = topo_order_queries(queries)

    # DFS from target back through deps to collect the reachable subgraph
    seen: Set[str] = set()
    stack = [target]
    while stack:
        cur = stack.pop()
        if cur in seen:
            continue
        seen.add(cur)
        stack.extend(edges[cur])

    # Keep only reachable nodes in the global topo order, ensure target is last
    chain = [n for n in order if n in seen and n != target]
    chain.append(target)
    return chain
