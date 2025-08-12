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

from typing import Dict, List, Set
import re
from collections import deque


def _normalize_query_name(name: str) -> str:
    """Keep the original name, but normalize how we compare."""
    return (name or "").strip()


def find_query_refs(m_text: str, known_names: Set[str]) -> Set[str]:
    """
    Return the set of query names referenced by this M text.

    We consider:
      - #"Some Query"
      - Bare identifiers that are an exact name of another query (after basic cleaning)

    NOTE: Do NOT try to remove self-references here; we filter them later when
    constructing the dependency edges (we know the current query name there).
    """
    refs: Set[str] = set()
    text = m_text or ""

    # 1) Explicit #"Name" references (quoted names with spaces etc.)
    for m in re.finditer(r'#"\s*([^"]+?)\s*"', text):
        nm = _normalize_query_name(m.group(1))
        if nm in known_names:
            refs.add(nm)

    # 2) Bare identifiers: tokens that match a known query name exactly.
    #    This is conservative to avoid false positives.
    tokens = set(re.findall(r'\b[A-Za-z_][A-Za-z0-9_\.]*\b', text))
    for t in tokens:
        if t in known_names:
            refs.add(t)

    return refs


def topo_order_queries(queries: Dict[str, str]) -> List[str]:
    """
    Return a global, deterministic topological order for all queries.
    Independent components are ordered by name.
    """
    names = {_normalize_query_name(n) for n in queries.keys()}
    edges = {n: set() for n in names}   # n -> deps
    rev   = {n: set() for n in names}   # dep -> users

    for name, text in queries.items():
        n = _normalize_query_name(name)
        deps = find_query_refs(text, names)
        # filter out self-edges here
        edges[n] = {d for d in deps if d in names and d != n}
        for d in edges[n]:
            rev[d].add(n)

    # Kahn's algorithm with deterministic queue
    indeg = {n: 0 for n in names}
    for n in names:
        for d in edges[n]:
            indeg[n] += 1

    q = deque(sorted([n for n in names if indeg[n] == 0]))
    order: List[str] = []
    while q:
        n = q.popleft()
        order.append(n)
        for user in sorted(rev[n]):  # sort for stability
            indeg[user] -= 1
            if indeg[user] == 0:
                q.append(user)

    # If a cycle exists, append remaining nodes deterministically
    if len(order) != len(names):
        for n in sorted(names):
            if n not in order:
                order.append(n)

    return order


def dependency_chain_for(target: str, queries: Dict[str, str]) -> List[str]:
    """
    Return the dependency chain for 'target': all transitive deps (in topo order) + target at the end.
    """
    target_n = _normalize_query_name(target)
    names = {_normalize_query_name(n) for n in queries.keys()}
    if target_n not in names:
        return [target]  # let caller handle "missing" gracefully

    # Build graph (name -> deps), and reverse graph
    edges = {n: set() for n in names}
    rev   = {n: set() for n in names}

    # Map normalized name -> original text
    text_by_norm = { _normalize_query_name(n): (queries.get(n) or "") for n in queries }

    for n in names:
        deps = find_query_refs(text_by_norm[n], names)
        edges[n] = {d for d in deps if d in names and d != n}
        for d in edges[n]:
            rev[d].add(n)

    # Global order to keep output stable
    order = topo_order_queries(queries)

    # DFS from target back through deps to collect the reachable subgraph
    seen: Set[str] = set()
    stack = [target_n]
    while stack:
        cur = stack.pop()
        if cur in seen:
            continue
        seen.add(cur)
        stack.extend(edges[cur])

    # Keep only reachable nodes in the global topo order, ensure target is last
    chain = [n for n in order if n in seen and n != target_n]
    chain.append(target_n)
    return chain
