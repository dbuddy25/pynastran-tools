#!/usr/bin/env python3
"""Find free (unconnected) edges in a shell mesh.

Free edges indicate mesh discontinuities, missing elements, or model
boundaries. Useful for mesh quality checks.

Requires: pip install pyNastran

Usage:
    python find_free_edges.py model.bdf
    python find_free_edges.py model.bdf --csv free_edges.csv
"""
import argparse
import sys
from collections import Counter
from pyNastran.bdf.bdf import BDF


def find_free_edges(model: BDF) -> list:
    """Find free edges in the shell mesh.

    A free edge is an element edge shared by only one element.

    Returns:
        List of (nid1, nid2) tuples representing free edges.
    """
    edge_count = Counter()

    for eid, elem in model.elements.items():
        if elem.type in ('CQUAD4', 'CQUAD8'):
            nodes = elem.node_ids
            corners = nodes[:4]  # first 4 are corners
            edges = [
                tuple(sorted([corners[0], corners[1]])),
                tuple(sorted([corners[1], corners[2]])),
                tuple(sorted([corners[2], corners[3]])),
                tuple(sorted([corners[3], corners[0]])),
            ]
            for edge in edges:
                edge_count[edge] += 1

        elif elem.type in ('CTRIA3', 'CTRIA6'):
            nodes = elem.node_ids
            corners = nodes[:3]
            edges = [
                tuple(sorted([corners[0], corners[1]])),
                tuple(sorted([corners[1], corners[2]])),
                tuple(sorted([corners[2], corners[0]])),
            ]
            for edge in edges:
                edge_count[edge] += 1

    free = [edge for edge, count in edge_count.items() if count == 1]
    return sorted(free)


def main() -> None:
    parser = argparse.ArgumentParser(
        description='Find free (unconnected) shell edges in a BDF model.')
    parser.add_argument('bdf', help='Path to the BDF file')
    parser.add_argument('--csv', default=None,
                        help='Export free edges to CSV file')
    args = parser.parse_args()

    model = BDF()
    model.read_bdf(args.bdf)

    edges = find_free_edges(model)

    sep = '=' * 50
    print(sep)
    print(f"FREE EDGE REPORT: {args.bdf}")
    print(sep)

    n_shells = sum(1 for e in model.elements.values()
                   if e.type in ('CQUAD4', 'CQUAD8', 'CTRIA3', 'CTRIA6'))
    print(f"  Shell elements: {n_shells:,}")
    print(f"  Free edges:     {len(edges):,}")

    if not edges:
        print("\n  No free edges found (mesh is closed).")
    else:
        # Collect unique nodes on free edges
        free_nodes = set()
        for n1, n2 in edges:
            free_nodes.add(n1)
            free_nodes.add(n2)
        print(f"  Nodes on free edges: {len(free_nodes):,}")

        print(f"\n  {'Node 1':>8s}  {'Node 2':>8s}")
        print(f"  {'-'*8}  {'-'*8}")
        for n1, n2 in edges[:50]:  # limit output
            print(f"  {n1:8d}  {n2:8d}")
        if len(edges) > 50:
            print(f"  ... and {len(edges) - 50} more")

    if args.csv and edges:
        with open(args.csv, 'w') as f:
            f.write("node1,node2\n")
            for n1, n2 in edges:
                f.write(f"{n1},{n2}\n")
        print(f"\n  Exported to {args.csv}")

    print(f"\n{sep}")


if __name__ == '__main__':
    main()
