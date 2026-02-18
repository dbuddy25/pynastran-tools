#!/usr/bin/env python3
"""Renumber nodes and elements in a BDF model with ID offsets.

Useful when merging models or reorganizing ID ranges.

Requires: pip install pyNastran

Usage:
    python renumber_model.py input.bdf output.bdf
    python renumber_model.py input.bdf output.bdf --nid-offset 10000 --eid-offset 20000
    python renumber_model.py input.bdf output.bdf --nid-start 1 --eid-start 1
"""
import argparse
import sys
from pyNastran.bdf.mesh_utils.bdf_renumber import bdf_renumber


def main() -> None:
    parser = argparse.ArgumentParser(
        description='Renumber nodes/elements in a Nastran BDF file.')
    parser.add_argument('input_bdf', help='Input BDF file')
    parser.add_argument('output_bdf', help='Output BDF file')
    parser.add_argument('--nid-start', type=int, default=None,
                        help='Starting node ID')
    parser.add_argument('--eid-start', type=int, default=None,
                        help='Starting element ID')
    parser.add_argument('--pid-start', type=int, default=None,
                        help='Starting property ID')
    parser.add_argument('--mid-start', type=int, default=None,
                        help='Starting material ID')
    parser.add_argument('--nid-offset', type=int, default=None,
                        help='Add offset to all node IDs')
    parser.add_argument('--eid-offset', type=int, default=None,
                        help='Add offset to all element IDs')
    parser.add_argument('--size', type=int, default=8, choices=[8, 16],
                        help='Card field width (8 or 16, default: 8)')
    args = parser.parse_args()

    # Build starting_id_dict if any start values given
    starting_id_dict = None
    if any([args.nid_start, args.eid_start, args.pid_start, args.mid_start]):
        starting_id_dict = {}
        if args.nid_start is not None:
            starting_id_dict['nid'] = args.nid_start
        if args.eid_start is not None:
            starting_id_dict['eid'] = args.eid_start
        if args.pid_start is not None:
            starting_id_dict['pid'] = args.pid_start
        if args.mid_start is not None:
            starting_id_dict['mid'] = args.mid_start

    # If offset mode, read + manually offset + write
    if args.nid_offset is not None or args.eid_offset is not None:
        from pyNastran.bdf.bdf import BDF

        model = BDF()
        model.read_bdf(args.input_bdf)

        nid_off = args.nid_offset or 0
        eid_off = args.eid_offset or 0

        print(f"Applying offsets: NID+{nid_off}, EID+{eid_off}")
        print(f"  Nodes before: {len(model.nodes)}")
        print(f"  Elements before: {len(model.elements)}")

        # Renumber nodes
        if nid_off != 0:
            old_nodes = dict(model.nodes)
            model.nodes.clear()
            for nid, node in old_nodes.items():
                node.nid += nid_off
                model.nodes[node.nid] = node

            # Update element node references
            for eid, elem in model.elements.items():
                elem.nodes = [n + nid_off for n in elem.node_ids]

            # Update rigid elements
            for eid, elem in model.rigid_elements.items():
                if hasattr(elem, 'Gmi'):
                    elem.Gmi = [n + nid_off for n in elem.Gmi]
                if hasattr(elem, 'gn'):
                    elem.gn += nid_off

        # Renumber elements
        if eid_off != 0:
            old_elems = dict(model.elements)
            model.elements.clear()
            for eid, elem in old_elems.items():
                elem.eid += eid_off
                model.elements[elem.eid] = elem

        model.write_bdf(args.output_bdf, size=args.size)
        print(f"  Written to {args.output_bdf}")

    else:
        # Use pyNastran's built-in renumber
        bdf_renumber(
            args.input_bdf, args.output_bdf,
            size=args.size,
            is_double=(args.size == 16),
            starting_id_dict=starting_id_dict,
        )
        print(f"Renumbered model written to {args.output_bdf}")


if __name__ == '__main__':
    main()
