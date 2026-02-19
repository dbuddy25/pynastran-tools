#!/usr/bin/env python3
"""Nastran Include File Renumbering Tool.

Standalone GUI that reads a Nastran BDF/DAT file with INCLUDE files, catalogs
entity types and ID ranges per file, lets the user set new start/end ranges,
validates, and renumbers all cards — including contact, loads, BCs, case
control — writing per-file output to a new directory.

Usage:
    python renumber_includes.py
"""
import json
import os
import re
import tkinter as tk
from collections import defaultdict
from tkinter import filedialog, messagebox

import customtkinter as ctk
import tksheet

from pyNastran.bdf.bdf import BDF

try:
    from bdf_utils import (
        IncludeFileParser, CARD_ENTITY_MAP, ENTITY_TYPES, ENTITY_LABELS,
        make_model,
    )
except ImportError:
    from preprocessing.bdf_utils import (
        IncludeFileParser, CARD_ENTITY_MAP, ENTITY_TYPES, ENTITY_LABELS,
        make_model,
    )

# Cards that pyNastran may not support — disable to avoid parse errors.
# Shorter list than mass_scale because the renumber tool needs most contact cards.
_CARDS_TO_SKIP = [
    'BCPROPS', 'BCTPARM', 'BCPARA', 'BOUTPUT',
]


# ═══════════════════════════════════════════════════════════════════════════════
# Section 2: MappingBuilder
# ═══════════════════════════════════════════════════════════════════════════════

class MappingBuilder:
    """Build old→new ID maps from file ownership and user-specified ranges."""

    def __init__(self, file_ids, ranges):
        """
        Args:
            file_ids: dict[filepath, dict[entity_type, set[int]]] from parser
            ranges: dict[filepath, dict[entity_type, (start_id, end_id)]]
        """
        self.file_ids = file_ids
        self.ranges = ranges
        self.maps = {}  # {entity_type: {old_id: new_id}}

    def build(self):
        """Build all ID maps. Returns dict[entity_type, dict[int, int]]."""
        self.maps = {etype: {} for etype in ENTITY_TYPES}

        for etype in ENTITY_TYPES:
            for filepath, ids_by_type in self.file_ids.items():
                ids = ids_by_type.get(etype, set())
                if not ids:
                    continue

                range_info = self.ranges.get(filepath, {}).get(etype)
                if range_info is None:
                    continue

                start_id, end_id = range_info
                sorted_ids = sorted(ids)

                for i, old_id in enumerate(sorted_ids):
                    new_id = start_id + i
                    self.maps[etype][old_id] = new_id

        return self.maps


# ═══════════════════════════════════════════════════════════════════════════════
# Section 3: Validator
# ═══════════════════════════════════════════════════════════════════════════════

class Validator:
    """Pre-apply and post-apply validation checks."""

    @staticmethod
    def validate_ranges(file_ids, ranges, include_set_ids=True):
        """Pre-apply validation. Returns list of error strings (empty = OK)."""
        errors = []

        for etype in ENTITY_TYPES:
            # Skip set IDs if not included
            if not include_set_ids and etype in ('spc_id', 'mpc_id', 'load_id'):
                continue

            # Collect all ranges for this entity type to check overlaps
            etype_ranges = []

            for filepath in file_ids:
                ids = file_ids[filepath].get(etype, set())
                if not ids:
                    continue

                fname = os.path.basename(filepath)
                range_info = ranges.get(filepath, {}).get(etype)

                if range_info is None:
                    errors.append(
                        f"{fname}/{ENTITY_LABELS.get(etype, etype)}: "
                        f"no range specified for {len(ids)} entities")
                    continue

                start_id, end_id = range_info
                count = len(ids)

                # Positive IDs
                if start_id < 1:
                    errors.append(
                        f"{fname}/{ENTITY_LABELS.get(etype, etype)}: "
                        f"start_id must be >= 1 (got {start_id})")
                if end_id < start_id:
                    errors.append(
                        f"{fname}/{ENTITY_LABELS.get(etype, etype)}: "
                        f"end_id ({end_id}) < start_id ({start_id})")
                    continue

                # Capacity check
                capacity = end_id - start_id + 1
                if capacity < count:
                    errors.append(
                        f"{fname}/{ENTITY_LABELS.get(etype, etype)}: "
                        f"range [{start_id}-{end_id}] has capacity {capacity} "
                        f"but {count} entities need renumbering")

                etype_ranges.append((start_id, end_id, fname))

            # Overlap check
            etype_ranges.sort()
            for i in range(len(etype_ranges) - 1):
                s1, e1, f1 = etype_ranges[i]
                s2, e2, f2 = etype_ranges[i + 1]
                if s2 <= e1:
                    errors.append(
                        f"{ENTITY_LABELS.get(etype, etype)}: "
                        f"ranges overlap between {f1} [{s1}-{e1}] "
                        f"and {f2} [{s2}-{e2}]")

        # CID 0 check
        for filepath in file_ids:
            cids = file_ids[filepath].get('cid', set())
            if 0 in cids:
                range_info = ranges.get(filepath, {}).get('cid')
                if range_info and range_info[0] != 0:
                    errors.append(
                        f"{os.path.basename(filepath)}/Coord ID: "
                        f"CID 0 (basic coord system) should not be remapped")

        return errors

    @staticmethod
    def post_validate(original_path, output_path):
        """Post-apply validation. Returns (warnings, errors) lists."""
        warnings = []
        errors = []

        try:
            model = BDF(mode='nx')
            model.read_bdf(output_path)
        except Exception as exc:
            errors.append(f"Could not re-read output file: {exc}")
            return warnings, errors

        try:
            orig = BDF(mode='nx')
            orig.read_bdf(original_path)
        except Exception as exc:
            warnings.append(f"Could not re-read original for comparison: {exc}")
            return warnings, errors

        # Count comparison
        checks = [
            ('nodes', 'nodes'), ('elements', 'elements'),
            ('properties', 'properties'), ('materials', 'materials'),
            ('coords', 'coords'),
        ]
        for label, attr in checks:
            orig_count = len(getattr(orig, attr, {}))
            new_count = len(getattr(model, attr, {}))
            if orig_count != new_count:
                errors.append(
                    f"{label} count mismatch: original={orig_count}, "
                    f"output={new_count}")

        # Connectivity check: every element node must exist
        for eid, elem in model.elements.items():
            try:
                nids = elem.node_ids
                for nid in nids:
                    if nid is not None and nid not in model.nodes and nid != 0:
                        errors.append(
                            f"Element {eid} ({elem.type}) references "
                            f"missing node {nid}")
                        break
            except Exception:
                pass

        return warnings, errors


# ═══════════════════════════════════════════════════════════════════════════════
# Section 4: CardRenumberer
# ═══════════════════════════════════════════════════════════════════════════════

class CardRenumberer:
    """Apply ID mappings to every card in the pyNastran BDF model."""

    def __init__(self, model, maps, include_set_ids=True):
        """
        Args:
            model: pyNastran BDF model (not cross-referenced)
            maps: dict[entity_type, dict[old_id, new_id]]
            include_set_ids: whether to renumber spc/mpc/load set IDs
        """
        self.model = model
        self.nid_map = maps.get('nid', {})
        self.eid_map = maps.get('eid', {})
        self.pid_map = maps.get('pid', {})
        self.mid_map = maps.get('mid', {})
        self.cid_map = maps.get('cid', {})
        self.spc_map = maps.get('spc_id', {}) if include_set_ids else {}
        self.mpc_map = maps.get('mpc_id', {}) if include_set_ids else {}
        self.load_map = maps.get('load_id', {}) if include_set_ids else {}
        self.contact_map = maps.get('contact_id', {})
        self.set_map = maps.get('set_id', {})
        self.method_map = maps.get('method_id', {})
        self.table_map = maps.get('table_id', {})

    def _m(self, id_map, old_id):
        """Map an ID, returning the original if not in the map."""
        if old_id is None or old_id == 0:
            return old_id
        return id_map.get(old_id, old_id)

    def _m_list(self, id_map, old_ids):
        """Map a list of IDs."""
        return [self._m(id_map, x) for x in old_ids]

    def apply(self):
        """Renumber all cards and rebuild model dicts."""
        self._renumber_nodes()
        self._renumber_elements()
        self._renumber_rigid_elements()
        self._renumber_masses()
        self._renumber_properties()
        self._renumber_materials()
        self._renumber_coords()
        self._renumber_spcs()
        self._renumber_mpcs()
        self._renumber_loads()
        self._renumber_contact()
        self._renumber_sets()
        self._renumber_methods()
        self._renumber_tables()
        self._renumber_suport()

    def _renumber_nodes(self):
        """Renumber GRID and SPOINT cards."""
        model = self.model
        new_nodes = {}
        for nid, node in model.nodes.items():
            new_nid = self._m(self.nid_map, nid)
            node.nid = new_nid
            node.cp = self._m(self.cid_map, node.cp)
            node.cd = self._m(self.cid_map, node.cd)
            new_nodes[new_nid] = node
        model.nodes = new_nodes

        if hasattr(model, 'spoints') and model.spoints:
            new_spoints = {}
            for sp_id, sp in model.spoints.items():
                new_id = self._m(self.nid_map, sp_id)
                if hasattr(sp, 'spoint'):
                    sp.spoint = new_id
                new_spoints[new_id] = sp
            model.spoints = new_spoints

    def _renumber_elements(self):
        """Renumber all element cards."""
        model = self.model
        new_elements = {}
        for eid, elem in model.elements.items():
            new_eid = self._m(self.eid_map, eid)
            elem.eid = new_eid
            etype = elem.type

            # Map nodes
            if hasattr(elem, 'nodes'):
                elem.nodes = self._m_list(self.nid_map, elem.nodes)

            # Map property ID (most elements)
            if etype not in ('CONROD',) and hasattr(elem, 'pid'):
                elem.pid = self._m(self.pid_map, elem.pid)

            # CONROD: has mid instead of pid
            if etype == 'CONROD' and hasattr(elem, 'mid'):
                elem.mid = self._m(self.mid_map, elem.mid)

            # CBAR/CBEAM: g0 orientation node
            if etype in ('CBAR', 'CBEAM'):
                if hasattr(elem, 'g0') and elem.g0 is not None:
                    if isinstance(elem.g0, int) and elem.g0 > 0:
                        elem.g0 = self._m(self.nid_map, elem.g0)

            # CBUSH: cid
            if etype == 'CBUSH' and hasattr(elem, 'cid'):
                elem.cid = self._m(self.cid_map, elem.cid)

            # theta_mcid for shells (if integer, it's a CID)
            if etype in ('CQUAD4', 'CQUAD8', 'CTRIA3', 'CTRIA6',
                         'CQUADR', 'CTRIAR'):
                if hasattr(elem, 'theta_mcid'):
                    mcid = elem.theta_mcid
                    if isinstance(mcid, int) and mcid > 0:
                        elem.theta_mcid = self._m(self.cid_map, mcid)

            new_elements[new_eid] = elem
        model.elements = new_elements

    def _renumber_rigid_elements(self):
        """Renumber RBE2, RBE3, RBAR."""
        model = self.model
        new_rigid = {}
        for eid, elem in model.rigid_elements.items():
            new_eid = self._m(self.eid_map, eid)
            elem.eid = new_eid
            etype = elem.type

            if etype == 'RBE2':
                if hasattr(elem, 'gn'):
                    elem.gn = self._m(self.nid_map, elem.gn)
                if hasattr(elem, 'Gmi'):
                    elem.Gmi = self._m_list(self.nid_map, elem.Gmi)

            elif etype == 'RBE3':
                if hasattr(elem, 'refgrid'):
                    elem.refgrid = self._m(self.nid_map, elem.refgrid)
                if hasattr(elem, 'Gijs'):
                    elem.Gijs = [self._m_list(self.nid_map, gij_list)
                                 for gij_list in elem.Gijs]

            elif etype == 'RBAR':
                if hasattr(elem, 'nodes'):
                    elem.nodes = self._m_list(self.nid_map, elem.nodes)
                # RBAR may store ga/gb instead of nodes
                if hasattr(elem, 'ga'):
                    elem.ga = self._m(self.nid_map, elem.ga)
                if hasattr(elem, 'gb'):
                    elem.gb = self._m(self.nid_map, elem.gb)

            new_rigid[new_eid] = elem
        model.rigid_elements = new_rigid

    def _renumber_masses(self):
        """Renumber CONM1, CONM2, CMASS1-4."""
        model = self.model
        new_masses = {}
        for eid, elem in model.masses.items():
            new_eid = self._m(self.eid_map, eid)
            elem.eid = new_eid
            etype = elem.type

            if etype in ('CONM2', 'CONM1'):
                if hasattr(elem, 'nid'):
                    elem.nid = self._m(self.nid_map, elem.nid)
                if hasattr(elem, 'cid'):
                    elem.cid = self._m(self.cid_map, elem.cid)

            elif etype in ('CMASS1', 'CMASS2'):
                if hasattr(elem, 'nodes'):
                    elem.nodes = self._m_list(self.nid_map, elem.nodes)
                if hasattr(elem, 'pid') and etype == 'CMASS1':
                    elem.pid = self._m(self.pid_map, elem.pid)

            elif etype in ('CMASS3', 'CMASS4'):
                # CMASS3/4 reference SPOINTs, not grid nodes
                if hasattr(elem, 'nodes'):
                    elem.nodes = self._m_list(self.nid_map, elem.nodes)

            new_masses[new_eid] = elem
        model.masses = new_masses

    def _renumber_properties(self):
        """Renumber all property cards."""
        model = self.model
        new_props = {}
        for pid, prop in model.properties.items():
            new_pid = self._m(self.pid_map, pid)
            prop.pid = new_pid
            ptype = prop.type

            if ptype == 'PSHELL':
                for attr in ('mid1', 'mid2', 'mid3', 'mid4'):
                    val = getattr(prop, attr, None)
                    if val is not None and val > 0:
                        setattr(prop, attr, self._m(self.mid_map, val))

            elif ptype in ('PCOMP', 'PCOMPG', 'PCOMPLS'):
                if hasattr(prop, 'mids'):
                    prop.mids = self._m_list(self.mid_map, prop.mids)

            elif ptype in ('PSOLID', 'PLSOLID'):
                if hasattr(prop, 'mid'):
                    prop.mid = self._m(self.mid_map, prop.mid)
                if hasattr(prop, 'cordm') and prop.cordm is not None:
                    if isinstance(prop.cordm, int) and prop.cordm > 0:
                        prop.cordm = self._m(self.cid_map, prop.cordm)

            elif ptype in ('PBAR', 'PBARL', 'PBEAM', 'PBEAML', 'PROD',
                           'PSHEAR', 'PWELD', 'PFAST', 'PVISC'):
                if hasattr(prop, 'mid'):
                    prop.mid = self._m(self.mid_map, prop.mid)

            # PBUSH, PBUSHT, PELAS, PDAMP, PGAP: no mid/cid refs to remap

            new_props[new_pid] = prop
        model.properties = new_props

    def _renumber_materials(self):
        """Renumber MAT1, MAT2, MAT8, MAT9, MAT10."""
        model = self.model
        new_mats = {}
        for mid, mat in model.materials.items():
            new_mid = self._m(self.mid_map, mid)
            mat.mid = new_mid
            new_mats[new_mid] = mat
        model.materials = new_mats

    def _renumber_coords(self):
        """Renumber CORD2R/C/S, CORD1R/C/S."""
        model = self.model
        new_coords = {}
        for cid, coord in model.coords.items():
            if cid == 0:
                new_coords[0] = coord
                continue
            new_cid = self._m(self.cid_map, cid)
            coord.cid = new_cid
            if hasattr(coord, 'rid') and coord.rid is not None:
                if isinstance(coord.rid, int):
                    coord.rid = self._m(self.cid_map, coord.rid)
            new_coords[new_cid] = coord
        model.coords = new_coords

    def _renumber_spcs(self):
        """Renumber SPC, SPC1, SPCADD."""
        model = self.model

        # SPCs are stored in model.spcs as {spc_id: [card, ...]}
        if hasattr(model, 'spcs'):
            new_spcs = {}
            for sid, spc_list in model.spcs.items():
                new_sid = self._m(self.spc_map, sid)
                for card in spc_list:
                    card.conid = new_sid
                    ctype = card.type
                    if ctype == 'SPC':
                        if hasattr(card, 'nodes'):
                            card.nodes = self._m_list(self.nid_map, card.nodes)
                        # Some versions use 'gids'
                        if hasattr(card, 'gids'):
                            card.gids = self._m_list(self.nid_map, card.gids)
                    elif ctype == 'SPC1':
                        if hasattr(card, 'nodes'):
                            card.nodes = self._m_list(self.nid_map, card.nodes)
                new_spcs.setdefault(new_sid, []).extend(spc_list)
            model.spcs = new_spcs

        # SPCADD
        if hasattr(model, 'spcadds'):
            new_spcadds = {}
            for sid, add_list in model.spcadds.items():
                new_sid = self._m(self.spc_map, sid)
                for card in add_list:
                    card.conid = new_sid
                    if hasattr(card, 'spc_ids'):
                        card.spc_ids = self._m_list(self.spc_map, card.spc_ids)
                    if hasattr(card, 'sets'):
                        card.sets = self._m_list(self.spc_map, card.sets)
                new_spcadds.setdefault(new_sid, []).extend(add_list)
            model.spcadds = new_spcadds

    def _renumber_mpcs(self):
        """Renumber MPC, MPCADD."""
        model = self.model

        if hasattr(model, 'mpcs'):
            new_mpcs = {}
            for sid, mpc_list in model.mpcs.items():
                new_sid = self._m(self.mpc_map, sid)
                for card in mpc_list:
                    card.conid = new_sid
                    if hasattr(card, 'nodes'):
                        card.nodes = self._m_list(self.nid_map, card.nodes)
                    if hasattr(card, 'gids'):
                        card.gids = self._m_list(self.nid_map, card.gids)
                new_mpcs.setdefault(new_sid, []).extend(mpc_list)
            model.mpcs = new_mpcs

        if hasattr(model, 'mpcadds'):
            new_mpcadds = {}
            for sid, add_list in model.mpcadds.items():
                new_sid = self._m(self.mpc_map, sid)
                for card in add_list:
                    card.conid = new_sid
                    if hasattr(card, 'mpc_ids'):
                        card.mpc_ids = self._m_list(self.mpc_map, card.mpc_ids)
                    if hasattr(card, 'sets'):
                        card.sets = self._m_list(self.mpc_map, card.sets)
                new_mpcadds.setdefault(new_sid, []).extend(add_list)
            model.mpcadds = new_mpcadds

    def _renumber_loads(self):
        """Renumber FORCE, MOMENT, PLOAD4, GRAV, LOAD, TEMP, TEMPD, DLOAD, etc."""
        model = self.model

        if hasattr(model, 'loads'):
            new_loads = {}
            for sid, load_list in model.loads.items():
                new_sid = self._m(self.load_map, sid)
                for card in load_list:
                    card.sid = new_sid
                    ltype = card.type

                    if ltype in ('FORCE', 'MOMENT'):
                        if hasattr(card, 'node'):
                            card.node = self._m(self.nid_map, card.node)
                        if hasattr(card, 'cid'):
                            card.cid = self._m(self.cid_map, card.cid)

                    elif ltype == 'PLOAD4':
                        if hasattr(card, 'eids'):
                            card.eids = self._m_list(self.eid_map, card.eids)
                        if hasattr(card, 'eid'):
                            card.eid = self._m(self.eid_map, card.eid)
                        if hasattr(card, 'g1') and card.g1:
                            card.g1 = self._m(self.nid_map, card.g1)
                        if hasattr(card, 'g34') and card.g34:
                            card.g34 = self._m(self.nid_map, card.g34)
                        if hasattr(card, 'cid'):
                            card.cid = self._m(self.cid_map, card.cid)

                    elif ltype == 'GRAV':
                        if hasattr(card, 'cid'):
                            card.cid = self._m(self.cid_map, card.cid)

                    elif ltype == 'LOAD':
                        if hasattr(card, 'load_ids'):
                            card.load_ids = self._m_list(
                                self.load_map, card.load_ids)

                    elif ltype == 'TEMP':
                        if hasattr(card, 'temperatures'):
                            new_temps = {}
                            for nid, temp in card.temperatures.items():
                                new_nid = self._m(self.nid_map, nid)
                                new_temps[new_nid] = temp
                            card.temperatures = new_temps

                    elif ltype == 'DLOAD':
                        if hasattr(card, 'load_ids'):
                            card.load_ids = self._m_list(
                                self.load_map, card.load_ids)

                    elif ltype == 'DAREA':
                        if hasattr(card, 'nodes'):
                            card.nodes = self._m_list(self.nid_map, card.nodes)
                        # Some versions use node_id
                        if hasattr(card, 'node_id'):
                            card.node_id = self._m(self.nid_map, card.node_id)

                    elif ltype in ('RLOAD1', 'RLOAD2', 'TLOAD1', 'TLOAD2'):
                        if hasattr(card, 'excite_id'):
                            card.excite_id = self._m(
                                self.load_map, card.excite_id)
                        if hasattr(card, 'tid') and isinstance(card.tid, int):
                            card.tid = self._m(self.table_map, card.tid)
                        if hasattr(card, 'tid1') and isinstance(card.tid1, int):
                            card.tid1 = self._m(self.table_map, card.tid1)

                    elif ltype in ('PLOAD', 'PLOAD2'):
                        if hasattr(card, 'nodes'):
                            card.nodes = self._m_list(self.nid_map, card.nodes)
                        if hasattr(card, 'eids'):
                            card.eids = self._m_list(self.eid_map, card.eids)
                        if hasattr(card, 'eid'):
                            card.eid = self._m(self.eid_map, card.eid)

                    elif ltype == 'RFORCE':
                        if hasattr(card, 'nid'):
                            card.nid = self._m(self.nid_map, card.nid)
                        if hasattr(card, 'cid'):
                            card.cid = self._m(self.cid_map, card.cid)

                new_loads.setdefault(new_sid, []).extend(load_list)
            model.loads = new_loads

        # DLOADS stored separately in some versions
        if hasattr(model, 'dloads'):
            new_dloads = {}
            for sid, dload_list in model.dloads.items():
                new_sid = self._m(self.load_map, sid)
                for card in dload_list:
                    card.sid = new_sid
                    if hasattr(card, 'load_ids'):
                        card.load_ids = self._m_list(
                            self.load_map, card.load_ids)
                new_dloads.setdefault(new_sid, []).extend(dload_list)
            model.dloads = new_dloads

        if hasattr(model, 'dload_entries'):
            new_entries = {}
            for sid, entry_list in model.dload_entries.items():
                new_sid = self._m(self.load_map, sid)
                for card in entry_list:
                    card.sid = new_sid
                    if hasattr(card, 'excite_id'):
                        card.excite_id = self._m(
                            self.load_map, card.excite_id)
                    if hasattr(card, 'tid') and isinstance(card.tid, int):
                        card.tid = self._m(self.table_map, card.tid)
                new_entries.setdefault(new_sid, []).extend(entry_list)
            model.dload_entries = new_entries

    def _renumber_contact(self):
        """Renumber BSURF, BSURFS, BCTSET, BCTADD, BCONP, BCBODY, BLSEG."""
        model = self.model

        # BSURF
        if hasattr(model, 'bsurf'):
            new_bsurf = {}
            for sid, card in model.bsurf.items():
                new_sid = self._m(self.contact_map, sid)
                card.sid = new_sid
                if hasattr(card, 'eids'):
                    card.eids = self._m_list(self.eid_map, card.eids)
                new_bsurf[new_sid] = card
            model.bsurf = new_bsurf

        # BSURFS
        if hasattr(model, 'bsurfs'):
            new_bsurfs = {}
            for sid, card in model.bsurfs.items():
                new_sid = self._m(self.contact_map, sid)
                card.sid = new_sid
                if hasattr(card, 'eids'):
                    card.eids = self._m_list(self.eid_map, card.eids)
                if hasattr(card, 'nodes'):
                    card.nodes = self._m_list(self.nid_map, card.nodes)
                new_bsurfs[new_sid] = card
            model.bsurfs = new_bsurfs

        # BCTSET
        if hasattr(model, 'bctsets'):
            new_bctsets = {}
            for sid, card in model.bctsets.items():
                new_sid = self._m(self.contact_map, sid)
                card.csid = new_sid
                if hasattr(card, 'sids'):
                    card.sids = self._m_list(self.contact_map, card.sids)
                if hasattr(card, 'tids'):
                    card.tids = self._m_list(self.contact_map, card.tids)
                new_bctsets[new_sid] = card
            model.bctsets = new_bctsets

        # BCTADD
        if hasattr(model, 'bctadds'):
            new_bctadds = {}
            for sid, card in model.bctadds.items():
                new_sid = self._m(self.contact_map, sid)
                card.csid = new_sid
                if hasattr(card, 'contact_sets'):
                    card.contact_sets = self._m_list(
                        self.contact_map, card.contact_sets)
                new_bctadds[new_sid] = card
            model.bctadds = new_bctadds

        # BCONP
        if hasattr(model, 'bconp'):
            new_bconp = {}
            for sid, card in model.bconp.items():
                new_sid = self._m(self.contact_map, sid)
                card.contact_id = new_sid
                if hasattr(card, 'slave'):
                    card.slave = self._m(self.contact_map, card.slave)
                if hasattr(card, 'master'):
                    card.master = self._m(self.contact_map, card.master)
                if hasattr(card, 'cid'):
                    card.cid = self._m(self.cid_map, card.cid)
                new_bconp[new_sid] = card
            model.bconp = new_bconp

        # BCBODY
        if hasattr(model, 'bcbodys'):
            new_bcbodys = {}
            for sid, card in model.bcbodys.items():
                new_sid = self._m(self.contact_map, sid)
                if hasattr(card, 'contact_id'):
                    card.contact_id = new_sid
                if hasattr(card, 'bsid'):
                    card.bsid = self._m(self.contact_map, card.bsid)
                new_bcbodys[new_sid] = card
            model.bcbodys = new_bcbodys

        # BLSEG
        if hasattr(model, 'blsegs'):
            new_blsegs = {}
            for sid, card in model.blsegs.items():
                new_sid = self._m(self.contact_map, sid)
                if hasattr(card, 'sid'):
                    card.sid = new_sid
                if hasattr(card, 'nodes'):
                    card.nodes = self._m_list(self.nid_map, card.nodes)
                new_blsegs[new_sid] = card
            model.blsegs = new_blsegs

    def _renumber_sets(self):
        """Renumber SET1, SET3."""
        model = self.model

        if hasattr(model, 'sets'):
            new_sets = {}
            for sid, card in model.sets.items():
                new_sid = self._m(self.set_map, sid)
                card.sid = new_sid

                # Heuristic: check if IDs look like nodes or elements
                if hasattr(card, 'ids') and card.ids:
                    id_set = set(card.ids)
                    node_overlap = len(id_set & set(self.nid_map.keys()))
                    elem_overlap = len(id_set & set(self.eid_map.keys()))

                    if node_overlap >= elem_overlap:
                        card.ids = self._m_list(self.nid_map, card.ids)
                    else:
                        card.ids = self._m_list(self.eid_map, card.ids)

                new_sets[new_sid] = card
            model.sets = new_sets

    def _renumber_methods(self):
        """Renumber EIGRL, EIGR."""
        model = self.model
        if hasattr(model, 'methods'):
            new_methods = {}
            for sid, card in model.methods.items():
                new_sid = self._m(self.method_map, sid)
                card.sid = new_sid
                new_methods[new_sid] = card
            model.methods = new_methods

    def _renumber_tables(self):
        """Renumber TABLED1, TABLEM1."""
        model = self.model
        if hasattr(model, 'tables'):
            new_tables = {}
            for tid, card in model.tables.items():
                new_tid = self._m(self.table_map, tid)
                card.tid = new_tid
                new_tables[new_tid] = card
            model.tables = new_tables

    def _renumber_suport(self):
        """Renumber SUPORT and SUPORT1 node references."""
        model = self.model

        if hasattr(model, 'suport') and model.suport:
            for card in model.suport:
                if hasattr(card, 'nodes'):
                    card.nodes = self._m_list(self.nid_map, card.nodes)
                if hasattr(card, 'IDs'):
                    card.IDs = self._m_list(self.nid_map, card.IDs)

        if hasattr(model, 'suport1'):
            new_suport1 = {}
            for sid, card in model.suport1.items():
                if hasattr(card, 'nodes'):
                    card.nodes = self._m_list(self.nid_map, card.nodes)
                if hasattr(card, 'IDs'):
                    card.IDs = self._m_list(self.nid_map, card.IDs)
                new_suport1[sid] = card
            model.suport1 = new_suport1


# ═══════════════════════════════════════════════════════════════════════════════
# Section 5: CaseControlRenumberer
# ═══════════════════════════════════════════════════════════════════════════════

class CaseControlRenumberer:
    """Update ID references in the case control deck."""

    # (keyword, map_key) pairs
    ENTRIES = [
        ('LOAD', 'load_id'),
        ('SPC', 'spc_id'),
        ('MPC', 'mpc_id'),
        ('METHOD', 'method_id'),
        ('CMETHOD', 'method_id'),
        ('DLOAD', 'load_id'),
        ('FREQ', 'table_id'),
        ('TSTEP', 'table_id'),
        ('SDAMP', 'table_id'),
        ('DEFORM', 'load_id'),
        ('SUPORT1', 'set_id'),
    ]

    # Pattern for TEMPERATURE(LOAD) = id, TEMPERATURE(INITIAL) = id
    TEMP_RE = re.compile(
        r'(TEMPERATURE\s*\(\s*(?:LOAD|INITIAL)\s*\)\s*=\s*)(\d+)',
        re.IGNORECASE)

    def __init__(self, maps, include_set_ids=True):
        self.maps = maps
        self.include_set_ids = include_set_ids

    def renumber_case_control(self, case_control_lines):
        """Renumber IDs in case control lines. Returns new list of lines."""
        new_lines = []
        for line in case_control_lines:
            new_line = self._process_line(line)
            new_lines.append(new_line)
        return new_lines

    def _process_line(self, line):
        """Process a single case control line."""
        stripped = line.strip().upper()

        for keyword, map_key in self.ENTRIES:
            if not self.include_set_ids and map_key in (
                    'spc_id', 'mpc_id', 'load_id'):
                continue

            id_map = self.maps.get(map_key, {})
            if not id_map:
                continue

            pattern = re.compile(
                rf'({keyword}\s*[=(]\s*)(\d+)', re.IGNORECASE)
            match = pattern.search(line)
            if match:
                old_id = int(match.group(2))
                new_id = id_map.get(old_id, old_id)
                line = line[:match.start(2)] + str(new_id) + line[match.end(2):]

        # TEMPERATURE entries
        for map_key in ('load_id',):
            if not self.include_set_ids:
                continue
            id_map = self.maps.get(map_key, {})
            if not id_map:
                continue
            match = self.TEMP_RE.search(line)
            if match:
                old_id = int(match.group(2))
                new_id = id_map.get(old_id, old_id)
                line = (line[:match.start(2)] + str(new_id)
                        + line[match.end(2):])

        return line


# ═══════════════════════════════════════════════════════════════════════════════
# Section 6: OutputWriter
# ═══════════════════════════════════════════════════════════════════════════════

class OutputWriter:
    """Write renumbered cards to per-file output preserving include structure."""

    # Card write order categories
    CARD_ORDER = [
        # 1. Coordinate systems
        ('CORD2R', 'CORD2C', 'CORD2S', 'CORD1R', 'CORD1C', 'CORD1S'),
        # 2. Nodes
        ('GRID', 'SPOINT'),
        # 3. Solid/shell/beam elements
        ('CHEXA', 'CPENTA', 'CTETRA', 'CQUAD4', 'CQUAD8', 'CTRIA3',
         'CTRIA6', 'CQUADR', 'CTRIAR', 'CSHEAR',
         'CBAR', 'CBEAM', 'CROD', 'CONROD', 'CBUSH',
         'CELAS1', 'CELAS2', 'CELAS3', 'CELAS4',
         'CDAMP1', 'CDAMP2', 'CDAMP3', 'CDAMP4',
         'CGAP', 'CWELD', 'CFAST', 'CVISC', 'PLOTEL',
         'CHBDYG', 'CHBDYE'),
        # 4. Rigid elements
        ('RBE2', 'RBE3', 'RBAR'),
        # 5. Mass elements
        ('CONM1', 'CONM2', 'CMASS1', 'CMASS2', 'CMASS3', 'CMASS4'),
        # 6. Properties
        ('PSHELL', 'PCOMP', 'PCOMPG', 'PCOMPLS', 'PSOLID', 'PLSOLID',
         'PBAR', 'PBARL', 'PBEAM', 'PBEAML', 'PROD',
         'PBUSH', 'PBUSHT', 'PELAS', 'PDAMP', 'PGAP',
         'PSHEAR', 'PWELD', 'PFAST', 'PVISC'),
        # 7. Materials
        ('MAT1', 'MAT2', 'MAT8', 'MAT9', 'MAT10'),
        # 8. Loads
        ('FORCE', 'MOMENT', 'PLOAD', 'PLOAD2', 'PLOAD4', 'GRAV',
         'RFORCE', 'TEMP', 'TEMPD', 'DAREA'),
        # 9. Load combinations
        ('LOAD', 'DLOAD'),
        # 10. Dynamic loads
        ('RLOAD1', 'RLOAD2', 'TLOAD1', 'TLOAD2'),
        # 11. Constraints
        ('SPC', 'SPC1', 'SPCADD', 'MPC', 'MPCADD'),
        # 12. Contact
        ('BSURF', 'BSURFS', 'BCTSET', 'BCTADD', 'BCONP', 'BCBODY',
         'BCTPARA', 'BCTPARM', 'BLSEG', 'BFRIC'),
        # 13. Sets
        ('SET1', 'SET3'),
        # 14. Methods
        ('EIGRL', 'EIGR'),
        # 15. Tables
        ('TABLED1', 'TABLEM1'),
    ]

    def __init__(self, model, parser, maps, case_renumberer, include_set_ids,
                 log_func=None):
        """
        Args:
            model: renumbered pyNastran BDF model (uncross-referenced)
            parser: IncludeFileParser with file ownership info
            maps: dict[entity_type, dict[old_id, new_id]]
            case_renumberer: CaseControlRenumberer instance
            include_set_ids: bool
            log_func: optional callable for diagnostic messages
        """
        self.model = model
        self.parser = parser
        self.maps = maps
        self.case_renumberer = case_renumberer
        self.include_set_ids = include_set_ids
        self._log = log_func or (lambda msg: None)

    # All model dicts that may contain cards, keyed by entity type
    _MODEL_DICTS = {
        'nid': [('nodes', False), ('spoints', False)],
        'eid': [('elements', False), ('rigid_elements', False),
                ('masses', False), ('plotels', False)],
        'pid': [('properties', False)],
        'mid': [('materials', False)],
        'cid': [('coords', False)],
        'spc_id': [('spcs', True), ('spcadds', True)],
        'mpc_id': [('mpcs', True), ('mpcadds', True)],
        'load_id': [('loads', True), ('dloads', True),
                    ('dload_entries', True)],
        'contact_id': [('bsurf', False), ('bsurfs', False),
                       ('bctsets', False), ('bctadds', False),
                       ('bconp', False), ('bcbodys', False),
                       ('blsegs', False), ('bfriction', False),
                       ('bctparas', False), ('bctparms', False)],
        'set_id': [('sets', False)],
        'method_id': [('methods', False)],
        'table_id': [('tables', False)],
    }

    def write(self, output_dir):
        """Write all files to output_dir. Returns list of written files."""
        os.makedirs(output_dir, exist_ok=True)

        main_path = self.parser.all_files[0]
        written_files = []

        # Write include files first (so we know their output paths)
        include_out_paths = {}
        for filepath in self.parser.all_files[1:]:
            rel = os.path.relpath(filepath, os.path.dirname(main_path))
            out_path = os.path.join(output_dir, rel)
            os.makedirs(os.path.dirname(out_path), exist_ok=True)

            self._write_include_file(filepath, out_path)
            include_out_paths[filepath] = out_path
            written_files.append(out_path)

        # Write main file
        main_out = os.path.join(output_dir, os.path.basename(main_path))
        self._write_main_file(main_path, main_out, include_out_paths)
        written_files.insert(0, main_out)

        return written_files

    def _get_new_ids_for_file(self, filepath, entity_type):
        """Get the set of NEW ids that belong to this file."""
        id_map = self.maps.get(entity_type, {})
        orig_ids = self.parser.file_ids[filepath].get(entity_type, set())
        return {id_map.get(old_id, old_id) for old_id in orig_ids}

    @staticmethod
    def _write_card_safe(card):
        """Write a single card, returning the string or None on failure."""
        try:
            return card.write_card(size=8)
        except Exception:
            try:
                return str(card)
            except Exception:
                return None

    def _write_ordered_cards(self, filepath):
        """Write cards via CARD_ORDER. Returns (lines, written_ids_by_etype)."""
        lines = []
        written_ids = defaultdict(set)  # {entity_type: set(new_ids)}

        for card_group in self.CARD_ORDER:
            group_lines = []
            for card_type in card_group:
                entity_type = CARD_ENTITY_MAP.get(card_type)
                cards = self._get_cards_for_file(filepath, card_type)
                for card in cards:
                    text = self._write_card_safe(card)
                    if text:
                        group_lines.append(text)
                        if entity_type:
                            cid = getattr(card, 'eid', None) or \
                                  getattr(card, 'nid', None) or \
                                  getattr(card, 'pid', None) or \
                                  getattr(card, 'mid', None) or \
                                  getattr(card, 'cid', None) or \
                                  getattr(card, 'sid', None) or \
                                  getattr(card, 'tid', None) or \
                                  getattr(card, 'conid', None) or \
                                  getattr(card, 'contact_id', None)
                            if cid is not None:
                                written_ids[entity_type].add(cid)
            if group_lines:
                lines.extend(group_lines)

        return lines, written_ids

    def _write_remaining_cards(self, filepath, written_ids):
        """Fallback: write any cards belonging to this file not yet written."""
        lines = []
        fname = os.path.basename(filepath)

        for etype in ENTITY_TYPES:
            new_ids = self._get_new_ids_for_file(filepath, etype)
            if not new_ids:
                continue

            already = written_ids.get(etype, set())
            remaining = new_ids - already
            if not remaining:
                continue

            # Search all model dicts for this entity type
            dict_specs = self._MODEL_DICTS.get(etype, [])
            for attr_name, is_list_dict in dict_specs:
                d = getattr(self.model, attr_name, None)
                if not d or not isinstance(d, dict):
                    continue
                for card_id, card_or_list in d.items():
                    if card_id not in remaining:
                        continue
                    cards = card_or_list if is_list_dict and \
                        isinstance(card_or_list, list) else [card_or_list]
                    for card in cards:
                        text = self._write_card_safe(card)
                        if text:
                            lines.append(text)
                            remaining.discard(card_id)

            if remaining:
                self._log(f"  WARNING: {fname}/{etype}: "
                          f"{len(remaining)} IDs not found in model dicts")

        return lines

    def _log_diagnostics(self, filepath, written_ids):
        """Log card counts per entity type for diagnostics."""
        fname = os.path.basename(filepath)
        for etype in ENTITY_TYPES:
            expected = self._get_new_ids_for_file(filepath, etype)
            n_written = len(written_ids.get(etype, set()))
            n_expected = len(expected)
            if n_expected == 0:
                continue
            if n_written != n_expected:
                self._log(f"  DIAG {fname}/{etype}: wrote {n_written}/"
                          f"{n_expected}")

    def _write_include_file(self, orig_path, out_path):
        """Write a single include file with its renumbered cards."""
        lines = [f'$ Renumbered from: {os.path.basename(orig_path)}\n']

        ordered_lines, written_ids = self._write_ordered_cards(orig_path)
        lines.extend(ordered_lines)

        # Fallback: catch any cards missed by CARD_ORDER
        remaining_lines = self._write_remaining_cards(orig_path, written_ids)
        if remaining_lines:
            lines.append('$ --- Fallback cards (not in CARD_ORDER) ---\n')
            lines.extend(remaining_lines)

        self._log_diagnostics(orig_path, written_ids)

        with open(out_path, 'w') as f:
            f.writelines(lines)

    def _write_main_file(self, orig_path, out_path, include_out_paths):
        """Write the main BDF file with updated case control and includes."""
        lines = []

        # 1. Executive control — copy verbatim from original
        exec_lines, case_lines = self._read_sections(orig_path)
        lines.extend(exec_lines)

        # 2. Case control — with IDs updated
        if case_lines:
            updated_case = self.case_renumberer.renumber_case_control(
                case_lines)
            lines.extend(updated_case)

        # 3. BEGIN BULK
        lines.append('BEGIN BULK\n')

        # 4. INCLUDE statements with updated paths
        main_dir = os.path.dirname(out_path)
        for orig_inc_path in self.parser.file_tree.get(orig_path, []):
            if orig_inc_path in include_out_paths:
                inc_out = include_out_paths[orig_inc_path]
                rel_path = os.path.relpath(inc_out, main_dir)
                lines.append(f"INCLUDE '{rel_path}'\n")

        # 5. Main file's own bulk data cards
        ordered_lines, written_ids = self._write_ordered_cards(orig_path)
        lines.extend(ordered_lines)

        # Fallback: catch any cards missed by CARD_ORDER
        remaining_lines = self._write_remaining_cards(orig_path, written_ids)
        if remaining_lines:
            lines.append('$ --- Fallback cards (not in CARD_ORDER) ---\n')
            lines.extend(remaining_lines)

        self._log_diagnostics(orig_path, written_ids)

        # 6. ENDDATA
        lines.append('ENDDATA\n')

        with open(out_path, 'w') as f:
            f.writelines(lines)

    def _get_cards_for_file(self, filepath, card_type):
        """Get all card objects of a given type that belong to the file."""
        cards = []
        model = self.model
        entity_type = CARD_ENTITY_MAP.get(card_type)
        if entity_type is None:
            return cards

        new_ids = self._get_new_ids_for_file(filepath, entity_type)
        if not new_ids:
            return cards

        # Map card type to model dict
        card_dicts = self._get_card_dict(card_type)
        for card_dict in card_dicts:
            for card_id, card in card_dict.items():
                if isinstance(card, list):
                    # Some dicts store lists (loads, spcs, mpcs)
                    for c in card:
                        if c.type == card_type and card_id in new_ids:
                            cards.append(c)
                else:
                    if card.type == card_type and card_id in new_ids:
                        cards.append(card)

        return cards

    def _get_card_dict(self, card_type):
        """Return the model dict(s) that store cards of this type."""
        model = self.model
        entity_type = CARD_ENTITY_MAP.get(card_type)

        if entity_type == 'nid':
            if card_type == 'SPOINT':
                return [getattr(model, 'spoints', {})]
            return [model.nodes]
        elif entity_type == 'eid':
            if card_type in ('RBE2', 'RBE3', 'RBAR'):
                return [model.rigid_elements]
            if card_type in ('CONM1', 'CONM2', 'CMASS1', 'CMASS2',
                             'CMASS3', 'CMASS4'):
                return [model.masses]
            if card_type == 'PLOTEL':
                return [getattr(model, 'plotels', {})]
            return [model.elements]
        elif entity_type == 'pid':
            return [model.properties]
        elif entity_type == 'mid':
            return [model.materials]
        elif entity_type == 'cid':
            return [model.coords]
        elif entity_type == 'spc_id':
            dicts = []
            if card_type == 'SPCADD':
                dicts.append(getattr(model, 'spcadds', {}))
            else:
                dicts.append(getattr(model, 'spcs', {}))
            return dicts
        elif entity_type == 'mpc_id':
            dicts = []
            if card_type == 'MPCADD':
                dicts.append(getattr(model, 'mpcadds', {}))
            else:
                dicts.append(getattr(model, 'mpcs', {}))
            return dicts
        elif entity_type == 'load_id':
            dicts = [getattr(model, 'loads', {})]
            if card_type == 'DLOAD':
                dicts.append(getattr(model, 'dloads', {}))
            if card_type in ('RLOAD1', 'RLOAD2', 'TLOAD1', 'TLOAD2'):
                dicts.append(getattr(model, 'dload_entries', {}))
            return dicts
        elif entity_type == 'contact_id':
            dicts = []
            attr_map = {
                'BSURF': 'bsurf', 'BSURFS': 'bsurfs',
                'BCTSET': 'bctsets', 'BCTADD': 'bctadds',
                'BCONP': 'bconp', 'BCBODY': 'bcbodys',
                'BLSEG': 'blsegs', 'BFRIC': 'bfriction',
                'BCTPARA': 'bctparas', 'BCTPARM': 'bctparms',
            }
            attr = attr_map.get(card_type)
            if attr:
                dicts.append(getattr(model, attr, {}))
            return dicts
        elif entity_type == 'set_id':
            return [getattr(model, 'sets', {})]
        elif entity_type == 'method_id':
            return [getattr(model, 'methods', {})]
        elif entity_type == 'table_id':
            return [getattr(model, 'tables', {})]

        return [{}]

    def _read_sections(self, filepath):
        """Read a BDF file and split into exec control lines and case control lines."""
        exec_lines = []
        case_lines = []

        if not os.path.isfile(filepath):
            return exec_lines, case_lines

        with open(filepath, 'r', errors='replace') as f:
            lines = f.readlines()

        section = 'exec'  # exec -> case -> bulk
        for line in lines:
            upper = line.strip().upper()
            if section == 'exec':
                exec_lines.append(line)
                if upper.startswith('CEND'):
                    section = 'case'
            elif section == 'case':
                if upper.startswith('BEGIN') and 'BULK' in upper:
                    section = 'bulk'
                    break
                case_lines.append(line)
            # Stop at BEGIN BULK

        return exec_lines, case_lines


# ═══════════════════════════════════════════════════════════════════════════════
# Section 7: GUI (CustomTkinter + tksheet)
# ═══════════════════════════════════════════════════════════════════════════════

class RenumberIncludesTool(ctk.CTkFrame):
    """Main GUI application for include file renumbering."""

    def __init__(self, parent=None):
        super().__init__(parent)

        self._bdf_path = None
        self._parser = None
        self._summary = None        # {filepath: {etype: (count, min, max)}}
        self._include_set_ids = tk.BooleanVar(value=True)

        # Row data for sheet: list of tuples linking rows to (filepath, [etypes])
        self._simple_row_map = []  # [(filepath, [etypes...]), ...]

        self._build_ui()

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top bar: input file ──
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill=tk.X, padx=10, pady=(10, 4))

        ctk.CTkLabel(top, text="Input BDF:").pack(side=tk.LEFT)
        self._path_var = tk.StringVar()
        ctk.CTkEntry(top, textvariable=self._path_var, width=480).pack(
            side=tk.LEFT, padx=6)
        ctk.CTkButton(top, text="Browse\u2026", width=90,
                      command=self._browse_input).pack(side=tk.LEFT)
        ctk.CTkButton(top, text="Scan", width=70,
                      command=self._scan).pack(side=tk.LEFT, padx=6)

        # ── Options row ──
        opt = ctk.CTkFrame(self, fg_color="transparent")
        opt.pack(fill=tk.X, padx=10, pady=2)
        ctk.CTkCheckBox(
            opt, text="Include set IDs (SPC/MPC/Load)",
            variable=self._include_set_ids).pack(side=tk.LEFT)

        # ── Start ID + Suggest Ranges ──
        suggest_frame = ctk.CTkFrame(self, fg_color="transparent")
        suggest_frame.pack(fill=tk.X, padx=10, pady=(4, 2))

        ctk.CTkLabel(suggest_frame, text="Start ID:").pack(
            side=tk.LEFT, padx=(0, 4))
        self._start_id_var = tk.StringVar(value="1")
        ctk.CTkEntry(suggest_frame, textvariable=self._start_id_var,
                      width=100).pack(side=tk.LEFT, padx=(0, 8))
        ctk.CTkButton(suggest_frame, text="Suggest Ranges", width=120,
                       command=self._suggest_ranges).pack(side=tk.LEFT)

        # ── Table container ──
        ctk.CTkLabel(self, text="Entity Summary & Range Editor:",
                     font=ctk.CTkFont(weight="bold")).pack(
            anchor=tk.W, padx=10, pady=(6, 2))

        self._table_container = ctk.CTkFrame(self)
        self._table_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=2)

        # Simple sheet (the only sheet)
        self._simple_sheet = tksheet.Sheet(
            self._table_container,
            headers=['File', 'Entity Types', 'Total Count',
                     'New Start', 'New End'],
            show_x_scrollbar=True, show_y_scrollbar=True,
            height=300)
        self._simple_sheet.enable_bindings(
            "single_select", "column_select", "row_select",
            "arrowkeys", "edit_cell", "copy", "paste",
            "column_width_resize")
        self._simple_sheet.readonly_columns(columns=[0, 1, 2])
        self._simple_sheet.pack(fill=tk.BOTH, expand=True)
        self._simple_sheet.bind("<<SheetModified>>",
                                 self._on_simple_sheet_modified)

        # ── Output dir ──
        out_bar = ctk.CTkFrame(self, fg_color="transparent")
        out_bar.pack(fill=tk.X, padx=10, pady=4)
        ctk.CTkLabel(out_bar, text="Output Dir:").pack(side=tk.LEFT)
        self._outdir_var = tk.StringVar()
        ctk.CTkEntry(out_bar, textvariable=self._outdir_var, width=480).pack(
            side=tk.LEFT, padx=6)
        ctk.CTkButton(out_bar, text="Browse\u2026", width=90,
                      command=self._browse_output).pack(side=tk.LEFT)

        # ── Action buttons ──
        btn_bar = ctk.CTkFrame(self, fg_color="transparent")
        btn_bar.pack(fill=tk.X, padx=10, pady=4)
        self._validate_btn = ctk.CTkButton(
            btn_bar, text="Validate", command=self._validate,
            state=tk.DISABLED, width=100)
        self._validate_btn.pack(side=tk.LEFT, padx=4)
        self._apply_btn = ctk.CTkButton(
            btn_bar, text="Apply Renumbering", command=self._apply,
            state=tk.DISABLED, width=140)
        self._apply_btn.pack(side=tk.LEFT, padx=4)
        self._save_cfg_btn = ctk.CTkButton(
            btn_bar, text="Save Config", command=self._save_config,
            state=tk.DISABLED, width=100)
        self._save_cfg_btn.pack(side=tk.LEFT, padx=4)
        self._load_cfg_btn = ctk.CTkButton(
            btn_bar, text="Load Config", command=self._load_config,
            width=100)
        self._load_cfg_btn.pack(side=tk.LEFT, padx=4)

        # ── Log pane ──
        ctk.CTkLabel(self, text="Log:").pack(anchor=tk.W, padx=10)
        self._log = ctk.CTkTextbox(self, height=120, state=tk.DISABLED,
                                   wrap=tk.WORD)
        self._log.pack(fill=tk.X, padx=10, pady=(0, 4))

        # ── Status bar ──
        self._status_var = tk.StringVar(value="Ready")
        self._status_label = ctk.CTkLabel(
            self, textvariable=self._status_var, anchor=tk.W)
        self._status_label.pack(fill=tk.X, padx=10, pady=(0, 6))

    # ── Suggest ranges / cascading ───────────────────────────────────────────

    def _suggest_ranges(self):
        """Auto-suggest round-number ranges starting from the Start ID."""
        import math

        if not self._simple_row_map:
            return

        try:
            cursor = int(self._start_id_var.get().strip())
        except (ValueError, TypeError):
            messagebox.showerror("Invalid Start ID",
                                 "Please enter a valid integer for Start ID.")
            return

        data = self._simple_sheet.get_sheet_data()
        for i, (filepath, etypes) in enumerate(self._simple_row_map):
            if i >= len(data):
                break
            try:
                total_count = int(str(data[i][2]).strip())
            except (ValueError, TypeError):
                total_count = 0
            if total_count == 0:
                data[i][3] = str(cursor)
                data[i][4] = str(cursor)
                cursor += 1
                continue

            raw_end = cursor + total_count - 1
            if raw_end <= 0:
                mag = 1
            else:
                mag = 10 ** int(math.floor(math.log10(raw_end)))
            rounded_end = int(math.ceil(raw_end / mag) * mag)
            data[i][3] = str(cursor)
            data[i][4] = str(rounded_end)
            cursor = rounded_end + 1

        self._simple_sheet.set_sheet_data(data)

    def _on_simple_sheet_modified(self, event=None):
        """Cascade New Start values when a New End cell is edited."""
        data = self._simple_sheet.get_sheet_data()
        if not data or not self._simple_row_map:
            return

        # Find which cells were edited — cascade from first changed New End
        # We just re-cascade all rows: each New Start = prev New End + 1
        for i in range(1, len(data)):
            try:
                prev_end = int(str(data[i - 1][4]).strip())
                data[i][3] = str(prev_end + 1)
            except (ValueError, TypeError):
                pass

        self._simple_sheet.set_sheet_data(data)

    def _auto_allocate(self, file_ranges):
        """Auto-allocate sub-ranges per entity type within a file's range.

        Args:
            file_ranges: dict[filepath, (start, end)]

        Returns:
            dict[filepath, dict[etype, (start, end)]]
        """
        result = {}
        for filepath, (file_start, file_end) in file_ranges.items():
            etypes_present = []
            for etype in ENTITY_TYPES:
                ids = self._parser.file_ids.get(filepath, {}).get(etype, set())
                if ids:
                    etypes_present.append(etype)

            if not etypes_present:
                continue

            total_range = file_end - file_start + 1
            n_types = len(etypes_present)
            block_size = total_range // n_types

            alloc = {}
            for j, etype in enumerate(etypes_present):
                block_start = file_start + j * block_size
                if j == n_types - 1:
                    block_end = file_end  # last block gets remainder
                else:
                    block_end = block_start + block_size - 1
                alloc[etype] = (block_start, block_end)

            result[filepath] = alloc
        return result

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _log_msg(self, msg):
        """Append a message to the log pane."""
        self._log.configure(state=tk.NORMAL)
        self._log.insert(tk.END, msg + '\n')
        self._log.see(tk.END)
        self._log.configure(state=tk.DISABLED)

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select Nastran BDF/DAT File",
            filetypes=[("BDF files", "*.bdf *.dat *.nas"), ("All", "*.*")])
        if path:
            self._path_var.set(path)

    def _browse_output(self):
        d = filedialog.askdirectory(title="Select Output Directory")
        if d:
            self._outdir_var.set(d)

    # ── Scan ─────────────────────────────────────────────────────────────────

    def _scan(self):
        """Parse the input BDF and populate the entity tables."""
        path = self._path_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Error", "Please select a valid BDF file.")
            return

        self._status_var.set(f"Scanning {os.path.basename(path)}\u2026")
        self.update_idletasks()

        try:
            parser = IncludeFileParser()
            parser.parse(path)
        except Exception as exc:
            messagebox.showerror("Scan Error", str(exc))
            self._status_var.set("Scan failed")
            return

        self._bdf_path = path
        self._parser = parser
        self._summary = parser.get_summary()

        self._log_msg(f"Scanned {os.path.basename(path)}: "
                      f"{len(parser.all_files) - 1} include files found")

        totals = defaultdict(int)
        for fp, etypes in self._summary.items():
            for etype, (count, _, _) in etypes.items():
                totals[etype] += count
        parts = [f"{totals[et]} {ENTITY_LABELS[et].lower()}s"
                 for et in ENTITY_TYPES if totals.get(et, 0) > 0]
        self._log_msg(f"Total: {', '.join(parts)}")

        self._populate_simple_sheet()

        self._validate_btn.configure(state=tk.NORMAL)
        self._apply_btn.configure(state=tk.NORMAL)
        self._save_cfg_btn.configure(state=tk.NORMAL)
        self._status_var.set("Scan complete -- fill in new ranges")

    def _populate_simple_sheet(self):
        """Build simple sheet data from scan results."""
        self._simple_row_map = []
        rows = []

        for filepath in self._parser.all_files:
            fname = os.path.basename(filepath)
            parent = os.path.basename(os.path.dirname(filepath))
            display = f"{parent}/{fname}" if parent else fname

            etypes = self._summary.get(filepath, {})
            if not etypes:
                continue

            type_labels = [ENTITY_LABELS.get(et, et)
                           for et in ENTITY_TYPES if et in etypes]
            total_count = sum(c for c, _, _ in etypes.values())
            all_mins = [mn for _, mn, _ in etypes.values()]
            all_maxs = [mx for _, _, mx in etypes.values()]
            overall_min = min(all_mins) if all_mins else 1
            overall_max = max(all_maxs) if all_maxs else 1

            etypes_present = [et for et in ENTITY_TYPES if et in etypes]

            rows.append([
                display,
                ", ".join(type_labels),
                str(total_count),
                str(overall_min),
                str(overall_max),
            ])
            self._simple_row_map.append((filepath, etypes_present))

        self._simple_sheet.set_sheet_data(rows)
        self._simple_sheet.set_all_cell_sizes_to_text()

    # ── Get ranges from current sheet ────────────────────────────────────────

    def _get_ranges(self):
        """Read ranges from the sheet.

        Returns dict[filepath, dict[etype, (start, end)]] or None on error.
        """
        return self._get_ranges_simple()

    def _get_ranges_simple(self):
        """Read ranges from simple sheet, auto-allocate to per-entity."""
        data = self._simple_sheet.get_sheet_data()
        file_ranges = {}
        for i, (filepath, etypes) in enumerate(self._simple_row_map):
            if i >= len(data):
                break
            row = data[i]
            try:
                start = int(str(row[3]).strip())
                end = int(str(row[4]).strip())
            except (ValueError, TypeError):
                messagebox.showerror(
                    "Invalid Range",
                    f"Invalid number in range for "
                    f"{os.path.basename(filepath)}")
                return None
            file_ranges[filepath] = (start, end)

        allocated = self._auto_allocate(file_ranges)

        # Flatten to standard ranges dict
        ranges = {}
        for filepath, alloc in allocated.items():
            ranges[filepath] = alloc
        return ranges

    # ── Validate / Apply ─────────────────────────────────────────────────────

    def _validate(self):
        """Run pre-apply validation and show results in the log."""
        ranges = self._get_ranges()
        if ranges is None:
            return

        include_sets = self._include_set_ids.get()
        errors = Validator.validate_ranges(
            self._parser.file_ids, ranges, include_sets)

        if errors:
            self._log_msg("--- Validation FAILED ---")
            for err in errors:
                self._log_msg(f"  ERROR: {err}")
            self._status_var.set(f"Validation failed: {len(errors)} error(s)")
        else:
            self._log_msg("--- Validation PASSED ---")
            self._status_var.set("Validation passed")

    def _apply(self):
        """Run the full renumbering pipeline."""
        ranges = self._get_ranges()
        if ranges is None:
            return

        outdir = self._outdir_var.get().strip()
        if not outdir:
            messagebox.showerror("Error", "Please select an output directory.")
            return

        include_sets = self._include_set_ids.get()

        errors = Validator.validate_ranges(
            self._parser.file_ids, ranges, include_sets)
        if errors:
            self._log_msg("--- Pre-apply validation FAILED ---")
            for err in errors:
                self._log_msg(f"  ERROR: {err}")
            self._status_var.set("Cannot apply: fix validation errors first")
            return

        self._status_var.set("Applying renumbering\u2026")
        self.update_idletasks()

        try:
            # 1. Build maps
            self._log_msg("Building ID maps\u2026")
            builder = MappingBuilder(self._parser.file_ids, ranges)
            maps = builder.build()

            total_mapped = sum(len(m) for m in maps.values())
            self._log_msg(f"  {total_mapped} IDs to renumber")

            # 2. Read model with pyNastran
            self._log_msg("Reading model with pyNastran\u2026")
            model = make_model(_CARDS_TO_SKIP)
            model.read_bdf(self._bdf_path)

            # 3. Renumber cards
            self._log_msg("Renumbering cards\u2026")
            renumberer = CardRenumberer(model, maps, include_sets)
            renumberer.apply()

            # 4. Prepare case control renumberer
            case_renumberer = CaseControlRenumberer(maps, include_sets)

            # 5. Uncross-reference before writing
            try:
                model.uncross_reference()
            except Exception:
                pass

            # 6. Write output
            self._log_msg(f"Writing to {outdir}\u2026")
            writer = OutputWriter(
                model, self._parser, maps, case_renumberer, include_sets,
                log_func=self._log_msg)
            written = writer.write(outdir)

            self._log_msg(f"  Wrote {len(written)} file(s)")
            for f in written:
                self._log_msg(f"    {f}")

            # 7. Post-validation
            self._log_msg("Running post-validation\u2026")
            main_out = written[0]
            warnings, post_errors = Validator.post_validate(
                self._bdf_path, main_out)
            for w in warnings:
                self._log_msg(f"  WARNING: {w}")
            for e in post_errors:
                self._log_msg(f"  POST-ERROR: {e}")

            if not post_errors:
                self._log_msg("--- Renumbering complete (no errors) ---")
                self._status_var.set("Done -- renumbering applied successfully")
            else:
                self._log_msg(
                    f"--- Renumbering complete with "
                    f"{len(post_errors)} post-validation error(s) ---")
                self._status_var.set(
                    f"Done with {len(post_errors)} post-validation error(s)")

        except Exception as exc:
            self._log_msg(f"FATAL: {exc}")
            self._status_var.set("Apply failed")
            messagebox.showerror("Error", str(exc))

    # ── Config save/load ─────────────────────────────────────────────────────

    def _save_config(self):
        """Save current ranges to a JSON file."""
        ranges = self._get_ranges()
        if ranges is None:
            return

        path = filedialog.asksaveasfilename(
            title="Save Configuration",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All", "*.*")])
        if not path:
            return

        config = {
            'input_bdf': self._bdf_path,
            'include_set_ids': self._include_set_ids.get(),
            'output_dir': self._outdir_var.get(),
        }

        simple_data = self._simple_sheet.get_sheet_data()
        simple_ranges = {}
        for i, (filepath, etypes) in enumerate(self._simple_row_map):
            if i >= len(simple_data):
                break
            row = simple_data[i]
            try:
                s = int(str(row[3]).strip())
                e = int(str(row[4]).strip())
                simple_ranges[os.path.basename(filepath)] = [s, e]
            except (ValueError, TypeError):
                pass
        config['simple_ranges'] = simple_ranges

        with open(path, 'w') as f:
            json.dump(config, f, indent=2)

        self._log_msg(f"Config saved to {path}")

    def _load_config(self):
        """Load ranges from a JSON file."""
        path = filedialog.askopenfilename(
            title="Load Configuration",
            filetypes=[("JSON files", "*.json"), ("All", "*.*")])
        if not path:
            return

        with open(path, 'r') as f:
            config = json.load(f)

        if 'include_set_ids' in config:
            self._include_set_ids.set(config['include_set_ids'])
        if 'output_dir' in config:
            self._outdir_var.set(config['output_dir'])

        # Apply simple ranges
        simple_cfg = config.get('simple_ranges', {})
        applied = 0
        if simple_cfg:
            simple_data = self._simple_sheet.get_sheet_data()
            for i, (filepath, etypes) in enumerate(self._simple_row_map):
                if i >= len(simple_data):
                    break
                fname = os.path.basename(filepath)
                if fname in simple_cfg:
                    s, e = simple_cfg[fname]
                    simple_data[i][3] = str(s)
                    simple_data[i][4] = str(e)
                    applied += 1
            self._simple_sheet.set_sheet_data(simple_data)

        self._log_msg(f"Config loaded from {path} ({applied} ranges applied)")


def main():
    import logging
    logging.getLogger("customtkinter").setLevel(logging.ERROR)
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    root.title("Nastran Include File Renumbering Tool")
    root.geometry("1200x780")
    app = RenumberIncludesTool(root)
    app.pack(fill=tk.BOTH, expand=True)
    root.mainloop()


if __name__ == '__main__':
    main()
