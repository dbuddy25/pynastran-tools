#!/usr/bin/env python3
"""BDF Partitioner — core algorithm and pyvista visualization.

Partitions a monolithic Nastran BDF into component-level include files.
Components are connected by RBE2-CBUSH-RBE2 interfaces and/or glue contact
(BCTABLE, BCPROP, BCPROPS, BSURF). Detects part boundaries via flood-fill.

No GUI dependencies — pure algorithm + optional pyvista preview.
"""
import os
import re
from collections import defaultdict, deque
from dataclasses import dataclass, field

from pyNastran.bdf.bdf import BDF

try:
    from bdf_utils import make_model, extract_card_info, CARD_ENTITY_MAP
except ImportError:
    from preprocessing.bdf_utils import (
        make_model, extract_card_info, CARD_ENTITY_MAP,
    )

# Contact cards unsupported by pyNastran — handle as passthrough text
_CARDS_TO_SKIP = [
    'BCPROP', 'BCPROPS', 'BCPARA', 'BOUTPUT', 'BGPARM',
]

# ── Data structures ────────────────────────────────────────────────────────


@dataclass
class RBE2Chain:
    """A boundary connection: RBE2 — CBUSH — RBE2."""
    cbush_eid: int
    cbush_nodes: tuple          # (GA, GB)
    rbe2_a_eid: int
    rbe2_a_ind_node: int        # RBE2-A independent node (= CBUSH GA or GB)
    rbe2_a_dep_nodes: list
    rbe2_b_eid: int
    rbe2_b_ind_node: int        # RBE2-B independent node
    rbe2_b_dep_nodes: list


@dataclass
class Part:
    """A connected component of the mesh."""
    part_id: int
    name: str                   # user-renameable, used in filenames
    element_ids: set = field(default_factory=set)
    node_ids: set = field(default_factory=set)
    property_ids: set = field(default_factory=set)


@dataclass
class Joint:
    """Connection between two parts via CBUSHes and/or glue contact."""
    part_a_id: int
    part_b_id: int
    chains: list = field(default_factory=list)           # list[RBE2Chain]
    contact_pairs: list = field(default_factory=list)     # list of contact info dicts
    pbush_pids: set = field(default_factory=set)          # PBUSH PIDs used by joint


@dataclass
class PartitionResult:
    """Full partitioning output."""
    parts: list                 # list[Part]
    joints: list                # list[Joint]
    warnings: list              # list[str]


# ── Partitioning algorithm ─────────────────────────────────────────────────


def partition_model(model):
    """Partition a cross-referenced BDF model into parts and joints.

    Returns a PartitionResult with parts, joints, and warnings.
    """
    warnings = []

    # Step 1: Build adjacency maps
    node_to_elems = defaultdict(set)  # nid -> set of eids
    elem_to_nodes = {}                # eid -> set of nids

    for eid, elem in model.elements.items():
        nids = _get_element_nodes(elem)
        elem_to_nodes[eid] = nids
        for nid in nids:
            node_to_elems[nid].add(eid)

    # NOTE: rigid_elements and masses are NOT added to adjacency.
    # They would bridge across CBUSH boundaries. Assigned post-flood-fill.

    # Step 2: Find RBE2-CBUSH-RBE2 boundary walls
    rbe2_by_ind_node = {}   # independent_nid -> RBE2 element
    for eid, rigid in model.rigid_elements.items():
        if rigid.type == 'RBE2':
            ind_nid = _rbe2_independent_node(rigid)
            if ind_nid is not None:
                rbe2_by_ind_node[ind_nid] = rigid

    chains = []
    wall_eids = set()   # eids that form boundary walls (CBUSH + paired RBE2s)
    wall_nodes = set()  # independent nodes of boundary RBE2s (GA, GB of CBUSH)

    for eid, elem in model.elements.items():
        if elem.type != 'CBUSH':
            continue
        ga, gb = _cbush_nodes(elem)
        if gb is None or gb == 0:
            continue  # grounded spring — not a boundary

        rbe2_a = rbe2_by_ind_node.get(ga)
        rbe2_b = rbe2_by_ind_node.get(gb)
        if rbe2_a is None or rbe2_b is None:
            continue  # not a full RBE2-CBUSH-RBE2 chain

        chain = RBE2Chain(
            cbush_eid=eid,
            cbush_nodes=(ga, gb),
            rbe2_a_eid=rbe2_a.eid,
            rbe2_a_ind_node=ga,
            rbe2_a_dep_nodes=_rbe2_dependent_nodes(rbe2_a),
            rbe2_b_eid=rbe2_b.eid,
            rbe2_b_ind_node=gb,
            rbe2_b_dep_nodes=_rbe2_dependent_nodes(rbe2_b),
        )
        chains.append(chain)
        wall_eids.add(eid)
        wall_eids.add(rbe2_a.eid)
        wall_eids.add(rbe2_b.eid)
        wall_nodes.add(ga)
        wall_nodes.add(gb)

    # Step 3: Flood-fill — BFS from each unvisited element
    visited_eids = set()
    raw_parts = []  # list of sets of eids

    all_eids = set(elem_to_nodes.keys()) - wall_eids
    for seed_eid in sorted(all_eids):
        if seed_eid in visited_eids:
            continue
        component = set()
        queue = deque([seed_eid])
        while queue:
            eid = queue.popleft()
            if eid in visited_eids or eid in wall_eids:
                continue
            visited_eids.add(eid)
            component.add(eid)
            for nid in elem_to_nodes.get(eid, set()):
                if nid in wall_nodes:
                    continue  # don't cross boundary
                for neighbor_eid in node_to_elems.get(nid, set()):
                    if neighbor_eid not in visited_eids and neighbor_eid not in wall_eids:
                        queue.append(neighbor_eid)
        if component:
            raw_parts.append(component)

    # Step 4: Build Part objects
    parts = []
    for i, eids in enumerate(raw_parts):
        # Collect property IDs first (needed for naming)
        property_ids = set()
        for eid in eids:
            pid = _get_element_pid(model, eid)
            if pid is not None:
                property_ids.add(pid)

        part = Part(
            part_id=i + 1,
            name=_derive_part_name(model, property_ids, i + 1),
            element_ids=eids,
            property_ids=property_ids,
        )
        # Collect nodes
        for eid in eids:
            part.node_ids.update(elem_to_nodes.get(eid, set()))
        # Also include RBE2 dependent nodes that belong to this part's mesh
        for chain in chains:
            if set(chain.rbe2_a_dep_nodes) & part.node_ids:
                part.node_ids.update(chain.rbe2_a_dep_nodes)
            if set(chain.rbe2_b_dep_nodes) & part.node_ids:
                part.node_ids.update(chain.rbe2_b_dep_nodes)
        parts.append(part)

    # Deduplicate part names
    name_counts = defaultdict(int)
    for part in parts:
        name_counts[part.name] += 1
    seen = defaultdict(int)
    for part in parts:
        if name_counts[part.name] > 1:
            seen[part.name] += 1
            part.name = f"{part.name}_{seen[part.name]}"

    # Build node-to-part lookup
    node_to_part = {}
    for part in parts:
        for nid in part.node_ids:
            node_to_part[nid] = part.part_id

    # Assign interior rigid elements to parts by node voting
    part_by_id = {p.part_id: p for p in parts}
    for eid, rigid in model.rigid_elements.items():
        if eid in wall_eids:
            continue  # boundary RBE2s go to joint files
        nids = _get_rigid_nodes(rigid)
        owner = _find_part_for_nodes(list(nids), node_to_part)
        if owner is not None:
            part_by_id[owner].element_ids.add(eid)

    # Assign mass elements to parts by node voting
    for eid, mass_elem in model.masses.items():
        nids = _get_mass_nodes(mass_elem)
        owner = _find_part_for_nodes(list(nids), node_to_part)
        if owner is not None:
            part_by_id[owner].element_ids.add(eid)

    # Step 5: Build joints from chains
    joint_map = {}  # (min_part_id, max_part_id) -> Joint
    for chain in chains:
        # Determine which part owns each RBE2's dependents
        part_a_id = _find_part_for_nodes(chain.rbe2_a_dep_nodes, node_to_part)
        part_b_id = _find_part_for_nodes(chain.rbe2_b_dep_nodes, node_to_part)
        if part_a_id is None or part_b_id is None:
            warnings.append(
                f"CBUSH {chain.cbush_eid}: could not assign both RBE2s to parts")
            continue
        if part_a_id == part_b_id:
            warnings.append(
                f"CBUSH {chain.cbush_eid}: both RBE2s in same Part_{part_a_id}")
            continue
        key = (min(part_a_id, part_b_id), max(part_a_id, part_b_id))
        if key not in joint_map:
            joint_map[key] = Joint(part_a_id=key[0], part_b_id=key[1])
        joint = joint_map[key]
        joint.chains.append(chain)
        # Collect PBUSH PID from the CBUSH
        cbush = model.elements.get(chain.cbush_eid)
        if cbush is not None:
            pid = _get_element_pid(model, chain.cbush_eid)
            if pid is not None:
                joint.pbush_pids.add(pid)

    # Step 6: Map contact surfaces to joints
    _assign_contact_to_joints(model, parts, joint_map, warnings)

    joints = sorted(joint_map.values(), key=lambda j: (j.part_a_id, j.part_b_id))

    # Orphan node check
    all_part_nodes = set()
    for p in parts:
        all_part_nodes.update(p.node_ids)
    for chain in chains:
        all_part_nodes.add(chain.cbush_nodes[0])
        all_part_nodes.add(chain.cbush_nodes[1])
    model_nodes = set(model.nodes.keys())
    orphans = model_nodes - all_part_nodes
    if orphans:
        warnings.append(f"{len(orphans)} orphan node(s) not assigned to any part")

    return PartitionResult(parts=parts, joints=joints, warnings=warnings)


# ── Helper functions ───────────────────────────────────────────────────────


def _get_element_nodes(elem):
    """Get node IDs from a structural element."""
    nids = set()
    try:
        for n in elem.node_ids:
            if n is not None and n > 0:
                nids.add(n)
    except AttributeError:
        try:
            for n in elem.nodes:
                nid = n if isinstance(n, int) else getattr(n, 'nid', None)
                if nid is not None and nid > 0:
                    nids.add(nid)
        except (AttributeError, TypeError):
            pass
    return nids


def _get_rigid_nodes(rigid):
    """Get all node IDs from a rigid element (RBE2, RBE3, RBAR, etc.)."""
    nids = set()
    if rigid.type == 'RBE2':
        ind = _rbe2_independent_node(rigid)
        if ind is not None:
            nids.add(ind)
        nids.update(_rbe2_dependent_nodes(rigid))
    elif rigid.type == 'RBE3':
        try:
            ref = rigid.refgrid if hasattr(rigid, 'refgrid') else rigid.ref_grid_id
            if ref is not None and ref > 0:
                nids.add(ref)
        except (AttributeError, TypeError):
            pass
        try:
            for groups in rigid.Gijs:
                for n in groups:
                    nid = n if isinstance(n, int) else getattr(n, 'nid', None)
                    if nid is not None and nid > 0:
                        nids.add(nid)
        except (AttributeError, TypeError):
            pass
    elif rigid.type == 'RBAR':
        for attr in ('ga', 'gb', 'Ga', 'Gb'):
            val = getattr(rigid, attr, None)
            if val is not None:
                nid = val if isinstance(val, int) else getattr(val, 'nid', None)
                if nid is not None and nid > 0:
                    nids.add(nid)
    return nids


def _get_mass_nodes(mass_elem):
    """Get node IDs from a mass element (CONM2, CMASS, etc.)."""
    nids = set()
    if mass_elem.type == 'CONM2':
        nid = mass_elem.nid
        if isinstance(nid, int):
            if nid > 0:
                nids.add(nid)
        else:
            nid_val = getattr(nid, 'nid', None)
            if nid_val is not None and nid_val > 0:
                nids.add(nid_val)
    elif mass_elem.type == 'CONM1':
        nid = getattr(mass_elem, 'nid', None)
        if nid is not None:
            nid_val = nid if isinstance(nid, int) else getattr(nid, 'nid', None)
            if nid_val is not None and nid_val > 0:
                nids.add(nid_val)
    else:
        # CMASS1, CMASS2, etc.
        for attr in ('g1', 'g2', 'nid', 'node_ids', 'nodes'):
            val = getattr(mass_elem, attr, None)
            if val is None:
                continue
            if isinstance(val, int):
                if val > 0:
                    nids.add(val)
            elif hasattr(val, '__iter__'):
                for n in val:
                    if isinstance(n, int) and n > 0:
                        nids.add(n)
                    elif hasattr(n, 'nid') and n.nid > 0:
                        nids.add(n.nid)
    return nids


def _rbe2_independent_node(rbe2):
    """Get the independent node of an RBE2."""
    gn = getattr(rbe2, 'gn', None)
    if gn is None:
        gn = getattr(rbe2, 'independent_node', None)
    if gn is None:
        return None
    if isinstance(gn, int):
        return gn
    return getattr(gn, 'nid', None)


def _rbe2_dependent_nodes(rbe2):
    """Get the dependent node list of an RBE2."""
    gm = getattr(rbe2, 'Gmi', None)
    if gm is None:
        gm = getattr(rbe2, 'dependent_nodes', None)
    if gm is None:
        return []
    result = []
    for n in gm:
        if isinstance(n, int):
            result.append(n)
        else:
            nid = getattr(n, 'nid', None)
            if nid is not None:
                result.append(nid)
    return result


def _cbush_nodes(elem):
    """Get (GA, GB) for a CBUSH element."""
    ga = gb = None
    try:
        nodes = elem.node_ids
        ga = nodes[0] if len(nodes) > 0 else None
        gb = nodes[1] if len(nodes) > 1 else None
    except (AttributeError, TypeError):
        try:
            ga = elem.ga if isinstance(elem.ga, int) else elem.ga.nid
            gb = elem.gb if isinstance(elem.gb, int) else (elem.gb.nid if elem.gb else None)
        except AttributeError:
            pass
    return ga, gb


def _get_element_pid(model, eid):
    """Get the property ID for an element, or None."""
    elem = model.elements.get(eid)
    if elem is None:
        elem = model.masses.get(eid)
    if elem is None:
        return None
    pid = getattr(elem, 'pid', None)
    if pid is None:
        return None
    if isinstance(pid, int):
        return pid if pid > 0 else None
    return getattr(pid, 'pid', None)


def _find_part_for_nodes(node_list, node_to_part):
    """Find the part that owns the majority of the given nodes."""
    votes = defaultdict(int)
    for nid in node_list:
        pid = node_to_part.get(nid)
        if pid is not None:
            votes[pid] += 1
    if not votes:
        return None
    return max(votes, key=votes.get)


def _derive_part_name(model, property_ids, part_id):
    """Derive a human-readable name from the first property's comment."""
    for pid in sorted(property_ids):
        prop = model.properties.get(pid)
        if prop is None:
            continue
        comment = getattr(prop, 'comment', '')
        name = _parse_comment_name(comment)
        if name:
            return name
    return f"Part_{part_id:03d}"


def _parse_comment_name(comment):
    """Extract a usable name from a property comment string."""
    if not comment:
        return None
    text = comment.strip().lstrip('$').strip()
    if not text:
        return None
    # Femap format: "Femap Property 10 : Wing_Skin PSHELL"
    if ':' in text:
        after_colon = text.split(':', 1)[1].strip()
        tokens = after_colon.split()
        if tokens:
            return tokens[0]
    # Generic: strip card types and IDs, take remaining text
    for card_type in ('PSHELL', 'PCOMP', 'PCOMPG', 'PSOLID', 'PBAR',
                       'PBARL', 'PBEAM', 'PROD', 'PBUSH'):
        text = text.replace(card_type, '').strip()
    text = re.sub(r'PID\s*=\s*\d+', '', text, flags=re.IGNORECASE).strip()
    text = re.sub(r'^\d+\s*', '', text).strip()
    if text:
        return text[:30].strip().rstrip('-_')
    return None


def _assign_contact_to_joints(model, parts, joint_map, warnings):
    """Map BSURF/BSURFS surfaces to part pairs and assign to joints."""
    # Build eid-to-part lookup
    eid_to_part = {}
    for part in parts:
        for eid in part.element_ids:
            eid_to_part[eid] = part.part_id

    # Build pid-to-part lookup
    pid_to_parts = defaultdict(set)
    for part in parts:
        for pid in part.property_ids:
            pid_to_parts[pid].add(part.part_id)

    # Map BSURF contact_id -> set of part_ids (via element ownership)
    contact_to_parts = defaultdict(set)

    # BSURF references elements
    for cid, bsurf in getattr(model, 'bsurfs', {}).items():
        eids = set()
        if hasattr(bsurf, 'eids') and bsurf.eids:
            eids.update(bsurf.eids)
        for eid in eids:
            part_id = eid_to_part.get(eid)
            if part_id is not None:
                contact_to_parts[cid].add(part_id)

    # BCTSET pairs surfaces — map to joints
    for cid, bctset in getattr(model, 'bctsets', {}).items():
        if not hasattr(bctset, 'contact_ids'):
            continue
        # Each BCTSET row pairs two surfaces
        try:
            sids = list(bctset.contact_ids)
        except (TypeError, AttributeError):
            continue
        for i in range(0, len(sids) - 1, 2):
            sid_a, sid_b = sids[i], sids[i + 1]
            parts_a = contact_to_parts.get(sid_a, set())
            parts_b = contact_to_parts.get(sid_b, set())
            for pa in parts_a:
                for pb in parts_b:
                    if pa != pb:
                        key = (min(pa, pb), max(pa, pb))
                        if key not in joint_map:
                            joint_map[key] = Joint(part_a_id=key[0], part_b_id=key[1])
                        joint_map[key].contact_pairs.append({
                            'bctset_id': cid,
                            'surf_a': sid_a,
                            'surf_b': sid_b,
                        })


# ── Merge parts ────────────────────────────────────────────────────────────


def merge_parts(result, part_ids_to_merge):
    """Merge multiple parts into one, absorbing joints between them.

    Args:
        result: PartitionResult to modify in-place
        part_ids_to_merge: list/set of part_id values to merge

    Returns:
        Updated PartitionResult
    """
    merge_set = set(part_ids_to_merge)
    if len(merge_set) < 2:
        return result

    # Find the parts to merge
    merging = [p for p in result.parts if p.part_id in merge_set]
    if len(merging) < 2:
        return result

    # Create merged part (keep lowest ID)
    base = min(merging, key=lambda p: p.part_id)
    merged = Part(
        part_id=base.part_id,
        name=base.name,
        element_ids=set(),
        node_ids=set(),
        property_ids=set(),
    )
    for p in merging:
        merged.element_ids.update(p.element_ids)
        merged.node_ids.update(p.node_ids)
        merged.property_ids.update(p.property_ids)

    # Absorb joints between merged parts — their elements become interior
    absorbed_joints = []
    remaining_joints = []
    for joint in result.joints:
        if joint.part_a_id in merge_set and joint.part_b_id in merge_set:
            absorbed_joints.append(joint)
        else:
            remaining_joints.append(joint)

    # Move absorbed CBUSH/RBE2 elements into the merged part
    for joint in absorbed_joints:
        for chain in joint.chains:
            merged.element_ids.add(chain.cbush_eid)
            merged.element_ids.add(chain.rbe2_a_eid)
            merged.element_ids.add(chain.rbe2_b_eid)
            merged.node_ids.add(chain.cbush_nodes[0])
            merged.node_ids.add(chain.cbush_nodes[1])
            merged.node_ids.update(chain.rbe2_a_dep_nodes)
            merged.node_ids.update(chain.rbe2_b_dep_nodes)
        merged.property_ids.update(joint.pbush_pids)

    # Re-key remaining joints that reference merged parts
    for joint in remaining_joints:
        if joint.part_a_id in merge_set:
            joint.part_a_id = merged.part_id
        if joint.part_b_id in merge_set:
            joint.part_b_id = merged.part_id
        # Ensure part_a < part_b
        if joint.part_a_id > joint.part_b_id:
            joint.part_a_id, joint.part_b_id = joint.part_b_id, joint.part_a_id

    # Rebuild parts list
    other_parts = [p for p in result.parts if p.part_id not in merge_set]
    result.parts = sorted(other_parts + [merged], key=lambda p: p.part_id)
    result.joints = sorted(remaining_joints, key=lambda j: (j.part_a_id, j.part_b_id))
    return result


# ── Output writer ──────────────────────────────────────────────────────────


_EXEC_CARD_NAMES = frozenset({
    'PARAM', 'EIGRL', 'EIGR', 'GRAV', 'LOAD', 'DLOAD',
    'TEMP', 'TEMPD', 'RFORCE',
})

_MATERIAL_CARD_NAMES = frozenset({
    'MAT1', 'MAT2', 'MAT8', 'MAT9', 'MAT10',
})

_PROPERTY_CARD_NAMES = frozenset({
    'PSHELL', 'PCOMP', 'PCOMPG', 'PSOLID', 'PBAR', 'PBARL',
    'PBEAM', 'PBEAML', 'PROD', 'PELAS', 'PDAMP', 'PGAP',
    'PWELD', 'PFAST', 'PVISC', 'PSHEAR', 'PLSOLID', 'PCOMPLS',
})

_COORD_CARD_NAMES = frozenset({
    'CORD2R', 'CORD2C', 'CORD2S', 'CORD1R', 'CORD1C', 'CORD1S',
})

_CONTACT_GLOBAL_CARDS = frozenset({
    'BCTPARA', 'BCTPARM',
})


def write_partition(model, result, output_dir, bdf_path, log_fn=None):
    """Write partitioned include files.

    Args:
        model: cross-referenced BDF model
        result: PartitionResult
        output_dir: directory to write output files
        bdf_path: path to original BDF (for extracting exec/case control)
        log_fn: optional callable(str) for progress messages

    Returns:
        dict with validation info: {'total_elems': int, 'total_nodes': int,
                                     'written_elems': int, 'written_nodes': int}
    """
    def log(msg):
        if log_fn:
            log_fn(msg)

    os.makedirs(output_dir, exist_ok=True)

    # Build lookup maps
    eid_to_part = {}
    nid_to_part = {}
    for part in result.parts:
        for eid in part.element_ids:
            eid_to_part[eid] = part.part_id
        for nid in part.node_ids:
            nid_to_part[nid] = part.part_id

    wall_eids = set()
    wall_nids = set()
    for joint in result.joints:
        for chain in joint.chains:
            wall_eids.add(chain.cbush_eid)
            wall_eids.add(chain.rbe2_a_eid)
            wall_eids.add(chain.rbe2_b_eid)
            wall_nids.add(chain.cbush_nodes[0])
            wall_nids.add(chain.cbush_nodes[1])

    joint_pbush_pids = set()
    for joint in result.joints:
        joint_pbush_pids.update(joint.pbush_pids)

    # Part name lookup
    part_names = {p.part_id: p.name for p in result.parts}

    # ── Write part files ──
    written_nodes = set()
    written_elems = set()

    for part in result.parts:
        fname = _safe_filename(part.name) + '.bdf'
        fpath = os.path.join(output_dir, fname)
        log(f"Writing {fname}")

        lines = []
        lines.append(f'$ Part: {part.name} (ID={part.part_id})\n')
        lines.append(f'$ Elements: {len(part.element_ids)}, '
                     f'Nodes: {len(part.node_ids)}\n')
        lines.append('$\n')

        # GRIDs
        lines.append('$ --- Nodes ---\n')
        for nid in sorted(part.node_ids):
            node = model.nodes.get(nid)
            if node is not None:
                lines.append(_write_card(node))
                written_nodes.add(nid)

        # Structural elements
        lines.append('$ --- Elements ---\n')
        for eid in sorted(part.element_ids):
            elem = model.elements.get(eid)
            if elem is not None:
                lines.append(_write_card(elem))
                written_elems.add(eid)

        # Interior rigid elements (RBE2/RBE3/RBAR not in wall)
        rigids_written = False
        for eid in sorted(part.element_ids):
            rigid = model.rigid_elements.get(eid)
            if rigid is not None:
                if not rigids_written:
                    lines.append('$ --- Rigid Elements ---\n')
                    rigids_written = True
                lines.append(_write_card(rigid))
                written_elems.add(eid)

        # Mass elements (CONM2, CMASS, etc.)
        masses_written = False
        for eid, mass_elem in sorted(model.masses.items()):
            mnids = _get_mass_nodes(mass_elem)
            if mnids and mnids.issubset(part.node_ids):
                if not masses_written:
                    lines.append('$ --- Mass Elements ---\n')
                    masses_written = True
                lines.append(_write_card(mass_elem))
                written_elems.add(eid)

        # SPCs — if all nodes in this part
        spcs_written = False
        for spc_id, spc_list in model.spcs.items():
            for spc in spc_list:
                spc_nids = _get_spc_nodes(spc)
                if spc_nids and spc_nids.issubset(part.node_ids):
                    if not spcs_written:
                        lines.append('$ --- SPCs ---\n')
                        spcs_written = True
                    lines.append(_write_card(spc))

        # Loads — if all nodes/elems in this part
        loads_written = False
        for load_id, load_list in model.loads.items():
            for load in load_list:
                if _load_belongs_to_part(load, part, eid_to_part):
                    if not loads_written:
                        lines.append('$ --- Loads ---\n')
                        loads_written = True
                    lines.append(_write_card(load))

        with open(fpath, 'w') as f:
            f.writelines(lines)

    # ── Write joint files ──
    for joint in result.joints:
        name_a = part_names.get(joint.part_a_id, f'Part_{joint.part_a_id}')
        name_b = part_names.get(joint.part_b_id, f'Part_{joint.part_b_id}')
        fname = _safe_filename(f'{name_a}-to-{name_b}') + '.bdf'
        fpath = os.path.join(output_dir, fname)
        log(f"Writing {fname}")

        lines = []
        lines.append(f'$ Joint: {name_a} <-> {name_b}\n')
        lines.append(f'$ Chains: {len(joint.chains)}, '
                     f'Contact pairs: {len(joint.contact_pairs)}\n')
        lines.append('$\n')

        # CBUSH elements
        if joint.chains:
            lines.append('$ --- CBUSH elements ---\n')
            for chain in sorted(joint.chains, key=lambda c: c.cbush_eid):
                elem = model.elements.get(chain.cbush_eid)
                if elem is not None:
                    lines.append(_write_card(elem))
                    written_elems.add(chain.cbush_eid)

        # RBE2 elements
        if joint.chains:
            lines.append('$ --- RBE2 elements ---\n')
            rbe2_eids = set()
            for chain in joint.chains:
                rbe2_eids.add(chain.rbe2_a_eid)
                rbe2_eids.add(chain.rbe2_b_eid)
            for eid in sorted(rbe2_eids):
                rigid = model.rigid_elements.get(eid)
                if rigid is not None:
                    lines.append(_write_card(rigid))
                    written_elems.add(eid)

        # PBUSH properties
        if joint.pbush_pids:
            lines.append('$ --- PBUSH properties ---\n')
            for pid in sorted(joint.pbush_pids):
                prop = model.properties.get(pid)
                if prop is not None:
                    lines.append(_write_card(prop))

        with open(fpath, 'w') as f:
            f.writelines(lines)

    # ── Write shared.bdf ──
    shared_path = os.path.join(output_dir, 'shared.bdf')
    log("Writing shared.bdf")
    lines = []
    lines.append('$ Shared: materials, properties, coordinate systems\n')
    lines.append('$\n')

    # Materials
    lines.append('$ --- Materials ---\n')
    for mid in sorted(model.materials.keys()):
        lines.append(_write_card(model.materials[mid]))

    # Properties (excluding PBUSH in joints)
    lines.append('$ --- Properties ---\n')
    for pid in sorted(model.properties.keys()):
        if pid in joint_pbush_pids:
            continue
        lines.append(_write_card(model.properties[pid]))

    # Coordinate systems
    if model.coords:
        lines.append('$ --- Coordinate Systems ---\n')
        for cid in sorted(model.coords.keys()):
            if cid == 0:
                continue  # basic CS
            lines.append(_write_card(model.coords[cid]))

    # Global contact parameters
    for attr_name in ('bctparas', 'bctparms'):
        container = getattr(model, attr_name, {})
        if container:
            lines.append(f'$ --- {attr_name.upper()} ---\n')
            for cid in sorted(container.keys()):
                lines.append(_write_card(container[cid]))

    # SPCs not fully in one part
    for spc_id, spc_list in model.spcs.items():
        for spc in spc_list:
            spc_nids = _get_spc_nodes(spc)
            if not spc_nids:
                lines.append(_write_card(spc))
                continue
            # Check if already written to a part
            assigned = False
            for part in result.parts:
                if spc_nids.issubset(part.node_ids):
                    assigned = True
                    break
            if not assigned:
                lines.append(_write_card(spc))

    with open(shared_path, 'w') as f:
        f.writelines(lines)

    # ── Write master.bdf ──
    master_path = os.path.join(output_dir, 'master.bdf')
    log("Writing master.bdf")
    exec_lines, case_lines = _extract_exec_case_control(bdf_path)

    lines = []
    lines.extend(exec_lines)
    lines.extend(case_lines)

    # BEGIN BULK
    lines.append('BEGIN BULK\n')

    # INCLUDEs
    lines.append("INCLUDE 'shared.bdf'\n")
    for part in result.parts:
        fname = _safe_filename(part.name) + '.bdf'
        lines.append(f"INCLUDE '{fname}'\n")
    for joint in result.joints:
        name_a = part_names.get(joint.part_a_id, f'Part_{joint.part_a_id}')
        name_b = part_names.get(joint.part_b_id, f'Part_{joint.part_b_id}')
        fname = _safe_filename(f'{name_a}-to-{name_b}') + '.bdf'
        lines.append(f"INCLUDE '{fname}'\n")

    # Exec-level cards (PARAM, EIGRL, etc.)
    exec_cards_written = False
    for attr_name, container in [
        ('params', model.params),
    ]:
        for key, card in container.items():
            if not exec_cards_written:
                lines.append('$ --- Parameters ---\n')
                exec_cards_written = True
            lines.append(_write_card(card))

    if hasattr(model, 'methods') and model.methods:
        lines.append('$ --- Methods ---\n')
        for mid, method in sorted(model.methods.items()):
            lines.append(_write_card(method))

    lines.append('ENDDATA\n')

    with open(master_path, 'w') as f:
        f.writelines(lines)

    # ── Passthrough contact cards (BCPROP, BCPROPS) ──
    _write_passthrough_contact(bdf_path, output_dir, result, eid_to_part,
                               part_names, log)

    # ── Validation ──
    total_elems = len(model.elements) + len(model.rigid_elements) + len(model.masses)
    total_nodes = len(model.nodes)

    return {
        'total_elems': total_elems,
        'total_nodes': total_nodes,
        'written_elems': len(written_elems),
        'written_nodes': len(written_nodes),
    }


def _write_card(card):
    """Write a card using write_card(size=8), stripping leading comments."""
    try:
        text = card.write_card(size=8)
    except Exception:
        try:
            text = str(card)
        except Exception:
            return ''
    # Strip pyNastran's auto-generated leading comment
    text_lines = text.split('\n')
    while text_lines and text_lines[0].strip().startswith('$'):
        text_lines.pop(0)
    text = '\n'.join(text_lines)
    if text and not text.endswith('\n'):
        text += '\n'
    return text


def _safe_filename(name):
    """Convert a part/joint name to a filesystem-safe filename."""
    return re.sub(r'[^\w\-.]', '_', name.lower().strip())


def _get_spc_nodes(spc):
    """Get node IDs referenced by an SPC/SPC1 card."""
    nids = set()
    if hasattr(spc, 'node_ids'):
        for n in spc.node_ids:
            if isinstance(n, int) and n > 0:
                nids.add(n)
            elif hasattr(n, 'nid') and n.nid > 0:
                nids.add(n.nid)
    elif hasattr(spc, 'nodes'):
        for n in spc.nodes:
            if isinstance(n, int) and n > 0:
                nids.add(n)
            elif hasattr(n, 'nid') and n.nid > 0:
                nids.add(n.nid)
    return nids


def _load_belongs_to_part(load, part, eid_to_part):
    """Check if a load card's referenced nodes/elements are all in one part."""
    # FORCE, MOMENT reference a single node
    if hasattr(load, 'node_id'):
        nid = load.node_id
        if isinstance(nid, int):
            return nid in part.node_ids
        nid = getattr(nid, 'nid', None)
        return nid is not None and nid in part.node_ids
    if hasattr(load, 'node'):
        nid = load.node
        if isinstance(nid, int):
            return nid in part.node_ids
        nid = getattr(nid, 'nid', None)
        return nid is not None and nid in part.node_ids
    # PLOAD4 references an element
    if hasattr(load, 'eid'):
        eid = load.eid
        if isinstance(eid, int):
            return eid_to_part.get(eid) == part.part_id
        eid = getattr(eid, 'eid', None)
        return eid is not None and eid_to_part.get(eid) == part.part_id
    return False


def _extract_exec_case_control(bdf_path):
    """Extract executive and case control sections from the main BDF."""
    exec_lines = []
    case_lines = []
    try:
        with open(bdf_path, 'r', errors='replace') as f:
            lines = f.readlines()
    except OSError:
        return exec_lines, case_lines

    in_exec = True
    in_case = False
    for line in lines:
        upper = line.strip().upper()
        if in_exec:
            exec_lines.append(line)
            if upper.startswith('CEND'):
                in_exec = False
                in_case = True
                continue
        elif in_case:
            if upper.startswith('BEGIN') and 'BULK' in upper:
                break
            case_lines.append(line)

    return exec_lines, case_lines


def _write_passthrough_contact(bdf_path, output_dir, result, eid_to_part,
                               part_names, log):
    """Scan raw BDF for BCPROP/BCPROPS cards and append to joint files.

    These cards reference PIDs. We map PIDs to parts via property ownership,
    and write them to the appropriate joint file.
    """
    # Build PID -> part_id lookup
    pid_to_part = {}
    for part in result.parts:
        for pid in part.property_ids:
            pid_to_part[pid] = part.part_id

    # Scan raw BDF for passthrough cards
    passthrough_cards = _collect_passthrough_cards(bdf_path)
    if not passthrough_cards:
        return

    # For each passthrough card, extract PIDs and assign to a joint
    for card_name, card_lines in passthrough_cards:
        pids = _extract_pids_from_passthrough(card_lines, card_name)
        part_ids = set()
        for pid in pids:
            pid_part = pid_to_part.get(pid)
            if pid_part is not None:
                part_ids.add(pid_part)

        if len(part_ids) < 2:
            # Can't determine joint — append to shared.bdf
            shared_path = os.path.join(output_dir, 'shared.bdf')
            with open(shared_path, 'a') as f:
                f.writelines(card_lines)
            continue

        # Pick the first pair
        sorted_ids = sorted(part_ids)
        pa, pb = sorted_ids[0], sorted_ids[1]
        for joint in result.joints:
            if joint.part_a_id == pa and joint.part_b_id == pb:
                name_a = part_names.get(pa, f'Part_{pa}')
                name_b = part_names.get(pb, f'Part_{pb}')
                fname = _safe_filename(f'{name_a}-to-{name_b}') + '.bdf'
                fpath = os.path.join(output_dir, fname)
                with open(fpath, 'a') as f:
                    f.writelines(card_lines)
                break


def _collect_passthrough_cards(bdf_path):
    """Collect BCPROP/BCPROPS card blocks from raw BDF text."""
    cards = []
    try:
        with open(bdf_path, 'r', errors='replace') as f:
            lines = f.readlines()
    except OSError:
        return cards

    in_bulk = False
    current_card = None
    current_lines = []

    for line in lines:
        upper = line.strip().upper()
        if not in_bulk:
            if upper.startswith('BEGIN') and 'BULK' in upper:
                in_bulk = True
            continue
        if upper.startswith('ENDDATA'):
            break

        stripped = line.strip()
        if not stripped or stripped.startswith('$'):
            continue

        first_char = stripped[0]
        if first_char.isalpha():
            # Flush previous
            if current_card and current_lines:
                cards.append((current_card, current_lines))
            card_name = stripped[:8].strip().upper().rstrip('*')
            if ',' in stripped:
                card_name = stripped.split(',')[0].strip().upper()
            if card_name in ('BCPROP', 'BCPROPS'):
                current_card = card_name
                current_lines = [line]
            else:
                current_card = None
                current_lines = []
        else:
            # Continuation
            if current_card:
                current_lines.append(line)

    if current_card and current_lines:
        cards.append((current_card, current_lines))

    return cards


def _extract_pids_from_passthrough(card_lines, card_name):
    """Extract PID values from BCPROP/BCPROPS raw lines."""
    pids = set()
    for i, line in enumerate(card_lines):
        if i == 0:
            # First line: PIDs start at field 2
            start_field = 2
        else:
            # Continuation: PIDs start at field 1
            start_field = 1

        stripped = line.rstrip('\n')
        if ',' in stripped:
            fields = stripped.split(',')
            for f in fields[start_field:]:
                try:
                    pids.add(int(f.strip()))
                except (ValueError, TypeError):
                    pass
        else:
            # Fixed format 8-char fields
            for fi in range(start_field, 9):
                col_start = fi * 8
                col_end = col_start + 8
                if col_start >= len(stripped):
                    break
                field_str = stripped[col_start:col_end].strip()
                if field_str:
                    try:
                        pids.add(int(field_str))
                    except (ValueError, TypeError):
                        pass
    return pids


# ── pyvista visualization ──────────────────────────────────────────────────


def build_pyvista_mesh(model, parts):
    """Build a pyvista UnstructuredGrid colored by part_id.

    Returns (mesh, available) where available=False if pyvista not installed.
    """
    try:
        import pyvista as pv
        import numpy as np
    except ImportError:
        return None, False

    # VTK cell type constants
    VTK_TRIANGLE = 5
    VTK_QUAD = 9
    VTK_TETRA = 10
    VTK_HEXAHEDRON = 12
    VTK_WEDGE = 13
    VTK_LINE = 3

    _ELEM_TYPE_MAP = {
        'CTRIA3': VTK_TRIANGLE,
        'CTRIA6': VTK_TRIANGLE,
        'CTRIAR': VTK_TRIANGLE,
        'CQUAD4': VTK_QUAD,
        'CQUAD8': VTK_QUAD,
        'CQUADR': VTK_QUAD,
        'CTETRA': VTK_TETRA,
        'CHEXA': VTK_HEXAHEDRON,
        'CPENTA': VTK_WEDGE,
        'CBAR': VTK_LINE,
        'CBEAM': VTK_LINE,
        'CROD': VTK_LINE,
        'CONROD': VTK_LINE,
        'CBUSH': VTK_LINE,
    }

    _EXPECTED_NODES = {
        VTK_TRIANGLE: 3,
        VTK_QUAD: 4,
        VTK_TETRA: 4,
        VTK_HEXAHEDRON: 8,
        VTK_WEDGE: 6,
        VTK_LINE: 2,
    }

    # Build eid -> part_id map
    eid_to_part = {}
    for part in parts:
        for eid in part.element_ids:
            eid_to_part[eid] = part.part_id

    # Collect grid points
    nid_to_idx = {}
    points = []
    for nid in sorted(model.nodes.keys()):
        node = model.nodes[nid]
        try:
            xyz = node.get_position()
        except Exception:
            xyz = getattr(node, 'xyz', [0., 0., 0.])
        nid_to_idx[nid] = len(points)
        points.append(xyz)

    if not points:
        return None, True

    points_arr = np.array(points, dtype=np.float64)

    cells = []
    cell_types = []
    cell_part_ids = []

    for eid, elem in model.elements.items():
        vtk_type = _ELEM_TYPE_MAP.get(elem.type)
        if vtk_type is None:
            continue

        nids = []
        try:
            for n in elem.node_ids:
                if n is not None and n > 0:
                    nids.append(n)
        except AttributeError:
            try:
                for n in elem.nodes:
                    nid = n if isinstance(n, int) else getattr(n, 'nid', None)
                    if nid is not None and nid > 0:
                        nids.append(nid)
            except (AttributeError, TypeError):
                continue

        expected = _EXPECTED_NODES.get(vtk_type, len(nids))
        nids = nids[:expected]
        if len(nids) < expected:
            continue

        indices = []
        valid = True
        for nid in nids:
            idx = nid_to_idx.get(nid)
            if idx is None:
                valid = False
                break
            indices.append(idx)
        if not valid:
            continue

        cells.append([len(indices)] + indices)
        cell_types.append(vtk_type)
        cell_part_ids.append(eid_to_part.get(eid, 0))

    if not cells:
        return None, True

    cells_flat = []
    for c in cells:
        cells_flat.extend(c)

    mesh = pv.UnstructuredGrid(
        np.array(cells_flat, dtype=np.int64),
        np.array(cell_types, dtype=np.uint8),
        points_arr,
    )
    mesh.cell_data['part_id'] = np.array(cell_part_ids, dtype=np.int32)

    return mesh, True


def show_partition_preview(mesh, parts):
    """Show pyvista preview — single mesh colored by part_id scalar."""
    try:
        import pyvista as pv
    except ImportError:
        return
    if mesh is None or mesh.n_cells == 0:
        return

    plotter = pv.Plotter(title="BDF Partition Preview")

    n_parts = len(parts)
    plotter.add_mesh(
        mesh,
        scalars='part_id',
        cmap='tab20' if n_parts <= 20 else 'turbo',
        show_edges=False,
        show_scalar_bar=True,
        scalar_bar_args={'title': 'Part ID'},
    )

    plotter.add_text(
        f"{n_parts} parts, {mesh.n_cells} elements",
        position='upper_left', font_size=10,
    )
    plotter.show()
