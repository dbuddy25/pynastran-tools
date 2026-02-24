"""Shared utilities for Nastran BDF tools.

Provides:
- IncludeFileParser: parse raw BDF text to discover includes and map card IDs
  to their source files (works independently of pyNastran's ifile tracking)
- CARD_ENTITY_MAP, ENTITY_TYPES, ENTITY_LABELS: card-to-entity-type mappings
- make_model(): create a BDF model with optional card disabling
"""
import os
import re
from collections import defaultdict

from pyNastran.bdf.bdf import BDF


# ── Card name → entity type mapping ──────────────────────────────────────────

CARD_ENTITY_MAP = {
    # Nodes
    'GRID': 'nid', 'SPOINT': 'nid',
    # Elements
    'CQUAD4': 'eid', 'CTRIA3': 'eid', 'CHEXA': 'eid', 'CPENTA': 'eid',
    'CTETRA': 'eid', 'CBAR': 'eid', 'CBEAM': 'eid', 'CROD': 'eid',
    'CONROD': 'eid', 'CBUSH': 'eid', 'CELAS1': 'eid', 'CELAS2': 'eid',
    'CELAS3': 'eid', 'CELAS4': 'eid',
    'CDAMP1': 'eid', 'CDAMP2': 'eid', 'CDAMP3': 'eid', 'CDAMP4': 'eid',
    'CGAP': 'eid', 'CQUAD8': 'eid',
    'CTRIA6': 'eid', 'CQUADR': 'eid', 'CTRIAR': 'eid', 'CSHEAR': 'eid',
    'PLOTEL': 'eid', 'CWELD': 'eid', 'CFAST': 'eid', 'CVISC': 'eid',
    'CHBDYG': 'eid', 'CHBDYE': 'eid',
    'RBE2': 'eid', 'RBE3': 'eid', 'RBAR': 'eid',
    'CONM1': 'eid', 'CONM2': 'eid',
    'CMASS1': 'eid', 'CMASS2': 'eid', 'CMASS3': 'eid', 'CMASS4': 'eid',
    # Properties
    'PSHELL': 'pid', 'PCOMP': 'pid', 'PCOMPG': 'pid', 'PSOLID': 'pid',
    'PBAR': 'pid', 'PBARL': 'pid', 'PBEAM': 'pid', 'PBEAML': 'pid',
    'PROD': 'pid', 'PBUSH': 'pid', 'PBUSHT': 'pid', 'PELAS': 'pid',
    'PDAMP': 'pid', 'PGAP': 'pid', 'PWELD': 'pid', 'PFAST': 'pid',
    'PVISC': 'pid', 'PSHEAR': 'pid', 'PLSOLID': 'pid', 'PCOMPLS': 'pid',
    # Materials
    'MAT1': 'mid', 'MAT2': 'mid', 'MAT8': 'mid', 'MAT9': 'mid',
    'MAT10': 'mid',
    # Coordinate systems
    'CORD2R': 'cid', 'CORD2C': 'cid', 'CORD2S': 'cid',
    'CORD1R': 'cid', 'CORD1C': 'cid', 'CORD1S': 'cid',
    # SPCs
    'SPC': 'spc_id', 'SPC1': 'spc_id', 'SPCADD': 'spc_id',
    # MPCs
    'MPC': 'mpc_id', 'MPCADD': 'mpc_id',
    # Loads
    'FORCE': 'load_id', 'MOMENT': 'load_id', 'PLOAD4': 'load_id',
    'GRAV': 'load_id', 'LOAD': 'load_id', 'TEMP': 'load_id',
    'TEMPD': 'load_id', 'RFORCE': 'load_id', 'RLOAD1': 'load_id',
    'RLOAD2': 'load_id', 'TLOAD1': 'load_id', 'TLOAD2': 'load_id',
    'DAREA': 'load_id', 'DLOAD': 'load_id', 'PLOAD': 'load_id',
    'PLOAD2': 'load_id',
    # Contact
    'BSURF': 'contact_id', 'BSURFS': 'contact_id', 'BCTSET': 'contact_id',
    'BCTADD': 'contact_id', 'BCONP': 'contact_id', 'BCBODY': 'contact_id',
    'BCTPARA': 'contact_id', 'BCTPARM': 'contact_id', 'BLSEG': 'contact_id',
    'BFRIC': 'contact_id',
    # Sets
    'SET1': 'set_id', 'SET3': 'set_id',
    # Methods
    'EIGRL': 'method_id', 'EIGR': 'method_id',
    # Tables
    'TABLED1': 'table_id', 'TABLEM1': 'table_id',
}

ENTITY_TYPES = [
    'nid', 'eid', 'pid', 'mid', 'cid',
    'spc_id', 'mpc_id', 'load_id', 'contact_id',
    'set_id', 'method_id', 'table_id',
]

# Only these geometric entity types get new ID assignments during renumbering.
# Non-geometric types (spc_id, mpc_id, load_id, etc.) keep their primary IDs
# but still have nid/eid/pid/mid/cid cross-references updated.
RENUMBER_TYPES = ['nid', 'eid', 'pid', 'mid', 'cid']

ENTITY_LABELS = {
    'nid': 'Node ID', 'eid': 'Element ID', 'pid': 'Property ID',
    'mid': 'Material ID', 'cid': 'Coord ID', 'spc_id': 'SPC ID',
    'mpc_id': 'MPC ID', 'load_id': 'Load ID', 'contact_id': 'Contact ID',
    'set_id': 'Set ID', 'method_id': 'Method ID', 'table_id': 'Table ID',
}


# ── Model creation ───────────────────────────────────────────────────────────

def make_model(cards_to_skip=None):
    """Create a BDF model with optional card disabling.

    Args:
        cards_to_skip: list of card names to disable (e.g. unsupported contact
            cards). Disabled cards are stored as rejected text and written
            back out unchanged.
    """
    model = BDF(mode='nx')
    if cards_to_skip:
        if hasattr(model, 'disable_cards'):
            model.disable_cards(cards_to_skip)
        else:
            for attr in ('_card_parser', '_card_parser_prepare'):
                parser = getattr(model, attr, None)
                if isinstance(parser, dict):
                    for card in cards_to_skip:
                        parser.pop(card, None)
    return model


# ── Card line parsing ─────────────────────────────────────────────────────────

def extract_card_info(line):
    """Extract card name and primary ID from a raw BDF line.

    Handles fixed-field (8-char or 16-char) and free-field (comma-delimited).
    Returns (name, id) or (None, None) for comments, continuations, blanks.
    """
    stripped = line.strip()
    if not stripped or stripped.startswith('$'):
        return None, None

    first_char = stripped[0]
    if first_char in ('+', '*') or not first_char.isalpha():
        return None, None

    if ',' in stripped:
        fields = stripped.split(',')
        card_name = fields[0].strip().upper()
        id_str = fields[1].strip() if len(fields) > 1 else ''
    else:
        card_name = stripped[:8].strip().upper()
        if card_name.endswith('*'):
            id_str = stripped[8:24].strip() if len(stripped) > 8 else ''
        else:
            id_str = stripped[8:16].strip() if len(stripped) > 8 else ''

    card_name = card_name.rstrip('*')

    try:
        card_id = int(id_str)
    except (ValueError, TypeError):
        return card_name, None

    return card_name, card_id


# ── Include file parser ──────────────────────────────────────────────────────

class IncludeFileParser:
    """Parse raw BDF text to discover includes and determine card ownership.

    pyNastran merges all includes into one model, losing file-of-origin.
    This parser reads raw text to map each card's primary ID to its source file.
    """

    _INCLUDE_RE = re.compile(
        r"^\s*INCLUDE\s+['\"]?(.+?)['\"]?\s*$", re.IGNORECASE | re.MULTILINE)

    def __init__(self):
        self.file_tree = {}          # filepath -> list of child include paths
        self.file_ids = {}           # filepath -> {entity_type: set(ids)}
        self.file_passthrough = {}   # filepath -> list of raw lines for unrecognized cards
        self.all_files = []          # ordered list of filepaths (main first)

    def parse(self, main_path):
        """Parse the main BDF and all includes, building file_ids map."""
        main_path = os.path.abspath(main_path)
        self.file_tree = {}
        self.file_ids = {}
        self.file_passthrough = {}
        self.all_files = []
        self._parse_file(main_path, is_main=True)

    def _parse_file(self, filepath, is_main=False):
        """Recursively parse a single file.

        Include files start in bulk data mode (no exec/case control).
        """
        filepath = os.path.abspath(filepath)
        if filepath in self.file_ids:
            return  # already parsed (avoid cycles)

        self.all_files.append(filepath)
        self.file_ids[filepath] = defaultdict(set)
        self.file_passthrough[filepath] = []
        self.file_tree[filepath] = []

        if not os.path.isfile(filepath):
            return

        with open(filepath, 'r', errors='replace') as f:
            lines = f.readlines()

        file_dir = os.path.dirname(filepath)
        in_bulk = not is_main
        past_exec = not is_main
        in_passthrough_card = False  # track continuations of unrecognized cards

        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            upper = stripped.upper()

            if not past_exec and upper.startswith('CEND'):
                past_exec = True
                i += 1
                continue
            if not in_bulk and upper.startswith('BEGIN') and 'BULK' in upper:
                in_bulk = True
                i += 1
                continue
            if in_bulk and upper.startswith('ENDDATA'):
                break

            inc_match = self._INCLUDE_RE.match(stripped)
            if inc_match:
                inc_path = inc_match.group(1)
                full_path = self._resolve_include(inc_path, file_dir)
                self.file_tree[filepath].append(full_path)
                self._parse_file(full_path, is_main=False)
                in_passthrough_card = False
                i += 1
                continue

            if in_bulk and stripped and not stripped.startswith('$'):
                card_name = self._extract_card_name(stripped)
                if card_name and (card_name[0] == '+' or card_name[0] == '*'):
                    # Continuation line — include if previous card was passthrough
                    if in_passthrough_card:
                        self.file_passthrough[filepath].append(line)
                elif card_name and CARD_ENTITY_MAP.get(card_name) is not None:
                    in_passthrough_card = False
                    self._classify_card(filepath, [line])
                elif card_name:
                    in_passthrough_card = True
                    self.file_passthrough[filepath].append(line)

            i += 1

    @staticmethod
    def _extract_card_name(stripped_line):
        """Extract card name from a stripped bulk data line."""
        if ',' in stripped_line:
            card_name = stripped_line.split(',')[0].strip().upper()
        else:
            card_name = stripped_line[:8].strip().upper()
        return card_name.rstrip('*')

    def _classify_card(self, filepath, card_lines):
        """Extract card name and primary ID, add to file_ids."""
        first_line = card_lines[0].strip()
        if not first_line or first_line.startswith('$'):
            return

        if ',' in first_line:
            fields = first_line.split(',')
            card_name = fields[0].strip().upper()
            id_str = fields[1].strip() if len(fields) > 1 else ''
        else:
            card_name = first_line[:8].strip().upper()
            if card_name.endswith('*'):
                id_str = first_line[8:24].strip() if len(first_line) > 8 else ''
            else:
                id_str = first_line[8:16].strip() if len(first_line) > 8 else ''

        if card_name.startswith('+') or card_name.startswith('*'):
            return

        card_name = card_name.rstrip('*')

        entity_type = CARD_ENTITY_MAP.get(card_name)
        if entity_type is None:
            return

        try:
            card_id = int(id_str)
            if card_id > 0:
                self.file_ids[filepath][entity_type].add(card_id)
        except (ValueError, TypeError):
            pass

    def _resolve_include(self, inc_path, base_dir):
        """Resolve an include path relative to the base directory."""
        inc_path = inc_path.strip().strip("'\"")
        if os.path.isabs(inc_path):
            return os.path.normpath(inc_path)
        return os.path.normpath(os.path.join(base_dir, inc_path))

    def get_summary(self):
        """Return summary dict: {filepath: {entity_type: (count, min_id, max_id)}}."""
        summary = {}
        for fp in self.all_files:
            summary[fp] = {}
            for etype in ENTITY_TYPES:
                ids = self.file_ids[fp].get(etype, set())
                if ids:
                    summary[fp][etype] = (len(ids), min(ids), max(ids))
        return summary
