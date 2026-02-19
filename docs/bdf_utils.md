# bdf_utils — Shared BDF Utilities

Shared module used by both `mass_scale.py` and `renumber_includes.py`.

## What It Provides

| Export | Description |
|--------|-------------|
| `CARD_ENTITY_MAP` | Dict mapping Nastran card names to entity types (nid, eid, pid, mid, etc.) |
| `ENTITY_TYPES` | Canonical ordered list of entity type strings |
| `ENTITY_LABELS` | Human-readable labels for each entity type |
| `make_model(cards_to_skip=None)` | Create a `pyNastran.bdf.BDF` model with optional card disabling |
| `IncludeFileParser` | Parse raw BDF text to discover includes and map card IDs to source files |

## IncludeFileParser

pyNastran merges all INCLUDE files into one model, losing file-of-origin info. This parser reads raw BDF text independently to:

1. Recursively discover all INCLUDE files (nested, quoted/unquoted paths)
2. Track which card IDs (by entity type) belong to which file
3. Build a `file_tree` (parent -> children) and `file_ids` (filepath -> {entity_type: set(ids)})

### Usage

```python
from bdf_utils import IncludeFileParser

parser = IncludeFileParser()
parser.parse('/path/to/main.bdf')

# Ordered list of all files (main first, then includes)
parser.all_files  # ['/path/to/main.bdf', '/path/to/mesh.dat', ...]

# Card ownership: which IDs are in which file
parser.file_ids   # {filepath: {'eid': {1,2,3}, 'nid': {10,20,30}, ...}}

# Summary with counts and ranges
parser.get_summary()  # {filepath: {'eid': (count, min_id, max_id), ...}}
```

## make_model

Creates a BDF model with optional card disabling. Each tool passes its own skip list:

```python
from bdf_utils import make_model

# Mass scale tool skips all contact cards (irrelevant for mass)
model = make_model(['BCPROPS', 'BCTPARM', 'BSURF', ...])

# Renumber tool skips fewer (needs contact cards for renumbering)
model = make_model(['BCPROPS', 'BCTPARM', 'BCPARA', 'BOUTPUT'])
```

Disabled cards are stored as rejected text and written back out unchanged.

## Source

See `preprocessing/bdf_utils.py`.
