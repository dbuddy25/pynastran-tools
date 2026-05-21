"""Shared ASD/FRF helpers shared between postprocessing modules.

These were originally inline in asd_overlay.py and are extracted here
so response_limiting.py can reuse them without duplication.
"""

import numpy as np


RESPONSE_TYPES = {
    "Acceleration": {
        "psd_attr": "accelerations",
        "rms_attr": "accelerations",
        "frf_attr": "accelerations",
        "id_attr": "node_gridtype",
        "entity_label": "Node",
        "dof_labels": ("T1 (X)", "T2 (Y)", "T3 (Z)"),
        "unit_choices": ["in/s²", "m/s²"],
        "unit_factors": {"in/s²": 386.089, "m/s²": 9.80665},
        "psd_units": "g²/Hz",
        "rms_units": "g",
        "rms_fmt": ".3g",
        "frf_units": "g/g",
        "input_label": "Input ASD",
    },
    "Displacement": {
        "psd_attr": "displacements",
        "rms_attr": "displacements",
        "frf_attr": "displacements",
        "id_attr": "node_gridtype",
        "entity_label": "Node",
        "dof_labels": ("T1 (X)", "T2 (Y)", "T3 (Z)"),
        "unit_choices": ["in", "mm", "m"],
        "unit_factors": {"in": 1.0, "mm": 0.0393701, "m": 39.3701},
        "psd_units": "in²/Hz",
        "rms_units": "in (RMS)",
        "rms_fmt": ".2e",
        "frf_units": "in/g",
        "input_label": "Input PSD",
    },
    "SPC Force": {
        "psd_attr": "spc_forces",
        "rms_attr": "spc_forces",
        "frf_attr": "spc_forces",
        "id_attr": "node_gridtype",
        "entity_label": "Node",
        "dof_labels": ("T1 (X)", "T2 (Y)", "T3 (Z)"),
        "unit_choices": ["lbf", "N"],
        "unit_factors": {"lbf": 1.0, "N": 0.224809},
        "psd_units": "lbf²/Hz",
        "rms_units": "lbf (RMS)",
        "rms_fmt": ".3g",
        "frf_units": "lbf/g",
        "input_label": "Input PSD",
    },
    "CBUSH Force": {
        "psd_attr": "cbush_force",
        "rms_attr": "cbush_force",
        "frf_attr": "cbush_force",
        "id_attr": "element",
        "entity_label": "Element",
        "dof_labels": ("F1", "F2", "F3", "M1", "M2", "M3"),
        "unit_choices": ["lbf", "N"],
        "unit_factors": {"lbf": 1.0, "N": 0.224809},
        "psd_units": "lbf²/Hz",
        "rms_units": "lbf (RMS)",
        "rms_fmt": ".3g",
        "frf_units": "lbf/g",
        "input_label": "Input PSD",
    },
}


def sc_int(key):
    """Normalize a pyNastran result dict key to a plain integer subcase ID."""
    return int(key[0]) if isinstance(key, tuple) else int(key)


def subcase_options(result_dict):
    """Return [(sc_id, display_label), ...] sorted by sc_id.

    Pulls SUBTITLE then LABEL from the pyNastran table (CASE CONTROL cards).
    Falls back to the bare integer string when neither is set.
    """
    seen = {}
    for key in sorted(result_dict.keys(), key=sc_int):
        sc = sc_int(key)
        if sc in seen:
            continue
        tbl = result_dict[key]
        sub = (getattr(tbl, "subtitle", "") or "").strip()
        lab = (getattr(tbl, "label", "") or "").strip()
        hint = sub or lab
        display = f"{sc} — {hint}" if hint else str(sc)
        seen[sc] = display
    return list(seen.items())


def lookup_subcase(result_dict, subcase_int):
    """Fetch a result table by integer subcase ID regardless of key format."""
    if subcase_int in result_dict:
        return result_dict[subcase_int]
    for key, val in result_dict.items():
        if sc_int(key) == subcase_int:
            return val
    return None


def parse_asd_text(text_str):
    """Parse 2-column ASD text (freq, g²/Hz). Returns (freqs, asds) arrays or (None, None)."""
    freqs, asds = [], []
    for line in text_str.splitlines():
        line = line.strip()
        if not line or line.startswith('#') or line.startswith('$'):
            continue
        parts = line.replace(',', ' ').split()
        if len(parts) < 2:
            continue
        try:
            freqs.append(float(parts[0]))
            asds.append(float(parts[1]))
        except ValueError:
            continue
    if len(freqs) < 2:
        return None, None
    freqs_arr = np.array(freqs)
    asds_arr = np.array(asds)
    order = np.argsort(freqs_arr)
    return freqs_arr[order], asds_arr[order]


def parse_asd_file(path):
    """Read and parse a 2-column ASD file. Returns (freqs, asds). Raises on error."""
    try:
        with open(path, encoding='utf-8') as f:
            text = f.read()
    except OSError as exc:
        raise OSError(f"Could not read {path}: {exc}") from exc
    freqs, asds = parse_asd_text(text)
    if freqs is None:
        raise ValueError(f"Need at least 2 frequency points in {path}")
    return freqs, asds


def interp_loglog(freqs_in, asd_in, query_freqs):
    """Log-log interpolate asd_in from freqs_in onto query_freqs. Out-of-range → 0."""
    result = np.zeros(len(query_freqs))
    for i, f in enumerate(query_freqs):
        if f < freqs_in[0] or f > freqs_in[-1]:
            result[i] = 0.0
            continue
        idx = int(np.searchsorted(freqs_in, f, side='right')) - 1
        idx = min(idx, len(freqs_in) - 2)
        fl, fh = freqs_in[idx], freqs_in[idx + 1]
        al, ah = asd_in[idx], asd_in[idx + 1]
        if fl <= 0 or fh <= 0 or al <= 0 or ah <= 0 or fl == fh:
            result[i] = al
        else:
            b = np.log(ah / al) / np.log(fh / fl)
            result[i] = al * (f / fl) ** b
    return result


def grms_loglog(freqs, asd):
    """Area under an ASD curve using analytical log-log segment integration (FEMCI)."""
    area = 0.0
    for i in range(len(freqs) - 1):
        fl, fh = float(freqs[i]), float(freqs[i + 1])
        al, ah = float(asd[i]), float(asd[i + 1])
        if fl <= 0 or fh <= 0 or al <= 0 or ah <= 0:
            continue
        log_f = np.log(fh / fl)
        b = np.log(ah / al) / log_f if log_f != 0 else 0.0
        if abs(b + 1.0) < 1e-6:
            area += al * fl * log_f
        else:
            area += (ah * fh - al * fl) / (b + 1.0)
    return area


def cumulative_grms_loglog(freqs, asd):
    """Cumulative RMS array using FEMCI log-log integration. cum[0] = 0."""
    cum_area = np.zeros(len(freqs))
    running = 0.0
    for i in range(len(freqs) - 1):
        fl, fh = float(freqs[i]), float(freqs[i + 1])
        al, ah = float(asd[i]), float(asd[i + 1])
        if fl <= 0 or fh <= 0 or al <= 0 or ah <= 0:
            cum_area[i + 1] = running
            continue
        log_f = np.log(fh / fl)
        b = np.log(ah / al) / log_f if log_f != 0 else 0.0
        if abs(b + 1.0) < 1e-6:
            running += al * fl * log_f
        else:
            running += (ah * fh - al * fl) / (b + 1.0)
        cum_area[i + 1] = running
    return np.sqrt(np.maximum(cum_area, 0.0))
