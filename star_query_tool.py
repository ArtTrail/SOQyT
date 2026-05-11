"""
Stellar Object Query Tool
Queries 5 astronomical databases: SIMBAD, AAVSO VSX, VizieR/2MASS, Gaia DR3, NASA Exoplanet Archive (NEA).
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import math
import statistics
import requests
import json
import os
import urllib.parse
from collections import defaultdict

# ──────────────────────────────────────────────────────────────
# Config persistence
# ──────────────────────────────────────────────────────────────

_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')

def _load_config():
    try:
        with open(_CONFIG_PATH, 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def _save_config(data: dict):
    try:
        existing = _load_config()
        existing.update(data)
        with open(_CONFIG_PATH, 'w') as f:
            json.dump(existing, f, indent=2)
    except Exception:
        pass

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# ──────────────────────────────────────────────────────────────
# Column definitions
# ──────────────────────────────────────────────────────────────

SIMBAD_COLS = [
    ('Name',    'Star Name',  500),
    ('RA_hms',  'RA (hms)',   200),
    ('Dec_dms', 'Dec (dms)',  200),
]

VSX_COLS = [
    ('Name',    'Star Name',  400),
    ('AUID',    'AUID',       160),
    ('RA_hms',  'RA (hms)',   200),
    ('Dec_dms', 'Dec (dms)',  200),
]

TMASS_COLS = [
    ('Name',    '2MASS ID',   400),
    ('RA_hms',  'RA (hms)',   200),
    ('Dec_dms', 'Dec (dms)',  200),
]

WISE_COLS   = [('Name','AllWISE ID',400),('RA_hms','RA (hms)',200),('Dec_dms','Dec (dms)',200)]
APASS_COLS  = [('Name','APASS ID',  400),('RA_hms','RA (hms)',200),('Dec_dms','Dec (dms)',200)]
TYCHO2_COLS = [('Name','Tycho-2 ID',400),('RA_hms','RA (hms)',200),('Dec_dms','Dec (dms)',200)]

GAIA_COLS = [
    ('source_id', 'Source ID',  300),
    ('RA_hms',    'RA (hms)',   200),
    ('Dec_dms',   'Dec (dms)', 200),
]

NEA_OVERVIEW_COLS = [
    ('pl_name',        'Planet',      220),
    ('hostname',       'Host Star',   200),
    ('discoverymethod','Method',      140),
    ('disc_year',      'Year',         70),
    ('disc_facility',  'Facility',    180),
]
NEA_ORBITAL_COLS = [
    ('pl_name',    'Planet',    220),
    ('pl_orbper',  'Period (d)', 120),
    ('pl_orbsmax', 'a (AU)',     100),
    ('pl_orbeccen','Ecc',         80),
    ('pl_orbincl', 'Incl (°)',   100),
    ('pl_tranmid', 'T0 (BJD)',   160),
]
NEA_PLANET_COLS = [
    ('pl_name',   'Planet',      220),
    ('pl_rade',   'Rp (RE)',     100),
    ('pl_bmasse', 'Mp (ME)',     100),
    ('pl_eqt',    'Teq (K)',     100),
    ('pl_trandep','Trans Depth', 120),
    ('pl_trandur','Duration (h)', 120),
]
NEA_STAR_COLS = [
    ('pl_name', 'Planet',    220),
    ('st_teff', 'Teff (K)',  100),
    ('st_logg', 'log g',      80),
    ('st_met',  '[Fe/H]',     80),
    ('st_rad',  'R* (R\u2609)',   100),
    ('st_mass', 'M* (M\u2609)',   100),
]
NEA_SYSTEM_COLS = [
    ('pl_name',    'Planet',      220),
    ('sy_dist',    'Dist (pc)',   110),
    ('sy_vmag',    'V mag',        90),
    ('sy_kmag',    'K mag',        90),
    ('sy_gaiamag', 'Gaia mag',     90),
    ('sy_pnum',    '# Planets',    90),
]

NUMERIC_COLS = {
    'Period', 'MaxMag', 'MinMag', 'Dist_arcsec', 'Score', 'GMag', 'N_refs',
    'pl_orbper', 'pl_orbsmax', 'pl_orbeccen', 'pl_orbincl', 'pl_rade', 'pl_bmasse',
    'pl_eqt', 'pl_trandep', 'pl_trandur', 'st_teff', 'st_logg', 'st_met',
    'st_rad', 'st_mass', 'sy_dist', 'sy_vmag', 'sy_kmag', 'sy_gaiamag',
    'disc_year', 'sy_pnum',
}

SPINNER_FRAMES = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']


# ──────────────────────────────────────────────────────────────
# Coordinate helpers
# ──────────────────────────────────────────────────────────────

def ra_hms_to_deg(s):
    """Convert RA string 'HH MM SS.ss' or 'HH:MM:SS.ss' to decimal degrees."""
    if s is None:
        return None
    s = str(s).strip().replace(':', ' ')
    parts = s.split()
    if len(parts) < 2:
        return None
    try:
        h = float(parts[0])
        m = float(parts[1]) if len(parts) > 1 else 0.0
        sec = float(parts[2]) if len(parts) > 2 else 0.0
        return (h + m / 60.0 + sec / 3600.0) * 15.0
    except (ValueError, IndexError):
        return None

def dec_dms_to_deg(s):
    """Convert Dec string '±DD MM SS.s' or '±DD:MM:SS.s' to decimal degrees."""
    if s is None:
        return None
    s = str(s).strip().replace(':', ' ')
    neg = s.startswith('-')
    s = s.lstrip('+-')
    parts = s.split()
    if not parts:
        return None
    try:
        d = float(parts[0])
        m = float(parts[1]) if len(parts) > 1 else 0.0
        sec = float(parts[2]) if len(parts) > 2 else 0.0
        val = d + m / 60.0 + sec / 3600.0
        return -val if neg else val
    except (ValueError, IndexError):
        return None

def parse_ra(s):
    """Parse RA in hms or decimal degrees."""
    if s is None:
        return None
    s = str(s).strip()
    # If it contains letters or colons or spaces with 3 parts → hms
    if ':' in s or (len(s.split()) == 3 and not s.replace('.', '').replace('-', '').replace(' ', '').isdigit()):
        return ra_hms_to_deg(s)
    try:
        return float(s)
    except ValueError:
        return ra_hms_to_deg(s)

def parse_dec(s):
    """Parse Dec in dms or decimal degrees."""
    if s is None:
        return None
    s = str(s).strip()
    if ':' in s or (len(s.split()) >= 2 and not s.replace('.', '').replace('-', '').replace('+', '').replace(' ', '').isdigit()):
        return dec_dms_to_deg(s)
    try:
        return float(s)
    except ValueError:
        return dec_dms_to_deg(s)

def deg_to_hms(deg):
    """Convert decimal degrees RA to HH MM SS.ss string."""
    if deg is None:
        return ''
    deg = float(deg) % 360.0
    h_total = deg / 15.0
    h = int(h_total)
    m_total = (h_total - h) * 60.0
    m = int(m_total)
    sec = (m_total - m) * 60.0
    return f"{h:02d} {m:02d} {sec:05.2f}"

def deg_to_dms(deg):
    """Convert decimal degrees Dec to ±DD MM SS.s string."""
    if deg is None:
        return ''
    neg = deg < 0
    deg = abs(float(deg))
    d = int(deg)
    m_total = (deg - d) * 60.0
    m = int(m_total)
    sec = (m_total - m) * 60.0
    sign = '-' if neg else '+'
    return f"{sign}{d:02d} {m:02d} {sec:04.1f}"

def arcsec_dist(ra1, dec1, ra2, dec2):
    """Great-circle distance in arcseconds between two (RA, Dec) points in degrees."""
    if None in (ra1, dec1, ra2, dec2):
        return None
    r1 = math.radians(ra1)
    r2 = math.radians(ra2)
    d1 = math.radians(dec1)
    d2 = math.radians(dec2)
    cos_angle = (math.sin(d1) * math.sin(d2) +
                 math.cos(d1) * math.cos(d2) * math.cos(r1 - r2))
    cos_angle = max(-1.0, min(1.0, cos_angle))
    return math.degrees(math.acos(cos_angle)) * 3600.0

def _parse_vsx_ra(s):
    """Parse VSX RA 'HH MM SS.ss' to decimal degrees."""
    return ra_hms_to_deg(s)

def _parse_vsx_dec(s):
    """Parse VSX Dec '±DD MM SS.s' to decimal degrees."""
    return dec_dms_to_deg(s)


# ──────────────────────────────────────────────────────────────
# Filter helpers
# ──────────────────────────────────────────────────────────────

def _try_float(val):
    """Return float or None."""
    if val is None or val == '':
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None

def _strip_flags(val):
    """Strip flag characters from a magnitude/period string, return numeric string or ''."""
    if val is None:
        return ''
    s = str(val).strip().lstrip('<>(: ').rstrip('):> ')
    # Remove single trailing/leading colon or parenthesis
    s = s.strip('():< ')
    try:
        float(s)
        return s
    except ValueError:
        # Try removing individual flag chars
        cleaned = ''
        for c in s:
            if c.isdigit() or c in '.+-':
                cleaned += c
        try:
            float(cleaned)
            return cleaned
        except ValueError:
            return ''

def _passes_filters(period_val, mag_val, period_min, period_max, mag_min, mag_max):
    """
    Apply period and magnitude filters. If the star has no data for that field,
    do NOT exclude it — show it without filtering on that field.
    """
    p = _try_float(period_val)
    m = _try_float(mag_val)

    if p is not None:
        if period_min is not None and p < period_min:
            return False
        if period_max is not None and p > period_max:
            return False

    if m is not None:
        if mag_min is not None and m < mag_min:
            return False
        if mag_max is not None and m > mag_max:
            return False

    return True


# ──────────────────────────────────────────────────────────────
# SIMBAD queries
# ──────────────────────────────────────────────────────────────

SIMBAD_TAP = "https://simbad.u-strasbg.fr/simbad/sim-tap/sync"

OTYPE_LABELS = {
    'RR*': 'RR Lyrae',
    'Ce*': 'Cepheid',
    'dS*': 'Delta Sct',
    'RV*': 'RV Tau',
    'LP*': 'Long-period',
    'SR*': 'Semi-regular',
    'Mi*': 'Mira',
    'Ell*': 'Ellipsoidal',
    'Ro*': 'Rotating',
    'EB*': 'Eclipsing bin.',
    'WV*': 'W Vir',
}

def _build_simbad_results(rows, period_min, period_max, mag_min, mag_max,
                           ra_center=None, dec_center=None, name_mode=False):
    """
    Process raw SIMBAD TAP rows into result dicts.
    SELECT columns: main_id(0), ra(1), dec(2), otype(3), vartyp(4),
                    period(5), vmax(6), vmin(7), magtyp(8), bibcode(9)
    Groups by main_id, computes median period/mag, counts bibcodes for N_refs.
    """
    grouped = defaultdict(list)
    for row in rows:
        grouped[row[0]].append(row)

    results = []
    for main_id, star_rows in grouped.items():
        r0      = star_rows[0]
        ra_val  = r0[1]
        dec_val = r0[2]
        otype   = (r0[3] or '').strip()
        vartype = (r0[4] or '').strip()

        periods_clean = []
        vmaxes_clean  = []
        vmins_clean   = []
        magtyp_val    = ''
        bibcodes      = set()

        for r in star_rows:
            if r[5] is not None:
                try: periods_clean.append(float(r[5]))
                except (ValueError, TypeError): pass
            if r[6] is not None:
                try: vmaxes_clean.append(float(r[6]))
                except (ValueError, TypeError): pass
            if r[7] is not None:
                try: vmins_clean.append(float(r[7]))
                except (ValueError, TypeError): pass
            if r[8]:
                magtyp_val = str(r[8])
            if r[9]:
                bibcodes.add(str(r[9]))

        period_med = statistics.median(periods_clean) if periods_clean else None
        vmax_med   = statistics.median(vmaxes_clean)  if vmaxes_clean  else None
        vmin_med   = statistics.median(vmins_clean)   if vmins_clean   else None

        period_str = f"{period_med:.4f}" if period_med is not None else ''
        maxmag_str = f"{vmax_med:.3f}"   if vmax_med   is not None else ''
        minmag_str = f"{vmin_med:.3f}"   if vmin_med   is not None else ''

        if not _passes_filters(period_str, maxmag_str, period_min, period_max, mag_min, mag_max):
            continue

        try:
            ra_f = float(ra_val) if ra_val is not None else None
        except (ValueError, TypeError):
            ra_f = None
        try:
            dec_f = float(dec_val) if dec_val is not None else None
        except (ValueError, TypeError):
            dec_f = None

        dist = arcsec_dist(ra_center, dec_center, ra_f, dec_f) if ra_center is not None else None
        dist_str = f"{dist:.1f}" if dist is not None else ''

        otype_label = OTYPE_LABELS.get(otype, otype)

        results.append({
            'Name':        main_id or '',
            'RA_hms':      deg_to_hms(ra_f),
            'Dec_dms':     deg_to_dms(dec_f),
            'OType_label': otype_label,
            'VarType':     vartype,
            'Period':      period_str,
            'MaxMag':      maxmag_str,
            'MinMag':      minmag_str,
            'MagBand':     magtyp_val,
            'Dist_arcsec': dist_str,
            'N_refs':      str(len(bibcodes)),
        })

    return results


def query_simbad(ra_deg, dec_deg, radius_arcmin, otype_filter,
                 period_min, period_max, mag_min, mag_max, status_callback):
    """Coordinate-based SIMBAD query (variable-star focused, INNER JOIN mesVar)."""
    radius_deg = radius_arcmin / 60.0

    otype_clause = ""
    if otype_filter and otype_filter != 'All':
        otype_clause = f"AND b.otype = '{otype_filter}'"

    adql = f"""
SELECT b.main_id, b.ra, b.dec, b.otype, v.vartyp, v.period, v.vmax, v.vmin, v.magtyp, v.bibcode
FROM basic b
INNER JOIN mesVar v ON b.oid = v.oidref
WHERE CONTAINS(POINT('ICRS', b.ra, b.dec),
               CIRCLE('ICRS', {ra_deg}, {dec_deg}, {radius_deg})) = 1
  AND v.period IS NOT NULL
  {otype_clause}
""".strip()

    if status_callback:
        status_callback("Querying SIMBAD...")
    params = {
        'REQUEST': 'doQuery',
        'LANG': 'ADQL',
        'FORMAT': 'json',
        'QUERY': adql,
    }
    r = requests.post(SIMBAD_TAP, data=params, timeout=60)
    r.raise_for_status()
    data = r.json().get('data', [])

    return _build_simbad_results(data, period_min, period_max, mag_min, mag_max,
                                  ra_center=ra_deg, dec_center=dec_deg)


def query_simbad_by_name(name, period_min, period_max, mag_min, mag_max,
                          status_callback, otype_filter=None):
    """Name-based SIMBAD query (LEFT JOIN mesVar so any star can be found)."""

    def _run_adql(adql):
        params = {
            'REQUEST': 'doQuery',
            'LANG': 'ADQL',
            'FORMAT': 'json',
            'QUERY': adql,
        }
        r = requests.post(SIMBAD_TAP, data=params, timeout=60)
        r.raise_for_status()
        return r.json().get('data', [])

    if status_callback:
        status_callback("Querying SIMBAD...")

    # Exact match first
    adql_exact = f"""
SELECT b.main_id, b.ra, b.dec, b.otype, v.vartyp, v.period, v.vmax, v.vmin, v.magtyp, v.bibcode
FROM basic b
LEFT JOIN mesVar v ON b.oid = v.oidref
JOIN ident i ON b.oid = i.oidref
WHERE i.id = '{name.replace("'", "''")}'
""".strip()

    data = _run_adql(adql_exact)

    if not data:
        # Fallback: LIKE search
        adql_like = f"""
SELECT b.main_id, b.ra, b.dec, b.otype, v.vartyp, v.period, v.vmax, v.vmin, v.magtyp, v.bibcode
FROM basic b
LEFT JOIN mesVar v ON b.oid = v.oidref
JOIN ident i ON b.oid = i.oidref
WHERE i.id LIKE '%{name.replace("'", "''")}%'
""".strip()
        data = _run_adql(adql_like)

    # Apply otype filter if specified
    if otype_filter and otype_filter != 'All' and data:
        data = [row for row in data if row[3] and row[3].strip() == otype_filter]

    return _build_simbad_results(data, period_min, period_max, mag_min, mag_max,
                                  ra_center=None, dec_center=None, name_mode=True)


# ──────────────────────────────────────────────────────────────
# AAVSO VSX queries
# ──────────────────────────────────────────────────────────────

VSX_BASE = "https://www.aavso.org/vsx/index.php"

def _parse_vsx_object(obj, ra_center=None, dec_center=None):
    """Parse a single VSX object dict into a result dict."""
    name = obj.get('Name', '') or ''
    auid = obj.get('AUID', '') or ''

    ra_raw  = obj.get('RA2000', '') or ''
    dec_raw = obj.get('Declination2000', '') or ''

    ra_deg  = _parse_vsx_ra(str(ra_raw).strip())
    dec_deg = _parse_vsx_dec(str(dec_raw).strip())

    vartype = obj.get('VariabilityType', '') or ''
    period_raw = obj.get('Period', '') or ''
    maxmag_raw = obj.get('MaxMag', '') or ''
    minmag_raw = obj.get('MinMag', '') or ''
    magband    = obj.get('Bands', '') or obj.get('Band', '') or ''

    period_str = _strip_flags(period_raw)
    maxmag_str = _strip_flags(maxmag_raw)
    minmag_str = _strip_flags(minmag_raw)

    dist = arcsec_dist(ra_center, dec_center, ra_deg, dec_deg) if ra_center is not None else None
    dist_str = f"{dist:.1f}" if dist is not None else ''

    return {
        'Name':        name,
        'AUID':        auid,
        'RA_hms':      deg_to_hms(ra_deg),
        'Dec_dms':     deg_to_dms(dec_deg),
        'VarType':     str(vartype),
        'Period':      period_str,
        'MaxMag':      maxmag_str,
        'MinMag':      minmag_str,
        'MagBand':     str(magband),
        'Dist_arcsec': dist_str,
        '_ra_deg':     ra_deg,
        '_dec_deg':    dec_deg,
    }


def query_vsx_by_name(name, period_min, period_max, mag_min, mag_max, status_callback):
    """Name-based VSX query."""
    if status_callback:
        status_callback("Querying AAVSO VSX...")

    url = VSX_BASE
    params = {
        'view': 'api.object',
        'format': 'json',
        'ident': name,
    }
    r = requests.get(url, params=params, timeout=30)
    r.raise_for_status()
    j = r.json()
    obj = j.get('VSXObject')
    if not obj:
        return []

    result = _parse_vsx_object(obj)
    if not _passes_filters(result['Period'], result['MaxMag'],
                           period_min, period_max, mag_min, mag_max):
        return []
    # Remove internal keys
    result.pop('_ra_deg', None)
    result.pop('_dec_deg', None)
    return [result]


def query_vsx(ra_deg, dec_deg, radius_arcmin, period_min, period_max,
              mag_min, mag_max, status_callback):
    """Coordinate-based VSX query."""
    if status_callback:
        status_callback("Querying AAVSO VSX...")

    radius_deg = radius_arcmin / 60.0
    params = {
        'view': 'api.list',
        'format': 'json',
        'ra': ra_deg,
        'dec': dec_deg,
        'radius': radius_deg,
        'coords': 'decimal',
    }
    r = requests.get(VSX_BASE, params=params, timeout=30)
    r.raise_for_status()
    j = r.json()

    raw = j.get('VSXObjects', {}).get('VSXObject', [])
    if isinstance(raw, dict):
        raw = [raw]

    results = []
    for obj in raw:
        result = _parse_vsx_object(obj, ra_center=ra_deg, dec_center=dec_deg)
        if not _passes_filters(result['Period'], result['MaxMag'],
                               period_min, period_max, mag_min, mag_max):
            continue
        result.pop('_ra_deg', None)
        result.pop('_dec_deg', None)
        results.append(result)

    return results


# ──────────────────────────────────────────────────────────────
# VizieR / GCVS queries
# ──────────────────────────────────────────────────────────────

VIZIER_TAP = "https://tapvizier.cds.unistra.fr/TAPVizieR/tap/sync"

def _tmass_name_from_coords(ra_f, dec_f):
    """Build a 2MASS-style designation from decimal degrees RA/Dec."""
    if ra_f is None or dec_f is None:
        return ''
    # RA → HHMMSS.ss
    h_tot = ra_f / 15.0
    h = int(h_tot); m_tot = (h_tot - h) * 60; m = int(m_tot); s = (m_tot - m) * 60
    ra_s = f"{h:02d}{m:02d}{s:05.2f}"
    # Dec → ±DDMMSS.s
    neg = dec_f < 0; dabs = abs(dec_f)
    d = int(dabs); m2_tot = (dabs - d) * 60; m2 = int(m2_tot); s2 = (m2_tot - m2) * 60
    dec_s = f"{'-' if neg else '+'}{d:02d}{m2:02d}{s2:04.1f}"
    return f"2MASS J{ra_s}{dec_s}"


def _tmass_row_to_result(row, ra_center=None, dec_center=None):
    """Convert a 2MASS TAP row to result dict.
    Row: (RAJ2000, DEJ2000, Jmag, Hmag, Kmag)  — 5 columns, no name column.
    """
    ra_raw  = row[0]
    dec_raw = row[1]
    jmag    = row[2]
    hmag    = row[3]
    kmag    = row[4]

    try:
        ra_f = float(ra_raw) if ra_raw is not None else None
    except (ValueError, TypeError):
        ra_f = None
    try:
        dec_f = float(dec_raw) if dec_raw is not None else None
    except (ValueError, TypeError):
        dec_f = None

    def _mag(v):
        try:
            return f"{float(v):.3f}" if v is not None else ''
        except (ValueError, TypeError):
            return ''

    dist = arcsec_dist(ra_center, dec_center, ra_f, dec_f) if ra_center is not None else None
    dist_str = f"{dist:.1f}" if dist is not None else ''

    return {
        'Name':        _tmass_name_from_coords(ra_f, dec_f),
        'RA_hms':      deg_to_hms(ra_f),
        'Dec_dms':     deg_to_dms(dec_f),
        'Jmag':        _mag(jmag),
        'Hmag':        _mag(hmag),
        'Kmag':        _mag(kmag),
        'Dist_arcsec': dist_str,
        '_ra_deg':     ra_f,
        '_dec_deg':    dec_f,
    }


def _parse_votable(text, service='TAP'):
    """
    Parse a VOTable XML response from a TAP service.
    Raises RuntimeError with the server's error message if QUERY_STATUS=ERROR.
    Returns a list of row-lists (TD text values, None for empty cells) on success.
    """
    import xml.etree.ElementTree as ET
    try:
        root = ET.fromstring(text)
    except ET.ParseError as e:
        raise RuntimeError(f"{service}: could not parse response XML: {e}")

    # Check for QUERY_STATUS ERROR (namespace-agnostic tag match)
    for elem in root.iter():
        local = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if local == 'INFO' and elem.get('name') == 'QUERY_STATUS':
            if elem.get('value', '').upper() == 'ERROR':
                msg = (elem.text or '').strip() or 'Unknown query error'
                raise RuntimeError(f"{service}: {msg}")

    # Extract TABLEDATA rows
    rows = []
    for elem in root.iter():
        local = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if local == 'TR':
            row = []
            for child in elem:
                cloc = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if cloc == 'TD':
                    row.append(child.text)   # None if cell is empty
            rows.append(row)
    return rows


def _parse_gaia_csv(text):
    """
    Parse a CSV response from the Gaia ESAC TAP server.
    Returns a list of dicts keyed by column name (from the CSV header row).
    Empty cells become None. Raises RuntimeError on server errors.
    """
    import csv, io
    lines = text.strip().splitlines()
    if not lines:
        return []
    first = lines[0].strip()
    if first.startswith('<'):
        raise RuntimeError(f"Gaia: {first[:200]}")
    reader = csv.DictReader(io.StringIO(text))
    return [{k: (v if v != '' else None) for k, v in row.items()} for row in reader]


def _vizier_post(adql):
    """POST an ADQL query to VizieR TAP, parse VOTable, return data rows."""
    params = {'REQUEST': 'doQuery', 'LANG': 'ADQL', 'FORMAT': 'votable', 'QUERY': adql}
    r = requests.post(VIZIER_TAP, data=params, timeout=60)
    # Don't raise_for_status first — VizieR returns 400 with a VOTable body
    # that contains the actual error message; let _parse_votable extract it.
    try:
        return _parse_votable(r.text, service='VizieR')
    except RuntimeError:
        raise
    except Exception:
        r.raise_for_status()   # fallback for truly non-XML responses
        return []


def _tmass_adql_cone(ra_deg, dec_deg, radius_deg):
    return f"""
SELECT t.RAJ2000, t.DEJ2000, t.Jmag, t.Hmag, t.Kmag
FROM "II/246/out" AS t
WHERE CONTAINS(POINT('ICRS', t.RAJ2000, t.DEJ2000),
               CIRCLE('ICRS', {ra_deg}, {dec_deg}, {radius_deg})) = 1
""".strip()


def query_tmass_by_name(name, radius_arcmin, status_callback):
    """Name-based 2MASS query: resolve via SIMBAD then cone search."""
    if status_callback:
        status_callback("Resolving name for VizieR/2MASS...")
    ra_deg, dec_deg = _resolve_name_simbad(name)
    if ra_deg is None:
        return []
    if status_callback:
        status_callback("Querying VizieR / 2MASS...")
    data = _vizier_post(_tmass_adql_cone(ra_deg, dec_deg, radius_arcmin / 60.0))
    results = []
    for row in data:
        r = _tmass_row_to_result(row, ra_center=ra_deg, dec_center=dec_deg)
        r.pop('_ra_deg', None); r.pop('_dec_deg', None)
        results.append(r)
    return results


def query_tmass(ra_deg, dec_deg, radius_arcmin, status_callback):
    """Coordinate-based 2MASS query via VizieR TAP."""
    if status_callback:
        status_callback("Querying VizieR / 2MASS...")
    radius_deg = radius_arcmin / 60.0
    data = _vizier_post(_tmass_adql_cone(ra_deg, dec_deg, radius_deg))
    results = []
    for row in data:
        r = _tmass_row_to_result(row, ra_center=ra_deg, dec_center=dec_deg)
        r.pop('_ra_deg', None); r.pop('_dec_deg', None)
        results.append(r)
    return results


# ── AllWISE ───────────────────────────────────────────────────

def _wise_name_from_coords(ra_f, dec_f):
    if ra_f is None or dec_f is None:
        return ''
    ra_h  = int(ra_f / 15)
    ra_m  = int((ra_f / 15 - ra_h) * 60)
    ra_s  = ((ra_f / 15 - ra_h) * 60 - ra_m) * 60
    sign  = '+' if dec_f >= 0 else '-'
    ad    = abs(dec_f)
    dd    = int(ad)
    dm    = int((ad - dd) * 60)
    ds    = ((ad - dd) * 60 - dm) * 60
    return f"WISE J{ra_h:02d}{ra_m:02d}{ra_s:05.2f}{sign}{dd:02d}{dm:02d}{ds:04.1f}"


def _wise_adql_cone(ra_deg, dec_deg, radius_deg):
    return f"""
SELECT t.RAJ2000, t.DEJ2000, t.W1mag, t.W2mag, t.W3mag, t.W4mag
FROM "II/328/allwise" AS t
WHERE CONTAINS(POINT('ICRS', t.RAJ2000, t.DEJ2000),
               CIRCLE('ICRS', {ra_deg}, {dec_deg}, {radius_deg})) = 1
""".strip()


def _wise_row_to_result(row, ra_center=None, dec_center=None):
    def _mag(v):
        try: return f"{float(v):.3f}" if v is not None else ''
        except: return ''
    try: ra_f  = float(row[0])
    except: ra_f = None
    try: dec_f = float(row[1])
    except: dec_f = None
    dist = arcsec_dist(ra_center, dec_center, ra_f, dec_f) if ra_center is not None else None
    return {
        'Name':        _wise_name_from_coords(ra_f, dec_f),
        'RA_hms':      deg_to_hms(ra_f),
        'Dec_dms':     deg_to_dms(dec_f),
        'W1mag':       _mag(row[2] if len(row) > 2 else None),
        'W2mag':       _mag(row[3] if len(row) > 3 else None),
        'W3mag':       _mag(row[4] if len(row) > 4 else None),
        'W4mag':       _mag(row[5] if len(row) > 5 else None),
        'Dist_arcsec': f"{dist:.1f}" if dist is not None else '',
    }


def query_wise_by_name(name, radius_arcmin, status_callback):
    if status_callback:
        status_callback("Resolving name for AllWISE...")
    ra_deg, dec_deg = _resolve_name_simbad(name)
    if ra_deg is None:
        return []
    if status_callback:
        status_callback("Querying AllWISE...")
    data = _vizier_post(_wise_adql_cone(ra_deg, dec_deg, radius_arcmin / 60.0))
    return [_wise_row_to_result(row, ra_center=ra_deg, dec_center=dec_deg) for row in data]


def query_wise(ra_deg, dec_deg, radius_arcmin, status_callback):
    if status_callback:
        status_callback("Querying AllWISE...")
    data = _vizier_post(_wise_adql_cone(ra_deg, dec_deg, radius_arcmin / 60.0))
    return [_wise_row_to_result(row, ra_center=ra_deg, dec_center=dec_deg) for row in data]


# ── APASS DR9 ─────────────────────────────────────────────────

def _apass_name_from_coords(ra_f, dec_f):
    if ra_f is None or dec_f is None:
        return ''
    sign = '+' if dec_f >= 0 else '-'
    return f"APASS J{ra_f:09.5f}{sign}{abs(dec_f):08.5f}"


def _apass_adql_cone(ra_deg, dec_deg, radius_deg):
    return f"""
SELECT t.RAJ2000, t.DEJ2000, t.Vmag, t.Bmag
FROM "II/336/apass9" AS t
WHERE CONTAINS(POINT('ICRS', t.RAJ2000, t.DEJ2000),
               CIRCLE('ICRS', {ra_deg}, {dec_deg}, {radius_deg})) = 1
""".strip()


def _apass_row_to_result(row, ra_center=None, dec_center=None):
    def _mag(v):
        try: return f"{float(v):.3f}" if v is not None else ''
        except: return ''
    try: ra_f  = float(row[0])
    except: ra_f = None
    try: dec_f = float(row[1])
    except: dec_f = None
    dist = arcsec_dist(ra_center, dec_center, ra_f, dec_f) if ra_center is not None else None
    return {
        'Name':        _apass_name_from_coords(ra_f, dec_f),
        'RA_hms':      deg_to_hms(ra_f),
        'Dec_dms':     deg_to_dms(dec_f),
        'Vmag':        _mag(row[2] if len(row) > 2 else None),
        'Bmag':        _mag(row[3] if len(row) > 3 else None),
        'Dist_arcsec': f"{dist:.1f}" if dist is not None else '',
    }


def query_apass_by_name(name, radius_arcmin, status_callback):
    if status_callback:
        status_callback("Resolving name for APASS...")
    ra_deg, dec_deg = _resolve_name_simbad(name)
    if ra_deg is None:
        return []
    if status_callback:
        status_callback("Querying APASS DR9...")
    data = _vizier_post(_apass_adql_cone(ra_deg, dec_deg, radius_arcmin / 60.0))
    return [_apass_row_to_result(row, ra_center=ra_deg, dec_center=dec_deg) for row in data]


def query_apass(ra_deg, dec_deg, radius_arcmin, status_callback):
    if status_callback:
        status_callback("Querying APASS DR9...")
    data = _vizier_post(_apass_adql_cone(ra_deg, dec_deg, radius_arcmin / 60.0))
    return [_apass_row_to_result(row, ra_center=ra_deg, dec_center=dec_deg) for row in data]


# ── Tycho-2 ───────────────────────────────────────────────────

def _tycho2_adql_cone(ra_deg, dec_deg, radius_deg):
    return f"""
SELECT t.RAmdeg, t.DEmdeg, t.BTmag, t.VTmag, t.pmRA, t.pmDE, t.TYC1, t.TYC2, t.TYC3
FROM "I/259/tyc2" AS t
WHERE CONTAINS(POINT('ICRS', t.RAmdeg, t.DEmdeg),
               CIRCLE('ICRS', {ra_deg}, {dec_deg}, {radius_deg})) = 1
""".strip()


def _tycho2_row_to_result(row, ra_center=None, dec_center=None):
    def _mag(v):
        try: return f"{float(v):.3f}" if v is not None else ''
        except: return ''
    def _pm(v):
        try: return f"{float(v):.1f}" if v is not None else ''
        except: return ''
    try: ra_f  = float(row[0])
    except: ra_f = None
    try: dec_f = float(row[1])
    except: dec_f = None
    # Build TYC designation from TYC1, TYC2, TYC3
    try:
        tyc1 = str(int(float(row[6]))) if row[6] is not None else ''
        tyc2 = str(int(float(row[7]))) if row[7] is not None else ''
        tyc3 = str(int(float(row[8]))) if row[8] is not None else ''
        tyc_name = f"TYC {tyc1}-{tyc2}-{tyc3}" if tyc1 else ''
    except Exception:
        tyc_name = ''
    dist = arcsec_dist(ra_center, dec_center, ra_f, dec_f) if ra_center is not None else None
    return {
        'Name':        tyc_name,
        'RA_hms':      deg_to_hms(ra_f),
        'Dec_dms':     deg_to_dms(dec_f),
        'BTmag':       _mag(row[2] if len(row) > 2 else None),
        'VTmag':       _mag(row[3] if len(row) > 3 else None),
        'pmRA':        _pm(row[4] if len(row) > 4 else None),
        'pmDE':        _pm(row[5] if len(row) > 5 else None),
        'Dist_arcsec': f"{dist:.1f}" if dist is not None else '',
    }


def query_tycho2_by_name(name, radius_arcmin, status_callback):
    if status_callback:
        status_callback("Resolving name for Tycho-2...")
    ra_deg, dec_deg = _resolve_name_simbad(name)
    if ra_deg is None:
        return []
    if status_callback:
        status_callback("Querying Tycho-2...")
    data = _vizier_post(_tycho2_adql_cone(ra_deg, dec_deg, radius_arcmin / 60.0))
    return [_tycho2_row_to_result(row, ra_center=ra_deg, dec_center=dec_deg) for row in data]


def query_tycho2(ra_deg, dec_deg, radius_arcmin, status_callback):
    if status_callback:
        status_callback("Querying Tycho-2...")
    data = _vizier_post(_tycho2_adql_cone(ra_deg, dec_deg, radius_arcmin / 60.0))
    return [_tycho2_row_to_result(row, ra_center=ra_deg, dec_center=dec_deg) for row in data]


# ──────────────────────────────────────────────────────────────
# Gaia DR3 queries
# ──────────────────────────────────────────────────────────────

GAIA_TAP = "https://gea.esac.esa.int/tap-server/tap/sync"

# Optional parameter groups — each entry: (key, label, [gaia_source columns])
GAIA_PARAM_OPTS = [
    ('bp_mag',   'BP Magnitude',     ['phot_bp_mean_mag']),
    ('rp_mag',   'RP Magnitude',     ['phot_rp_mean_mag']),
    ('parallax', 'Parallax',         ['parallax', 'parallax_error']),
    ('pm',       'Proper Motion',    ['pmra', 'pmra_error', 'pmdec', 'pmdec_error']),
    ('rv',       'Radial Velocity',  ['radial_velocity', 'radial_velocity_error']),
    ('ruwe',     'RUWE',             ['ruwe']),
]

def _resolve_name_simbad(name):
    """Resolve object name via SIMBAD to get (ra_deg, dec_deg). Returns (None, None) if not found."""
    adql = f"""
SELECT b.ra, b.dec
FROM basic b
JOIN ident i ON b.oid = i.oidref
WHERE i.id = '{name.replace("'", "''")}'
""".strip()
    params = {
        'REQUEST': 'doQuery',
        'LANG': 'ADQL',
        'FORMAT': 'json',
        'QUERY': adql,
    }
    try:
        r = requests.post(SIMBAD_TAP, data=params, timeout=30)
        r.raise_for_status()
        data = r.json().get('data', [])
        if data and data[0][0] is not None:
            return float(data[0][0]), float(data[0][1])
    except Exception:
        pass
    return None, None


def query_gaia(ra_deg, dec_deg, radius_arcmin, status_callback, extra_cols=None):
    """Coordinate-based Gaia DR3 query — any stellar object in gaia_source."""
    if status_callback:
        status_callback("Querying Gaia DR3...")

    radius_deg = radius_arcmin / 60.0
    base_cols = ['source_id', 'ra', 'dec', 'phot_g_mean_mag']
    select = ', '.join(base_cols + (extra_cols or []))
    adql = f"""
SELECT {select}
FROM gaiadr3.gaia_source
WHERE ra  BETWEEN {ra_deg  - radius_deg} AND {ra_deg  + radius_deg}
  AND dec BETWEEN {dec_deg - radius_deg} AND {dec_deg + radius_deg}
""".strip()

    params = {
        'REQUEST': 'doQuery',
        'LANG': 'ADQL',
        'FORMAT': 'csv',
        'MAXREC': '2000',
        'QUERY': adql,
    }
    r = requests.post(GAIA_TAP, data=params, timeout=120)
    r.raise_for_status()
    data = _parse_gaia_csv(r.text)
    return _gaia_rows_to_results(data, ra_center=ra_deg, dec_center=dec_deg)


def _gaia_rows_to_results(rows, ra_center=None, dec_center=None):
    """Convert list of Gaia CSV dicts (keyed by column name) to result dicts for display."""
    def _f(v, fmt='.3f'):
        try:
            return format(float(v), fmt) if v is not None else ''
        except (ValueError, TypeError):
            return ''

    results = []
    for row in rows:
        ra_f  = float(row['ra'])  if row.get('ra')  is not None else None
        dec_f = float(row['dec']) if row.get('dec') is not None else None
        dist  = arcsec_dist(ra_center, dec_center, ra_f, dec_f) if ra_center is not None else None
        results.append({
            'source_id':   str(row.get('source_id') or ''),
            'RA_hms':      deg_to_hms(ra_f),
            'Dec_dms':     deg_to_dms(dec_f),
            'GMag':        _f(row.get('phot_g_mean_mag')),
            'BPMag':       _f(row.get('phot_bp_mean_mag')),
            'RPMag':       _f(row.get('phot_rp_mean_mag')),
            'Parallax':    _f(row.get('parallax'),       '.4f'),
            'PlxErr':      _f(row.get('parallax_error'), '.4f'),
            'PMRA':        _f(row.get('pmra'),           '.3f'),
            'PMRAErr':     _f(row.get('pmra_error'),     '.3f'),
            'PMDec':       _f(row.get('pmdec'),          '.3f'),
            'PMDecErr':    _f(row.get('pmdec_error'),    '.3f'),
            'RV':          _f(row.get('radial_velocity'),       '.2f'),
            'RVErr':       _f(row.get('radial_velocity_error'), '.2f'),
            'RUWE':        _f(row.get('ruwe'), '.3f'),
            'Dist_arcsec': f"{dist:.1f}" if dist is not None else '',
        })
    return results


def query_gaia_by_name(name, radius_arcmin, status_callback, extra_cols=None):
    """Name-based Gaia DR3 query.
    If name looks like a Gaia source_id (all digits or 'Gaia DR3 XXXXX'), do a direct
    source_id lookup.  Otherwise resolve via SIMBAD and do a cone search.
    """
    if status_callback:
        status_callback("Querying Gaia DR3...")

    extra = extra_cols or []

    # Detect raw source_id or "Gaia DR3 XXXXX" prefix
    clean = name.strip()
    source_id_int = None
    if clean.lower().startswith('gaia dr3 '):
        candidate = clean[9:].strip()
        if candidate.isdigit():
            source_id_int = int(candidate)
    elif clean.isdigit() and len(clean) >= 10:
        source_id_int = int(clean)

    if source_id_int is not None:
        base_cols = ['gs.source_id', 'gs.ra', 'gs.dec', 'gs.phot_g_mean_mag']
        extra_prefixed = [f'gs.{c}' for c in extra]
        select = ', '.join(base_cols + extra_prefixed)
        adql = f"""
SELECT {select}
FROM gaiadr3.gaia_source AS gs
WHERE gs.source_id = {source_id_int}
""".strip()
        params = {
            'REQUEST': 'doQuery', 'LANG': 'ADQL', 'FORMAT': 'csv',
            'MAXREC': '10', 'QUERY': adql,
        }
        r = requests.post(GAIA_TAP, data=params, timeout=60)
        r.raise_for_status()
        data = _parse_gaia_csv(r.text)
        return _gaia_rows_to_results(data)

    # Resolve via SIMBAD, then box filter against gaia_source.
    if status_callback:
        status_callback("Resolving name via SIMBAD for Gaia...")
    ra_deg, dec_deg = _resolve_name_simbad(name)
    if ra_deg is None:
        return []

    radius_deg = radius_arcmin / 60.0
    base_cols = ['source_id', 'ra', 'dec', 'phot_g_mean_mag']
    select = ', '.join(base_cols + extra)
    adql = f"""
SELECT {select}
FROM gaiadr3.gaia_source
WHERE ra  BETWEEN {ra_deg  - radius_deg} AND {ra_deg  + radius_deg}
  AND dec BETWEEN {dec_deg - radius_deg} AND {dec_deg + radius_deg}
""".strip()
    params = {
        'REQUEST': 'doQuery', 'LANG': 'ADQL', 'FORMAT': 'csv',
        'MAXREC': '200', 'QUERY': adql,
    }
    r = requests.post(GAIA_TAP, data=params, timeout=60)
    r.raise_for_status()
    data = _parse_gaia_csv(r.text)
    return _gaia_rows_to_results(data, ra_center=ra_deg, dec_center=dec_deg)


# ──────────────────────────────────────────────────────────────
# NASA Exoplanet Archive queries
# ──────────────────────────────────────────────────────────────

NEA_TAP = "https://exoplanetarchive.ipac.caltech.edu/TAP/sync"

_NEA_SELECT = """
pl_name, hostname, discoverymethod, disc_year, disc_facility,
pl_orbper, pl_orbsmax, pl_orbeccen, pl_orbincl, pl_tranmid, pl_imppar,
pl_rade, pl_bmasse, pl_eqt, pl_trandep, pl_trandur, pl_dens,
st_teff, st_logg, st_met, st_rad, st_mass, st_lum, st_age,
sy_snum, sy_pnum, sy_dist, sy_vmag, sy_kmag, sy_gaiamag,
ra, dec
""".strip()


def _parse_nea_csv(text):
    """Parse NEA TAP CSV response. Returns list of dicts keyed by column name."""
    import csv, io
    lines = text.strip().splitlines()
    if not lines:
        return []
    first = lines[0].strip()
    if first.startswith('<'):
        raise RuntimeError(f"NEA: {first[:300]}")
    reader = csv.DictReader(io.StringIO(text))
    return [{k.strip(): (v.strip() if v and v.strip() != '' else None)
             for k, v in row.items()} for row in reader]


def _nea_fmt_row(raw):
    """Normalize a raw NEA CSV dict for display (numbers formatted to reasonable precision)."""
    def _fmt(v, decimals=3):
        if v is None:
            return ''
        try:
            f = float(v)
            return f'{f:.{decimals}f}'
        except (ValueError, TypeError):
            return str(v)

    return {
        'pl_name':        raw.get('pl_name') or '',
        'hostname':       raw.get('hostname') or '',
        'discoverymethod':raw.get('discoverymethod') or '',
        'disc_year':      raw.get('disc_year') or '',
        'disc_facility':  raw.get('disc_facility') or '',
        'pl_orbper':      _fmt(raw.get('pl_orbper'), 5),
        'pl_orbsmax':     _fmt(raw.get('pl_orbsmax'), 5),
        'pl_orbeccen':    _fmt(raw.get('pl_orbeccen'), 4),
        'pl_orbincl':     _fmt(raw.get('pl_orbincl'), 2),
        'pl_tranmid':     _fmt(raw.get('pl_tranmid'), 4),
        'pl_imppar':      _fmt(raw.get('pl_imppar'), 3),
        'pl_rade':        _fmt(raw.get('pl_rade'), 3),
        'pl_bmasse':      _fmt(raw.get('pl_bmasse'), 3),
        'pl_eqt':         _fmt(raw.get('pl_eqt'), 0),
        'pl_trandep':     _fmt(raw.get('pl_trandep'), 6),
        'pl_trandur':     _fmt(raw.get('pl_trandur'), 4),
        'pl_dens':        _fmt(raw.get('pl_dens'), 3),
        'st_teff':        _fmt(raw.get('st_teff'), 0),
        'st_logg':        _fmt(raw.get('st_logg'), 3),
        'st_met':         _fmt(raw.get('st_met'), 3),
        'st_rad':         _fmt(raw.get('st_rad'), 3),
        'st_mass':        _fmt(raw.get('st_mass'), 3),
        'st_lum':         _fmt(raw.get('st_lum'), 3),
        'st_age':         _fmt(raw.get('st_age'), 2),
        'sy_snum':        raw.get('sy_snum') or '',
        'sy_pnum':        raw.get('sy_pnum') or '',
        'sy_dist':        _fmt(raw.get('sy_dist'), 1),
        'sy_vmag':        _fmt(raw.get('sy_vmag'), 2),
        'sy_kmag':        _fmt(raw.get('sy_kmag'), 2),
        'sy_gaiamag':     _fmt(raw.get('sy_gaiamag'), 2),
        'ra':             raw.get('ra') or '',
        'dec':            raw.get('dec') or '',
    }


def _nea_query(where_clause, status_callback):
    """Run a NEA TAP query and return list of formatted planet dicts."""
    adql = f"SELECT {_NEA_SELECT} FROM pscomppars WHERE {where_clause}"
    params = {
        'REQUEST': 'doQuery',
        'LANG':    'ADQL',
        'FORMAT':  'csv',
        'QUERY':   adql,
    }
    r = requests.post(NEA_TAP, data=params, timeout=60)
    r.raise_for_status()
    raw_rows = _parse_nea_csv(r.text)
    return [_nea_fmt_row(row) for row in raw_rows]


def query_nea_by_name(name, status_callback):
    """Search NEA by hostname or planet name."""
    if status_callback:
        status_callback("Querying NASA Exoplanet Archive...")
    safe = name.replace("'", "''")
    where = f"hostname = '{safe}' OR pl_name = '{safe}'"
    results = _nea_query(where, status_callback)
    if not results:
        # Try partial match on hostname
        where2 = f"hostname LIKE '%{safe}%'"
        results = _nea_query(where2, status_callback)
    return results


def query_nea(ra_deg, dec_deg, radius_arcmin, status_callback):
    """Search NEA by coordinate box."""
    if status_callback:
        status_callback("Querying NASA Exoplanet Archive...")
    r = radius_arcmin / 60.0
    where = (f"ra BETWEEN {ra_deg - r} AND {ra_deg + r} "
             f"AND dec BETWEEN {dec_deg - r} AND {dec_deg + r}")
    return _nea_query(where, status_callback)


# ──────────────────────────────────────────────────────────────
# Main App
# ──────────────────────────────────────────────────────────────

BG    = "#1e1e2e"
PANEL = "#2a2a3e"
ACC   = "#7c9fd4"
FG    = "#cdd6f4"
ENT   = "#313244"
SEL   = "#45475a"

BANNER_WORK_BG = "#3a3210"
BANNER_WORK_FG = "#f9e2af"
BANNER_OK_BG   = "#1a3a1a"
BANNER_OK_FG   = "#a6e3a1"
BANNER_ERR_BG  = "#3a1f10"
BANNER_ERR_FG  = "#fab387"
BANNER_NS_BG   = PANEL
BANNER_NS_FG   = "#888aaa"


class StarQueryApp(tk.Tk):

    TABS = [
        ('simbad', 'SIMBAD',    SIMBAD_COLS),
        ('vsx',    'AAVSO VSX', VSX_COLS),
        ('vizier', 'VizieR',    None),
        ('gaia',   'Gaia DR3',  GAIA_COLS),
        ('nea',    'NEA',       None),
    ]

    VIZIER_SUBTABS = [
        ('tmass',  '2MASS',    TMASS_COLS),
        ('wise',   'AllWISE',  WISE_COLS),
        ('apass',  'APASS',    APASS_COLS),
        ('tycho2', 'Tycho-2',  TYCHO2_COLS),
    ]

    NEA_SUBTABS = [
        ('nea_overview', 'Overview', NEA_OVERVIEW_COLS),
        ('nea_orbital',  'Orbital',  NEA_ORBITAL_COLS),
        ('nea_planet',   'Planet',   NEA_PLANET_COLS),
        ('nea_star',     'Star',     NEA_STAR_COLS),
        ('nea_system',   'System',   NEA_SYSTEM_COLS),
    ]

    def __init__(self):
        super().__init__()
        self.title("Stellar Object Query Tool")
        self.geometry("1700x1200")
        self.minsize(1200, 900)
        self.configure(bg=BG)

        _all_keys = (
            [k for k, _, _ in self.TABS if k not in ('vizier', 'nea')] +
            [k for k, _, _ in self.VIZIER_SUBTABS] +
            [k for k, _, _ in self.NEA_SUBTABS]
        )
        self._results      = {k: [] for k in _all_keys}
        self._sort_col     = {k: None for k in _all_keys}
        self._sort_asc     = {k: True for k in _all_keys}
        self._anim_job     = {k: None for k in _all_keys}
        self._anim_idx     = {k: 0    for k in _all_keys}
        self._banner_label = {}
        self._banner_outer = {}
        self._progress_bar = {}
        self._trees        = {}
        self._tab_cols     = {}
        self._notebook     = None
        self._vizier_notebook = None
        self._nea_notebook    = None
        self._cfg          = _load_config()

        self._build_ui()
        self.bind('<Return>', lambda e: self._run_query())

    # ──────────────────────────────────────────────────────────
    # UI construction
    # ──────────────────────────────────────────────────────────

    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use('default')

        style.configure('TFrame',       background=BG)
        style.configure('Panel.TFrame', background=PANEL)

        style.configure('TLabel',  background=BG,    foreground=FG, font=('Segoe UI', 11))
        style.configure('Panel.TLabel', background=PANEL, foreground=FG, font=('Segoe UI', 11))
        style.configure('Header.TLabel', background=PANEL, foreground=ACC,
                        font=('Segoe UI', 11, 'bold'))

        style.configure('TRadiobutton', background=PANEL, foreground=FG,
                        font=('Segoe UI', 11), indicatorcolor=ENT, selectcolor=ACC)
        style.map('TRadiobutton',
                  background=[('active', PANEL)],
                  foreground=[('active', FG)])

        style.configure('TCheckbutton', background=PANEL, foreground=FG,
                        font=('Segoe UI', 11))
        style.map('TCheckbutton',
                  background=[('active', PANEL)],
                  foreground=[('active', FG)])

        style.configure('TEntry', fieldbackground=ENT, foreground=FG,
                        insertcolor=FG, font=('Segoe UI', 11))

        style.configure('TCombobox', fieldbackground=ENT, foreground=FG,
                        background=ENT, selectbackground=SEL, font=('Segoe UI', 11))
        style.map('TCombobox',
                  fieldbackground=[('readonly', ENT)],
                  foreground=[('readonly', FG)])

        style.configure('TSeparator', background=SEL)

        style.configure('TNotebook', background=BG, borderwidth=0)
        style.configure('TNotebook.Tab', background=PANEL, foreground=FG,
                        padding=[12, 6], font=('Segoe UI', 11))
        style.map('TNotebook.Tab',
                  background=[('selected', ACC)],
                  foreground=[('selected', BG)])

        style.configure('Treeview', background=BG, fieldbackground=BG,
                        foreground=FG, font=('Segoe UI', 18),
                        rowheight=44, borderwidth=0)
        style.configure('Treeview.Heading', background=PANEL, foreground=ACC,
                        font=('Segoe UI', 11, 'bold'), relief='flat')
        style.map('Treeview',
                  background=[('selected', SEL)],
                  foreground=[('selected', FG)])
        style.map('Treeview.Heading',
                  background=[('active', SEL)])

        style.configure('Vertical.TScrollbar',   background=PANEL, troughcolor=BG,
                        arrowcolor=FG,  borderwidth=0)
        style.configure('Horizontal.TScrollbar', background=PANEL, troughcolor=BG,
                        arrowcolor=FG,  borderwidth=0)

        style.configure('Accent.TButton', background=ACC, foreground=BG,
                        font=('Segoe UI', 11, 'bold'), padding=[8, 6])
        style.map('Accent.TButton',
                  background=[('active', FG)],
                  foreground=[('active', BG)])

        style.configure('TButton', background=PANEL, foreground=FG,
                        font=('Segoe UI', 11), padding=[8, 6])
        style.map('TButton',
                  background=[('active', SEL)],
                  foreground=[('active', FG)])

        style.configure('TProgressbar', troughcolor=ENT, background=ACC, borderwidth=0)

    def _build_ui(self):
        self._setup_styles()
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # Left panel (fixed ~230px)
        left = tk.Frame(self, bg=PANEL, width=230)
        left.grid(row=0, column=0, sticky='nsew', padx=(8, 0), pady=8)
        left.grid_propagate(False)
        self._build_left_panel(left)

        # Right panel
        right = ttk.Frame(self)
        right.grid(row=0, column=1, sticky='nsew', padx=8, pady=8)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)
        self._build_right_panel(right)

        # Status bar
        self._status_var = tk.StringVar(value="Ready.")
        status_bar = tk.Label(self, textvariable=self._status_var,
                              bg=PANEL, fg=FG, font=('Segoe UI', 10),
                              anchor='w', padx=8)
        status_bar.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 4), padx=8)

    def _build_left_panel(self, parent):
        pad = {'padx': 10, 'pady': 3}

        # Header
        tk.Label(parent, text="SEARCH PARAMETERS", bg=PANEL, fg=ACC,
                 font=('Segoe UI', 11, 'bold')).pack(fill='x', padx=10, pady=(10, 4))

        ttk.Separator(parent, orient='horizontal').pack(fill='x', padx=10, pady=2)

        # Mode radio buttons
        self._mode_var = tk.StringVar(value='name')
        mode_frame = tk.Frame(parent, bg=PANEL)
        mode_frame.pack(fill='x', **pad)
        tk.Radiobutton(mode_frame, text="By Name", variable=self._mode_var, value='name',
                       bg=PANEL, fg=FG, selectcolor=ENT, activebackground=PANEL,
                       activeforeground=FG, font=('Segoe UI', 11),
                       command=self._toggle_mode).pack(side='left', padx=(0, 12))
        tk.Radiobutton(mode_frame, text="By Coordinates", variable=self._mode_var, value='coords',
                       bg=PANEL, fg=FG, selectcolor=ENT, activebackground=PANEL,
                       activeforeground=FG, font=('Segoe UI', 11),
                       command=self._toggle_mode).pack(side='left')

        # Name entry
        tk.Label(parent, text="Star Name:", bg=PANEL, fg=FG,
                 font=('Segoe UI', 10)).pack(fill='x', padx=10, pady=(6, 0))
        self._name_var = tk.StringVar()
        self._name_entry = ttk.Entry(parent, textvariable=self._name_var)
        self._name_entry.pack(fill='x', **pad)
        self._name_hint = tk.Label(parent, text='e.g. "RR Lyr", "V* AB Aur"',
                                   bg=PANEL, fg=BANNER_NS_FG, font=('Segoe UI', 9, 'italic'))
        self._name_hint.pack(fill='x', padx=10, pady=(0, 2))

        # RA / Dec / Radius entries
        tk.Label(parent, text="RA (deg or hms):", bg=PANEL, fg=FG,
                 font=('Segoe UI', 10)).pack(fill='x', padx=10, pady=(6, 0))
        self._ra_var = tk.StringVar()
        self._ra_entry = ttk.Entry(parent, textvariable=self._ra_var)
        self._ra_entry.pack(fill='x', **pad)

        tk.Label(parent, text="Dec (deg or dms):", bg=PANEL, fg=FG,
                 font=('Segoe UI', 10)).pack(fill='x', padx=10, pady=(4, 0))
        self._dec_var = tk.StringVar()
        self._dec_entry = ttk.Entry(parent, textvariable=self._dec_var)
        self._dec_entry.pack(fill='x', **pad)

        tk.Label(parent, text="Search Radius (arcmin):", bg=PANEL, fg=FG,
                 font=('Segoe UI', 10)).pack(fill='x', padx=10, pady=(4, 0))
        self._radius_var = tk.StringVar(value=self._cfg.get('last_radius', '0.1'))
        self._radius_entry = ttk.Entry(parent, textvariable=self._radius_var)
        self._radius_entry.pack(fill='x', **pad)

        ttk.Separator(parent, orient='horizontal').pack(fill='x', padx=10, pady=6)

        # Data sources
        tk.Label(parent, text="DATA SOURCES", bg=PANEL, fg=ACC,
                 font=('Segoe UI', 11, 'bold')).pack(fill='x', **pad)

        self._src_vars = {}
        for key, label, _ in self.TABS:
            var = tk.BooleanVar(value=True)
            self._src_vars[key] = var
            tk.Checkbutton(parent, text=label, variable=var,
                           bg=PANEL, fg=FG, selectcolor=ENT,
                           activebackground=PANEL, activeforeground=FG,
                           anchor='w', font=('Segoe UI', 11)).pack(fill='x', padx=14, pady=1)

        ttk.Separator(parent, orient='horizontal').pack(fill='x', padx=10, pady=6)

        tk.Label(parent, text="GAIA PARAMETERS", bg=PANEL, fg=ACC,
                 font=('Segoe UI', 11, 'bold')).pack(fill='x', **pad)

        self._gaia_param_vars = {}
        for key, label, _ in GAIA_PARAM_OPTS:
            var = tk.BooleanVar(value=True)
            self._gaia_param_vars[key] = var
            tk.Checkbutton(parent, text=label, variable=var,
                           bg=PANEL, fg=FG, selectcolor=ENT,
                           activebackground=PANEL, activeforeground=FG,
                           anchor='w', font=('Segoe UI', 11)).pack(fill='x', padx=14, pady=1)

        ttk.Separator(parent, orient='horizontal').pack(fill='x', padx=10, pady=6)

        # Filters
        tk.Label(parent, text="FILTERS", bg=PANEL, fg=ACC,
                 font=('Segoe UI', 11, 'bold')).pack(fill='x', **pad)

        tk.Label(parent, text="Object Type (SIMBAD only):", bg=PANEL, fg=FG,
                 font=('Segoe UI', 10)).pack(fill='x', padx=10, pady=(4, 0))
        otype_vals = ['All', 'RR*', 'Ce*', 'dS*', 'RV*', 'LP*', 'SR*', 'Mi*', 'Ell*', 'Ro*', 'EB*', 'WV*']
        self._otype_var = tk.StringVar(value='All')
        self._otype_combo = ttk.Combobox(parent, textvariable=self._otype_var,
                                          values=otype_vals, state='readonly')
        self._otype_combo.pack(fill='x', **pad)

        # Period range
        tk.Label(parent, text="Period range (days):", bg=PANEL, fg=FG,
                 font=('Segoe UI', 10)).pack(fill='x', padx=10, pady=(6, 0))
        pf = tk.Frame(parent, bg=PANEL)
        pf.pack(fill='x', padx=10, pady=2)
        self._period_min_var = tk.StringVar()
        self._period_max_var = tk.StringVar()
        ttk.Entry(pf, textvariable=self._period_min_var, width=7).pack(side='left')
        tk.Label(pf, text=' – ', bg=PANEL, fg=FG, font=('Segoe UI', 10)).pack(side='left')
        ttk.Entry(pf, textvariable=self._period_max_var, width=7).pack(side='left')

        # Magnitude range
        tk.Label(parent, text="Max Mag range:", bg=PANEL, fg=FG,
                 font=('Segoe UI', 10)).pack(fill='x', padx=10, pady=(6, 0))
        mf = tk.Frame(parent, bg=PANEL)
        mf.pack(fill='x', padx=10, pady=2)
        self._mag_min_var = tk.StringVar()
        self._mag_max_var = tk.StringVar()
        ttk.Entry(mf, textvariable=self._mag_min_var, width=7).pack(side='left')
        tk.Label(mf, text=' – ', bg=PANEL, fg=FG, font=('Segoe UI', 10)).pack(side='left')
        ttk.Entry(mf, textvariable=self._mag_max_var, width=7).pack(side='left')

        ttk.Separator(parent, orient='horizontal').pack(fill='x', padx=10, pady=8)

        # Buttons
        ttk.Button(parent, text="▶  Run Query", style='Accent.TButton',
                   command=self._run_query).pack(fill='x', padx=10, pady=2)
        ttk.Button(parent, text="⬇  Export to Excel",
                   command=self._export_excel).pack(fill='x', padx=10, pady=2)
        ttk.Button(parent, text="⬇  Export to CSV",
                   command=self._export_csv).pack(fill='x', padx=10, pady=2)
        ttk.Button(parent, text="✕  Clear Results",
                   command=self._clear_all).pack(fill='x', padx=10, pady=(2, 8))

        self._toggle_mode()

    def _build_right_panel(self, parent):
        nb = ttk.Notebook(parent)
        nb.grid(row=0, column=0, sticky='nsew')
        self._notebook = nb
        nb.bind('<<NotebookTabChanged>>', self._on_tab_changed)
        parent.rowconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)   # detail panel gets equal share
        parent.columnconfigure(0, weight=1)

        for key, label, cols in self.TABS:
            if key == 'vizier':
                self._build_vizier_tab(label)
            elif key == 'nea':
                self._build_nea_tab(label)
            else:
                self._build_tab(key, label, cols)

        # Detail panel — takes bottom half
        detail_outer = tk.Frame(parent, bg=PANEL, bd=0)
        detail_outer.grid(row=1, column=0, sticky='nsew', pady=(6, 0))
        detail_outer.columnconfigure(0, weight=1)
        detail_outer.rowconfigure(1, weight=1)

        tk.Label(detail_outer, text="Selected Star Detail", bg=PANEL, fg=ACC,
                 font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=8, pady=(4, 0))

        detail_inner = tk.Frame(detail_outer, bg=ENT)
        detail_inner.grid(row=1, column=0, sticky='nsew', padx=8, pady=(2, 6))
        detail_inner.columnconfigure(0, weight=1)
        detail_inner.rowconfigure(0, weight=1)

        self._detail_text = tk.Text(detail_inner, bg=ENT, fg=FG,
                                     font=('Consolas', 17), relief='flat',
                                     state='disabled', wrap='word', padx=8, pady=4,
                                     insertbackground=FG)
        detail_vsb = ttk.Scrollbar(detail_inner, orient='vertical',
                                    command=self._detail_text.yview)
        self._detail_text.configure(yscrollcommand=detail_vsb.set)
        self._detail_text.grid(row=0, column=0, sticky='nsew')
        detail_vsb.grid(row=0, column=1, sticky='ns')

    def _build_vizier_tab(self, outer_label):
        """Build the VizieR outer tab containing a sub-notebook for each VizieR catalog."""
        tab_frame = ttk.Frame(self._notebook)
        tab_frame.columnconfigure(0, weight=1)
        tab_frame.rowconfigure(0, weight=1)
        self._notebook.add(tab_frame, text=outer_label)

        sub_nb = ttk.Notebook(tab_frame)
        sub_nb.grid(row=0, column=0, sticky='nsew')
        self._vizier_notebook = sub_nb
        sub_nb.bind('<<NotebookTabChanged>>', self._on_vizier_subtab_changed)

        for key, label, cols in self.VIZIER_SUBTABS:
            self._build_tab(key, label, cols, notebook=sub_nb)

    def _build_nea_tab(self, outer_label):
        """Build the NEA outer tab containing a sub-notebook for each parameter category."""
        tab_frame = ttk.Frame(self._notebook)
        tab_frame.columnconfigure(0, weight=1)
        tab_frame.rowconfigure(0, weight=1)
        self._notebook.add(tab_frame, text=outer_label)

        sub_nb = ttk.Notebook(tab_frame)
        sub_nb.grid(row=0, column=0, sticky='nsew')
        self._nea_notebook = sub_nb
        sub_nb.bind('<<NotebookTabChanged>>', self._on_nea_subtab_changed)

        for key, label, cols in self.NEA_SUBTABS:
            self._build_tab(key, label, cols, notebook=sub_nb)

    def _build_tab(self, key, label, cols, notebook=None):
        nb = notebook if notebook is not None else self._notebook
        tab_frame = ttk.Frame(nb)
        tab_frame.columnconfigure(0, weight=1)
        tab_frame.rowconfigure(1, weight=1)
        nb.add(tab_frame, text=label)

        # Banner frame
        banner_outer = tk.Frame(tab_frame, bg=PANEL, height=48)
        banner_outer.grid(row=0, column=0, sticky='ew')
        banner_outer.grid_propagate(False)
        banner_outer.columnconfigure(0, weight=1)
        self._banner_outer[key] = banner_outer

        lbl = tk.Label(banner_outer, text="Ready", bg=PANEL, fg=BANNER_NS_FG,
                       font=('Segoe UI', 13, 'bold'), anchor='w', padx=12)
        lbl.grid(row=0, column=0, sticky='ew', pady=4)
        self._banner_label[key] = lbl

        pb = ttk.Progressbar(banner_outer, mode='indeterminate', length=200)
        # Don't grid it yet — shown only during query
        self._progress_bar[key] = pb

        # Treeview
        tree_frame = ttk.Frame(tab_frame)
        tree_frame.grid(row=1, column=0, sticky='nsew')
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        col_ids = [c[0] for c in cols]
        tree = ttk.Treeview(tree_frame, columns=col_ids, show='headings',
                            selectmode='browse')
        tree.tag_configure('odd',  background='#25253a')
        tree.tag_configure('even', background=BG)

        for cid, cheader, cwidth in cols:
            tree.heading(cid, text=cheader,
                         command=lambda c=cid, k=key: self._sort(k, c))
            tree.column(cid, width=cwidth, minwidth=50, stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient='vertical',   command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal',  command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        tree.bind('<<TreeviewSelect>>', lambda e, k=key: self._on_double_click(k, e))

        self._trees[key] = tree
        self._tab_cols[key] = cols

    # ──────────────────────────────────────────────────────────
    # Banner methods
    # ──────────────────────────────────────────────────────────

    def _stop_spinner(self, key):
        if self._anim_job[key] is not None:
            try:
                self.after_cancel(self._anim_job[key])
            except Exception:
                pass
            self._anim_job[key] = None

    def _start_spinner(self, key, msg):
        self._anim_idx[key] = 0

        def tick():
            frame = SPINNER_FRAMES[self._anim_idx[key] % len(SPINNER_FRAMES)]
            self._anim_idx[key] += 1
            try:
                self._banner_label[key].config(text=f"{frame}  {msg}")
                self._anim_job[key] = self.after(100, tick)
            except Exception:
                pass

        tick()

    def _banner_working(self, key, msg="Querying..."):
        self._stop_spinner(key)
        outer = self._banner_outer[key]
        outer.config(bg=BANNER_WORK_BG)
        lbl = self._banner_label[key]
        lbl.config(bg=BANNER_WORK_BG, fg=BANNER_WORK_FG)

        pb = self._progress_bar[key]
        pb.grid(row=0, column=1, padx=(0, 12), pady=8, sticky='e')
        pb.start(15)

        self._start_spinner(key, msg)

    def _banner_done(self, key, n):
        self._stop_spinner(key)
        pb = self._progress_bar[key]
        pb.stop()
        pb.grid_remove()

        if n > 0:
            outer = self._banner_outer[key]
            outer.config(bg=BANNER_OK_BG)
            lbl = self._banner_label[key]
            lbl.config(bg=BANNER_OK_BG, fg=BANNER_OK_FG,
                       text=f"✓  {n} result{'s' if n != 1 else ''}")
        else:
            outer = self._banner_outer[key]
            outer.config(bg=BANNER_ERR_BG)
            lbl = self._banner_label[key]
            lbl.config(bg=BANNER_ERR_BG, fg=BANNER_ERR_FG, text="✗  No results")

    def _banner_ready(self, key):
        self._stop_spinner(key)
        pb = self._progress_bar[key]
        pb.stop()
        pb.grid_remove()
        outer = self._banner_outer[key]
        outer.config(bg=PANEL)
        lbl = self._banner_label[key]
        lbl.config(bg=PANEL, fg=BANNER_NS_FG, text="Ready")

    def _banner_not_searched(self, key):
        self._stop_spinner(key)
        pb = self._progress_bar[key]
        pb.stop()
        pb.grid_remove()
        outer = self._banner_outer[key]
        outer.config(bg=BANNER_NS_BG)
        lbl = self._banner_label[key]
        lbl.config(bg=BANNER_NS_BG, fg=BANNER_NS_FG, text="Not searched")

    def _banner_error(self, key, msg):
        self._stop_spinner(key)
        pb = self._progress_bar[key]
        pb.stop()
        pb.grid_remove()
        outer = self._banner_outer[key]
        outer.config(bg=BANNER_ERR_BG)
        lbl = self._banner_label[key]
        # Truncate long error messages
        short_msg = msg[:80] + '...' if len(msg) > 80 else msg
        lbl.config(bg=BANNER_ERR_BG, fg=BANNER_ERR_FG, text=f"✗  Error: {short_msg}")

    # ──────────────────────────────────────────────────────────
    # Mode toggle
    # ──────────────────────────────────────────────────────────

    def _toggle_mode(self):
        mode = self._mode_var.get()
        if mode == 'name':
            self._name_entry.config(state='normal')
            self._name_hint.config(fg=BANNER_NS_FG)
            self._ra_entry.config(state='disabled')
            self._dec_entry.config(state='disabled')
            self._radius_entry.config(state='normal')   # always editable
        else:
            self._name_entry.config(state='disabled')
            self._name_hint.config(fg='#555577')
            self._ra_entry.config(state='normal')
            self._dec_entry.config(state='normal')
            self._radius_entry.config(state='normal')

    # ──────────────────────────────────────────────────────────
    # Query execution
    # ──────────────────────────────────────────────────────────

    def _get_filters(self):
        def _f(var):
            v = var.get().strip()
            return _try_float(v)

        return {
            'period_min': _f(self._period_min_var),
            'period_max': _f(self._period_max_var),
            'mag_min':    _f(self._mag_min_var),
            'mag_max':    _f(self._mag_max_var),
            'otype':      self._otype_var.get(),
        }

    def _run_query(self):
        mode   = self._mode_var.get()
        flt    = self._get_filters()

        if mode == 'name':
            name = self._name_var.get().strip()
            if not name:
                messagebox.showwarning("Input Required", "Please enter a star name.")
                return
            ra_deg = dec_deg = None
            try:
                radius = float(self._radius_var.get().strip())
            except (ValueError, TypeError):
                radius = 1.0   # default 1 arcmin for name searches
        else:
            ra_str     = self._ra_var.get().strip()
            dec_str    = self._dec_str = self._dec_var.get().strip()
            radius_str = self._radius_var.get().strip()
            if not ra_str or not dec_str:
                messagebox.showwarning("Input Required", "Please enter RA and Dec.")
                return
            ra_deg = parse_ra(ra_str)
            dec_deg = parse_dec(dec_str)
            if ra_deg is None or dec_deg is None:
                messagebox.showerror("Parse Error", "Could not parse RA or Dec. Please check format.")
                return
            try:
                radius = float(radius_str)
            except ValueError:
                radius = 30.0
            _save_config({'last_radius': radius_str})
            name = None

        self._set_status("Running query...")

        # Collect selected Gaia extra columns
        gaia_extra = []
        for opt_key, _, col_list in GAIA_PARAM_OPTS:
            if self._gaia_param_vars[opt_key].get():
                gaia_extra.extend(col_list)

        for key, label, cols in self.TABS:
            if not self._src_vars[key].get():
                if key == 'vizier':
                    for vk, vlabel, _ in self.VIZIER_SUBTABS:
                        self._clear_tab(vk)
                        self._banner_not_searched(vk)
                        self._update_tab_text(vk, vlabel, None)
                elif key == 'nea':
                    for nk, nlabel, _ in self.NEA_SUBTABS:
                        self._clear_tab(nk)
                        self._banner_not_searched(nk)
                        self._update_tab_text(nk, nlabel, None)
                else:
                    self._clear_tab(key)
                    self._banner_not_searched(key)
                    self._update_tab_text(key, label, None)
                continue

            if key == 'vizier':
                for vk, vlabel, _ in self.VIZIER_SUBTABS:
                    self._clear_tab(vk, silent=True)
                    self._banner_working(vk)
                    self._update_tab_text(vk, vlabel, None)
                    thread = threading.Thread(
                        target=self._query_worker,
                        args=(vk, vlabel, mode, name, ra_deg, dec_deg, radius, flt, gaia_extra),
                        daemon=True,
                    )
                    thread.start()
            elif key == 'nea':
                for nk, nlabel, _ in self.NEA_SUBTABS:
                    self._clear_tab(nk, silent=True)
                    self._banner_working(nk)
                    self._update_tab_text(nk, nlabel, None)
                thread = threading.Thread(
                    target=self._query_worker,
                    args=(key, label, mode, name, ra_deg, dec_deg, radius, flt, gaia_extra),
                    daemon=True,
                )
                thread.start()
            else:
                self._clear_tab(key, silent=True)
                self._banner_working(key)
                self._update_tab_text(key, label, None)
                thread = threading.Thread(
                    target=self._query_worker,
                    args=(key, label, mode, name, ra_deg, dec_deg, radius, flt, gaia_extra),
                    daemon=True,
                )
                thread.start()

    def _query_worker(self, key, label, mode, name, ra_deg, dec_deg, radius, flt, gaia_extra=None):
        pm  = flt['period_min']
        pM  = flt['period_max']
        mm  = flt['mag_min']
        mM  = flt['mag_max']
        ot  = flt['otype']

        def status_cb(msg):
            self.after(0, self._set_status, msg)

        try:
            if key == 'simbad':
                if mode == 'name':
                    results = query_simbad_by_name(name, pm, pM, mm, mM, status_cb, otype_filter=ot)
                else:
                    results = query_simbad(ra_deg, dec_deg, radius, ot, pm, pM, mm, mM, status_cb)

            elif key == 'vsx':
                if mode == 'name':
                    results = query_vsx_by_name(name, pm, pM, mm, mM, status_cb)
                else:
                    results = query_vsx(ra_deg, dec_deg, radius, pm, pM, mm, mM, status_cb)

            elif key == 'tmass':
                if mode == 'name':
                    results = query_tmass_by_name(name, radius, status_cb)
                else:
                    results = query_tmass(ra_deg, dec_deg, radius, status_cb)

            elif key == 'wise':
                if mode == 'name':
                    results = query_wise_by_name(name, radius, status_cb)
                else:
                    results = query_wise(ra_deg, dec_deg, radius, status_cb)

            elif key == 'apass':
                if mode == 'name':
                    results = query_apass_by_name(name, radius, status_cb)
                else:
                    results = query_apass(ra_deg, dec_deg, radius, status_cb)

            elif key == 'tycho2':
                if mode == 'name':
                    results = query_tycho2_by_name(name, radius, status_cb)
                else:
                    results = query_tycho2(ra_deg, dec_deg, radius, status_cb)

            elif key == 'gaia':
                if mode == 'name':
                    results = query_gaia_by_name(name, radius, status_cb, extra_cols=gaia_extra)
                else:
                    results = query_gaia(ra_deg, dec_deg, radius, status_cb, extra_cols=gaia_extra)

            elif key == 'nea':
                if mode == 'name':
                    results = query_nea_by_name(name, status_cb)
                else:
                    results = query_nea(ra_deg, dec_deg, radius, status_cb)
                self.after(0, self._populate_nea, results)
                for nk, nlabel, _ in self.NEA_SUBTABS:
                    self.after(0, self._update_tab_text, nk, nlabel, len(results))
                return

            else:
                results = []

            self.after(0, self._populate, key, results)
            self.after(0, self._update_tab_text, key, label, len(results))

        except Exception as e:
            if key == 'nea':
                for nk, nlabel, _ in self.NEA_SUBTABS:
                    self.after(0, self._banner_error, nk, str(e))
                    self.after(0, self._update_tab_text, nk, nlabel, 0)
            else:
                self.after(0, self._banner_error, key, str(e))
                self.after(0, self._update_tab_text, key, label, 0)
            self.after(0, self._set_status, f"Error querying {label}: {e}")

    # ──────────────────────────────────────────────────────────
    # Populate / sort / clear
    # ──────────────────────────────────────────────────────────

    def _populate(self, key, results):
        self._results[key] = results
        self._refresh_tree(key)
        self._banner_done(key, len(results))
        self._set_status(f"{key.upper()}: {len(results)} result(s) loaded.")
        if len(results) == 1:
            tree = self._trees[key]
            children = tree.get_children()
            if children:
                tree.selection_set(children[0])
                tree.focus(children[0])
                try:
                    active = self._active_tab_key()
                except Exception:
                    active = None
                if active == key:
                    self._show_row_detail(key, 0)

    def _populate_nea(self, results):
        """Populate all NEA sub-tab treeviews with the shared results list."""
        n = len(results)
        for key, label, _ in self.NEA_SUBTABS:
            self._results[key] = results
            self._refresh_tree(key)
            self._banner_done(key, n)
        self._set_status(f"NEA: {n} planet(s) loaded.")
        if n == 1:
            try:
                active = self._active_tab_key()
            except Exception:
                active = None
            for key, _, _ in self.NEA_SUBTABS:
                tree = self._trees[key]
                children = tree.get_children()
                if children:
                    tree.selection_set(children[0])
                    tree.focus(children[0])
            if active in {k for k, _, _ in self.NEA_SUBTABS}:
                self._show_row_detail(active, 0)

    def _refresh_tree(self, key):
        tree = self._trees[key]
        cols = self._tab_cols[key]
        col_ids = [c[0] for c in cols]

        tree.delete(*tree.get_children())

        results = self._results[key]
        for i, row in enumerate(results):
            values = [row.get(cid, '') for cid in col_ids]
            tag = 'odd' if i % 2 == 1 else 'even'
            tree.insert('', 'end', iid=str(i), values=values, tags=(tag,))

    def _clear_tab(self, key, silent=False):
        tree = self._trees[key]
        tree.delete(*tree.get_children())
        self._results[key] = []
        self._sort_col[key] = None
        self._sort_asc[key] = True
        # Reset headings (remove sort arrows)
        for cid, cheader, _ in self._tab_cols[key]:
            self._trees[key].heading(cid, text=cheader)
        # Clear detail
        self._set_detail('')
        if not silent:
            self._banner_ready(key)

    def _clear_all(self):
        for key, label, cols in self.TABS:
            if key == 'vizier':
                for vk, vlabel, _ in self.VIZIER_SUBTABS:
                    self._clear_tab(vk)
                    self._update_tab_text(vk, vlabel, None)
            elif key == 'nea':
                for nk, nlabel, _ in self.NEA_SUBTABS:
                    self._clear_tab(nk)
                    self._update_tab_text(nk, nlabel, None)
            else:
                self._clear_tab(key)
                self._update_tab_text(key, label, None)
        self._set_status("Results cleared.")

    def _sort(self, key, col):
        if self._sort_col[key] == col:
            self._sort_asc[key] = not self._sort_asc[key]
        else:
            self._sort_col[key] = col
            self._sort_asc[key] = True

        asc = self._sort_asc[key]
        is_numeric = col in NUMERIC_COLS

        def sort_key(row):
            val = row.get(col, '')
            if is_numeric:
                try:
                    return (0, float(val))
                except (ValueError, TypeError):
                    return (1, 0.0)  # empty/invalid sorts last
            else:
                return (0 if val else 1, str(val).lower())

        self._results[key].sort(key=sort_key, reverse=not asc)
        self._refresh_tree(key)

        # Update headings to show sort arrow
        for cid, cheader, _ in self._tab_cols[key]:
            if cid == col:
                arrow = '▲' if asc else '▼'
                self._trees[key].heading(cid, text=f"{cheader} {arrow}")
            else:
                self._trees[key].heading(cid, text=cheader)

    def _update_tab_text(self, key, base_label, n):
        vizier_keys = {k for k, _, _ in self.VIZIER_SUBTABS}
        nea_keys    = {k for k, _, _ in self.NEA_SUBTABS}

        if key in vizier_keys:
            sub_nb = self._vizier_notebook
            for i, (k, lbl, _) in enumerate(self.VIZIER_SUBTABS):
                if k == key:
                    sub_nb.tab(i, text=lbl if n is None else f"{lbl} ({n})")
                    break
            total = sum(len(self._results[k]) for k, _, _ in self.VIZIER_SUBTABS)
            nb = self._notebook
            for i, (k, lbl, _) in enumerate(self.TABS):
                if k == 'vizier':
                    nb.tab(i, text=f"VizieR ({total})")
                    break

        elif key in nea_keys:
            sub_nb = self._nea_notebook
            for i, (k, lbl, _) in enumerate(self.NEA_SUBTABS):
                if k == key:
                    sub_nb.tab(i, text=lbl if n is None else f"{lbl} ({n})")
                    break
            # All NEA sub-tabs have the same count — use this one to update the outer tab
            nb = self._notebook
            for i, (k, lbl, _) in enumerate(self.TABS):
                if k == 'nea':
                    nb.tab(i, text='NEA' if n is None else f"NEA ({n})")
                    break

        else:
            nb = self._notebook
            for i, (k, lbl, _) in enumerate(self.TABS):
                if k == key:
                    nb.tab(i, text=base_label if n is None else f"{base_label} ({n})")
                    break

    def _active_tab_key(self):
        nb = self._notebook
        idx = nb.index(nb.select())
        key = self.TABS[idx][0]
        if key == 'vizier' and self._vizier_notebook is not None:
            try:
                sub_idx = self._vizier_notebook.index(self._vizier_notebook.select())
                return self.VIZIER_SUBTABS[sub_idx][0]
            except Exception:
                return self.VIZIER_SUBTABS[0][0]
        if key == 'nea' and self._nea_notebook is not None:
            try:
                sub_idx = self._nea_notebook.index(self._nea_notebook.select())
                return self.NEA_SUBTABS[sub_idx][0]
            except Exception:
                return self.NEA_SUBTABS[0][0]
        return key

    # ──────────────────────────────────────────────────────────
    # Detail panel
    # ──────────────────────────────────────────────────────────

    def _set_detail(self, text):
        self._detail_text.config(state='normal')
        self._detail_text.delete('1.0', 'end')
        if text:
            self._detail_text.insert('1.0', text)
        self._detail_text.config(state='disabled')

    def _show_row_detail(self, key, idx):
        if idx >= len(self._results[key]):
            return
        row = self._results[key][idx]
        basic_text = self._format_detail(key, row)
        if key == 'simbad':
            cached = row.get('_extended_detail')
            if cached:
                self._set_detail(cached)
            else:
                self._set_detail(basic_text + "\n\n  Loading extended data…")
                main_id = row.get('Name', '')
                if main_id:
                    threading.Thread(
                        target=self._fetch_simbad_detail,
                        args=(main_id, basic_text, row),
                        daemon=True,
                    ).start()
        else:
            self._set_detail(basic_text)

    def _on_double_click(self, key, event):
        """Called on <<TreeviewSelect>> (single click)."""
        try:
            if self._active_tab_key() != key:
                return   # selection change on a background tab — ignore
        except Exception:
            pass
        tree = self._trees[key]
        sel  = tree.selection()
        if not sel:
            return
        try:
            idx = int(sel[0])
        except ValueError:
            return
        self._show_row_detail(key, idx)

    def _on_vizier_subtab_changed(self, event=None):
        """Update detail panel when the user switches VizieR sub-tabs."""
        try:
            sub_idx = self._vizier_notebook.index(self._vizier_notebook.select())
            key = self.VIZIER_SUBTABS[sub_idx][0]
        except Exception:
            return
        tree = self._trees[key]
        sel  = tree.selection()
        if sel:
            try:
                self._show_row_detail(key, int(sel[0]))
            except (ValueError, IndexError):
                self._set_detail('')
        elif len(self._results[key]) == 1:
            children = tree.get_children()
            if children:
                tree.selection_set(children[0])
                self._show_row_detail(key, 0)
        else:
            self._set_detail('')

    def _on_nea_subtab_changed(self, event=None):
        """Update detail panel when the user switches NEA sub-tabs."""
        try:
            sub_idx = self._nea_notebook.index(self._nea_notebook.select())
            key = self.NEA_SUBTABS[sub_idx][0]
        except Exception:
            return
        tree = self._trees[key]
        sel  = tree.selection()
        if sel:
            try:
                self._show_row_detail(key, int(sel[0]))
            except (ValueError, IndexError):
                self._set_detail('')
        elif len(self._results[key]) == 1:
            children = tree.get_children()
            if children:
                tree.selection_set(children[0])
                self._show_row_detail(key, 0)
        else:
            self._set_detail('')

    def _on_tab_changed(self, event=None):
        """Update detail panel when the user switches tabs."""
        try:
            nb  = self._notebook
            idx = nb.index(nb.select())
            key = self.TABS[idx][0]
        except Exception:
            return
        if key == 'vizier':
            self._on_vizier_subtab_changed()
            return
        if key == 'nea':
            self._on_nea_subtab_changed()
            return
        tree = self._trees[key]
        sel  = tree.selection()
        if sel:
            try:
                self._show_row_detail(key, int(sel[0]))
            except (ValueError, IndexError):
                self._set_detail('')
        elif len(self._results[key]) == 1:
            children = tree.get_children()
            if children:
                tree.selection_set(children[0])
                self._show_row_detail(key, 0)
        else:
            self._set_detail('')

    def _fetch_simbad_detail(self, main_id, basic_text, result_row=None):
        """Background: fetch extended SIMBAD data then update detail panel."""
        safe_id = main_id.replace("'", "''")
        params_base = {'REQUEST': 'doQuery', 'LANG': 'ADQL', 'FORMAT': 'json'}

        def _q(adql):
            try:
                r = requests.post(SIMBAD_TAP,
                                  data={**params_base, 'QUERY': adql}, timeout=30)
                r.raise_for_status()
                return r.json().get('data', [])
            except Exception:
                return []

        # Extended basic fields
        adql_ext = f"""
SELECT b.sp_type, b.plx_value, b.plx_err, b.pmra, b.pmdec
FROM basic b
WHERE b.main_id = '{safe_id}'
""".strip()

        # All identifiers (no ORDER BY — some ADQL engines reject it here)
        adql_ident = f"""
SELECT i.id FROM ident i
JOIN basic b ON i.oidref = b.oid
WHERE b.main_id = '{safe_id}'
""".strip()

        # References (distinct bibcodes from mesVar, up to 5)
        adql_refs = f"""
SELECT DISTINCT v.bibcode FROM mesVar v
JOIN basic b ON v.oidref = b.oid
WHERE b.main_id = '{safe_id}'
""".strip()

        ext_data   = _q(adql_ext)
        ident_data = _q(adql_ident)
        refs_data  = _q(adql_refs)

        lines = [basic_text, '']

        # Spectral type, parallax, proper motion
        if ext_data:
            row = ext_data[0]
            sp_type  = str(row[0]).strip() if row[0] else '—'
            plx_val  = row[1]
            plx_err  = row[2]
            pmra_v   = row[3]
            pmdec_v  = row[4]

            plx_str = '—'
            if plx_val is not None:
                try:
                    pv = float(plx_val)
                    pe = float(plx_err) if plx_err is not None else None
                    dist_pc = 1000.0 / pv if pv > 0 else None
                    plx_str = f"{pv:.3f}"
                    if pe is not None:
                        plx_str += f" ± {pe:.3f} mas"
                    else:
                        plx_str += " mas"
                    if dist_pc is not None:
                        plx_str += f"  (≈ {dist_pc:.0f} pc)"
                except (ValueError, TypeError):
                    pass

            pm_str = '—'
            if pmra_v is not None and pmdec_v is not None:
                try:
                    pm_str = f"RA {float(pmra_v):+.2f}  Dec {float(pmdec_v):+.2f}  mas/yr"
                except (ValueError, TypeError):
                    pass

            lines.append(f"Spectral Type:  {sp_type}")
            lines.append(f"Parallax:       {plx_str}")
            lines.append(f"Proper Motion:  {pm_str}")

        # Identifiers
        if ident_data:
            ids = [str(r[0]).strip() for r in ident_data if r[0]]
            ids = [i for i in ids if i]
            if ids:
                lines.append('')
                lines.append('Other Names:')
                # wrap into lines of ~3 per line
                for i in range(0, len(ids), 3):
                    lines.append('  ' + '   |   '.join(ids[i:i+3]))

        # References
        if refs_data:
            bibcodes = [str(r[0]).strip() for r in refs_data if r[0]][:5]
            if bibcodes:
                lines.append('')
                lines.append('References (≤5):')
                for b in bibcodes:
                    lines.append(f'  {b}')

        full_text = '\n'.join(lines)
        if result_row is not None:
            result_row['_extended_detail'] = full_text
        self.after(0, self._set_detail, full_text)

    def _dash(self, val):
        return val if val else '—'

    def _format_detail(self, key, row):
        d = self._dash
        if key == 'simbad':
            return (
                f"Name: {d(row.get('Name'))}    RA: {d(row.get('RA_hms'))}    "
                f"Dec: {d(row.get('Dec_dms'))}    Type: {d(row.get('OType_label'))}    "
                f"Var Sub-type: {d(row.get('VarType'))}\n"
                f"Period: {d(row.get('Period'))} d    "
                f"Max Mag: {d(row.get('MaxMag'))}    Min Mag: {d(row.get('MinMag'))}    "
                f"Band: {d(row.get('MagBand'))}    "
                f"Dist: {d(row.get('Dist_arcsec'))}\"    N refs: {d(row.get('N_refs'))}"
            )
        elif key == 'vsx':
            return (
                f"Name: {d(row.get('Name'))}    AUID: {d(row.get('AUID'))}    "
                f"RA: {d(row.get('RA_hms'))}    Dec: {d(row.get('Dec_dms'))}\n"
                f"Var Type: {d(row.get('VarType'))}    Period: {d(row.get('Period'))} d    "
                f"Max Mag: {d(row.get('MaxMag'))}    Min Mag: {d(row.get('MinMag'))}    "
                f"Band: {d(row.get('MagBand'))}    Dist: {d(row.get('Dist_arcsec'))}\""
            )
        elif key == 'tmass':
            return (
                f"Name: {d(row.get('Name'))}    RA: {d(row.get('RA_hms'))}    "
                f"Dec: {d(row.get('Dec_dms'))}\n"
                f"J: {d(row.get('Jmag'))}    H: {d(row.get('Hmag'))}    "
                f"K: {d(row.get('Kmag'))}    Dist: {d(row.get('Dist_arcsec'))}\""
            )
        elif key == 'wise':
            return (
                f"Name: {d(row.get('Name'))}    RA: {d(row.get('RA_hms'))}    "
                f"Dec: {d(row.get('Dec_dms'))}\n"
                f"W1: {d(row.get('W1mag'))}    W2: {d(row.get('W2mag'))}    "
                f"W3: {d(row.get('W3mag'))}    W4: {d(row.get('W4mag'))}    "
                f"Dist: {d(row.get('Dist_arcsec'))}\""
            )
        elif key == 'apass':
            return (
                f"Name: {d(row.get('Name'))}    RA: {d(row.get('RA_hms'))}    "
                f"Dec: {d(row.get('Dec_dms'))}\n"
                f"V: {d(row.get('Vmag'))}    B: {d(row.get('Bmag'))}    "
                f"Dist: {d(row.get('Dist_arcsec'))}\""
            )
        elif key == 'tycho2':
            return (
                f"Name: {d(row.get('Name'))}    RA: {d(row.get('RA_hms'))}    "
                f"Dec: {d(row.get('Dec_dms'))}\n"
                f"BT: {d(row.get('BTmag'))}    VT: {d(row.get('VTmag'))}    "
                f"PM RA: {d(row.get('pmRA'))} mas/yr    PM Dec: {d(row.get('pmDE'))} mas/yr    "
                f"Dist: {d(row.get('Dist_arcsec'))}\""
            )
        elif key == 'gaia':
            lines = [
                f"Source ID: {d(row.get('source_id'))}    RA: {d(row.get('RA_hms'))}    "
                f"Dec: {d(row.get('Dec_dms'))}    Dist: {d(row.get('Dist_arcsec'))}\"",
                f"G Mag: {d(row.get('GMag'))}"
                + (f"    BP Mag: {d(row.get('BPMag'))}" if row.get('BPMag') else '')
                + (f"    RP Mag: {d(row.get('RPMag'))}" if row.get('RPMag') else ''),
            ]
            if row.get('Parallax'):
                lines.append(f"Parallax: {d(row.get('Parallax'))} ± {d(row.get('PlxErr'))} mas")
            if row.get('PMRA') or row.get('PMDec'):
                lines.append(
                    f"Proper Motion:  RA: {d(row.get('PMRA'))} ± {d(row.get('PMRAErr'))} mas/yr"
                    f"    Dec: {d(row.get('PMDec'))} ± {d(row.get('PMDecErr'))} mas/yr"
                )
            if row.get('RV'):
                lines.append(f"Radial Velocity: {d(row.get('RV'))} ± {d(row.get('RVErr'))} km/s")
            if row.get('RUWE'):
                lines.append(f"RUWE: {d(row.get('RUWE'))}")
            return '\n'.join(lines)
        elif key in {'nea_overview', 'nea_orbital', 'nea_planet', 'nea_star', 'nea_system'}:
            return self._format_detail_nea(row)
        return ''

    def _format_detail_nea(self, row):
        d = self._dash
        def _v(k, unit='', fmt=None):
            v = row.get(k)
            if not v:
                return '\u2014'
            try:
                fv = float(v)
                return (format(fv, fmt) if fmt else str(v)) + (f' {unit}' if unit else '')
            except (ValueError, TypeError):
                return str(v) + (f' {unit}' if unit else '')

        deg_sym  = '\u00b0'
        gcm3_sym = 'g/cm\u00b3'
        rsun_sym = 'R\u2609'
        msun_sym = 'M\u2609'
        lsun_sym = 'L\u2609'
        hr_orb   = '\u2500\u2500\u2500 ORBITAL PARAMETERS \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500'
        hr_pl    = '\u2500\u2500\u2500 PLANET PARAMETERS \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500'
        hr_st    = '\u2500\u2500\u2500 STELLAR PARAMETERS \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500'
        hr_sys   = '\u2500\u2500\u2500 SYSTEM \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500'
        lines = [
            f"Planet: {d(row.get('pl_name'))}    Host: {d(row.get('hostname'))}",
            f"Discovery: {d(row.get('discoverymethod'))} | {d(row.get('disc_year'))} | {d(row.get('disc_facility'))}",
            "",
            hr_orb,
            f"Period:       {_v('pl_orbper', 'd')}    T0 (BJD):  {_v('pl_tranmid')}",
            f"Semi-major:   {_v('pl_orbsmax', 'AU')}    Incl:      {_v('pl_orbincl', deg_sym)}",
            f"Eccentricity: {_v('pl_orbeccen')}    Impact:    {_v('pl_imppar')}",
            "",
            hr_pl,
            f"Radius:  {_v('pl_rade', 'RE')}    Mass:    {_v('pl_bmasse', 'ME')}",
            f"Teq:     {_v('pl_eqt', 'K')}    Density: {_v('pl_dens', gcm3_sym)}",
            f"Trans Depth: {_v('pl_trandep')}    Duration: {_v('pl_trandur', 'h')}",
            "",
            hr_st,
            f"Teff:  {_v('st_teff', 'K')}    log g: {_v('st_logg')}    [Fe/H]: {_v('st_met')}",
            f"R*:    {_v('st_rad', rsun_sym)}    M*:    {_v('st_mass', msun_sym)}    L*: {_v('st_lum', lsun_sym)}",
            f"Age:   {_v('st_age', 'Gyr')}",
            "",
            hr_sys,
            f"Distance: {_v('sy_dist', 'pc')}    # Stars: {_v('sy_snum')}    # Planets: {_v('sy_pnum')}",
            f"V mag: {_v('sy_vmag')}    K mag: {_v('sy_kmag')}    Gaia mag: {_v('sy_gaiamag')}",
        ]
        return '\n'.join(lines)

    # ──────────────────────────────────────────────────────────
    # Export
    # ──────────────────────────────────────────────────────────

    def _export_excel(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("Missing Library", "openpyxl is not installed.\nRun: pip install openpyxl")
            return

        key = self._active_tab_key()
        results = self._results[key]
        if not results:
            messagebox.showinfo("No Results", "No results to export for the active tab.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx'), ('All files', '*.*')],
            title="Export to Excel",
        )
        if not path:
            return

        cols = self._tab_cols[key]
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = key.upper()
        ws.append([cheader for _, cheader, _ in cols])
        for row in results:
            ws.append([row.get(cid, '') for cid, _, _ in cols])
        wb.save(path)
        self._set_status(f"Exported {len(results)} rows to {os.path.basename(path)}")

    def _export_csv(self):
        import csv

        key = self._active_tab_key()
        results = self._results[key]
        if not results:
            messagebox.showinfo("No Results", "No results to export for the active tab.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')],
            title="Export to CSV",
        )
        if not path:
            return

        cols = self._tab_cols[key]
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow([cheader for _, cheader, _ in cols])
            for row in results:
                writer.writerow([row.get(cid, '') for cid, _, _ in cols])

        self._set_status(f"Exported {len(results)} rows to {os.path.basename(path)}")

    # ──────────────────────────────────────────────────────────
    # Status bar
    # ──────────────────────────────────────────────────────────

    def _set_status(self, msg):
        self._status_var.set(msg)


# ──────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app = StarQueryApp()
    app.mainloop()
