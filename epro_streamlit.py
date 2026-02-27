"""
ePRO Chart Generator v9.2 — Streamlit Edition (Config-Aware)
============================================================
Run with:  streamlit run epro_streamlit_v9.py
Requires:  pip install streamlit pandas numpy matplotlib openpyxl
"""
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyBboxPatch, Rectangle
from pathlib import Path
from collections import defaultdict
from io import BytesIO
import re
import textwrap
import tempfile
import os
import json

# ═══════════════════════════════════════════════════════════════
# 1. CONFIGURATION & STYLE
# ═══════════════════════════════════════════════════════════════

FAVORABLE_THRESHOLD = 0.70
CENTER_FILTER = "VCS"

METADATA_COLS = {"STUDY REFERENCE", "SUBJECT ID", "STATUS", "STUDY CENTER ABBREV",
                 "RANDOMISATION ID", "START DATE", "LAST UPDATE DATE", "OCCURRENCE NO"}

COL_SUBJECT = "SUBJECT ID"
COL_STATUS = "STATUS"
COL_CENTER = "STUDY CENTER ABBREV"
COL_STUDY = "STUDY REFERENCE"

POSITIVE_KEYWORDS = [
    "YES", "AGREE", "EXCELLENT", "GOOD", "BETTER", "HIGH", "HIGHER", "ABOVE",
    "ALWAYS", "OFTEN", "FREQUENTLY", "IMPORTANT", "SATISFIED", "LIKE", "LIKE ME",
    "TRUE", "DEFINITELY", "WILL", "VERY", "EXTREMELY", "MUCH", "ALMOST ALWAYS",
    "EXCEPTIONAL", "FAIRLY", "MET EXPECTATIONS", "PROBABLY"
]
NEGATIVE_KEYWORDS = [
    "NO", "NOT", "DONT", "WONT", "DISAGREE", "NEVER", "POOR", "WORSE", "LOW",
    "LOWER", "BELOW", "RARELY", "SELDOM", "UNFAIR", "DISSATISFIED"
]

COLORS = {
    'favorable': '#0173B2',    'unfavorable': '#DE8F05',  'excellent': '#029E73',
    'warning': '#D55E00',      'moderate': '#CC78BC',     'neutral': '#949494',
    'bg_card': '#F8F9FA',      'text_main': '#333333',    'text_sub': '#666666',
    'scale_colors': ['#0173B2', '#56B4E9', '#029E73', '#F0E442',
                     '#DE8F05', '#D55E00', '#CC78BC', '#949494']
}
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Calibri', 'DejaVu Sans', 'Arial']
plt.rcParams['font.size'] = 10
plt.rcParams['axes.spines.top'] = False
plt.rcParams['axes.spines.right'] = False
plt.rcParams['axes.spines.left'] = False
plt.rcParams['axes.grid'] = True
plt.rcParams['grid.alpha'] = 0.3


# ═══════════════════════════════════════════════════════════════
# 2. DATA STRUCTURES
# ═══════════════════════════════════════════════════════════════

class QuestionInfo:
    def __init__(self, var_name, question_text, col_index, levels=None):
        self.var_name = var_name
        self.question_text = question_text
        self.col_index = col_index
        self.levels = levels or {}
        self.is_scaled = len(self.levels) > 0
        self.is_open_ended = not self.is_scaled
        self.is_multi_select = False
        self.fav_mask = []
        self.neutral_mask = None
        self.q_number = self._extract_q_number()

    def _extract_q_number(self):
        name = self.var_name
        m = re.search(r'_Q(\d+\w*)', name, re.IGNORECASE)
        if m: return m.group(1).lstrip('0') or '0'
        m = re.search(r'Q(\d+\w*)', name, re.IGNORECASE)
        if m: return m.group(1).lstrip('0') or '0'
        return name

    @property
    def scale_signature(self):
        if not self.is_scaled: return "OPEN"
        return tuple(self.levels[k] for k in sorted(self.levels.keys()))


class TimepointData:
    def __init__(self, name, sheet_name, file_path):
        self.name = name
        self.sheet_name = sheet_name
        self.file_path = file_path
        self.questions = []
        self.subject_data = None
        self.included_subjects = []
        self.non_completed = []
        self.dropped_ids = []
        self.deviation_subjects = {}
        self.study_refs = []
        self.n_total = 0
        self.n_completed = 0
        self.n_included = 0
        self.randomization_groups = {}
        self.needs_unrandomization = False
        self.randomization_source = ""


# ═══════════════════════════════════════════════════════════════
# 3. CONFIG IMPORT — Load VBA-exported JSON
# ═══════════════════════════════════════════════════════════════

def load_config_from_json(json_bytes):
    try:
        json_str = json_bytes.decode('utf-8')
    except UnicodeDecodeError:
        json_str = json_bytes.decode('latin-1')

    config = json.loads(json_str)
    settings = config.get("settings", {})
    tp_configs = config.get("timepoints", [])

    timepoints = []
    for tp_cfg in tp_configs:
        tp = TimepointData(
            name=tp_cfg["name"],
            sheet_name=tp_cfg["sheet_name"],
            file_path=tp_cfg.get("file_path", ""),
        )
        tp.study_refs = tp_cfg.get("study_refs", [])
        tp.n_total = int(tp_cfg.get("total_rows", 0))
        tp.n_completed = int(tp_cfg.get("completed_rows", 0))
        tp.dropped_ids = tp_cfg.get("excluded_subjects", [])
        tp.deviation_subjects = tp_cfg.get("deviation_subjects", {})
        tp.needs_unrandomization = tp_cfg.get("needs_unrandomization", False)
        tp.randomization_source = tp_cfg.get("randomization_source", "")
        tp.randomization_groups = tp_cfg.get("randomization_groups", {})

        for q_cfg in tp_cfg.get("questions", []):
            levels = {}
            for k, v in q_cfg.get("levels", {}).items():
                try:
                    levels[int(k)] = v
                except ValueError:
                    pass

            qi = QuestionInfo(
                var_name=q_cfg["var_name"],
                question_text=q_cfg.get("question_text", ""),
                col_index=int(q_cfg.get("col_index", 0)),
                levels=levels,
            )
            qi.is_multi_select = q_cfg.get("is_multi_select", False)

            fav_raw = q_cfg.get("fav_mask", "(none)")
            if fav_raw in ("(none)", "(multi-select)", ""):
                qi.fav_mask = []
            else:
                qi.fav_mask = [int(x.strip()) for x in fav_raw.split(",") if x.strip().isdigit()]

            neutral_raw = q_cfg.get("neutral_mask", "(none)")
            if neutral_raw not in ("(none)", ""):
                try:
                    qi.neutral_mask = int(neutral_raw)
                except ValueError:
                    qi.neutral_mask = None

            tp.questions.append(qi)
        timepoints.append(tp)

    return config, settings, timepoints


def compute_n_included_from_config(tp, settings):
    completed_only = settings.get("completed_only", False)
    n = tp.n_completed if completed_only else tp.n_total
    n -= len(tp.dropped_ids)
    for sid, reason in tp.deviation_subjects.items():
        if "No Data" in reason or reason.startswith("[EXCLUDE_N]"):
            n -= 1
    return max(n, 0)


# ═══════════════════════════════════════════════════════════════
# 4. SHEET DETECTION & CLASSIFICATION
# ═══════════════════════════════════════════════════════════════

def classify_sheet(excel_path, sheet_name):
    try:
        df_peek = pd.read_excel(excel_path, sheet_name=sheet_name, nrows=8, header=None)
    except Exception:
        return {'type': 'unknown', 'headers': [], 'preview_rows': [], 'n_cols': 0, 'n_rows': 0}

    if df_peek.empty or df_peek.shape[1] < 2:
        return {'type': 'unknown', 'headers': [], 'preview_rows': [], 'n_cols': 0, 'n_rows': 0}

    headers = [str(v).strip() for v in df_peek.iloc[0] if pd.notna(v)]
    headers_upper = [h.upper() for h in headers]

    if len(headers_upper) >= 3 and headers_upper[:3] == ["VARIABLE NAME", "OPTION NAME", "OPTION VALUE"]:
        preview_df = df_peek.iloc[1:6].copy()
        preview_df.columns = list(df_peek.iloc[0])
        return {
            'type': 'option_values',
            'headers': headers,
            'preview_df': preview_df,
            'n_cols': df_peek.shape[1],
            'n_rows': None,
        }

    has_subject = COL_SUBJECT in headers_upper
    has_status = COL_STATUS in headers_upper
    quest_name = any(x in sheet_name.lower() for x in ["questionna", "quest", " q "])

    if has_subject and has_status:
        sheet_type = 'questionnaire'
    elif quest_name and (has_subject or has_status):
        sheet_type = 'questionnaire'
    else:
        sheet_type = 'unknown'

    row2 = [str(v).strip() if pd.notna(v) else "" for v in df_peek.iloc[1]] if df_peek.shape[0] > 1 else []
    preview_df = df_peek.iloc[2:6].copy() if df_peek.shape[0] > 2 else pd.DataFrame()
    if not preview_df.empty:
        preview_df.columns = list(df_peek.iloc[0])

    return {
        'type': sheet_type,
        'headers': headers,
        'question_texts': row2,
        'preview_df': preview_df,
        'n_cols': df_peek.shape[1],
        'n_rows': None,
    }


def detect_questionnaire_sheets(excel_path):
    xls = pd.ExcelFile(excel_path)
    results = []
    for sheet_name in xls.sheet_names:
        info = classify_sheet(excel_path, sheet_name)
        if info['type'] == 'questionnaire':
            results.append((sheet_name, sheet_name))
    return results


def find_option_values_sheet(excel_path, questionnaire_sheet_name):
    xls = pd.ExcelFile(excel_path)

    candidate = questionnaire_sheet_name + "1"
    if candidate in xls.sheet_names:
        info = classify_sheet(excel_path, candidate)
        if info['type'] == 'option_values':
            return candidate

    if len(questionnaire_sheet_name) > 2 and questionnaire_sheet_name[1] == '-':
        prefix = questionnaire_sheet_name[:2]
        for sn in xls.sheet_names:
            if sn.startswith(prefix) and "option" in sn.lower():
                return sn

    base_lower = questionnaire_sheet_name.lower()
    for sn in xls.sheet_names:
        sn_lower = sn.lower()
        if sn_lower == base_lower:
            continue
        if sn_lower.startswith(base_lower[:min(len(base_lower), 10)]):
            info = classify_sheet(excel_path, sn)
            if info['type'] == 'option_values':
                return sn

    for sn in xls.sheet_names:
        if "option" in sn.lower() and "value" in sn.lower():
            return sn

    for sn in xls.sheet_names:
        if sn == questionnaire_sheet_name:
            continue
        info = classify_sheet(excel_path, sn)
        if info['type'] == 'option_values':
            return sn

    return None


def auto_pair_sheets(excel_path):
    xls = pd.ExcelFile(excel_path)
    all_sheet_info = {}
    for sn in xls.sheet_names:
        all_sheet_info[sn] = classify_sheet(excel_path, sn)

    quest_sheets = detect_questionnaire_sheets(excel_path)
    pairs = []
    unpaired_quest = []
    used_ov = set()

    for sheet_name, display in quest_sheets:
        ov = find_option_values_sheet(excel_path, sheet_name)
        if ov and ov not in used_ov:
            pairs.append((sheet_name, ov))
            used_ov.add(ov)
        else:
            unpaired_quest.append(sheet_name)

    return pairs, unpaired_quest, all_sheet_info


# ═══════════════════════════════════════════════════════════════
# 5. PARSING LOGIC (Manual Workflow)
# ═══════════════════════════════════════════════════════════════

def build_scale_library(excel_path, ov_sheet_name):
    df = pd.read_excel(excel_path, sheet_name=ov_sheet_name, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    scales = {}
    for _, row in df.iterrows():
        var, opt, val = str(row[df.columns[0]]), str(row[df.columns[1]]), row[df.columns[2]]
        try: val_int = int(float(val))
        except: continue
        norm = re.sub(r'_1$', '', var.strip())
        if norm not in scales: scales[norm] = {}
        scales[norm][val_int] = opt.strip()
    return scales


def determine_favorable_mask(levels):
    if not levels: return []
    mask = []
    for idx in sorted(levels.keys()):
        label = levels[idx].upper().replace("'", "")
        padded = f" {label} "
        if any(f" {kw} " in padded for kw in NEGATIVE_KEYWORDS): continue
        if any(f" {kw} " in padded for kw in POSITIVE_KEYWORDS): mask.append(idx)
    return mask


def load_timepoint(excel_path, sheet_name, ov_sheet_name, global_exclusions):
    tp_name = re.sub(r'\bquestionnai(re)?\b', '', sheet_name, flags=re.IGNORECASE)
    tp_name = re.sub(r'\s+', ' ', tp_name).strip(' -_')
    tp = TimepointData(tp_name, sheet_name, excel_path)
    scales = build_scale_library(excel_path, ov_sheet_name)

    df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    headers = [str(v).strip() for v in df_raw.iloc[0]]
    q_texts = [str(v).strip() if pd.notna(v) else "" for v in df_raw.iloc[1]]

    df_data = df_raw.iloc[2:].copy()
    df_data.columns = headers
    df_data = df_data.reset_index(drop=True)

    c_subj = next((h for h in headers if h.upper() == COL_SUBJECT), None)
    c_stat = next((h for h in headers if h.upper() == COL_STATUS), None)
    c_cent = next((h for h in headers if h.upper() == COL_CENTER), None)
    c_ref = next((h for h in headers if h.upper() == COL_STUDY), None)

    if not (c_subj and c_stat): return None
    if c_ref: tp.study_refs = [r for r in df_data[c_ref].dropna().astype(str).unique() if r.strip()]

    def fmt_sid(v):
        try: return f"{int(float(v)):04d}"
        except: return str(v).strip()

    df_data['_SID'] = df_data[c_subj].apply(fmt_sid)
    df_data['_STATUS'] = df_data[c_stat].astype(str).str.upper().str.strip()

    if c_cent:
        df_data['_CENTER'] = df_data[c_cent].astype(str).str.upper().str.strip()
        df_center = df_data[df_data['_CENTER'] == CENTER_FILTER].copy()
    else:
        df_center = df_data.copy()

    if global_exclusions:
        present_exclusions = df_center[df_center['_SID'].isin(global_exclusions)]['_SID'].tolist()
        tp.dropped_ids.extend(present_exclusions)
        df_center = df_center[~df_center['_SID'].isin(global_exclusions)].copy()

    tp.n_total = len(df_center)
    tp.subject_data = df_center

    for i, var in enumerate(headers):
        if not var or var.upper() in {k.upper() for k in METADATA_COLS} or var.startswith('.'): continue
        norm = re.sub(r'_1$', '', var)
        levels = scales.get(norm, scales.get(var, {}))
        q_text = re.sub(r'^[\d\.\)\-\s]+', '', q_texts[i]).strip()
        if q_text.startswith('"'): q_text = q_text[1:-1]
        qi = QuestionInfo(var, q_text, i, levels)
        if qi.is_scaled: qi.fav_mask = determine_favorable_mask(levels)
        tp.questions.append(qi)

    completed_mask = df_center['_STATUS'] == 'COMPLETED'
    tp.n_completed = completed_mask.sum()
    tp.included_subjects = df_center[completed_mask]['_SID'].tolist()

    non_completed_df = df_center[~completed_mask].copy()
    q_vars = [q.var_name for q in tp.questions if q.is_scaled]
    for _, row in non_completed_df.iterrows():
        answered = sum(1 for qv in q_vars if pd.notna(row.get(qv, "")) and str(row.get(qv, "")).strip())
        tp.non_completed.append({
            'sid': row['_SID'], 'status': row['_STATUS'],
            'answered': answered, 'total_q': len(q_vars),
            'coverage_pct': (answered / len(q_vars) * 100) if q_vars else 0
        })

    tp.n_included = len(tp.included_subjects)
    return tp


def detect_neutral_index(levels, fav_mask):
    neutral_keywords = ['NEITHER', 'NEUTRAL', 'UNDECIDED', 'NO OPINION', 'NOT SURE',
                        'UNCERTAIN', 'MIXED', "DON'T KNOW", 'DONT KNOW']
    sorted_keys = sorted(levels.keys())

    for idx in sorted_keys:
        label_upper = levels[idx].upper()
        if any(kw in label_upper for kw in neutral_keywords):
            return idx

    if len(sorted_keys) % 2 == 1:
        mid = sorted_keys[len(sorted_keys) // 2]
        if mid not in fav_mask and mid != sorted_keys[0]:
            return mid

    return None


def compute_stats(tp, q, split_neutral=False):
    if not q.is_scaled: return None
    df = tp.subject_data[tp.subject_data['_SID'].isin(tp.included_subjects)]
    counts = {k: 0 for k in q.levels.keys()}
    for val in df[q.var_name]:
        try: v = int(float(val)); counts[v] = counts.get(v, 0) + 1
        except: pass
    n = tp.n_included

    if q.neutral_mask is None and split_neutral:
        q.neutral_mask = detect_neutral_index(q.levels, q.fav_mask)

    display_counts = dict(counts)

    if split_neutral and q.neutral_mask is not None:
        neutral_idx = q.neutral_mask
        if neutral_idx in counts and counts[neutral_idx] > 0:
            neutral_total = counts[neutral_idx]
            to_agree = neutral_total // 2
            to_disagree = neutral_total - to_agree

            agree_target = None
            for idx in sorted(q.fav_mask, reverse=True):
                if idx != neutral_idx:
                    agree_target = idx
                    break

            disagree_target = None
            for idx in sorted(q.levels.keys()):
                if idx not in q.fav_mask and idx != neutral_idx:
                    disagree_target = idx
                    break

            if agree_target is not None:
                display_counts[agree_target] = display_counts.get(agree_target, 0) + to_agree
            if disagree_target is not None:
                display_counts[disagree_target] = display_counts.get(disagree_target, 0) + to_disagree
            display_counts[neutral_idx] = 0

    strict_fav = sum(counts[k] for k in q.fav_mask)
    neutral_contribution = 0
    if split_neutral and q.neutral_mask is not None and q.neutral_mask in counts:
        neutral_contribution = counts[q.neutral_mask] // 2
    total_fav = strict_fav + neutral_contribution

    fav_pct = (total_fav / n * 100) if n else 0

    return {
        'level_pcts': {k: (v / n * 100 if n else 0) for k, v in display_counts.items()},
        'fav_pct': fav_pct,
        'unfav_pct': 100 - fav_pct,
        'n': n
    }


def compute_stats_from_config(tp, q, excel_path, settings):
    if not q.is_scaled or q.is_multi_select:
        return None

    try:
        df_raw = pd.read_excel(excel_path, sheet_name=tp.sheet_name, header=None)
    except Exception:
        return None

    headers = [str(v).strip() for v in df_raw.iloc[0]]
    df_data = df_raw.iloc[2:].copy()
    df_data.columns = headers
    df_data = df_data.reset_index(drop=True)

    c_subj = next((h for h in headers if h.upper() == COL_SUBJECT), None)
    c_stat = next((h for h in headers if h.upper() == COL_STATUS), None)
    c_cent = next((h for h in headers if h.upper() == COL_CENTER), None)

    if not c_subj:
        return None

    def fmt_sid(v):
        try: return f"{int(float(v)):04d}"
        except: return str(v).strip()

    df_data['_SID'] = df_data[c_subj].apply(fmt_sid)

    center = settings.get("center_filter", CENTER_FILTER)
    if c_cent:
        df_data = df_data[df_data[c_cent].astype(str).str.upper().str.strip() == center]

    excluded = set(tp.dropped_ids)
    df_data = df_data[~df_data['_SID'].isin(excluded)]

    dev_exclude = set()
    for sid, reason in tp.deviation_subjects.items():
        if "No Data" in reason or reason.startswith("[EXCLUDE_N]"):
            dev_exclude.add(sid)
    df_data = df_data[~df_data['_SID'].isin(dev_exclude)]

    if settings.get("completed_only", False) and c_stat:
        df_data = df_data[df_data[c_stat].astype(str).str.upper().str.strip() == "COMPLETED"]

    n = len(df_data)
    if n == 0:
        return {'level_pcts': {}, 'fav_pct': 0, 'unfav_pct': 100, 'n': 0}

    var_col = None
    for h in headers:
        norm_h = re.sub(r'_1$', '', h)
        if h == q.var_name or norm_h == q.var_name:
            var_col = h
            break

    if var_col is None:
        return None

    counts = {k: 0 for k in q.levels.keys()}
    for val in df_data[var_col]:
        try:
            v = int(float(val))
            if v in counts:
                counts[v] += 1
        except:
            pass

    split_neutral = settings.get("split_neutral", False)
    display_counts = dict(counts)

    if split_neutral and q.neutral_mask is not None:
        neutral_idx = q.neutral_mask
        if neutral_idx in counts and counts[neutral_idx] > 0:
            neutral_total = counts[neutral_idx]
            to_agree = neutral_total // 2
            to_disagree = neutral_total - to_agree

            agree_target = None
            for idx in sorted(q.fav_mask, reverse=True):
                if idx != neutral_idx:
                    agree_target = idx
                    break

            disagree_target = None
            for idx in sorted(q.levels.keys()):
                if idx not in q.fav_mask and idx != neutral_idx:
                    disagree_target = idx
                    break

            if agree_target is not None:
                display_counts[agree_target] = display_counts.get(agree_target, 0) + to_agree
            if disagree_target is not None:
                display_counts[disagree_target] = display_counts.get(disagree_target, 0) + to_disagree
            display_counts[neutral_idx] = 0

    strict_fav = sum(counts.get(k, 0) for k in q.fav_mask)
    neutral_contribution = 0
    if split_neutral and q.neutral_mask is not None and q.neutral_mask in counts:
        neutral_contribution = counts[q.neutral_mask] // 2
    total_fav = strict_fav + neutral_contribution

    fav_pct = (total_fav / n * 100) if n else 0

    return {
        'level_pcts': {k: (v / n * 100 if n else 0) for k, v in display_counts.items()},
        'fav_pct': fav_pct,
        'unfav_pct': 100 - fav_pct,
        'n': n
    }


# ═══════════════════════════════════════════════════════════════
# 6. TITLE CLEANING
# ═══════════════════════════════════════════════════════════════

def build_chart_title(tp, text, is_topline):
    study = tp.study_refs[0] if tp.study_refs else "Study"
    tl = " TopLine" if is_topline else ""
    clean_name = re.sub(r'^\d+-?' + re.escape(study) + r'\s*', '', tp.name).strip()
    clean_name = re.sub(r'\bMasca\w*\b', 'Mascara', clean_name)
    clean_name = re.sub(r'\bEyela\w*\b', 'Eyelash', clean_name)
    return f"{study}: {clean_name}{tl} {text}"


def clean_chart_title(chart_id, suggested_title):
    cleaned = re.sub(r'\bMasca\w*\b', 'Mascara', suggested_title)
    cleaned = re.sub(r'\bEyela\w*\b', 'Eyelash', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned


# ═══════════════════════════════════════════════════════════════
# 7. CHART GENERATION
# ═══════════════════════════════════════════════════════════════

def make_bars_rounded(ax, pad=0.1, rounding_size=0.3):
    new_patches = []
    for patch in reversed(ax.patches):
        bb = patch.get_bbox()
        color = patch.get_facecolor()
        p_bbox = FancyBboxPatch(
            (bb.xmin, bb.ymin), abs(bb.width), abs(bb.height),
            boxstyle=f"round,pad={-pad},rounding_size={rounding_size}",
            ec="none", fc=color, mutation_aspect=0.5
        )
        patch.remove()
        new_patches.append(p_bbox)
    for patch in new_patches: ax.add_patch(patch)


def create_dashboard_page(tp, all_stats, is_topline, threshold_pct, custom_title=None):
    scaled_qs = [q for q in tp.questions if q.is_scaled and not q.is_multi_select and q.var_name in all_stats]
    if not scaled_qs: return None

    thresh_val = threshold_pct
    n_included = all_stats[scaled_qs[0].var_name]['n'] if scaled_qs else tp.n_included

    fig = plt.figure(figsize=(11, 8.5))
    fig.patch.set_facecolor('white')

    title = custom_title if custom_title else build_chart_title(tp, "Summary", is_topline)
    fig.text(0.05, 0.92, title, fontsize=18, weight='bold', color=COLORS['text_main'])
    fig.text(0.05, 0.88, f"Center: {CENTER_FILTER} | n={n_included} Included Subjects",
             fontsize=11, color=COLORS['text_sub'])

    # Color key / legend
    legend_y = 0.84
    fig.patches.append(Rectangle((0.05, legend_y), 0.015, 0.012, transform=fig.transFigure,
                                  color=COLORS['excellent'], zorder=2, clip_on=False))
    fig.text(0.07, legend_y + 0.002, f"≥ {int(thresh_val)}% Favorable", fontsize=9, color=COLORS['text_main'])

    fig.patches.append(Rectangle((0.22, legend_y), 0.015, 0.012, transform=fig.transFigure,
                                  color=COLORS['warning'], zorder=2, clip_on=False))
    fig.text(0.24, legend_y + 0.002, f"< {int(thresh_val)}% Favorable", fontsize=9, color=COLORS['text_main'])

    fig.text(0.05, 0.80,
             f"Dashed line indicates target favorable threshold ({int(thresh_val)}%)",
             fontsize=10, color=COLORS['text_sub'])

    ax = fig.add_axes([0.05, 0.12, 0.9, 0.64])
    x_pos = np.arange(len(scaled_qs))
    labels = [f"Q{q.q_number}" for q in scaled_qs]
    vals = [all_stats[q.var_name]['fav_pct'] for q in scaled_qs]
    colors = [COLORS['excellent'] if v >= thresh_val else COLORS['warning'] for v in vals]

    bars = ax.bar(x_pos, vals, color=colors, width=0.6, alpha=0.9)
    make_bars_rounded(ax, pad=0.05, rounding_size=0.2)
    ax.set_xticks(x_pos)
    ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=9)
    ax.set_ylabel("Favorable Response (%)")
    ax.set_ylim(0, 105)

    ax.axhline(thresh_val, color='#000000', linestyle='--', alpha=0.9, linewidth=1.0, zorder=5)


    return fig


def create_detailed_bars(tp, all_stats, is_topline, custom_title=None):
    scaled_qs = [q for q in tp.questions if q.is_scaled and not q.is_multi_select and q.var_name in all_stats]
    sorted_qs = sorted(scaled_qs, key=lambda q: all_stats[q.var_name]['fav_pct'], reverse=False)
    if not sorted_qs: return None

    fig, ax = plt.subplots(figsize=(12, max(6, len(sorted_qs) * 0.4 + 2)))
    for i, q in enumerate(sorted_qs):
        stats = all_stats[q.var_name]
        left = 0
        for j, key in enumerate(sorted(q.levels.keys())):
            pct = stats['level_pcts'].get(key, 0)
            c = COLORS['scale_colors'][j % len(COLORS['scale_colors'])]
            ax.barh(i, pct, left=left, height=0.7, color=c, edgecolor='white', linewidth=1)
            if pct > 6:
                ax.text(left + pct / 2, i, f"{pct:.2f}", ha='center', va='center', color='white', fontsize=8)
            left += pct

    ax.set_yticks(np.arange(len(sorted_qs)))
    ax.set_yticklabels([f"Q{q.q_number}" for q in sorted_qs], fontsize=10, weight='bold')

    title = custom_title if custom_title else build_chart_title(tp, "Detailed Breakdown", is_topline)
    ax.set_title(title, fontsize=14, weight='bold', loc='left')
    if sorted_qs:
        handles = [Rectangle((0, 0), 1, 1, color=COLORS['scale_colors'][i]) for i in
                   range(len(sorted_qs[0].levels))]
        labels = [sorted_qs[0].levels[k] for k in sorted(sorted_qs[0].levels.keys())]
        ax.legend(handles, labels, loc='upper center', bbox_to_anchor=(0.5, 1.05), ncol=4, frameon=False)
    plt.tight_layout()
    return fig


def create_comparison_page(timepoints, all_tp_stats):
    if len(timepoints) < 2: return None

    common = None
    tp_q_map = {}

    for tp in timepoints:
        q_map = {
            q.q_number: all_tp_stats[tp.name][q.var_name]['fav_pct']
            for q in tp.questions
            if q.is_scaled and not q.is_multi_select
            and q.var_name in all_tp_stats.get(tp.name, {})
        }
        tp_q_map[tp.name] = q_map

        if common is None:
            common = set(q_map.keys())
        else:
            common = common & set(q_map.keys())

    if not common: return None

    def sort_key(s):
        nums = re.findall(r'\d+', s)
        return int(nums[0]) if nums else 0

    sorted_qs = sorted(common, key=sort_key)

    contrast_colors = ['#0173B2', '#DE8F05', '#029E73', '#D55E00', '#CC78BC', '#8C564B']

    fig, ax = plt.subplots(figsize=(16, max(6, len(sorted_qs) * 0.5 + 2)))
    x = np.arange(len(sorted_qs))
    width = 0.8 / len(timepoints)

    for i, tp in enumerate(timepoints):
        vals = [tp_q_map[tp.name].get(qn, 0) for qn in sorted_qs]
        bar_color = contrast_colors[i % len(contrast_colors)]
        ax.bar(x + (i - len(timepoints) / 2 + 0.5) * width, vals, width, label=tp.name,
               color=bar_color, alpha=0.9, edgecolor='white', linewidth=1)

    make_bars_rounded(ax, pad=0.02, rounding_size=0.1)

    ax.set_xticks(x)
    ax.set_xticklabels([f"Q{q}" for q in sorted_qs], fontsize=10, weight='bold')
    ax.set_ylabel("Favorable Response (%)")
    ax.set_ylim(0, 105)

    ax.set_title("Cross-Timepoint Comparison (Common Questions Only)", fontsize=14, weight='bold', loc='left')
    ax.legend(loc='upper right', frameon=True, facecolor='white', framealpha=1)

    ax.set_axisbelow(True)
    ax.yaxis.grid(True, color='#EEEEEE')
    ax.xaxis.grid(False)

    plt.tight_layout()
    return fig


# ═══════════════════════════════════════════════════════════════
# 8. STREAMLIT APP
# ═══════════════════════════════════════════════════════════════

def init_session_state():
    defaults = {
        'step': 0,
        'timepoints': [],
        'all_tp_stats': {},
        'chart_titles': {},
        'file_processed': False,
        'config_loaded': False,
        'config_settings': {},
        'pdf_bytes': None,
        'pdf_name': '',
        'tp_names_confirmed': False,
        'needs_manual_pairing': False,
        'auto_pairs': [],
        'unpaired_quest': [],
        'all_sheet_info': {},
        'manual_pairs_confirmed': False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def render_sidebar():
    with st.sidebar:
        st.markdown("## 📊 ePRO Chart Gen v9.2")
        st.caption("Config-Aware Edition")
        st.divider()

        if st.session_state.config_loaded:
            st.success("Config loaded")
            steps = [
                "1️⃣ Upload & Configure",
                "2️⃣ Review & Generate",
            ]
        else:
            steps = [
                "1️⃣ Upload & Configure",
                "2️⃣ Verify Scales",
                "3️⃣ Subject Inclusion",
                "4️⃣ Review Titles",
                "5️⃣ Generate PDF",
            ]

        for i, label in enumerate(steps):
            disabled = i > st.session_state.step + 1
            if st.sidebar.button(label, key=f"nav_{i}", disabled=disabled, use_container_width=True):
                st.session_state.step = i
                st.rerun()


def _render_sheet_preview(sheet_name, info):
    badge = {
        'questionnaire': '📋 Questionnaire',
        'option_values': '📖 Option Values',
        'unknown': '❓ Unknown',
    }.get(info['type'], '❓ Unknown')

    st.markdown(f"**{sheet_name}** — {badge}")
    st.caption(f"{info['n_cols']} columns | Headers: {', '.join(info['headers'][:6])}{'...' if len(info['headers']) > 6 else ''}")

    preview_df = info.get('preview_df', pd.DataFrame())
    if not preview_df.empty:
        display_cols = list(preview_df.columns[:5])
        st.dataframe(preview_df[display_cols].head(3), use_container_width=True, hide_index=True)


def _render_manual_pairing_ui(excel_path):
    all_info = st.session_state.all_sheet_info
    auto_pairs = st.session_state.auto_pairs
    unpaired = st.session_state.unpaired_quest

    all_sheets = list(all_info.keys())
    ov_sheets = [sn for sn, info in all_info.items() if info['type'] == 'option_values']
    quest_sheets = [sn for sn, info in all_info.items() if info['type'] == 'questionnaire']

    if auto_pairs:
        st.success(f"Auto-detected {len(auto_pairs)} pair(s):")
        for q_sheet, ov_sheet in auto_pairs:
            st.caption(f"  📋 {q_sheet}  ↔  📖 {ov_sheet}")

    if unpaired:
        st.warning(f"{len(unpaired)} questionnaire sheet(s) could not be auto-paired with an option-values sheet.")

    st.divider()

    st.subheader("All Sheets in Workbook")
    st.caption("Expand any sheet to preview its contents, then use the pairing controls below.")

    for sn in all_sheets:
        info = all_info[sn]
        with st.expander(f"{sn}", expanded=(sn in unpaired)):
            _render_sheet_preview(sn, info)

    st.divider()

    st.subheader("Manual Sheet Pairing")
    st.caption(
        "Select which sheets are questionnaire data and pair each with its option-values sheet. "
        "You can add pairs beyond what was auto-detected, or override auto-detected pairs."
    )

    n_pairs = st.number_input(
        "Number of timepoints (questionnaire → option-values pairs)",
        min_value=1, max_value=10,
        value=max(len(auto_pairs) + len(unpaired), 1),
        key="n_manual_pairs"
    )

    prefilled_pairs = list(auto_pairs)
    for up in unpaired:
        prefilled_pairs.append((up, ""))

    manual_pairs = []
    for i in range(int(n_pairs)):
        st.markdown(f"**Pair {i + 1}**")
        col1, col2 = st.columns(2)

        default_quest = prefilled_pairs[i][0] if i < len(prefilled_pairs) else all_sheets[0]
        default_ov = prefilled_pairs[i][1] if i < len(prefilled_pairs) and prefilled_pairs[i][1] else None

        with col1:
            quest_idx = all_sheets.index(default_quest) if default_quest in all_sheets else 0
            q_sheet = st.selectbox(
                "Questionnaire sheet",
                all_sheets,
                index=quest_idx,
                key=f"manual_q_{i}"
            )

        with col2:
            ov_options = ["(none — skip)"] + all_sheets
            if default_ov and default_ov in all_sheets:
                ov_idx = ov_options.index(default_ov)
            else:
                ov_idx = 0
            ov_sheet = st.selectbox(
                "Option values sheet",
                ov_options,
                index=ov_idx,
                key=f"manual_ov_{i}"
            )

        if ov_sheet != "(none — skip)":
            manual_pairs.append((q_sheet, ov_sheet))

    return manual_pairs


def step_upload():
    st.header("Upload & Configure")

    import_info = (
        "Upload the VBA-exported JSON configuration **and** the matching ePRO Excel workbook.\n\n"
        "The JSON contains all scale decisions, exclusions, and masks, while the workbook "
        "contains the actual survey responses used to compute statistics and generate charts.\n\n"
        "If JSON is not compatible, a properly formatted text file exported from the Excel UserForm can be used instead."
    )

    mode = st.radio(
        "Workflow Mode",
        [
            "Import VBA Config + ePRO Workbook ",
            "Manual Workbook Import"
        ],
        horizontal=True
    )

    if mode.startswith("Import VBA Config"):
        st.caption(import_info)

        config_file = st.file_uploader("Upload JSON configuration (or compatible text file)", type=["json", "txt"])
        workbook_file = st.file_uploader(
            "Upload ePRO Workbook (required)", type=["xlsx", "xls"],
            help="Must be the same workbook used in the Excel UserForm export"
        )

        if config_file and workbook_file and st.button("Load Config + Data", type="primary"):
            with st.spinner("Loading configuration..."):
                config, settings, timepoints = load_config_from_json(config_file.read())

                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                    tmp.write(workbook_file.read())
                    tmp_path = tmp.name

                all_tp_stats = {}
                for tp in timepoints:
                    tp.n_included = compute_n_included_from_config(tp, settings)
                    stats = {}
                    for q in tp.questions:
                        if q.is_scaled and not q.is_multi_select:
                            s = compute_stats_from_config(tp, q, tmp_path, settings)
                            if s:
                                stats[q.var_name] = s
                    all_tp_stats[tp.name] = stats

                is_topline = settings.get("is_topline", False)
                chart_titles = {}
                for tp in timepoints:
                    chart_id = f"{tp.name}_dashboard"
                    raw_title = build_chart_title(tp, "Summary", is_topline)
                    chart_titles[chart_id] = clean_chart_title(chart_id, raw_title)

                    if tp.needs_unrandomization and tp.randomization_groups:
                        for grp_name in tp.randomization_groups:
                            chart_id = f"{tp.name}_{grp_name}_dashboard"
                            raw_title = build_chart_title(tp, f"{grp_name} Summary", is_topline)
                            chart_titles[chart_id] = clean_chart_title(chart_id, raw_title)

                st.session_state.timepoints = timepoints
                st.session_state.all_tp_stats = all_tp_stats
                st.session_state.chart_titles = chart_titles
                st.session_state.config_loaded = True
                st.session_state.config_settings = settings
                st.session_state.is_topline = is_topline
                st.session_state.tmp_path = tmp_path
                st.session_state.file_processed = True
                st.session_state.pdf_name = f"{Path(workbook_file.name).stem}{'_TPL' if is_topline else ''}_Charts_v9.pdf"
                st.session_state.step = 1
                st.rerun()

    else:
        # === MANUAL WORKFLOW ===
        st.subheader("Manual ePRO Workbook Import")
        st.caption(
            "Upload only the raw ePRO Excel workbook. "
            "Streamlit will detect questions, scales, and compute stats manually."
        )
        uploaded = st.file_uploader("Upload ePRO Workbook (required)", type=["xlsx", "xls"])

        col1, col2, col3 = st.columns(3)
        with col1:
            is_topline = st.checkbox("TopLine Report", value=False)
        with col2:
            split_neutral = st.checkbox("Split Neutral", value=False,
                help="When enabled, neutral/midpoint responses (e.g. \"Neither Agree nor Disagree\") "
                     "are split 50/50 between the favorable and unfavorable sides rather than being "
                     "counted as a standalone category. This is common in Likert-scale analysis where "
                     "the midpoint is considered ambivalent rather than truly neutral.")
        with col3:
            exclusions_str = st.text_input("Enter Dropped Subjects here", placeholder="e.g. 0042, 0091")

        if uploaded and st.button("Process Workbook", type="primary"):
            global_exclusions = [x.strip() for x in exclusions_str.split(',') if x.strip()] if exclusions_str else []
            normalized = []
            for ex in global_exclusions:
                try:
                    normalized.append(f"{int(ex):04d}")
                except ValueError:
                    normalized.append(ex)
            global_exclusions = normalized

            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(uploaded.read())
                tmp_path = tmp.name

            with st.spinner("Scanning workbook sheets..."):
                pairs, unpaired, all_sheet_info = auto_pair_sheets(tmp_path)

            st.session_state.tmp_path = tmp_path
            st.session_state.is_topline = is_topline
            st.session_state.split_neutral = split_neutral
            st.session_state.global_exclusions = global_exclusions
            st.session_state.uploaded_name = uploaded.name

            if pairs and not unpaired:
                _process_manual_pairs(pairs, tmp_path, global_exclusions, is_topline, uploaded.name)
            else:
                st.session_state.auto_pairs = pairs
                st.session_state.unpaired_quest = unpaired
                st.session_state.all_sheet_info = all_sheet_info
                st.session_state.needs_manual_pairing = True
                st.session_state.manual_pairs_confirmed = False
                st.rerun()

        # === MANUAL SHEET PAIRING FALLBACK ===
        if st.session_state.get('needs_manual_pairing') and not st.session_state.get('manual_pairs_confirmed'):
            st.divider()
            st.subheader("⚠️ Sheet Pairing Required")
            st.info(
                "Auto-detection couldn't fully pair all questionnaire sheets with their option-values sheets. "
                "Please review the sheets below and confirm the correct pairings."
            )

            manual_pairs = _render_manual_pairing_ui(st.session_state.tmp_path)

            if st.button("Confirm Pairings & Process", type="primary"):
                if not manual_pairs:
                    st.error("Please configure at least one questionnaire ↔ option-values pair.")
                    return

                tmp_path = st.session_state.tmp_path
                global_exclusions = st.session_state.get('global_exclusions', [])
                is_topline = st.session_state.get('is_topline', False)
                uploaded_name = st.session_state.get('uploaded_name', 'workbook.xlsx')

                st.session_state.needs_manual_pairing = False
                st.session_state.manual_pairs_confirmed = True

                _process_manual_pairs(manual_pairs, tmp_path, global_exclusions, is_topline, uploaded_name)

        # --- Timepoint Naming (shown after workbook is processed) ---
        if st.session_state.file_processed and not st.session_state.config_loaded and not st.session_state.get('tp_names_confirmed', False):
            st.divider()
            st.subheader("Confirm Timepoint Names")
            st.caption("Review and edit the detected timepoint names. These will appear in all chart titles.")

            timepoints = st.session_state.timepoints
            is_topline = st.session_state.get('is_topline', False)
            generic_patterns = ['epro data', 'epro', 'data', 'sheet']

            updated_names = []
            for i, tp in enumerate(timepoints):
                is_generic = any(p in tp.name.lower() for p in generic_patterns) or len(tp.name.strip()) < 3
                default_val = tp.name
                label = f"Timepoint {i + 1}"
                if is_generic:
                    label += " generic name detected"
                new_name = st.text_input(label, value=default_val, key=f"tp_name_{i}")
                updated_names.append(new_name)

            if st.button("Confirm Names", type="primary"):
                old_stats = dict(st.session_state.all_tp_stats)
                new_stats = {}
                for i, tp in enumerate(timepoints):
                    old_name = tp.name
                    tp.name = updated_names[i]
                    new_stats[tp.name] = old_stats.get(old_name, {})

                st.session_state.all_tp_stats = new_stats

                chart_titles = {}
                for tp in timepoints:
                    chart_id = f"{tp.name}_dashboard"
                    raw_title = build_chart_title(tp, "Summary", is_topline)
                    chart_titles[chart_id] = clean_chart_title(chart_id, raw_title)

                st.session_state.chart_titles = chart_titles
                st.session_state.tp_names_confirmed = True
                st.session_state.step = 1
                st.rerun()


def _process_manual_pairs(pairs, tmp_path, global_exclusions, is_topline, uploaded_name):
    timepoints = []
    all_tp_stats = {}

    with st.spinner("Processing timepoints..."):
        split_neutral = st.session_state.get('split_neutral', False)
        for quest_sheet, ov_sheet in pairs:
            tp = load_timepoint(tmp_path, quest_sheet, ov_sheet, global_exclusions)
            if not tp:
                continue
            stats = {}
            for q in tp.questions:
                if q.is_scaled:
                    stats[q.var_name] = compute_stats(tp, q, split_neutral=split_neutral)
            timepoints.append(tp)
            all_tp_stats[tp.name] = stats

    if not timepoints:
        st.error("No valid timepoints could be loaded from the selected pairs. "
                 "Please check that each questionnaire sheet has SUBJECT ID and STATUS columns, "
                 "and each option-values sheet has VARIABLE NAME / OPTION NAME / OPTION VALUE columns.")
        return

    st.session_state.timepoints = timepoints
    st.session_state.all_tp_stats = all_tp_stats
    st.session_state.is_topline = is_topline
    st.session_state.file_processed = True
    st.session_state.config_loaded = False
    st.session_state.config_settings = {}
    st.session_state.pdf_name = f"{Path(uploaded_name).stem}{'_TPL' if is_topline else ''}_Charts_v9.pdf"
    st.session_state.chart_titles = {}
    st.session_state.tp_names_confirmed = False
    st.session_state.step = 0
    st.rerun()


def step_config_review_and_generate():
    st.header("Review & Generate")

    timepoints = st.session_state.timepoints
    all_tp_stats = st.session_state.all_tp_stats
    titles = st.session_state.chart_titles
    settings = st.session_state.config_settings
    is_topline = settings.get("is_topline", False)
    threshold = settings.get("favorable_threshold", 0.70)
    threshold_pct = threshold * 100

    st.subheader("Imported Configuration")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Timepoints", len(timepoints))
    col2.metric("Threshold", f"{threshold_pct:.0f}%")
    col3.metric("Completed Only", "Yes" if settings.get("completed_only") else "No")
    col4.metric("Split Neutral", "Yes" if settings.get("split_neutral") else "No")

    for tp in timepoints:
        tp_stats = all_tp_stats.get(tp.name, {})
        n_val = tp_stats[next(iter(tp_stats))]['n'] if tp_stats else tp.n_included
        scaled_count = sum(1 for q in tp.questions if q.is_scaled and not q.is_multi_select)

        with st.expander(f"{tp.name} — n={n_val}, {scaled_count} scaled questions", expanded=False):
            c1, c2, c3 = st.columns(3)
            c1.metric("Total VCS", tp.n_total)
            c2.metric("Excluded", len(tp.dropped_ids))
            c3.metric("Deviations", len(tp.deviation_subjects))

            if tp.dropped_ids:
                st.caption(f"Dropped: {', '.join(sorted(tp.dropped_ids))}")
            if tp.deviation_subjects:
                for sid, reason in sorted(tp.deviation_subjects.items()):
                    display_reason = reason.replace("[EXCLUDE_N]", "")
                    st.caption(f"  {sid}: {display_reason}")

            if tp.needs_unrandomization:
                st.info(f"Randomization: {tp.randomization_source} — "
                        f"{len(tp.randomization_groups)} group(s)")

    st.divider()

    st.subheader("Chart Titles")
    updated_titles = {}
    for chart_id, title in titles.items():
        new_title = st.text_input(chart_id, value=title, key=f"title_{chart_id}")
        updated_titles[chart_id] = new_title
    st.session_state.chart_titles = updated_titles

    st.divider()

    if st.session_state.pdf_bytes:
        st.success("PDF generated successfully!")
        st.download_button("Download PDF", data=st.session_state.pdf_bytes,
                           file_name=st.session_state.pdf_name, mime="application/pdf", type="primary")
        if st.button("🔄 Regenerate"):
            st.session_state.pdf_bytes = None
            st.rerun()
        return

    if st.button("Generate PDF", type="primary"):
        _generate_pdf(timepoints, all_tp_stats, updated_titles, is_topline, threshold_pct)


def _generate_pdf(timepoints, all_tp_stats, titles, is_topline, threshold_pct):
    pdf_buffer = BytesIO()
    progress = st.progress(0, text="Generating charts...")
    total = len(timepoints) + (1 if len(timepoints) >= 2 else 0)
    chart_count = 0

    with PdfPages(pdf_buffer) as pdf:
        for tp in timepoints:
            stats = all_tp_stats.get(tp.name, {})

            fig = create_dashboard_page(tp, stats, is_topline, threshold_pct,
                                        custom_title=titles.get(f"{tp.name}_dashboard"))
            if fig: pdf.savefig(fig); plt.close(fig)
            chart_count += 1
            progress.progress(chart_count / total, text=f"{tp.name} dashboard...")

        if len(timepoints) >= 2:
            fig = create_comparison_page(timepoints, all_tp_stats)
            if fig: pdf.savefig(fig); plt.close(fig)

    progress.progress(1.0, text="Done!")
    st.session_state.pdf_bytes = pdf_buffer.getvalue()
    st.rerun()


# === Manual workflow steps ===

def step_scales():
    st.header("Verify Favorable Logic")
    st.caption("Review and adjust which response values count as 'favorable' for each scale group. "
               "Use per-question overrides for negatively worded or special-case questions.")
    timepoints = st.session_state.timepoints

    for tp in timepoints:
        st.subheader(f"{tp.name}")
        groups = defaultdict(list)
        for q in tp.questions:
            if q.is_scaled: groups[q.scale_signature].append(q)

        for sig, q_list in groups.items():
            ex_q = q_list[0]
            scale_str = " | ".join([f"**{k}**: {v}" for k, v in sorted(ex_q.levels.items())])
            with st.expander(f"Scale Group — {len(q_list)} questions", expanded=True):
                st.markdown(f"**Scale:** {scale_str}")

                current_fav = ", ".join(map(str, ex_q.fav_mask))
                group_key = f"fav_group_{tp.name}_{id(ex_q)}"
                new_fav = st.text_input("Favorable values (applies to all in group)", value=current_fav, key=group_key)
                try:
                    parsed = [int(x.strip()) for x in new_fav.split(',') if x.strip()]
                    group_mask = [p for p in parsed if p in ex_q.levels]
                except:
                    group_mask = ex_q.fav_mask

                for q in q_list:
                    q.fav_mask = list(group_mask)

                q_data = []
                for q in q_list:
                    fav_str = ", ".join(map(str, q.fav_mask)) if q.fav_mask else "None"
                    q_data.append({"Q#": f"Q{q.q_number}", "Favorable": fav_str,
                                   "Question": q.question_text[:80]})
                st.dataframe(pd.DataFrame(q_data), use_container_width=True, hide_index=True)

                if len(q_list) > 1:
                    override_key = f"show_overrides_{tp.name}_{id(ex_q)}"
                    show_overrides = st.checkbox("Show per-question overrides", key=override_key, value=False)

                    if show_overrides:
                        st.caption("Leave blank or match the group value to keep the default. "
                                   "Enter different values to override for that question.")
                        for q in q_list:
                            col1, col2 = st.columns([3, 2])
                            with col1:
                                st.markdown(f"**Q{q.q_number}:** {q.question_text[:60]}")
                            with col2:
                                q_key = f"fav_q_{tp.name}_{q.var_name}"
                                q_current = ", ".join(map(str, q.fav_mask))
                                q_new = st.text_input(
                                    f"Favorable for Q{q.q_number}",
                                    value=q_current,
                                    key=q_key,
                                    label_visibility="collapsed"
                                )
                                try:
                                    q_parsed = [int(x.strip()) for x in q_new.split(',') if x.strip()]
                                    q_valid = [p for p in q_parsed if p in q.levels]
                                    if q_valid:
                                        q.fav_mask = q_valid
                                except:
                                    pass

    if st.button("Confirm Scales", type="primary"):
        split_neutral = st.session_state.get('split_neutral', False)
        for tp in st.session_state.timepoints:
            stats = {}
            for q in tp.questions:
                if q.is_scaled: stats[q.var_name] = compute_stats(tp, q, split_neutral=split_neutral)
            st.session_state.all_tp_stats[tp.name] = stats
        st.session_state.step = 2
        st.rerun()


def step_subjects():
    st.header("Subject Inclusion")
    timepoints = st.session_state.timepoints
    any_non_completed = False

    for tp in timepoints:
        if not tp.non_completed: continue
        any_non_completed = True
        st.subheader(f"{tp.name}")
        st.markdown(f"**Completed:** {tp.n_completed} / {tp.n_total}")

        non_comp_with_data = [s for s in tp.non_completed if s['coverage_pct'] > 0]
        if not non_comp_with_data:
            st.info("No non-completed subjects with data.")
            continue

        for s in sorted(non_comp_with_data, key=lambda x: x['sid']):
            col1, col2, col3, col4 = st.columns([1, 2, 2, 3])
            key = f"inc_{tp.name}_{s['sid']}"
            with col1: include = st.checkbox("Include", key=key, value=False)
            with col2: st.markdown(f"**{s['sid']}**")
            with col3: st.caption(s['status'])
            with col4: st.progress(s['coverage_pct'] / 100, text=f"{s['coverage_pct']:.0f}% data")
            if include and s['sid'] not in tp.included_subjects:
                tp.included_subjects.append(s['sid'])
                tp.n_included = len(tp.included_subjects)

    if not any_non_completed:
        st.success("All subjects completed across all timepoints.")

    st.divider()
    st.subheader("Subject Summary")
    for tp in timepoints:
        with st.expander(f"{tp.name} — {len(tp.included_subjects)} included, {len(tp.dropped_ids)} excluded", expanded=True):
            col_inc, col_exc = st.columns(2)
            with col_inc:
                st.markdown("**Included Subjects**")
                if tp.included_subjects:
                    st.text(", ".join(sorted(tp.included_subjects)))
                else:
                    st.caption("None")
            with col_exc:
                st.markdown("**Excluded / Dropped Subjects**")
                if tp.dropped_ids:
                    st.text(", ".join(sorted(set(tp.dropped_ids))))
                else:
                    st.caption("None")

    if st.button("Confirm Subjects", type="primary"):
        split_neutral = st.session_state.get('split_neutral', False)
        for tp in st.session_state.timepoints:
            stats = {}
            for q in tp.questions:
                if q.is_scaled: stats[q.var_name] = compute_stats(tp, q, split_neutral=split_neutral)
            st.session_state.all_tp_stats[tp.name] = stats
        st.session_state.step = 3
        st.rerun()


def step_titles():
    st.header("Review Chart Titles")
    titles = st.session_state.chart_titles
    updated = {}
    for chart_id, title in titles.items():
        new_title = st.text_input(chart_id, value=title, key=f"title_{chart_id}")
        updated[chart_id] = new_title

    if st.button("Confirm Titles", type="primary"):
        st.session_state.chart_titles = updated
        st.session_state.step = 4
        st.rerun()


def step_generate_manual():
    st.header("Generate PDF")
    timepoints = st.session_state.timepoints
    all_tp_stats = st.session_state.all_tp_stats
    titles = st.session_state.chart_titles
    is_topline = st.session_state.get('is_topline', False)
    threshold_pct = FAVORABLE_THRESHOLD * 100

    col1, col2, col3 = st.columns(3)
    col1.metric("Timepoints", len(timepoints))
    col2.metric("Total Charts", len(titles))
    col3.metric("Center", CENTER_FILTER)

    tp_cols = st.columns(len(timepoints))
    for i, tp in enumerate(timepoints):
        tp_cols[i].metric(f"n — {tp.name}", tp.n_included)

    if st.session_state.pdf_bytes:
        st.success("PDF generated!")
        st.download_button("Download PDF", data=st.session_state.pdf_bytes,
                           file_name=st.session_state.pdf_name, mime="application/pdf", type="primary")
        if st.button("Regenerate"):
            st.session_state.pdf_bytes = None
            st.rerun()
        return

    if st.button("Generate PDF", type="primary"):
        _generate_pdf(timepoints, all_tp_stats, titles, is_topline, threshold_pct)


# ═══════════════════════════════════════════════════════════════
# 9. MAIN APP ENTRY
# ═══════════════════════════════════════════════════════════════

def main():
    st.set_page_config(page_title="ePRO Chart Generator v9.2", page_icon="📊", layout="wide")

    init_session_state()
    render_sidebar()

    if st.session_state.config_loaded:
        step_funcs = [step_upload, step_config_review_and_generate]
    else:
        step_funcs = [step_upload, step_scales, step_subjects, step_titles, step_generate_manual]

    current = min(st.session_state.step, len(step_funcs) - 1)
    step_funcs[current]()


if __name__ == "__main__":
    main()



