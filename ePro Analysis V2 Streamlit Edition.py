"""
ePRO Chart Generator v8 — Streamlit Edition
=============================================
Run with:  streamlit run epro_streamlit.py
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
    'trendline': '#1f77b4',
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
        self.fav_mask = []
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
        self.study_refs = []
        self.n_total = 0
        self.n_completed = 0
        self.n_included = 0


# ═══════════════════════════════════════════════════════════════
# 3. PARSING LOGIC
# ═══════════════════════════════════════════════════════════════

def detect_questionnaire_sheets(excel_path):
    xls = pd.ExcelFile(excel_path)
    results = []
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, nrows=2, header=None)
        except: continue
        if df.empty or df.shape[1] < 3: continue
        headers = [str(v).strip().upper() for v in df.iloc[0] if pd.notna(v)]
        if headers[:3] == ["VARIABLE NAME", "OPTION NAME", "OPTION VALUE"]: continue
        is_quest = any(x in sheet_name.lower() for x in ["questionna", "quest", " q "])
        header_match = sum(1 for k in [COL_STUDY, COL_SUBJECT, COL_STATUS] if k in headers)
        if is_quest or header_match >= 2:
            results.append((sheet_name, sheet_name))
    return results


def find_option_values_sheet(excel_path, questionnaire_sheet_name):
    xls = pd.ExcelFile(excel_path)
    if len(questionnaire_sheet_name) > 2 and questionnaire_sheet_name[1] == '-':
        prefix = questionnaire_sheet_name[:2]
        for sn in xls.sheet_names:
            if sn.startswith(prefix) and "option" in sn.lower(): return sn
    for sn in xls.sheet_names:
        if "option" in sn.lower() and "value" in sn.lower(): return sn
    return None


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


def compute_stats(tp, q):
    if not q.is_scaled: return None
    df = tp.subject_data[tp.subject_data['_SID'].isin(tp.included_subjects)]
    counts = {k: 0 for k in q.levels.keys()}
    for val in df[q.var_name]:
        try: v = int(float(val)); counts[v] = counts.get(v, 0) + 1
        except: pass
    n = tp.n_included
    fav = sum(counts[k] for k in q.fav_mask)
    return {
        'level_pcts': {k: (v / n * 100 if n else 0) for k, v in counts.items()},
        'fav_pct': (fav / n * 100) if n else 0,
        'unfav_pct': 100 - ((fav / n * 100) if n else 0),
        'n': n
    }


# ═══════════════════════════════════════════════════════════════
# 4. TITLE CLEANING
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
# 5. CHART GENERATION
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


def create_dashboard_page(tp, all_stats, is_topline, custom_title=None):
    scaled_qs = [q for q in tp.questions if q.is_scaled and q.var_name in all_stats]
    if not scaled_qs: return None

    fav_vals = [all_stats[q.var_name]['fav_pct'] for q in scaled_qs]
    avg_fav = np.mean(fav_vals) if fav_vals else 0
    thresh_val = FAVORABLE_THRESHOLD * 100

    fig = plt.figure(figsize=(11, 8.5))
    fig.patch.set_facecolor('white')

    title = custom_title if custom_title else build_chart_title(tp, "Summary", is_topline)
    fig.text(0.05, 0.92, title, fontsize=18, weight='bold', color=COLORS['text_main'])
    fig.text(0.05, 0.88, f"Center: {CENTER_FILTER} | n={tp.n_included} Included Subjects",
             fontsize=11, color=COLORS['text_sub'])

    cards = [
        {'label': 'Total Enrolled', 'val': str(tp.n_total + len(tp.dropped_ids)),
         'sub': 'VCS Center', 'x': 0.05, 'c': COLORS['favorable']},
        {'label': 'Completed', 'val': str(tp.n_completed),
         'sub': f'{(tp.n_completed / (tp.n_total + len(tp.dropped_ids)) * 100):.1f}% Rate', 'x': 0.29,
         'c': COLORS['excellent']},
        {'label': 'Avg Favorable', 'val': f"{avg_fav:.1f}%", 'sub': 'All Questions', 'x': 0.53,
         'c': COLORS['text_main']},
        {'label': f'Questions >{int(thresh_val)}%',
         'val': str(sum(1 for x in fav_vals if x >= thresh_val)), 'sub': 'Favorable Rate', 'x': 0.77,
         'c': COLORS['warning']},
    ]

    card_y = 0.72
    for card in cards:
        rect = Rectangle((card['x'], card_y), 0.18, 0.12, transform=fig.transFigure,
                          color=COLORS['bg_card'], zorder=1, clip_on=False)
        line = Rectangle((card['x'], card_y + 0.115), 0.18, 0.005, transform=fig.transFigure,
                          color=card['c'], zorder=2, clip_on=False)
        fig.patches.extend([rect, line])
        fig.text(card['x'] + 0.02, card_y + 0.08, card['label'], fontsize=10, color=COLORS['text_sub'])
        fig.text(card['x'] + 0.02, card_y + 0.035, card['val'], fontsize=22, weight='bold',
                 color=COLORS['text_main'])
        fig.text(card['x'] + 0.02, card_y + 0.015, card['sub'], fontsize=9, color=COLORS['text_sub'])

    fig.text(0.05, 0.70,
             f"Dashed line indicates target favorable threshold ({int(thresh_val)}%)",
             fontsize=10, color=COLORS['text_sub'])

    ax = fig.add_axes([0.05, 0.25, 0.9, 0.40])
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

    z = np.polyfit(x_pos, vals, 1)
    p = np.poly1d(z)
    ax.plot(x_pos, p(x_pos), color=COLORS['trendline'], linestyle='-', linewidth=2, label='Trendline')
    ax.axhline(thresh_val, color=COLORS['neutral'], linestyle='--', alpha=0.5, linewidth=1)
    ax.legend(loc='upper left', frameon=False, fontsize=9)

    all_present_ids = set(tp.subject_data['_SID'])
    inc_ids = set(tp.included_subjects)
    excluded = sorted(list((all_present_ids - inc_ids) | set(tp.dropped_ids)))

    if excluded:
        drop_str = ", ".join(excluded)
        fig.text(0.05, 0.08, f"Dropped/Excluded Subjects ({len(excluded)}):",
                 fontsize=9, weight='bold', color='red')
        fig.text(0.05, 0.05, textwrap.fill(drop_str, 120), fontsize=8,
                 color=COLORS['text_main'], va='top')
    else:
        fig.text(0.05, 0.05, "No subjects excluded from analysis.",
                 fontsize=9, color=COLORS['excellent'])

    return fig


def create_ranked_chart(tp, all_stats, is_topline, custom_title=None):
    scaled_qs = [q for q in tp.questions if q.is_scaled and q.var_name in all_stats]
    sorted_qs = sorted(scaled_qs, key=lambda q: all_stats[q.var_name]['fav_pct'], reverse=False)
    n_qs = len(sorted_qs)

    fig, ax = plt.subplots(figsize=(11, max(6, n_qs * 0.4 + 2)))
    labels = [f"Q{q.q_number}" for q in sorted_qs]
    vals = [all_stats[q.var_name]['fav_pct'] for q in sorted_qs]
    colors = [COLORS['excellent'] if v >= FAVORABLE_THRESHOLD * 100 else COLORS['warning'] for v in vals]

    ax.barh(np.arange(n_qs), vals, color=colors, height=0.6)
    make_bars_rounded(ax)
    ax.set_yticks(np.arange(n_qs))
    ax.set_yticklabels(labels, fontsize=10, weight='bold')
    ax.set_xlabel("Favorable Response (%)")

    title = custom_title if custom_title else build_chart_title(tp, "Ranked Performance", is_topline)
    ax.set_title(title, fontsize=14, weight='bold', loc='left')

    ax.axvline(FAVORABLE_THRESHOLD * 100, color=COLORS['excellent'], linestyle='--', linewidth=1)
    for i, v in enumerate(vals):
        ax.text(v + 1, i, f"{v:.1f}%", va='center', fontsize=9)

    plt.tight_layout()
    return fig


def create_detailed_bars(tp, all_stats, is_topline, custom_title=None):
    scaled_qs = [q for q in tp.questions if q.is_scaled and q.var_name in all_stats]
    sorted_qs = sorted(scaled_qs, key=lambda q: all_stats[q.var_name]['fav_pct'], reverse=False)
    fig, ax = plt.subplots(figsize=(12, max(6, len(sorted_qs) * 0.4 + 2)))
    for i, q in enumerate(sorted_qs):
        stats = all_stats[q.var_name]
        left = 0
        for j, key in enumerate(sorted(q.levels.keys())):
            pct = stats['level_pcts'].get(key, 0)
            c = COLORS['scale_colors'][j % len(COLORS['scale_colors'])]
            ax.barh(i, pct, left=left, height=0.7, color=c, edgecolor='white', linewidth=1)
            if pct > 6:
                ax.text(left + pct / 2, i, f"{int(pct)}", ha='center', va='center', color='white', fontsize=8)
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


def create_diverging_chart(tp, all_stats, is_topline, custom_title=None):
    scaled_qs = [q for q in tp.questions if q.is_scaled and q.var_name in all_stats]
    sorted_qs = sorted(scaled_qs, key=lambda q: all_stats[q.var_name]['fav_pct'], reverse=False)
    n_qs = len(sorted_qs)
    fig, ax = plt.subplots(figsize=(11, max(6, n_qs * 0.4 + 2)))
    y = np.arange(n_qs)

    favs = [all_stats[q.var_name]['fav_pct'] for q in sorted_qs]
    unfavs = [-all_stats[q.var_name]['unfav_pct'] for q in sorted_qs]
    labels = [f"Q{q.q_number}" for q in sorted_qs]

    ax.barh(y, favs, color=COLORS['favorable'], height=0.6, label='Favorable')
    ax.barh(y, unfavs, color=COLORS['unfavorable'], height=0.6, label='Unfavorable')

    make_bars_rounded(ax)
    ax.axvline(0, color='black', linewidth=0.8)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=10, weight='bold')

    title = custom_title if custom_title else build_chart_title(tp, "Diverging View", is_topline)
    ax.set_title(title, fontsize=14, weight='bold', loc='left')

    ax.set_xlim(-100, 100)
    ax.set_xticks([-100, -50, 0, 50, 100])
    ax.set_xticklabels(['100', '50', '0', '50', '100'])

    plt.tight_layout()
    return fig


def create_comparison_page(timepoints, all_tp_stats):
    if len(timepoints) < 2: return None
    common = None
    tp_q_map = {}
    for tp in timepoints:
        q_map = {q.q_number: all_tp_stats[tp.name][q.var_name]['fav_pct']
                 for q in tp.questions if q.is_scaled}
        tp_q_map[tp.name] = q_map
        common = set(q_map.keys()) if common is None else common & set(q_map.keys())
    if not common: return None
    sorted_qs = sorted(common, key=lambda x: (len(x), x))
    fig, ax = plt.subplots(figsize=(16, max(6, len(sorted_qs) * 0.5 + 2)))
    x = np.arange(len(sorted_qs))
    width = 0.8 / len(timepoints)
    for i, tp in enumerate(timepoints):
        vals = [tp_q_map[tp.name].get(qn, 0) for qn in sorted_qs]
        ax.bar(x + (i - len(timepoints) / 2 + 0.5) * width, vals, width, label=tp.name,
               color=COLORS['scale_colors'][i % len(COLORS['scale_colors'])], alpha=0.9)
    make_bars_rounded(ax, pad=0.02, rounding_size=0.1)
    ax.set_xticks(x)
    ax.set_xticklabels([f"Q{q}" for q in sorted_qs])
    ax.set_title("Cross-Timepoint Comparison", fontsize=14, weight='bold', loc='left')
    ax.legend()
    plt.tight_layout()
    return fig


# ═══════════════════════════════════════════════════════════════
# 6. STREAMLIT APP
# ═══════════════════════════════════════════════════════════════

def init_session_state():
    """Initialize all session state variables."""
    defaults = {
        'step': 0,
        'timepoints': [],
        'all_tp_stats': {},
        'chart_titles': {},
        'file_processed': False,
        'scales_confirmed': False,
        'subjects_confirmed': False,
        'titles_confirmed': False,
        'pdf_bytes': None,
        'pdf_name': '',
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def go_to_step(n):
    st.session_state.step = n


def render_sidebar():
    """Render the step navigation sidebar."""
    with st.sidebar:
        st.markdown("## 📊 ePRO Chart Gen v8")
        st.caption("Streamlit Edition")
        st.divider()

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


def step_upload():
    """Step 1: File upload and configuration."""
    st.header("📁 Upload & Configure")
    st.caption("Upload your ePRO workbook and set analysis options.")

    uploaded = st.file_uploader("Upload ePRO Workbook", type=["xlsx", "xls"])

    col1, col2 = st.columns(2)
    with col1:
        is_topline = st.toggle("TopLine Report", value=False)
    with col2:
        exclusions_str = st.text_input("Global Subject Exclusions", placeholder="e.g. 0042, 0091",
                                       help="These subjects will be excluded from ALL timepoints")

    if uploaded and st.button("🔍 Process Workbook", type="primary"):
        global_exclusions = [x.strip() for x in exclusions_str.split(',') if x.strip()] if exclusions_str else []

        # Save uploaded file to temp
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(uploaded.read())
            tmp_path = tmp.name

        with st.spinner("Reading workbook..."):
            sheets = detect_questionnaire_sheets(tmp_path)
            if not sheets:
                st.error("No questionnaire sheets found in workbook.")
                os.unlink(tmp_path)
                return

            timepoints = []
            all_tp_stats = {}

            for sheet_name, _ in sheets:
                ov = find_option_values_sheet(tmp_path, sheet_name)
                if not ov: continue
                tp = load_timepoint(tmp_path, sheet_name, ov, global_exclusions)
                if not tp: continue

                stats = {}
                for q in tp.questions:
                    if q.is_scaled:
                        stats[q.var_name] = compute_stats(tp, q)

                timepoints.append(tp)
                all_tp_stats[tp.name] = stats

            if not timepoints:
                st.error("No valid timepoints found.")
                os.unlink(tmp_path)
                return

            # Build chart titles
            chart_titles = {}
            for tp in timepoints:
                for suffix, text in [('dashboard', 'Summary'), ('ranked', 'Ranked Performance'),
                                     ('detailed', 'Detailed Breakdown'), ('diverging', 'Diverging View')]:
                    chart_id = f"{tp.name}_{suffix}"
                    raw_title = build_chart_title(tp, text, is_topline)
                    chart_titles[chart_id] = clean_chart_title(chart_id, raw_title)

            # Save to session
            st.session_state.timepoints = timepoints
            st.session_state.all_tp_stats = all_tp_stats
            st.session_state.chart_titles = chart_titles
            st.session_state.is_topline = is_topline
            st.session_state.tmp_path = tmp_path
            st.session_state.file_processed = True
            st.session_state.pdf_name = f"{Path(uploaded.name).stem}{'_TPL' if is_topline else ''}_Charts_v8.pdf"
            st.session_state.step = 1
            st.rerun()


def step_scales():
    """Step 2: Verify favorable logic."""
    st.header("⚖️ Verify Favorable Logic")
    st.caption("Review and adjust which response values count as 'favorable' for each scale group.")

    timepoints = st.session_state.timepoints

    for tp in timepoints:
        st.subheader(f"📋 {tp.name}")
        groups = defaultdict(list)
        for q in tp.questions:
            if q.is_scaled: groups[q.scale_signature].append(q)

        for sig, q_list in groups.items():
            ex_q = q_list[0]
            scale_str = " | ".join([f"**{k}**: {v}" for k, v in sorted(ex_q.levels.items())])

            with st.expander(f"Scale Group — {len(q_list)} questions", expanded=True):
                st.markdown(f"**Scale:** {scale_str}")

                # Show questions in this group
                q_data = []
                for q in q_list:
                    fav_str = ", ".join(map(str, q.fav_mask)) if q.fav_mask else "None"
                    q_data.append({"Q#": f"Q{q.q_number}", "Favorable": fav_str,
                                   "Question": q.question_text[:80]})
                st.dataframe(pd.DataFrame(q_data), use_container_width=True, hide_index=True)

                # Edit favorable values
                current_fav = ", ".join(map(str, ex_q.fav_mask))
                key = f"fav_{tp.name}_{id(ex_q)}"
                new_fav = st.text_input(f"Favorable values for this group", value=current_fav, key=key,
                                        help="Comma-separated values, e.g. 4, 5")

                # Parse and apply
                try:
                    parsed = [int(x.strip()) for x in new_fav.split(',') if x.strip()]
                    valid = [p for p in parsed if p in ex_q.levels]
                    if valid:
                        for q in q_list:
                            q.fav_mask = valid
                except:
                    pass

    if st.button("✅ Confirm Scales", type="primary"):
        # Recompute stats with updated masks
        for tp in st.session_state.timepoints:
            stats = {}
            for q in tp.questions:
                if q.is_scaled:
                    stats[q.var_name] = compute_stats(tp, q)
            st.session_state.all_tp_stats[tp.name] = stats

        st.session_state.scales_confirmed = True
        st.session_state.step = 2
        st.rerun()


def step_subjects():
    """Step 3: Subject inclusion."""
    st.header("👥 Subject Inclusion")
    st.caption("Review non-completed subjects and optionally include them in the analysis.")

    timepoints = st.session_state.timepoints
    any_non_completed = False

    for tp in timepoints:
        if not tp.non_completed: continue
        any_non_completed = True

        st.subheader(f"📋 {tp.name}")
        st.markdown(f"**Completed:** {tp.n_completed} / {tp.n_total}")

        non_comp_with_data = [s for s in tp.non_completed if s['coverage_pct'] > 0]
        if not non_comp_with_data:
            st.info("No non-completed subjects with data to include.")
            continue

        for s in sorted(non_comp_with_data, key=lambda x: x['sid']):
            col1, col2, col3, col4 = st.columns([1, 2, 2, 3])
            key = f"inc_{tp.name}_{s['sid']}"
            with col1:
                include = st.checkbox("Include", key=key, value=False)
            with col2:
                st.markdown(f"**{s['sid']}**")
            with col3:
                st.caption(s['status'])
            with col4:
                st.progress(s['coverage_pct'] / 100, text=f"{s['coverage_pct']:.0f}% data")

            if include and s['sid'] not in tp.included_subjects:
                tp.included_subjects.append(s['sid'])
                tp.n_included = len(tp.included_subjects)

    if not any_non_completed:
        st.success("All subjects across all timepoints are completed. Nothing to review here.")

    if st.button("✅ Confirm Subjects", type="primary"):
        # Recompute stats with updated inclusions
        for tp in st.session_state.timepoints:
            stats = {}
            for q in tp.questions:
                if q.is_scaled:
                    stats[q.var_name] = compute_stats(tp, q)
            st.session_state.all_tp_stats[tp.name] = stats

        st.session_state.subjects_confirmed = True
        st.session_state.step = 3
        st.rerun()


def step_titles():
    """Step 4: Review and edit chart titles."""
    st.header("✏️ Review Chart Titles")
    st.caption("Auto-cleaned titles are shown below. Edit any that need adjustment.")

    titles = st.session_state.chart_titles
    updated = {}

    for chart_id, title in titles.items():
        new_title = st.text_input(chart_id, value=title, key=f"title_{chart_id}")
        updated[chart_id] = new_title

    if st.button("✅ Confirm Titles", type="primary"):
        st.session_state.chart_titles = updated
        st.session_state.titles_confirmed = True
        st.session_state.step = 4
        st.rerun()


def step_generate():
    """Step 5: Generate and download PDF."""
    st.header("🚀 Generate PDF")

    timepoints = st.session_state.timepoints
    all_tp_stats = st.session_state.all_tp_stats
    titles = st.session_state.chart_titles
    is_topline = st.session_state.get('is_topline', False)

    # Summary cards
    total_charts = len(titles)
    total_subjects = sum(tp.n_included for tp in timepoints)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Timepoints", len(timepoints))
    col2.metric("Total Charts", total_charts)
    col3.metric("Included Subjects", total_subjects)
    col4.metric("Center", CENTER_FILTER)

    st.divider()

    # Show timepoint summaries
    with st.expander("📊 Timepoint Details"):
        for tp in timepoints:
            st.markdown(f"**{tp.name}** — {tp.n_included} included, "
                        f"{len(tp.dropped_ids)} dropped, {tp.n_completed} completed")

    if st.session_state.pdf_bytes:
        st.success("✅ PDF generated successfully!")
        st.download_button(
            label="📥 Download PDF",
            data=st.session_state.pdf_bytes,
            file_name=st.session_state.pdf_name,
            mime="application/pdf",
            type="primary",
        )
        if st.button("🔄 Regenerate"):
            st.session_state.pdf_bytes = None
            st.rerun()
        return

    if st.button("🚀 Generate PDF", type="primary"):
        pdf_buffer = BytesIO()
        progress = st.progress(0, text="Generating charts...")

        with PdfPages(pdf_buffer) as pdf:
            chart_count = 0
            total = len(timepoints) * 4 + (1 if len(timepoints) >= 2 else 0)

            for tp in timepoints:
                stats = all_tp_stats[tp.name]

                fig = create_dashboard_page(tp, stats, is_topline,
                                            custom_title=titles.get(f"{tp.name}_dashboard"))
                if fig: pdf.savefig(fig); plt.close(fig)
                chart_count += 1
                progress.progress(chart_count / total, text=f"Generating {tp.name} dashboard...")

                fig = create_ranked_chart(tp, stats, is_topline,
                                          custom_title=titles.get(f"{tp.name}_ranked"))
                if fig: pdf.savefig(fig); plt.close(fig)
                chart_count += 1
                progress.progress(chart_count / total, text=f"Generating {tp.name} ranked...")

                fig = create_detailed_bars(tp, stats, is_topline,
                                           custom_title=titles.get(f"{tp.name}_detailed"))
                if fig: pdf.savefig(fig); plt.close(fig)
                chart_count += 1
                progress.progress(chart_count / total, text=f"Generating {tp.name} detailed...")

                fig = create_diverging_chart(tp, stats, is_topline,
                                             custom_title=titles.get(f"{tp.name}_diverging"))
                if fig: pdf.savefig(fig); plt.close(fig)
                chart_count += 1
                progress.progress(chart_count / total, text=f"Generating {tp.name} diverging...")

            if len(timepoints) >= 2:
                fig = create_comparison_page(timepoints, all_tp_stats)
                if fig: pdf.savefig(fig); plt.close(fig)
                chart_count += 1
                progress.progress(1.0, text="Comparison chart complete!")

        progress.progress(1.0, text="✅ Done!")
        st.session_state.pdf_bytes = pdf_buffer.getvalue()
        st.rerun()


# ═══════════════════════════════════════════════════════════════
# 7. MAIN APP ENTRY
# ═══════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="ePRO Chart Generator v8",
        page_icon="📊",
        layout="wide",
    )

    init_session_state()
    render_sidebar()

    step_funcs = [step_upload, step_scales, step_subjects, step_titles, step_generate]
    step_funcs[st.session_state.step]()


if __name__ == "__main__":
    main()