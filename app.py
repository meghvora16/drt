import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pyimpspec as pyi
import io, zipfile, os

# ── Page config ────────────────────────────────────────────────
st.set_page_config(
    page_title="EIS · DRT Analyser",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ─────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=Space+Mono:wght@400;700&display=swap');

:root {
    --bg:        #0a0e1a;
    --surface:   #111827;
    --surface2:  #1c2537;
    --border:    #2a3a54;
    --accent:    #00d4ff;
    --accent2:   #ff6b35;
    --accent3:   #a855f7;
    --text:      #e2e8f0;
    --muted:     #64748b;
    --green:     #22c55e;
    --red:       #ef4444;
    --yellow:    #f59e0b;
}

html, body, [data-testid="stAppViewContainer"] {
    background: var(--bg) !important;
    color: var(--text) !important;
    font-family: 'Space Mono', monospace;
}

[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border);
}

h1, h2, h3 { font-family: 'Syne', sans-serif !important; }

.stButton > button {
    background: linear-gradient(135deg, var(--accent), #0099cc) !important;
    color: #0a0e1a !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 6px !important;
    padding: 0.6rem 1.4rem !important;
    letter-spacing: 0.05em !important;
    transition: all 0.2s !important;
}
.stButton > button:hover { transform: translateY(-1px); box-shadow: 0 4px 20px rgba(0,212,255,0.4) !important; }

.metric-card {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 0.8rem;
}
.metric-label {
    font-family: 'Space Mono', monospace;
    font-size: 0.72rem;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin-bottom: 0.3rem;
}
.metric-value {
    font-family: 'Syne', sans-serif;
    font-size: 1.6rem;
    font-weight: 800;
    color: var(--accent);
}
.metric-sub {
    font-size: 0.72rem;
    color: var(--muted);
    margin-top: 0.1rem;
}

.section-header {
    font-family: 'Syne', sans-serif;
    font-size: 1.1rem;
    font-weight: 700;
    color: var(--accent);
    text-transform: uppercase;
    letter-spacing: 0.12em;
    border-bottom: 1px solid var(--border);
    padding-bottom: 0.5rem;
    margin: 1.5rem 0 1rem 0;
}

.badge {
    display: inline-block;
    padding: 0.2rem 0.7rem;
    border-radius: 4px;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.08em;
    margin-right: 0.4rem;
}
.badge-cut  { background:#c0392b22; color:#ef4444; border:1px solid #c0392b55; }
.badge-perf { background:#8e44ad22; color:#a855f7; border:1px solid #8e44ad55; }
.badge-bulk { background:#1a527622; color:#00d4ff; border:1px solid #2471a355; }
.badge-near { background:#922b2122; color:#ff6b35; border:1px solid #922b2155; }
.badge-far  { background:#1e844922; color:#22c55e; border:1px solid #1e844955; }

.stTabs [data-baseweb="tab"] {
    font-family: 'Syne', sans-serif !important;
    font-weight: 600 !important;
    color: var(--muted) !important;
    border-bottom: 2px solid transparent !important;
}
.stTabs [aria-selected="true"] {
    color: var(--accent) !important;
    border-bottom: 2px solid var(--accent) !important;
}
.stTabs [data-baseweb="tab-panel"] { padding-top: 1rem !important; }

[data-testid="stFileUploader"] {
    background: var(--surface2) !important;
    border: 1px dashed var(--border) !important;
    border-radius: 8px !important;
}

.stSelectbox > div, .stMultiselect > div {
    background: var(--surface2) !important;
    border-color: var(--border) !important;
}

.info-box {
    background: rgba(0,212,255,0.07);
    border-left: 3px solid var(--accent);
    padding: 0.8rem 1rem;
    border-radius: 0 8px 8px 0;
    margin: 0.8rem 0;
    font-size: 0.85rem;
    color: var(--text);
}
.warn-box {
    background: rgba(245,158,11,0.07);
    border-left: 3px solid var(--yellow);
    padding: 0.8rem 1rem;
    border-radius: 0 8px 8px 0;
    margin: 0.8rem 0;
    font-size: 0.85rem;
}
.peak-chip {
    display: inline-block;
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 20px;
    padding: 0.25rem 0.8rem;
    font-size: 0.78rem;
    margin: 0.2rem;
    color: var(--text);
}
</style>
""", unsafe_allow_html=True)

# ── Helpers ────────────────────────────────────────────────────
PLOTLY_LAYOUT = dict(
    paper_bgcolor='#111827', plot_bgcolor='#111827',
    font=dict(family='Space Mono', color='#e2e8f0', size=11),
    xaxis=dict(gridcolor='#1c2537', linecolor='#2a3a54', zerolinecolor='#2a3a54'),
    yaxis=dict(gridcolor='#1c2537', linecolor='#2a3a54', zerolinecolor='#2a3a54'),
    legend=dict(bgcolor='#1c2537', bordercolor='#2a3a54', borderwidth=1),
    margin=dict(l=60,r=30,t=50,b=50),
)

def assign_process(tau):
    if   tau < 1e-5:  return 'HF oxide / contact loop'
    elif tau < 1e-4:  return 'Double-layer / HF oxide'
    elif tau < 1e-3:  return 'Outer passive oxide layer'
    elif tau < 1e-2:  return 'Charge transfer — outer film'
    elif tau < 1e-1:  return 'Charge transfer — inner Cr₂O₃'
    elif tau < 1.0:   return 'Ion migration / film repair'
    else:             return 'Slow dissolution / diffusion'

def load_eis_file(uploaded_file):
    df = pd.read_excel(uploaded_file, header=0)
    df.columns = ['Index','Freq','Zreal','Zimag','Zmod','Phase','Time']
    return df

def run_drt(df):
    pyi_df = pd.DataFrame({
        'frequency': df['Freq'].values,
        "z'":        df['Zreal'].values,
        "-z''":      df['Zimag'].values,
    })
    ds  = pyi.dataframe_to_data_sets(pyi_df, path='eis', label='EIS')[0]
    drt = pyi.calculate_drt(ds, method='tr-nnls')
    return ds, drt

def get_eis_params(df):
    rs  = df.loc[df['Freq']==100000,'Zreal'].values[0] if 100000 in df['Freq'].values else df.iloc[0]['Zreal']
    rp  = df.loc[df['Freq']==0.1,'Zreal'].values[0] - rs if 0.1 in df['Freq'].values else df.iloc[-1]['Zreal'] - rs
    zml = df.loc[df['Freq']==0.1,'Zmod'].values[0]  if 0.1 in df['Freq'].values else df.iloc[-1]['Zmod']
    phl = df.loc[df['Freq']==0.1,'Phase'].values[0] if 0.1 in df['Freq'].values else df.iloc[-1]['Phase']
    return dict(Rs=rs, Rp=rp, Zmod_lf=zml, Phase_lf=phl)

PT_COLORS = {
    1:'#ef4444', 2:'#f97316', 3:'#3b82f6', 4:'#2563eb',
    5:'#22c55e', 6:'#16a34a', 7:'#a855f7', 8:'#7c3aed',
    9:'#f59e0b', 10:'#d97706',
}

# ── Session state ──────────────────────────────────────────────
if 'data' not in st.session_state:
    st.session_state.data = {}       # {pt_num: {df, ds, drt, params}}
if 'sample_name' not in st.session_state:
    st.session_state.sample_name = 'My Sample'
if 'near_pts' not in st.session_state:
    st.session_state.near_pts = [1,2,9,10]
if 'zone_map' not in st.session_state:
    st.session_state.zone_map = {}   # {pt: 'Cut'|'Perf'|'Bulk'}

# ── Sidebar ────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div style="font-family:Syne;font-size:1.5rem;font-weight:800;color:#00d4ff;letter-spacing:0.05em;">⚡ EIS · DRT<br><span style="font-size:0.9rem;color:#64748b;font-weight:400;">Analyser</span></div>', unsafe_allow_html=True)
    st.markdown('---')

    st.markdown('<div class="section-header">Sample Config</div>', unsafe_allow_html=True)
    st.session_state.sample_name = st.text_input('Sample name', value=st.session_state.sample_name)

    material = st.selectbox('Material', ['Stainless Steel 316L','Stainless Steel 304','Carbon Steel','Aluminium Alloy','Custom'])
    treatment = st.selectbox('Surface treatment', [
        'Laser Cut Only',
        'Brushed + Pickled',
        'Brushed + Pickled + Passivated',
        'As-received',
        'Custom',
    ])

    st.markdown('<div class="section-header">Point Geometry</div>', unsafe_allow_html=True)
    near_input = st.multiselect('Near laser cut (★)', options=list(range(1,11)), default=[1,2,9,10])
    st.session_state.near_pts = near_input

    zone_type = st.radio('Zone type for near points', ['Cut edge + Perf hole (1,2=Cut | 9,10=Perf)', 'All near = same zone'])
    use_split = zone_type.startswith('Cut edge')

    st.markdown('<div class="section-header">DRT Settings</div>', unsafe_allow_html=True)
    drt_method = st.selectbox('DRT Method', ['tr-nnls','tr-rbf','bht'], help='TR-NNLS recommended for most cases')
    show_norm  = st.checkbox('Normalise DRT (shape comparison)', value=False)

    st.markdown('<div class="section-header">Upload Files</div>', unsafe_allow_html=True)
    st.markdown('<div class="info-box">Upload one .xlsx per measurement point. Name format: eis1.xlsx, eis2.xlsx etc.</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader('EIS files (.xlsx)', type='xlsx', accept_multiple_files=True)

    if uploaded and st.button('▶  Run DRT Analysis', use_container_width=True):
        progress = st.progress(0)
        status   = st.empty()
        st.session_state.data = {}

        for i, f in enumerate(uploaded):
            # extract point number from filename
            name = f.name
            num_str = ''.join(filter(str.isdigit, name.replace('.xlsx','')))
            pt_num  = int(num_str) if num_str else (i+1)

            status.markdown(f'<div class="info-box">Processing point {pt_num}…</div>', unsafe_allow_html=True)
            try:
                df     = load_eis_file(f)
                ds, drt = run_drt(df)
                params  = get_eis_params(df)

                # zone assignment
                if use_split:
                    if pt_num in [1,2]:   zone = 'Cut'
                    elif pt_num in [9,10]: zone = 'Perf'
                    elif pt_num in near_input: zone = 'Near'
                    else: zone = 'Bulk'
                else:
                    zone = 'Near' if pt_num in near_input else 'Bulk'

                st.session_state.data[pt_num] = dict(df=df, ds=ds, drt=drt, params=params, zone=zone)
            except Exception as e:
                st.error(f'Error on {name}: {e}')

            progress.progress((i+1)/len(uploaded))

        status.empty(); progress.empty()
        st.success(f'✓  {len(st.session_state.data)} points loaded')

# ── Main area ──────────────────────────────────────────────────
data = st.session_state.data
near_pts = st.session_state.near_pts

# Header
st.markdown(f"""
<div style="display:flex;align-items:baseline;gap:1rem;margin-bottom:0.5rem;">
  <span style="font-family:Syne;font-size:2rem;font-weight:800;color:#e2e8f0;">{st.session_state.sample_name}</span>
  <span style="font-family:Space Mono;font-size:0.8rem;color:#64748b;">{treatment} · {material}</span>
</div>
""", unsafe_allow_html=True)

if not data:
    # Landing state
    col1,col2,col3 = st.columns(3)
    for col,title,desc,icon in [
        (col1,'Upload','Load up to 10 EIS .xlsx files from sidebar','📁'),
        (col2,'Configure','Set point geometry — cut edge, perf hole, bulk','⚙️'),
        (col3,'Analyse','DRT, Nyquist, Bode, spatial maps and rankings','📊'),
    ]:
        with col:
            st.markdown(f"""
            <div class="metric-card" style="text-align:center;padding:2rem;">
              <div style="font-size:2.5rem;margin-bottom:0.8rem;">{icon}</div>
              <div style="font-family:Syne;font-size:1.1rem;font-weight:700;color:#00d4ff;">{title}</div>
              <div style="font-size:0.82rem;color:#64748b;margin-top:0.4rem;">{desc}</div>
            </div>""", unsafe_allow_html=True)
    st.stop()

pts_sorted = sorted(data.keys())
far_pts    = [p for p in pts_sorted if p not in near_pts]

# ── Tabs ───────────────────────────────────────────────────────
tab_overview, tab_drt, tab_nyquist, tab_bode, tab_spatial, tab_compare, tab_table = st.tabs([
    '📊 Overview', '🌊 DRT', '🔵 Nyquist', '📈 Bode', '🗺 Spatial', '⚖️ Compare', '📋 Data Table'
])

# ═══════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ═══════════════════════════════════════════════════════════════
with tab_overview:
    st.markdown('<div class="section-header">Key Parameters — All Points</div>', unsafe_allow_html=True)

    # Metric cards top row
    cols = st.columns(4)
    all_rp = [data[p]['params']['Rp'] for p in pts_sorted]
    all_rs = [data[p]['params']['Rs'] for p in pts_sorted]
    all_rprs = [data[p]['params']['Rp']/data[p]['params']['Rs'] for p in pts_sorted]

    near_rp  = np.mean([data[p]['params']['Rp'] for p in pts_sorted if p in near_pts]) if any(p in near_pts for p in pts_sorted) else 0
    far_rp   = np.mean([data[p]['params']['Rp'] for p in pts_sorted if p not in near_pts]) if any(p not in near_pts for p in pts_sorted) else 0
    ratio    = far_rp/near_rp if near_rp > 0 else 0

    rs_cv = np.std(all_rs)/np.mean(all_rs)*100 if all_rs else 0
    rp_cv = np.std(all_rp)/np.mean(all_rp)*100 if all_rp else 0

    def fmt_val(v):
        if v >= 1e6:   return f'{v/1e6:.2f} MΩ'
        elif v >= 1e3: return f'{v/1e3:.1f} kΩ'
        return f'{v:.0f} Ω'

    with cols[0]:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Avg Rs</div><div class="metric-value">{fmt_val(np.mean(all_rs))}</div><div class="metric-sub">CV = {rs_cv:.1f}%</div></div>', unsafe_allow_html=True)
    with cols[1]:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Avg Rp (near)</div><div class="metric-value">{fmt_val(near_rp)}</div><div class="metric-sub">★ points {near_pts}</div></div>', unsafe_allow_html=True)
    with cols[2]:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Avg Rp (far)</div><div class="metric-value">{fmt_val(far_rp)}</div><div class="metric-sub">bulk points</div></div>', unsafe_allow_html=True)
    with cols[3]:
        col_ratio = '#22c55e' if ratio < 1.5 else ('#f59e0b' if ratio < 3 else '#ef4444')
        verdict   = 'Cut BEFORE proc.' if ratio < 1.5 else 'Cut AFTER proc.'
        st.markdown(f'<div class="metric-card"><div class="metric-label">Far/Near Rp Ratio</div><div class="metric-value" style="color:{col_ratio};">{ratio:.2f}×</div><div class="metric-sub">{verdict}</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-header">Per-Point Summary</div>', unsafe_allow_html=True)

    # Summary table
    rows = []
    for p in pts_sorted:
        d  = data[p]
        pm = d['params']
        taus, gammas = d['drt'].get_peaks()
        dom_tau = taus[np.argmax(gammas)] if len(taus)>0 else 0
        rows.append({
            'Point': p,
            'Zone': d['zone'],
            'Rs (kΩ)': round(pm['Rs']/1000,1),
            'Rp (kΩ)': round(pm['Rp']/1000,1),
            'Rp/Rs': round(pm['Rp']/pm['Rs'],1),
            '|Z|@0.1Hz (kΩ)': round(pm['Zmod_lf']/1000,1),
            'Phase@0.1Hz (°)': round(pm['Phase_lf'],1),
            '# DRT peaks': len(taus),
            'Dom. τ (s)': round(float(dom_tau),4) if dom_tau else '-',
        })

    df_summary = pd.DataFrame(rows)
    st.dataframe(df_summary, use_container_width=True, hide_index=True)

    # Rs/Rp gradient insight
    rs_verdict = 'Thermal oxide gradient present → laser cut likely LAST step (no processing)' if rs_cv > 20 else 'Rs uniform → thermal oxide removed by pickling/processing'
    rp_verdict = 'Rp scattered → patchy passive film' if rp_cv > 30 else 'Rp uniform → homogeneous surface'
    st.markdown(f'<div class="info-box">Rs CV = {rs_cv:.1f}% — {rs_verdict}<br>Rp CV = {rp_cv:.1f}% — {rp_verdict}</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# TAB 2 — DRT
# ═══════════════════════════════════════════════════════════════
with tab_drt:
    st.markdown('<div class="section-header">Distribution of Relaxation Times (TR-NNLS)</div>', unsafe_allow_html=True)

    col_ctrl, col_main = st.columns([1,3])
    with col_ctrl:
        pts_select = st.multiselect('Points to show', pts_sorted, default=pts_sorted, key='drt_pts')
        show_peaks = st.checkbox('Mark peaks', value=True)
        show_avg   = st.checkbox('Show group average', value=True)
        log_gamma  = st.checkbox('Log γ axis', value=False)

    ref_taus = data[pts_sorted[0]]['drt'].get_time_constants()

    fig_drt = go.Figure()

    for p in pts_select:
        d      = data[p]
        taus   = d['drt'].get_time_constants()
        gammas = d['drt'].get_gammas()
        norm   = gammas.max() if (show_norm and gammas.max()>0) else 1
        color  = PT_COLORS.get(p,'#888')
        ls     = 'solid' if p in near_pts else 'dash'

        fig_drt.add_trace(go.Scatter(
            x=taus, y=gammas/norm,
            mode='lines', name=f'Pt {p} {"★" if p in near_pts else ""}',
            line=dict(color=color, width=2, dash=ls),
            hovertemplate=f'Pt {p}<br>τ=%{{x:.3e}} s<br>γ=%{{y:.2e}}<extra></extra>'
        ))

        if show_peaks:
            ptaus, pgammas = d['drt'].get_peaks()
            if len(ptaus) > 0:
                fig_drt.add_trace(go.Scatter(
                    x=ptaus, y=pgammas/norm,
                    mode='markers', showlegend=False,
                    marker=dict(color=color, size=9, symbol='circle',
                                line=dict(color='white',width=1.5)),
                    hovertemplate=f'Pt {p} peak<br>τ=%{{x:.3e}} s<br>γ=%{{y:.2e}}<br>{[assign_process(t) for t in ptaus]}<extra></extra>'
                ))

    # Group averages
    if show_avg and len(near_pts)>0 and len(far_pts)>0:
        near_sel = [p for p in near_pts if p in pts_select and p in data]
        far_sel  = [p for p in far_pts  if p in pts_select and p in data]
        if near_sel:
            ng = np.array([data[p]['drt'].get_gammas() for p in near_sel]).mean(0)
            norm_n = ng.max() if (show_norm and ng.max()>0) else 1
            fig_drt.add_trace(go.Scatter(x=ref_taus, y=ng/norm_n, mode='lines',
                name='Near avg ★', line=dict(color='#ff6b35',width=3.5,dash='solid'),
                hovertemplate='Near avg<br>τ=%{x:.3e}<br>γ=%{y:.2e}<extra></extra>'))
        if far_sel:
            fg = np.array([data[p]['drt'].get_gammas() for p in far_sel]).mean(0)
            norm_f = fg.max() if (show_norm and fg.max()>0) else 1
            fig_drt.add_trace(go.Scatter(x=ref_taus, y=fg/norm_f, mode='lines',
                name='Far avg', line=dict(color='#00d4ff',width=3.5,dash='dot'),
                hovertemplate='Far avg<br>τ=%{x:.3e}<br>γ=%{y:.2e}<extra></extra>'))

    # τ region vertical lines
    tau_regions = [(2e-6,'HF'),(1e-4,'Outer\noxide'),(5e-3,'Charge\ntransfer'),(8e-2,'Ion\nmigration'),(1.3,'Dissolution')]
    for tau,lbl in tau_regions:
        fig_drt.add_vline(x=tau, line_color='#2a3a54', line_dash='dot', line_width=1)
        fig_drt.add_annotation(x=np.log10(tau), y=1.02, xref='x', yref='paper',
                               text=lbl, showarrow=False, font=dict(size=9,color='#64748b'),
                               xanchor='left')

    y_label = 'γ / γ_max (normalised)' if show_norm else 'γ (Ω)'
    fig_drt.update_layout(**PLOTLY_LAYOUT,
        title='Distribution of Relaxation Times — TR-NNLS',
        xaxis=dict(**PLOTLY_LAYOUT['xaxis'], type='log', title='Time Constant τ (s)'),
        yaxis=dict(**PLOTLY_LAYOUT['yaxis'], type='log' if log_gamma else 'linear', title=y_label),
        height=500, legend=dict(**PLOTLY_LAYOUT['legend'], orientation='v'))
    st.plotly_chart(fig_drt, use_container_width=True)

    # Peak detail table
    st.markdown('<div class="section-header">DRT Peak Assignments</div>', unsafe_allow_html=True)
    peak_rows = []
    for p in pts_select:
        d = data[p]
        taus, gammas = d['drt'].get_peaks()
        for i, idx in enumerate(np.argsort(taus)):
            tau   = float(taus[idx])
            gamma = float(gammas[idx])
            peak_rows.append({
                'Point': p, 'Zone': d['zone'], 'Peak #': i+1,
                'τ (s)': f'{tau:.4e}', 'τ (ms)': round(tau*1000,4),
                'γ (Ω)': round(gamma,0),
                'Process': assign_process(tau),
            })
    if peak_rows:
        st.dataframe(pd.DataFrame(peak_rows), use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════════
# TAB 3 — NYQUIST
# ═══════════════════════════════════════════════════════════════
with tab_nyquist:
    st.markdown('<div class="section-header">Nyquist Plot</div>', unsafe_allow_html=True)

    pts_ny = st.multiselect('Points', pts_sorted, default=pts_sorted, key='ny_pts')
    unit_ny= st.radio('Unit', ['Ω','kΩ','MΩ'], horizontal=True, key='ny_unit')
    div_ny = {'Ω':1,'kΩ':1e3,'MΩ':1e6}[unit_ny]

    fig_ny = go.Figure()
    for p in pts_ny:
        d  = data[p]
        df = d['df']
        ls = 'solid' if p in near_pts else 'dash'
        fig_ny.add_trace(go.Scatter(
            x=df['Zreal']/div_ny, y=df['Zimag']/div_ny,
            mode='lines+markers', name=f'Pt {p} {"★" if p in near_pts else ""}',
            line=dict(color=PT_COLORS.get(p,'#888'), width=2, dash=ls),
            marker=dict(size=4),
            hovertemplate=f"Pt {p}<br>Z'=%{{x:.3f}} {unit_ny}<br>-Z''=%{{y:.3f}} {unit_ny}<extra></extra>"
        ))
    fig_ny.update_layout(**PLOTLY_LAYOUT,
        title='Nyquist Plot',
        xaxis=dict(**PLOTLY_LAYOUT['xaxis'], title=f"Z' ({unit_ny})"),
        yaxis=dict(**PLOTLY_LAYOUT['yaxis'], title=f"-Z'' ({unit_ny})"),
        height=520)
    st.plotly_chart(fig_ny, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# TAB 4 — BODE
# ═══════════════════════════════════════════════════════════════
with tab_bode:
    st.markdown('<div class="section-header">Bode Plots</div>', unsafe_allow_html=True)
    pts_bd = st.multiselect('Points', pts_sorted, default=pts_sorted, key='bd_pts')

    fig_bd = make_subplots(rows=1,cols=2,
                            subplot_titles=['|Z| vs Frequency','Phase vs Frequency'])

    for p in pts_bd:
        d  = data[p]
        df = d['df']
        ls = 'solid' if p in near_pts else 'dash'
        color = PT_COLORS.get(p,'#888')
        show_leg = True

        fig_bd.add_trace(go.Scatter(
            x=df['Freq'], y=df['Zmod']/1000,
            mode='lines', name=f'Pt {p} {"★" if p in near_pts else ""}',
            line=dict(color=color,width=2,dash=ls), showlegend=show_leg,
            hovertemplate=f'Pt {p}<br>f=%{{x:.1f}} Hz<br>|Z|=%{{y:.2f}} kΩ<extra></extra>'
        ), row=1,col=1)

        fig_bd.add_trace(go.Scatter(
            x=df['Freq'], y=df['Phase'],
            mode='lines', name=f'Pt {p}',
            line=dict(color=color,width=2,dash=ls), showlegend=False,
            hovertemplate=f'Pt {p}<br>f=%{{x:.1f}} Hz<br>Phase=%{{y:.2f}}°<extra></extra>'
        ), row=1,col=2)

    fig_bd.update_xaxes(type='log',title='Frequency (Hz)',gridcolor='#1c2537',linecolor='#2a3a54')
    fig_bd.update_yaxes(gridcolor='#1c2537',linecolor='#2a3a54')
    fig_bd.update_yaxes(type='log',title='|Z| (kΩ)',row=1,col=1)
    fig_bd.update_yaxes(title='−Phase (°)',row=1,col=2)
    fig_bd.update_layout(paper_bgcolor='#111827',plot_bgcolor='#111827',
                          font=dict(family='Space Mono',color='#e2e8f0',size=11),
                          height=480,showlegend=True,
                          legend=dict(bgcolor='#1c2537',bordercolor='#2a3a54',borderwidth=1))
    st.plotly_chart(fig_bd, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# TAB 5 — SPATIAL MAP
# ═══════════════════════════════════════════════════════════════
with tab_spatial:
    st.markdown('<div class="section-header">Spatial Maps — Rs, Rp, Rp/Rs across Sample</div>', unsafe_allow_html=True)

    pt_nums = pts_sorted
    rs_vals   = [data[p]['params']['Rs']/1000       for p in pt_nums]
    rp_vals   = [data[p]['params']['Rp']/1000       for p in pt_nums]
    rprs_vals = [data[p]['params']['Rp']/data[p]['params']['Rs'] for p in pt_nums]
    pt_colors_spatial = ['#ef4444' if p in [1,2] else ('#a855f7' if p in [9,10] else '#00d4ff') for p in pt_nums]
    xlabels   = [f'Pt {p}{"★" if p in near_pts else ""}' for p in pt_nums]

    fig_sp = make_subplots(rows=3,cols=1,shared_xaxes=True,vertical_spacing=0.08,
                            subplot_titles=['Rs (kΩ) — Solution/HF Resistance',
                                            'Rp (kΩ) — Polarisation Resistance',
                                            'Rp/Rs — Passive Film Quality (electrolyte-normalised)'])

    for row,vals,label in [(1,rs_vals,'Rs (kΩ)'),(2,rp_vals,'Rp (kΩ)'),(3,rprs_vals,'Rp/Rs')]:
        avg_near = np.mean([vals[pt_nums.index(p)] for p in near_pts if p in pt_nums]) if near_pts else 0
        avg_far  = np.mean([vals[pt_nums.index(p)] for p in pt_nums if p not in near_pts]) if far_pts else 0

        fig_sp.add_trace(go.Bar(
            x=xlabels, y=vals, marker_color=pt_colors_spatial,
            marker_line_color='#2a3a54', marker_line_width=1,
            text=[f'{v:.1f}' for v in vals], textposition='outside',
            textfont=dict(size=9,color='#e2e8f0'),
            showlegend=False,
            hovertemplate='%{x}<br>'+label+'=%{y:.2f}<extra></extra>'
        ), row=row, col=1)

        if avg_near:
            fig_sp.add_hline(y=avg_near, line_color='#ff6b35', line_dash='dash', line_width=1.5,
                             annotation_text=f'Near avg {avg_near:.1f}',
                             annotation_font=dict(color='#ff6b35',size=9), row=row, col=1)
        if avg_far:
            fig_sp.add_hline(y=avg_far, line_color='#00d4ff', line_dash='dash', line_width=1.5,
                             annotation_text=f'Far avg {avg_far:.1f}',
                             annotation_font=dict(color='#00d4ff',size=9), row=row, col=1)

    fig_sp.update_xaxes(gridcolor='#1c2537', linecolor='#2a3a54')
    fig_sp.update_yaxes(gridcolor='#1c2537', linecolor='#2a3a54')
    fig_sp.update_layout(paper_bgcolor='#111827', plot_bgcolor='#111827',
                          font=dict(family='Space Mono',color='#e2e8f0',size=10),
                          height=750, showlegend=False,
                          margin=dict(l=70,r=60,t=60,b=40))
    st.plotly_chart(fig_sp, use_container_width=True)

    rs_cv_sp = np.std(rs_vals)/np.mean(rs_vals)*100
    timing = ('⚠️ STEEP Rs gradient (CV {:.1f}%) → thermal oxide intact → no pickling done yet'.format(rs_cv_sp)
              if rs_cv_sp > 20 else
              '✓ Flat Rs (CV {:.1f}%) → thermal oxide removed → surface has been pickled'.format(rs_cv_sp))
    st.markdown(f'<div class="info-box">{timing}</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# TAB 6 — COMPARE ZONES
# ═══════════════════════════════════════════════════════════════
with tab_compare:
    st.markdown('<div class="section-header">Zone Comparison — DRT Shape</div>', unsafe_allow_html=True)

    zones_present = list(set(data[p]['zone'] for p in pts_sorted))
    zone_colors_map = {'Cut':'#ef4444','Perf':'#a855f7','Near':'#ff6b35','Bulk':'#00d4ff','Far':'#22c55e'}

    fig_cmp = go.Figure()
    ref_taus = data[pts_sorted[0]]['drt'].get_time_constants()

    for zone in zones_present:
        zone_pts = [p for p in pts_sorted if data[p]['zone']==zone]
        if not zone_pts: continue
        g_arr = np.array([data[p]['drt'].get_gammas() for p in zone_pts])
        avg_g = g_arr.mean(0); std_g = g_arr.std(0)
        norm  = avg_g.max() if avg_g.max()>0 else 1
        color = zone_colors_map.get(zone,'#888')

        fig_cmp.add_trace(go.Scatter(
            x=ref_taus, y=avg_g/norm, mode='lines',
            name=f'{zone} (avg, n={len(zone_pts)})',
            line=dict(color=color,width=3),
            hovertemplate=f'{zone} avg<br>τ=%{{x:.3e}}<br>γ/γmax=%{{y:.3f}}<extra></extra>'
        ))
        fig_cmp.add_trace(go.Scatter(
            x=np.concatenate([ref_taus, ref_taus[::-1]]),
            y=np.concatenate([(avg_g+std_g)/norm, (avg_g-std_g)[::-1]/norm]),
            fill='toself', fillcolor=color.replace('#',f'rgba(') + '18)',
            line=dict(color='rgba(0,0,0,0)'), showlegend=False,
            hoverinfo='skip'
        ))

    for tau,lbl in tau_regions:
        fig_cmp.add_vline(x=tau, line_color='#2a3a54', line_dash='dot', line_width=1)

    fig_cmp.update_layout(**PLOTLY_LAYOUT,
        title='DRT Shape by Zone — Average ± 1σ (normalised)',
        xaxis=dict(**PLOTLY_LAYOUT['xaxis'], type='log', title='τ (s)'),
        yaxis=dict(**PLOTLY_LAYOUT['yaxis'], title='γ / γ_max'),
        height=480)
    st.plotly_chart(fig_cmp, use_container_width=True)

    # Rp/Rs by zone
    st.markdown('<div class="section-header">Rp/Rs by Zone — Ranking Metric</div>', unsafe_allow_html=True)
    zone_rprs = {}
    for zone in zones_present:
        zone_pts = [p for p in pts_sorted if data[p]['zone']==zone]
        vals = [data[p]['params']['Rp']/data[p]['params']['Rs'] for p in zone_pts]
        zone_rprs[zone] = (np.mean(vals), np.std(vals), zone_pts)

    fig_rank = go.Figure()
    for zone,(avg,std,zpts) in zone_rprs.items():
        color = zone_colors_map.get(zone,'#888')
        fig_rank.add_trace(go.Bar(
            x=[zone], y=[avg],
            error_y=dict(type='data',array=[std],visible=True,color=color),
            marker_color=color, marker_line_color='white', marker_line_width=1.5,
            text=[f'{avg:.0f}'], textposition='outside',
            textfont=dict(size=13,color=color,family='Syne'),
            name=zone, showlegend=False,
            hovertemplate=f'{zone}<br>Rp/Rs = {avg:.1f} ± {std:.1f}<br>Points: {zpts}<extra></extra>'
        ))

    fig_rank.update_layout(**PLOTLY_LAYOUT,
        title='Average Rp/Rs per Zone — Higher = Better Passive Film',
        xaxis=dict(**PLOTLY_LAYOUT['xaxis'],title='Zone'),
        yaxis=dict(**PLOTLY_LAYOUT['yaxis'],title='Rp / Rs (dimensionless)'),
        height=380)
    st.plotly_chart(fig_rank, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# TAB 7 — DATA TABLE
# ═══════════════════════════════════════════════════════════════
with tab_table:
    st.markdown('<div class="section-header">Complete Data Table — All Points & All Frequencies</div>', unsafe_allow_html=True)

    pt_sel = st.selectbox('Select point', pts_sorted)
    df_show = data[pt_sel]['df'].copy()
    df_show.columns = ['Index','Freq (Hz)',"Z' (Ω)","-Z'' (Ω)",'|Z| (Ω)','Phase (°)','Time (s)']
    st.dataframe(df_show, use_container_width=True, hide_index=True)

    # Download button
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for p in pts_sorted:
            d  = data[p]
            df = d['df'].copy()
            df.columns = ['Index','Freq','Zreal','Zimag','Zmod','Phase','Time']

            # add DRT peaks sheet
            taus, gammas = d['drt'].get_peaks()
            pk_df = pd.DataFrame({
                'Peak': range(1,len(taus)+1),
                'tau (s)': taus,
                'gamma (Ohm)': gammas,
                'Process': [assign_process(t) for t in taus],
            })
            df.to_excel(writer, sheet_name=f'Pt{p}_EIS', index=False)
            pk_df.to_excel(writer, sheet_name=f'Pt{p}_DRT', index=False)

        # Summary sheet
        summ = []
        for p in pts_sorted:
            pm = data[p]['params']
            taus,gammas = data[p]['drt'].get_peaks()
            summ.append({
                'Point': p, 'Zone': data[p]['zone'],
                'Rs (Ohm)': round(pm['Rs'],1),
                'Rp (Ohm)': round(pm['Rp'],1),
                'Rp/Rs': round(pm['Rp']/pm['Rs'],1),
                '|Z|@0.1Hz': round(pm['Zmod_lf'],1),
                'Phase@0.1Hz': round(pm['Phase_lf'],1),
                'DRT peaks': len(taus),
            })
        pd.DataFrame(summ).to_excel(writer, sheet_name='Summary', index=False)

    buf.seek(0)
    st.download_button(
        label='⬇  Download Full Results (.xlsx)',
        data=buf,
        file_name=f'{st.session_state.sample_name.replace(" ","_")}_DRT_results.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )
