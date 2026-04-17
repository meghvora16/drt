import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pyimpspec as pyi
import io

st.set_page_config(page_title="EIS · DRT Analyser", page_icon="⚡",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=Space+Mono:wght@400;700&display=swap');
:root{--bg:#0a0e1a;--surface:#111827;--surface2:#1c2537;--border:#2a3a54;
      --accent:#00d4ff;--text:#e2e8f0;--muted:#64748b;}
html,body,[data-testid="stAppViewContainer"]{background:var(--bg)!important;
  color:var(--text)!important;font-family:'Space Mono',monospace;}
[data-testid="stSidebar"]{background:var(--surface)!important;border-right:1px solid var(--border);}
h1,h2,h3{font-family:'Syne',sans-serif!important;}
.stButton>button{background:linear-gradient(135deg,#00d4ff,#0099cc)!important;
  color:#0a0e1a!important;font-family:'Syne',sans-serif!important;font-weight:700!important;
  border:none!important;border-radius:6px!important;padding:0.6rem 1.4rem!important;}
.mc{background:var(--surface2);border:1px solid var(--border);border-radius:10px;
    padding:1.2rem 1.4rem;margin-bottom:0.8rem;}
.ml{font-size:0.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:0.1em;}
.mv{font-family:'Syne',sans-serif;font-size:1.6rem;font-weight:800;color:var(--accent);}
.ms{font-size:0.72rem;color:var(--muted);}
.sh{font-family:'Syne',sans-serif;font-size:1.1rem;font-weight:700;color:var(--accent);
    text-transform:uppercase;letter-spacing:0.12em;border-bottom:1px solid var(--border);
    padding-bottom:0.5rem;margin:1.5rem 0 1rem 0;}
.ib{background:rgba(0,212,255,0.07);border-left:3px solid #00d4ff;
    padding:0.8rem 1rem;border-radius:0 8px 8px 0;margin:0.8rem 0;font-size:0.85rem;}
</style>
""", unsafe_allow_html=True)

# ── Plot helpers ───────────────────────────────────────────────
_BG  = '#111827'
_AX  = dict(gridcolor='#1c2537', linecolor='#2a3a54', zerolinecolor='#2a3a54')
_LG  = dict(bgcolor='#1c2537', bordercolor='#2a3a54', borderwidth=1)
_FNT = dict(family='Space Mono', color='#e2e8f0', size=11)
_MRG = dict(l=60, r=40, t=50, b=50)

def _layout(title='', height=480, xkw=None, ykw=None, **kw):
    return dict(paper_bgcolor=_BG, plot_bgcolor=_BG, font=_FNT,
                margin=_MRG, legend=_LG, title=title, height=height,
                xaxis={**_AX, **(xkw or {})},
                yaxis={**_AX, **(ykw or {})}, **kw)

# ── Domain helpers ─────────────────────────────────────────────
def proc(tau):
    if   tau < 1e-5: return 'HF oxide / contact loop'
    elif tau < 1e-4: return 'Double-layer / HF oxide'
    elif tau < 1e-3: return 'Outer passive oxide layer'
    elif tau < 1e-2: return 'Charge transfer — outer film'
    elif tau < 1e-1: return 'Charge transfer — inner Cr2O3'
    elif tau < 1.0:  return 'Ion migration / film repair'
    else:            return 'Slow dissolution / diffusion'

def load_eis(f):
    df = pd.read_excel(f, header=0)
    df.columns = ['Index','Freq','Zreal','Zimag','Zmod','Phase','Time']
    return df

def run_drt(df, method='tr-nnls'):
    pdf = pd.DataFrame({'frequency': df['Freq'].values,
                        "z'": df['Zreal'].values, "-z''": df['Zimag'].values})
    ds  = pyi.dataframe_to_data_sets(pdf, path='eis', label='EIS')[0]
    return ds, pyi.calculate_drt(ds, method=method)

def get_params(df):
    hf = df['Freq'] == 100000
    lf = df['Freq'] == 0.1
    rs  = df.loc[hf,'Zreal'].values[0] if hf.any() else df.iloc[0]['Zreal']
    rp  = (df.loc[lf,'Zreal'].values[0] - rs) if lf.any() else (df.iloc[-1]['Zreal'] - rs)
    zml = df.loc[lf,'Zmod'].values[0]  if lf.any() else df.iloc[-1]['Zmod']
    phl = df.loc[lf,'Phase'].values[0] if lf.any() else df.iloc[-1]['Phase']
    return dict(Rs=rs, Rp=rp, Zmod_lf=zml, Phase_lf=phl)

def fmt(v):
    if v >= 1e6:   return f'{v/1e6:.2f} MΩ'
    elif v >= 1e3: return f'{v/1e3:.1f} kΩ'
    return f'{v:.0f} Ω'

PC = {1:'#ef4444',2:'#f97316',3:'#3b82f6',4:'#2563eb',5:'#22c55e',
      6:'#16a34a',7:'#a855f7',8:'#7c3aed',9:'#f59e0b',10:'#d97706'}
ZC = {'Cut':'#ef4444','Perf':'#a855f7','Near':'#ff6b35','Bulk':'#00d4ff','Far':'#22c55e'}
TR = [(2e-6,'HF'),(1e-4,'Outer'),(5e-3,'Xfer'),(8e-2,'Migration'),(1.3,'Dissolution')]

# ── Session state ──────────────────────────────────────────────
for k,v in [('data',{}),('sname','My Sample'),('near',[1,2,9,10])]:
    if k not in st.session_state: st.session_state[k] = v

# ── Sidebar ────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div style="font-family:Syne;font-size:1.5rem;font-weight:800;'
                'color:#00d4ff;">⚡ EIS · DRT<br>'
                '<span style="font-size:0.85rem;color:#64748b;font-weight:400;">'
                'Analyser</span></div>', unsafe_allow_html=True)
    st.markdown('---')

    st.markdown('<div class="sh">Sample</div>', unsafe_allow_html=True)
    st.session_state.sname = st.text_input('Sample name', value=st.session_state.sname)
    material  = st.selectbox('Material', ['SS 316L','SS 304','Carbon Steel','Aluminium','Custom'])
    treatment = st.selectbox('Treatment',
        ['Laser Cut Only','Brushed + Pickled','B+P + Passivated','As-received','Custom'])

    st.markdown('<div class="sh">Geometry</div>', unsafe_allow_html=True)
    near_in   = st.multiselect('Near laser cut ★', options=list(range(1,11)), default=[1,2,9,10])
    st.session_state.near = near_in
    split_geo = st.checkbox('1,2=Cut edge | 9,10=Perf hole | rest=Bulk', value=True)

    st.markdown('<div class="sh">DRT Settings</div>', unsafe_allow_html=True)
    drt_meth = st.selectbox('Method', ['tr-nnls','tr-rbf','bht'])
    norm_drt = st.checkbox('Normalise DRT', value=False)

    st.markdown('<div class="sh">Upload</div>', unsafe_allow_html=True)
    st.markdown('<div class="ib">One .xlsx per point — name as eis1.xlsx, eis2.xlsx …</div>',
                unsafe_allow_html=True)
    uploaded = st.file_uploader('EIS files', type='xlsx', accept_multiple_files=True)

    if uploaded and st.button('▶  Run DRT Analysis', use_container_width=True):
        prog = st.progress(0); msg = st.empty()
        st.session_state.data = {}
        for i, f in enumerate(uploaded):
            num_str = ''.join(c for c in f.name.replace('.xlsx','') if c.isdigit())
            pnum    = int(num_str) if num_str else (i+1)
            msg.markdown(f'<div class="ib">Processing point {pnum}…</div>', unsafe_allow_html=True)
            try:
                df      = load_eis(f)
                ds, drt = run_drt(df, drt_meth)
                params  = get_params(df)
                if split_geo:
                    if   pnum in [1,2]:       zone = 'Cut'
                    elif pnum in [9,10]:       zone = 'Perf'
                    elif pnum in near_in:      zone = 'Near'
                    else:                      zone = 'Bulk'
                else:
                    zone = 'Near' if pnum in near_in else 'Bulk'
                st.session_state.data[pnum] = dict(df=df,ds=ds,drt=drt,params=params,zone=zone)
            except Exception as e:
                st.error(f'{f.name}: {e}')
            prog.progress((i+1)/len(uploaded))
        msg.empty(); prog.empty()
        st.success(f'✓ {len(st.session_state.data)} points loaded')

# ── Main ───────────────────────────────────────────────────────
data     = st.session_state.data
near_pts = st.session_state.near

st.markdown(
    f'<div style="display:flex;align-items:baseline;gap:1rem;margin-bottom:0.5rem;">'
    f'<span style="font-family:Syne;font-size:2rem;font-weight:800;color:#e2e8f0;">'
    f'{st.session_state.sname}</span>'
    f'<span style="font-family:Space Mono;font-size:0.8rem;color:#64748b;">'
    f'{treatment} · {material}</span></div>', unsafe_allow_html=True)

if not data:
    for col,(icon,title,desc) in zip(st.columns(3),[
        ('📁','Upload','Load .xlsx EIS files from sidebar'),
        ('⚙️','Configure','Set geometry — cut edge, perf, bulk'),
        ('📊','Analyse','DRT, Nyquist, Bode, spatial maps'),
    ]):
        with col:
            st.markdown(
                f'<div class="mc" style="text-align:center;padding:2rem;">'
                f'<div style="font-size:2.5rem;margin-bottom:0.8rem;">{icon}</div>'
                f'<div style="font-family:Syne;font-size:1.1rem;font-weight:700;'
                f'color:#00d4ff;">{title}</div>'
                f'<div style="font-size:0.82rem;color:#64748b;margin-top:0.4rem;">'
                f'{desc}</div></div>', unsafe_allow_html=True)
    st.stop()

pts   = sorted(data.keys())
farps = [p for p in pts if p not in near_pts]

(t_ov,t_drt,t_ny,t_bo,t_sp,t_cmp,t_tb) = st.tabs([
    '📊 Overview','🌊 DRT','🔵 Nyquist','📈 Bode','🗺 Spatial','⚖️ Compare','📋 Data Table'])

# ── OVERVIEW ──────────────────────────────────────────────────
with t_ov:
    st.markdown('<div class="sh">Key Parameters</div>', unsafe_allow_html=True)
    all_rs = [data[p]['params']['Rs'] for p in pts]
    all_rp = [data[p]['params']['Rp'] for p in pts]
    rs_cv  = np.std(all_rs)/np.mean(all_rs)*100
    rp_cv  = np.std(all_rp)/np.mean(all_rp)*100
    near_v = [data[p]['params']['Rp'] for p in pts if p in near_pts]
    far_v  = [data[p]['params']['Rp'] for p in pts if p not in near_pts]
    nr     = np.mean(near_v) if near_v else 1
    fr     = np.mean(far_v)  if far_v  else 1
    ratio  = fr/nr

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(f'<div class="mc"><div class="ml">Avg Rs</div>'
                         f'<div class="mv">{fmt(np.mean(all_rs))}</div>'
                         f'<div class="ms">CV={rs_cv:.1f}%</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="mc"><div class="ml">Avg Rp Near ★</div>'
                         f'<div class="mv">{fmt(nr)}</div>'
                         f'<div class="ms">pts {near_pts}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="mc"><div class="ml">Avg Rp Far</div>'
                         f'<div class="mv">{fmt(fr)}</div>'
                         f'<div class="ms">bulk points</div></div>', unsafe_allow_html=True)
    with c4:
        col = '#22c55e' if ratio<1.5 else ('#f59e0b' if ratio<3 else '#ef4444')
        vrd = 'Cut BEFORE processing' if ratio<1.5 else 'Cut AFTER processing'
        st.markdown(f'<div class="mc"><div class="ml">Far/Near Rp</div>'
                    f'<div class="mv" style="color:{col};">{ratio:.2f}×</div>'
                    f'<div class="ms">{vrd}</div></div>', unsafe_allow_html=True)

    rows = []
    for p in pts:
        pm = data[p]['params']
        taus, gams = data[p]['drt'].get_peaks()
        dom = float(taus[np.argmax(gams)]) if len(taus)>0 else 0
        rows.append({'Pt':p,'Zone':data[p]['zone'],'Rs(kΩ)':round(pm['Rs']/1000,1),
                     'Rp(kΩ)':round(pm['Rp']/1000,1),'Rp/Rs':round(pm['Rp']/pm['Rs'],1),
                     '|Z|@0.1Hz(kΩ)':round(pm['Zmod_lf']/1000,1),
                     'Phase@0.1Hz':round(pm['Phase_lf'],1),'DRT peaks':len(taus),
                     'Dom τ(s)':round(dom,4) if dom else '-'})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.markdown(
        f'<div class="ib">'
        f'Rs CV={rs_cv:.1f}% — {"⚠ Steep gradient → thermal oxide intact" if rs_cv>20 else "✓ Flat → thermal oxide removed"}<br>'
        f'Rp CV={rp_cv:.1f}% — {"⚠ Scattered → patchy passive film" if rp_cv>30 else "✓ Uniform → homogeneous surface"}'
        f'</div>', unsafe_allow_html=True)

# ── DRT ───────────────────────────────────────────────────────
with t_drt:
    st.markdown('<div class="sh">Distribution of Relaxation Times</div>', unsafe_allow_html=True)
    cc, cm = st.columns([1,3])
    with cc:
        sel_d  = st.multiselect('Points', pts, default=pts, key='dp')
        mk_pk  = st.checkbox('Mark peaks', value=True)
        grp_av = st.checkbox('Group averages', value=True)
        log_g  = st.checkbox('Log γ axis', value=False)

    ref_t = data[pts[0]]['drt'].get_time_constants()
    fd    = go.Figure()

    for p in sel_d:
        d  = data[p]
        t  = d['drt'].get_time_constants()
        g  = d['drt'].get_gammas()
        n  = (g.max() if g.max()>0 else 1) if norm_drt else 1
        c  = PC.get(p,'#888')
        ls = 'solid' if p in near_pts else 'dash'
        fd.add_trace(go.Scatter(x=t, y=g/n, mode='lines',
            name=f'Pt {p}{"★" if p in near_pts else ""}',
            line=dict(color=c, width=2, dash=ls),
            hovertemplate=f'Pt {p}<br>τ=%{{x:.3e}}<br>γ=%{{y:.2e}}<extra></extra>'))
        if mk_pk:
            pt2,pg2 = d['drt'].get_peaks()
            if len(pt2)>0:
                fd.add_trace(go.Scatter(x=pt2, y=pg2/n, mode='markers', showlegend=False,
                    marker=dict(color=c, size=9, symbol='circle',
                                line=dict(color='white', width=1.5)),
                    hovertemplate=f'Pt {p} peak<br>τ=%{{x:.3e}}<br>γ=%{{y:.2e}}<extra></extra>'))

    if grp_av:
        ns = [p for p in near_pts if p in sel_d and p in data]
        fs = [p for p in farps    if p in sel_d and p in data]
        if ns:
            ng = np.array([data[p]['drt'].get_gammas() for p in ns]).mean(0)
            nn = ng.max() if (norm_drt and ng.max()>0) else 1
            fd.add_trace(go.Scatter(x=ref_t, y=ng/nn, mode='lines', name='Near avg ★',
                line=dict(color='#ff6b35', width=3.5)))
        if fs:
            fg = np.array([data[p]['drt'].get_gammas() for p in fs]).mean(0)
            fn = fg.max() if (norm_drt and fg.max()>0) else 1
            fd.add_trace(go.Scatter(x=ref_t, y=fg/fn, mode='lines', name='Far avg',
                line=dict(color='#00d4ff', width=3.5, dash='dot')))

    for tau,lbl in TR:
        fd.add_vline(x=tau, line_color='#2a3a54', line_dash='dot', line_width=1)
        fd.add_annotation(x=np.log10(tau), y=1.02, xref='x', yref='paper',
                          text=lbl, showarrow=False, font=dict(size=8,color='#64748b'), xanchor='left')

    yl = 'γ/γ_max' if norm_drt else 'γ (Ω)'
    fd.update_layout(**_layout('Distribution of Relaxation Times — TR-NNLS', height=500,
                               xkw=dict(type='log', title='Time Constant τ (s)'),
                               ykw=dict(type='log' if log_g else 'linear', title=yl)))
    st.plotly_chart(fd, use_container_width=True)

    st.markdown('<div class="sh">Peak Assignments</div>', unsafe_allow_html=True)
    pk_rows = []
    for p in sel_d:
        taus,gams = data[p]['drt'].get_peaks()
        for i,idx in enumerate(np.argsort(taus)):
            t2=float(taus[idx]); g2=float(gams[idx])
            pk_rows.append({'Pt':p,'Zone':data[p]['zone'],'#':i+1,
                            'τ(s)':f'{t2:.4e}','τ(ms)':round(t2*1000,4),
                            'γ(Ω)':round(g2,0),'Process':proc(t2)})
    if pk_rows:
        st.dataframe(pd.DataFrame(pk_rows), use_container_width=True, hide_index=True)

# ── NYQUIST ───────────────────────────────────────────────────
with t_ny:
    st.markdown('<div class="sh">Nyquist Plot</div>', unsafe_allow_html=True)
    sel_n = st.multiselect('Points', pts, default=pts, key='np')
    unit  = st.radio('Unit', ['Ω','kΩ','MΩ'], horizontal=True, key='nu')
    dv    = {'Ω':1,'kΩ':1e3,'MΩ':1e6}[unit]
    fn    = go.Figure()
    for p in sel_n:
        df=data[p]['df']; ls='solid' if p in near_pts else 'dash'
        fn.add_trace(go.Scatter(x=df['Zreal']/dv, y=df['Zimag']/dv,
            mode='lines+markers', name=f'Pt {p}{"★" if p in near_pts else ""}',
            line=dict(color=PC.get(p,'#888'), width=2, dash=ls), marker=dict(size=4),
            hovertemplate=f"Pt {p}<br>Z'=%{{x:.3f}} {unit}<br>-Z''=%{{y:.3f}} {unit}<extra></extra>"))
    fn.update_layout(**_layout('Nyquist Plot', height=520,
                               xkw=dict(title=f"Z' ({unit})"),
                               ykw=dict(title=f"-Z'' ({unit})")))
    st.plotly_chart(fn, use_container_width=True)

# ── BODE ─────────────────────────────────────────────────────
with t_bo:
    st.markdown('<div class="sh">Bode Plots</div>', unsafe_allow_html=True)
    sel_b = st.multiselect('Points', pts, default=pts, key='bp')
    fb    = make_subplots(rows=1,cols=2,subplot_titles=['|Z| vs Frequency','Phase vs Frequency'])
    for p in sel_b:
        df=data[p]['df']; c=PC.get(p,'#888'); ls='solid' if p in near_pts else 'dash'
        fb.add_trace(go.Scatter(x=df['Freq'],y=df['Zmod']/1000,mode='lines',
            name=f'Pt {p}{"★" if p in near_pts else ""}',
            line=dict(color=c,width=2,dash=ls),
            hovertemplate=f'Pt {p}<br>f=%{{x:.1f}} Hz<br>|Z|=%{{y:.2f}} kΩ<extra></extra>'),
            row=1,col=1)
        fb.add_trace(go.Scatter(x=df['Freq'],y=df['Phase'],mode='lines',
            name=f'Pt {p}',showlegend=False,line=dict(color=c,width=2,dash=ls),
            hovertemplate=f'Pt {p}<br>f=%{{x:.1f}} Hz<br>Phase=%{{y:.2f}}°<extra></extra>'),
            row=1,col=2)
    fb.update_xaxes(type='log',title='Frequency (Hz)',gridcolor='#1c2537',linecolor='#2a3a54')
    fb.update_yaxes(gridcolor='#1c2537',linecolor='#2a3a54')
    fb.update_yaxes(type='log',title='|Z| (kΩ)',row=1,col=1)
    fb.update_yaxes(title='−Phase (°)',row=1,col=2)
    fb.update_layout(paper_bgcolor=_BG,plot_bgcolor=_BG,font=_FNT,height=480,
                     showlegend=True,legend=_LG,margin=_MRG)
    st.plotly_chart(fb, use_container_width=True)

# ── SPATIAL ───────────────────────────────────────────────────
with t_sp:
    st.markdown('<div class="sh">Spatial Maps — Rs, Rp, Rp/Rs</div>', unsafe_allow_html=True)
    rs_v  = [data[p]['params']['Rs']/1000 for p in pts]
    rp_v  = [data[p]['params']['Rp']/1000 for p in pts]
    rr_v  = [data[p]['params']['Rp']/data[p]['params']['Rs'] for p in pts]
    bcols = ['#ef4444' if p in [1,2] else ('#a855f7' if p in [9,10] else '#00d4ff') for p in pts]
    xlbl  = [f'Pt {p}{"★" if p in near_pts else ""}' for p in pts]

    fs = make_subplots(rows=3,cols=1,shared_xaxes=True,vertical_spacing=0.08,
                       subplot_titles=['Rs (kΩ) — Thermal oxide gradient indicator',
                                       'Rp (kΩ) — Polarisation resistance',
                                       'Rp/Rs — Film quality (electrolyte-normalised)'])
    for row,vals,lbl in [(1,rs_v,'Rs'),(2,rp_v,'Rp'),(3,rr_v,'Rp/Rs')]:
        nn = np.mean([vals[pts.index(p)] for p in near_pts if p in pts]) if any(p in near_pts for p in pts) else None
        ff = np.mean([vals[pts.index(p)] for p in pts if p not in near_pts]) if farps else None
        fs.add_trace(go.Bar(x=xlbl,y=vals,marker_color=bcols,
            marker_line_color='#2a3a54',marker_line_width=1,
            text=[f'{v:.1f}' for v in vals],textposition='outside',
            textfont=dict(size=9,color='#e2e8f0'),showlegend=False,
            hovertemplate='%{x}<br>'+lbl+'=%{y:.2f}<extra></extra>'),row=row,col=1)
        if nn: fs.add_hline(y=nn,line_color='#ff6b35',line_dash='dash',line_width=1.5,
            annotation_text=f'Near avg {nn:.1f}',annotation_font=dict(color='#ff6b35',size=9),row=row,col=1)
        if ff: fs.add_hline(y=ff,line_color='#00d4ff',line_dash='dash',line_width=1.5,
            annotation_text=f'Far avg {ff:.1f}',annotation_font=dict(color='#00d4ff',size=9),row=row,col=1)
    fs.update_xaxes(gridcolor='#1c2537',linecolor='#2a3a54')
    fs.update_yaxes(gridcolor='#1c2537',linecolor='#2a3a54')
    fs.update_layout(paper_bgcolor=_BG,plot_bgcolor=_BG,font=_FNT,height=780,
                     showlegend=False,margin=dict(l=70,r=70,t=60,b=40))
    st.plotly_chart(fs, use_container_width=True)

    cv_r = np.std(rs_v)/np.mean(rs_v)*100
    msg  = (f'⚠ Rs CV={cv_r:.1f}% — Steep gradient → thermal oxide intact → no pickling'
            if cv_r>20 else f'✓ Rs CV={cv_r:.1f}% — Flat → thermal oxide removed by pickling')
    st.markdown(f'<div class="ib">{msg}</div>', unsafe_allow_html=True)

# ── COMPARE ───────────────────────────────────────────────────
with t_cmp:
    st.markdown('<div class="sh">Zone DRT Comparison</div>', unsafe_allow_html=True)
    zones = sorted(set(data[p]['zone'] for p in pts))
    ref_t = data[pts[0]]['drt'].get_time_constants()
    fc    = go.Figure()

    for z in zones:
        zp    = [p for p in pts if data[p]['zone']==z]
        ga    = np.array([data[p]['drt'].get_gammas() for p in zp])
        avg_g = ga.mean(0); std_g = ga.std(0)
        nm    = avg_g.max() if avg_g.max()>0 else 1
        col   = ZC.get(z,'#888')
        fc.add_trace(go.Scatter(x=ref_t,y=avg_g/nm,mode='lines',
            name=f'{z} (n={len(zp)})',line=dict(color=col,width=3),
            hovertemplate=f'{z}<br>τ=%{{x:.3e}}<br>γ/max=%{{y:.3f}}<extra></extra>'))
        xf = np.concatenate([ref_t,ref_t[::-1]])
        yf = np.concatenate([(avg_g+std_g)/nm,((avg_g-std_g)/nm)[::-1]])
        fc.add_trace(go.Scatter(x=xf,y=yf,fill='toself',
            fillcolor=col+'22',line=dict(color='rgba(0,0,0,0)'),showlegend=False,hoverinfo='skip'))

    for tau,_ in TR:
        fc.add_vline(x=tau,line_color='#2a3a54',line_dash='dot',line_width=1)

    fc.update_layout(**_layout('DRT by Zone — Average ± 1σ (normalised)', height=480,
                               xkw=dict(type='log',title='τ (s)'),
                               ykw=dict(title='γ / γ_max')))
    st.plotly_chart(fc, use_container_width=True)

    st.markdown('<div class="sh">Rp/Rs Ranking</div>', unsafe_allow_html=True)
    fr = go.Figure()
    for z in zones:
        zp  = [p for p in pts if data[p]['zone']==z]
        v   = [data[p]['params']['Rp']/data[p]['params']['Rs'] for p in zp]
        avg = np.mean(v); std = np.std(v); col = ZC.get(z,'#888')
        fr.add_trace(go.Bar(x=[z],y=[avg],
            error_y=dict(type='data',array=[std],visible=True,color=col),
            marker_color=col,marker_line_color='white',marker_line_width=1.5,
            text=[f'{avg:.0f}'],textposition='outside',
            textfont=dict(size=13,color=col,family='Syne'),
            name=z,showlegend=False,
            hovertemplate=f'{z}<br>Rp/Rs={avg:.1f}±{std:.1f}<br>pts:{zp}<extra></extra>'))
    fr.update_layout(**_layout('Average Rp/Rs per Zone — Higher = Better Film', height=380,
                               xkw=dict(title='Zone'),
                               ykw=dict(title='Rp / Rs')))
    st.plotly_chart(fr, use_container_width=True)

# ── DATA TABLE ────────────────────────────────────────────────
with t_tb:
    st.markdown('<div class="sh">Raw Data & Export</div>', unsafe_allow_html=True)
    psel   = st.selectbox('Select point', pts)
    df_sh  = data[psel]['df'].copy()
    df_sh.columns = ['Index','Freq (Hz)',"Z' (Ω)","-Z'' (Ω)",'|Z| (Ω)','Phase (°)','Time (s)']
    st.dataframe(df_sh, use_container_width=True, hide_index=True)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as wr:
        summ = []
        for p in pts:
            pm=data[p]['params']; taus,gams=data[p]['drt'].get_peaks()
            summ.append({'Pt':p,'Zone':data[p]['zone'],
                         'Rs(Ohm)':round(pm['Rs'],1),'Rp(Ohm)':round(pm['Rp'],1),
                         'Rp/Rs':round(pm['Rp']/pm['Rs'],1),
                         '|Z|@0.1Hz':round(pm['Zmod_lf'],1),
                         'Phase@0.1Hz':round(pm['Phase_lf'],1),'DRT peaks':len(taus)})
        pd.DataFrame(summ).to_excel(wr, sheet_name='Summary', index=False)
        for p in pts:
            d2=data[p]; df2=d2['df'].copy()
            df2.columns=['Index','Freq','Zreal','Zimag','Zmod','Phase','Time']
            df2.to_excel(wr, sheet_name=f'Pt{p}_EIS', index=False)
            taus,gams=d2['drt'].get_peaks()
            pd.DataFrame({'Peak':range(1,len(taus)+1),'tau(s)':taus,
                          'gamma(Ohm)':gams,'Process':[proc(t) for t in taus]}
                         ).to_excel(wr, sheet_name=f'Pt{p}_DRT', index=False)
    buf.seek(0)
    st.download_button('⬇  Download Full Results (.xlsx)', data=buf,
        file_name=f'{st.session_state.sname.replace(" ","_")}_DRT.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True)
