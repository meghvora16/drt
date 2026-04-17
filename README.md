# EIS · DRT Analyser — Streamlit App

A professional electrochemical impedance spectroscopy (EIS) and Distribution of Relaxation Times (DRT) analysis tool for stainless steel laser cut samples.

## Features

- **DRT Analysis** — TR-NNLS method (pyimpspec)
- **Nyquist & Bode plots** — interactive Plotly
- **Spatial maps** — Rs, Rp, Rp/Rs across measurement points
- **Zone comparison** — Cut edge vs Perforated hole vs Bulk
- **Laser cut timing diagnostic** — before or after processing
- **Full data export** — Excel with EIS + DRT peaks per point

## Setup & Run

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
streamlit run app.py
```

## How to Use

1. Open the app in your browser (usually http://localhost:8501)
2. In the **sidebar**:
   - Enter your sample name and treatment
   - Set which points are near the laser cut (default: 1, 2, 9, 10)
   - Choose zone geometry (cut edge / perf hole / bulk)
   - Upload your `.xlsx` EIS files (one per measurement point)
   - Click **▶ Run DRT Analysis**
3. Explore results across 7 tabs:
   - **Overview** — key parameters table and diagnostic verdict
   - **DRT** — interactive DRT curves with peak assignments
   - **Nyquist** — complex impedance plane
   - **Bode** — magnitude and phase vs frequency
   - **Spatial** — Rs/Rp/Rp·Rs maps across the sample
   - **Compare** — zone-averaged DRT shapes and Rp/Rs ranking
   - **Data Table** — raw data + full Excel export

## File Format

Each `.xlsx` file must have columns in this order:
```
Index | Frequency (Hz) | Z' (Ω) | -Z'' (Ω) | Z (Ω) | -Phase (°) | Time (s)
```

Name files as `eis1.xlsx`, `eis2.xlsx`, … `eis10.xlsx`  
(the point number is extracted from the filename).

## Interpretation Guide

| Metric | What it tells you |
|---|---|
| Rs CV% > 20% | Thermal oxide intact → no pickling done yet |
| Rs CV% < 10% | Thermal oxide removed by pickling |
| Rp Far/Near > 2× | Laser cut done AFTER processing |
| Rp Far/Near < 1.5× | Laser cut done BEFORE processing |
| Rp/Rs (overall) | Passive film quality — higher = better |
| DRT peaks > 6 | Rich, well-structured Cr₂O₃ film |
| DRT peaks < 4 | Simple film — thick but fewer layers |
