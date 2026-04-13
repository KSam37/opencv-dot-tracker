import pandas as pd
import xlsxwriter
import os

val_dir = os.path.dirname(os.path.abspath(__file__))
tracker_csv = os.path.join(val_dir, "Sample Side 1 Data.csv")
opencv_csv  = os.path.join(val_dir, "Instron - side - 1.csv")
out_path    = os.path.join(val_dir, "Tracker vs openCV Comparison.xlsx")

# ── Load data ────────────────────────────────────────────────────────────────
tracker_raw = pd.read_csv(tracker_csv, skiprows=2, names=["t", "L", "extra"])
tracker_raw = tracker_raw[tracker_raw["L"].notna()].copy()
tracker_raw["t"] = tracker_raw["t"].astype(float)
tracker_raw["L"] = tracker_raw["L"].astype(float) * 1000   # m → mm
tracker = tracker_raw[["t", "L"]].reset_index(drop=True)

ocv_raw = pd.read_csv(opencv_csv)
ocv_raw.columns = ["t", "d"]
cal = tracker["L"].iloc[0] / ocv_raw["d"].iloc[0]          # px → mm
ocv_raw["d_mm"] = ocv_raw["d"] * cal

merged = []
for _, row in tracker.iterrows():
    idx = (ocv_raw["t"] - row["t"]).abs().idxmin()
    merged.append((row["t"], row["L"], ocv_raw.loc[idx, "d_mm"]))

df = pd.DataFrame(merged, columns=["time", "tracker_mm", "opencv_mm"])
df["pct_diff"] = (df["tracker_mm"] - df["opencv_mm"]).abs() / df["tracker_mm"] * 100
n = len(df)
print(f"Tracker: {len(tracker)} pts  |  openCV: {len(ocv_raw)} pts  |  Merged: {n} pts")

# ── Workbook ─────────────────────────────────────────────────────────────────
wb  = xlsxwriter.Workbook(out_path)
ws  = wb.add_worksheet("Comparison")

# Formats
fmt_title   = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F4E79", "font_name": "Arial"})
fmt_hdr     = wb.add_format({"bold": True, "font_size": 11, "font_color": "white",
                              "bg_color": "#1F4E79", "align": "center", "border": 1, "font_name": "Arial"})
fmt_num     = wb.add_format({"num_format": "0.0000", "align": "center", "border": 1, "font_name": "Arial", "font_size": 10})
fmt_time    = wb.add_format({"num_format": "0.000",  "align": "center", "border": 1, "font_name": "Arial", "font_size": 10})
fmt_pct     = wb.add_format({"num_format": '0.000"%"', "align": "center", "border": 1, "font_name": "Arial", "font_size": 10})
fmt_stat_hdr= wb.add_format({"bold": True, "font_size": 13, "font_color": "white",
                              "bg_color": "#1F4E79", "align": "center", "font_name": "Arial"})
fmt_stat_lbl= wb.add_format({"bold": True, "font_size": 10, "bg_color": "#D6E4F0",
                              "align": "left", "border": 1, "font_name": "Arial"})
fmt_stat_val= wb.add_format({"font_size": 10, "align": "center", "border": 1, "font_name": "Arial"})
fmt_stat_key= wb.add_format({"bold": True, "font_size": 11, "font_color": "#1F4E79",
                              "align": "center", "border": 1, "font_name": "Arial"})

# Column widths
ws.set_column("A:A", 14)
ws.set_column("B:B", 16)
ws.set_column("C:C", 16)
ws.set_column("D:D", 18)
ws.set_column("E:E", 4)
ws.set_column("F:F", 32)
ws.set_column("G:G", 18)
ws.set_row(0, 28)
ws.set_row(1, 6)

# ── Title ────────────────────────────────────────────────────────────────────
ws.merge_range("A1:D1", "Tracker vs openCV Performance Comparison — Instron Side 1", fmt_title)

# ── Column headers ───────────────────────────────────────────────────────────
HDR_ROW = 2   # 0-indexed
ws.write(HDR_ROW, 0, "Time (s)",       fmt_hdr)
ws.write(HDR_ROW, 1, "Tracker (mm)",   fmt_hdr)
ws.write(HDR_ROW, 2, "openCV (mm)",    fmt_hdr)
ws.write(HDR_ROW, 3, "% Difference",   fmt_hdr)

# ── Data rows ────────────────────────────────────────────────────────────────
DATA_ROW0 = 3   # 0-indexed, so Excel row 4
for i, row in df.iterrows():
    r = DATA_ROW0 + i
    ws.write(r, 0, round(row["time"],       4), fmt_time)
    ws.write(r, 1, round(row["tracker_mm"], 4), fmt_num)
    ws.write(r, 2, round(row["opencv_mm"],  4), fmt_num)
    # Formula in column D (0-indexed=3), Excel row = r+1
    ws.write_formula(r, 3, f"=ABS(C{r+1}-B{r+1})/B{r+1}*100", fmt_pct)

LAST_ROW = DATA_ROW0 + n - 1  # 0-indexed last data row
er = LAST_ROW + 1              # Excel row number of last data row (1-indexed)
sr = DATA_ROW0 + 1             # Excel row number of first data row (1-indexed)

# ── Stats panel ──────────────────────────────────────────────────────────────
ws.merge_range("F1:G1", "Statistical Analysis", fmt_stat_hdr)
ws.merge_range("F2:G2", "openCV vs Manual Tracker",
               wb.add_format({"italic": True, "font_size": 10, "font_color": "#2E75B6",
                               "align": "center", "font_name": "Arial"}))

stats = [
    ("Number of Matched Points",      f"=COUNTA(A{sr}:A{er})",                                         "0",         False),
    ("Time Range (s)",                 f"=MAX(A{sr}:A{er})-MIN(A{sr}:A{er})",                           "0.00",      False),
    (None, None, None, False),
    ("Mean % Difference",              f"=AVERAGE(D{sr}:D{er})",                                        '0.000"%"',  True),
    ("Median % Difference",            f"=MEDIAN(D{sr}:D{er})",                                         '0.000"%"',  False),
    ("Max % Difference",               f"=MAX(D{sr}:D{er})",                                            '0.000"%"',  False),
    ("Min % Difference",               f"=MIN(D{sr}:D{er})",                                            '0.000"%"',  False),
    ("Std Dev % Difference",           f"=STDEV(D{sr}:D{er})",                                          '0.000"%"',  False),
    (None, None, None, False),
    ("RMSE (mm)",                      f"=SQRT(SUMPRODUCT((C{sr}:C{er}-B{sr}:B{er})^2)/COUNTA(B{sr}:B{er}))", "0.0000", True),
    (None, None, None, False),
    ("Correlation Coefficient (R)",    f"=CORREL(B{sr}:B{er},C{sr}:C{er})",                             "0.000000",  True),
    ("R-squared",                      f"=CORREL(B{sr}:B{er},C{sr}:C{er})^2",                           "0.000000",  True),
    (None, None, None, False),
    ("Tracker Mean Distance (mm)",     f"=AVERAGE(B{sr}:B{er})",                                        "0.000",     False),
    ("openCV Mean Distance (mm)",      f"=AVERAGE(C{sr}:C{er})",                                        "0.000",     False),
    ("Tracker Distance Range (mm)",    f"=MAX(B{sr}:B{er})-MIN(B{sr}:B{er})",                           "0.000",     False),
    ("openCV Distance Range (mm)",     f"=MAX(C{sr}:C{er})-MIN(C{sr}:C{er})",                           "0.000",     False),
]

for i, (label, formula, fmt_str, highlight) in enumerate(stats):
    r = 2 + i   # 0-indexed, starts at row 3 (Excel row 3)
    if label is None:
        continue
    val_fmt = fmt_stat_key if highlight else fmt_stat_val
    if fmt_str:
        val_fmt = wb.add_format({"num_format": fmt_str, "align": "center", "border": 1,
                                  "font_name": "Arial", "font_size": 11 if highlight else 10,
                                  "bold": highlight, "font_color": "#1F4E79" if highlight else "#000000"})
    ws.write(r, 5, label,   fmt_stat_lbl)
    ws.write_formula(r, 6, formula, val_fmt)

# ── Chart 1: Distance vs Time ─────────────────────────────────────────────────
chart1 = wb.add_chart({"type": "scatter", "subtype": "straight"})
chart1.set_title({"name": "Distance vs Time — Tracker vs openCV"})
chart1.set_x_axis({"name": "Time (s)", "min": 0, "max": df["time"].max()})
chart1.set_y_axis({"name": "Distance (mm)"})
chart1.set_size({"width": 560, "height": 360})
chart1.set_legend({"position": "bottom"})

chart1.add_series({
    "name":       "Tracker (manual)",
    "categories": ["Comparison", DATA_ROW0, 0, LAST_ROW, 0],
    "values":     ["Comparison", DATA_ROW0, 1, LAST_ROW, 1],
    "line":       {"color": "#1F4E79", "width": 1.5},
})
chart1.add_series({
    "name":       "openCV (automated)",
    "categories": ["Comparison", DATA_ROW0, 0, LAST_ROW, 0],
    "values":     ["Comparison", DATA_ROW0, 2, LAST_ROW, 2],
    "line":       {"color": "#FF6600", "width": 1.5, "dash_type": "dash"},
})
ws.insert_chart("F22", chart1)

# ── Chart 2: % Difference ────────────────────────────────────────────────────
chart2 = wb.add_chart({"type": "scatter", "subtype": "straight"})
chart2.set_title({"name": "% Difference Over Time"})
chart2.set_x_axis({"name": "Time (s)", "min": 0, "max": df["time"].max()})
chart2.set_y_axis({"name": "% Difference"})
chart2.set_size({"width": 560, "height": 300})
chart2.set_legend({"none": True})

chart2.add_series({
    "name":       "% Difference",
    "categories": ["Comparison", DATA_ROW0, 0, LAST_ROW, 0],
    "values":     ["Comparison", DATA_ROW0, 3, LAST_ROW, 3],
    "line":       {"color": "#C00000", "width": 1.5},
})
ws.insert_chart("F39", chart2)

wb.close()
print(f"Saved: {out_path}")
