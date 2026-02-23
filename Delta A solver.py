"""
CIDS correction, validation, and chirality decomposition
-------------------------------------------------------
UPDATED to support headers in this exact format:

  "CIDS-<sample ID>-<concentration>"
  "CD-<sample ID>-<concentration>"
  "ST-<sample ID>-<concentration>"

Examples:
  CIDS-DVC-0.25      CD-DVC-0.25      ST-DVC-0.25
  CIDS-LVC-0.25      CD-LVC-0.25      ST-LVC-0.25
  CIDS-D-VC-0.25     CD-D-VC-0.25     ST-D-VC-0.25   (sample ID may contain dashes)

Key behavior:
- The code groups by BOTH sample ID and concentration, so D-VC-0.25 and L-VC-0.25 are treated as two separate datasets.
- The concentration token is used ONLY for labeling/output (NOT in the equation), consistent with your manual calc (c = 1).

Equation (your corrected form):
  denom = 10^ST - 1
  CIDS = l_T*ΔA - (CD - ΔA)/denom
  ΔA   = (CIDS + CD/denom) / (l_T + 1/denom)

Outputs:
- Sheet 1: "CIDS_Decomposition" (full)
- Sheet 2: "DeltaA_only" (wavelength + ΔA columns only)

Dependencies:
  pip install numpy pandas openpyxl matplotlib
"""

import re
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure


APP_TITLE = "CIDS correction, validation, and chirality decomposition"
MIN_ST = 0.01


def _sanitize_columns(cols):
    return [str(c).strip() for c in cols]


def _to_float_conc(token: str):
    """
    Extract the first numeric value from a concentration token string.
    Accepts: "0.25", "0.25 pM", "0.25pM", "2", "2.0", "1e-3", etc.
    Returns (float_value, pretty_label_string) or (None, None).
    """
    if token is None:
        return None, None
    s = str(token).strip().replace("_", " ").replace(",", ".")
    m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
    if not m:
        return None, None
    try:
        v = float(m.group(0))
        label = re.sub(r"\s+", " ", str(token).strip())
        return v, label
    except Exception:
        return None, None


def _detect_wavelength_column(columns):
    cols = _sanitize_columns(columns)
    low = [c.lower() for c in cols]
    for i, c in enumerate(low):
        if ("wavelength" in c) or (re.search(r"\bwave\b", c) is not None) or ("nm" in c):
            return cols[i]
    return cols[0]


def _slug(s: str) -> str:
    """Safe tag for column names."""
    s = str(s).strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^0-9A-Za-z._\-]+", "_", s)  # keep dash/dot/underscore
    return s


def _parse_kind_sample_conc(colname: str):
    """
    Parse header: KIND-<sampleID>-<conc>

    - KIND is CIDS/CD/ST (case-insensitive)
    - sampleID may contain dashes, so we split from the RIGHT:
        first chunk = KIND
        last chunk  = concentration token
        middle      = sampleID (can include dashes)
    Returns:
      (kind_upper, sample_id_str, conc_float, conc_label_str) OR None if no match
    """
    m = re.match(r"^\s*(CIDS|CD|ST)\s*[-_]\s*(.+)\s*$", colname, flags=re.IGNORECASE)
    if not m:
        return None
    kind = m.group(1).upper()
    rest = m.group(2).strip()

    # Split from rightmost separator to get concentration token
    # Allow either '-' or '_' between sampleID and concentration, but your preferred is '-'
    # We'll split on the last '-' if present; otherwise last '_' as fallback.
    if "-" in rest:
        sample_part, conc_part = rest.rsplit("-", 1)
    elif "_" in rest:
        sample_part, conc_part = rest.rsplit("_", 1)
    else:
        return None

    sample_id = sample_part.strip()
    conc_val, conc_label = _to_float_conc(conc_part.strip())
    if not sample_id or conc_val is None:
        return None

    return kind, sample_id, conc_val, conc_label


def _group_columns_by_sample_and_conc(columns):
    """
    Build groups keyed by (sampleID, conc_float).

    Returns:
      wl_col: str
      groups: list of dicts, each dict:
        {
          "sample_id": str,
          "conc_val": float,
          "conc_label": str,
          "key_label": str,          # for legends and output tags
          "CIDS": colname,
          "CD": colname,
          "ST": colname
        }
    Keeps only complete triplets (CIDS+CD+ST).
    """
    cols = _sanitize_columns(columns)
    wl_col = _detect_wavelength_column(cols)

    found = {}  # (kind, sample_id, conc_val) -> (colname, conc_label)
    for col in cols:
        if col == wl_col:
            continue
        parsed = _parse_kind_sample_conc(col)
        if parsed is None:
            continue
        kind, sample_id, conc_val, conc_label = parsed
        found[(kind, sample_id, conc_val)] = (col, conc_label)

    # Find complete triplets
    keys = sorted(set((sid, c) for (k, sid, c) in found.keys()))
    groups = []
    for sid, cval in keys:
        if (("CIDS", sid, cval) in found) and (("CD", sid, cval) in found) and (("ST", sid, cval) in found):
            cids_col, lab1 = found[("CIDS", sid, cval)]
            cd_col, lab2 = found[("CD", sid, cval)]
            st_col, lab3 = found[("ST", sid, cval)]
            conc_label = lab1 or lab2 or lab3 or f"{cval:g}"
            # Human label used in legends
            key_label = f"{sid} @ {conc_label}"
            groups.append({
                "sample_id": sid,
                "conc_val": cval,
                "conc_label": conc_label,
                "key_label": key_label,
                "CIDS": cids_col,
                "CD": cd_col,
                "ST": st_col
            })

    # Sort groups: by sample_id then by concentration
    groups.sort(key=lambda g: (g["sample_id"].lower(), g["conc_val"]))
    return wl_col, groups


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1200x820")

        self.df = None
        self.wl_col = None
        self.groups = None  # list[dict]
        self.outputs = None

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(top, text="Load Excel", command=self.load_excel).pack(side=tk.LEFT)
        ttk.Button(top, text="Solve + Plot", command=self.solve_and_plot).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(top, text="Save Output Excel", command=self.save_output).pack(side=tk.LEFT, padx=(8, 0))

        self.status_var = tk.StringVar(value="Ready. Load an Excel file to begin.")
        ttk.Label(top, textvariable=self.status_var).pack(side=tk.LEFT, padx=(12, 0))

        mid = ttk.Frame(self, padding=(10, 0, 10, 10))
        mid.pack(side=tk.TOP, fill=tk.X)

        conf = ttk.LabelFrame(mid, text="Required confirmations", padding=10)
        conf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.var_cd_ok = tk.BooleanVar(value=False)
        self.var_cids_ok = tk.BooleanVar(value=False)

        ttk.Checkbutton(
            conf,
            text="Confirm input CD is straylight-corrected ΔE spectra (CD-<sample>-<conc>).",
            variable=self.var_cd_ok
        ).pack(anchor="w")

        ttk.Checkbutton(
            conf,
            text="Confirm input CIDS is fully K-factor-normalized (CIDS-<sample>-<conc>).",
            variable=self.var_cids_ok
        ).pack(anchor="w", pady=(4, 0))

        par = ttk.LabelFrame(mid, text="Parameters", padding=10)
        par.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))

        row = ttk.Frame(par)
        row.pack(fill=tk.X)
        ttk.Label(row, text="l_T (constant path length):").pack(side=tk.LEFT)
        self.lt_entry = ttk.Entry(row, width=12)
        self.lt_entry.insert(0, "0.83")
        self.lt_entry.pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(row, text="(your default: 0.83)").pack(side=tk.LEFT, padx=(6, 0))

        info = ttk.Frame(par)
        info.pack(fill=tk.X, pady=(10, 0))
        self.group_info_var = tk.StringVar(value="Detected groups: (none yet)")
        ttk.Label(info, textvariable=self.group_info_var).pack(anchor="w")

        plot_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        plot_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.fig = Figure(figsize=(9, 6), dpi=100)
        self.ax1 = self.fig.add_subplot(311)
        self.ax2 = self.fig.add_subplot(312)
        self.ax3 = self.fig.add_subplot(313)

        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        toolbar = NavigationToolbar2Tk(self.canvas, plot_frame)
        toolbar.update()

    def _set_status(self, msg):
        self.status_var.set(msg)
        self.update_idletasks()

    def load_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return

        try:
            df = pd.read_excel(path, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Failed to read Excel", f"Could not read the Excel file.\n\n{e}")
            return

        df.columns = _sanitize_columns(df.columns)
        self.df = df

        wl_col, groups = _group_columns_by_sample_and_conc(df.columns)
        self.wl_col = wl_col
        self.groups = groups
        self.outputs = None

        if len(groups) == 0:
            self.group_info_var.set("Detected groups: 0 (need headers like CIDS-DVC-0.25, CD-DVC-0.25, ST-DVC-0.25)")
            messagebox.showerror(
                "No matched groups found",
                "I couldn't find any complete triplets of headers:\n"
                "  CIDS-<sampleID>-<conc>\n"
                "  CD-<sampleID>-<conc>\n"
                "  ST-<sampleID>-<conc>\n\n"
                "Example:\n"
                "  CIDS-DVC-0.25, CD-DVC-0.25, ST-DVC-0.25\n"
            )
            return

        preview = []
        for g in groups[:8]:
            preview.append(g["key_label"])
        more = "" if len(groups) <= 8 else f" … (+{len(groups)-8} more)"
        self.group_info_var.set(f"Detected groups: {len(groups)}  |  e.g., " + "; ".join(preview) + more)

        self._set_status(f"Loaded {path.split('/')[-1]} | WL='{wl_col}' | groups={len(groups)}")

    def _read_params(self):
        if not self.var_cd_ok.get():
            messagebox.showerror(
                "Confirmation required",
                "Terminate: The CD spectra in the input file must be straylight-corrected ΔE spectra."
            )
            return None
        if not self.var_cids_ok.get():
            messagebox.showerror(
                "Confirmation required",
                "Terminate: The experimental CIDS spectra in the input file must be fully K-factor-normalized."
            )
            return None

        try:
            lt = float(self.lt_entry.get().strip())
            if not np.isfinite(lt) or lt <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Invalid l_T", "Please enter a valid positive number for l_T.")
            return None

        if self.df is None or self.groups is None or len(self.groups) == 0:
            messagebox.showerror("Missing groups", "Load a file with headers like CIDS-<sample>-<conc>, CD-<sample>-<conc>, ST-<sample>-<conc> first.")
            return None

        return lt

    def solve_and_plot(self):
        if self.df is None:
            messagebox.showinfo("No file", "Load an Excel file first.")
            return
        if self.groups is None or len(self.groups) == 0:
            messagebox.showinfo("No groups", "No matched groups detected (need CIDS-<sample>-<conc>, CD-<sample>-<conc>, ST-<sample>-<conc>).")
            return

        lt = self._read_params()
        if lt is None:
            return

        wl = pd.to_numeric(self.df[self.wl_col], errors="coerce").to_numpy(dtype=float)
        M = wl.shape[0]
        G = len(self.groups)

        cd = np.full((M, G), np.nan, dtype=float)
        cids = np.full((M, G), np.nan, dtype=float)
        st = np.full((M, G), np.nan, dtype=float)
        labels = []
        tags = []  # safe identifiers for column names

        for j, g in enumerate(self.groups):
            labels.append(g["key_label"])
            # tag includes both sample and conc so it is unique
            tag = f"{_slug(g['sample_id'])}_{g['conc_val']:g}"
            tags.append(tag)

            cd[:, j] = pd.to_numeric(self.df[g["CD"]], errors="coerce").to_numpy(dtype=float)
            cids[:, j] = pd.to_numeric(self.df[g["CIDS"]], errors="coerce").to_numpy(dtype=float)
            st[:, j] = pd.to_numeric(self.df[g["ST"]], errors="coerce").to_numpy(dtype=float)

        # Clamp ST
        st_clamped = np.where(np.isfinite(st), st, np.nan)
        st_clamped = np.where(st_clamped < MIN_ST, MIN_ST, st_clamped)

        # denom = 10^ST - 1
        denom = np.power(10.0, st_clamped) - 1.0

        # Solve ΔA with c=1 (matches your manual calculation)
        numerator = cids + (cd / denom)
        x = lt + (1.0 / denom)
        deltaA = np.where(np.isfinite(x) & (x != 0.0), numerator / x, np.nan)

        # Predict CIDS for validation
        cids_pred = (lt) * deltaA - (cd - deltaA) / denom

        resid = cids_pred - cids
        mspe = np.nanmean(resid * resid, axis=1)  # (M,)

        # Output sheets
        out = pd.DataFrame({self.wl_col: wl, "MSPE_CIDS": mspe})
        deltaA_only = pd.DataFrame({self.wl_col: wl})

        for j, g in enumerate(self.groups):
            tag = tags[j]
            out[f"SampleID_{tag}"] = g["sample_id"]
            out[f"Conc_{tag}"] = g["conc_val"]
            out[f"CD_{tag}"] = cd[:, j]
            out[f"CIDS_obs_{tag}"] = cids[:, j]
            out[f"ST_{tag}"] = st_clamped[:, j]
            out[f"DeltaA_CC_{tag}"] = deltaA[:, j]
            out[f"CD_minus_DeltaA_{tag}"] = cd[:, j] - deltaA[:, j]
            out[f"CIDS_pred_{tag}"] = cids_pred[:, j]
            out[f"denom_(10^ST-1)_{tag}"] = denom[:, j]

            deltaA_only[f"DeltaA_CC_{tag}"] = deltaA[:, j]

        out["l_T"] = lt
        deltaA_only["l_T"] = lt

        self.outputs = {
            "wl": wl,
            "labels": labels,
            "tags": tags,
            "cd": cd,
            "cids_obs": cids,
            "cids_pred": cids_pred,
            "deltaA": deltaA,
            "mspe": mspe,
            "out_df": out,
            "deltaA_only_df": deltaA_only,
        }

        self._plot()
        self._set_status(f"Solved ΔA(λ) for {G} groups (sampleID + conc). Ready to save output Excel.")

    def _plot(self):
        if not self.outputs:
            return

        wl = self.outputs["wl"]
        labels = self.outputs["labels"]
        cd = self.outputs["cd"]
        deltaA = self.outputs["deltaA"]
        cids_obs = self.outputs["cids_obs"]
        cids_pred = self.outputs["cids_pred"]
        mspe = self.outputs["mspe"]

        self.ax1.clear()
        self.ax2.clear()
        self.ax3.clear()

        # Ax1: CD, ΔA, CD-ΔA
        for j, lab in enumerate(labels):
            self.ax1.plot(wl, cd[:, j], linestyle="--", label=f"CD ({lab})")
            self.ax1.plot(wl, deltaA[:, j], linestyle="-", label=f"ΔA ({lab})")
            self.ax1.plot(wl, cd[:, j] - deltaA[:, j], linestyle=":", label=f"CD-ΔA ({lab})")

        self.ax1.set_ylabel("ΔE (dimensionless)")
        self.ax1.set_title("CD, ΔA, and CD-ΔA (grouped by sampleID + concentration)")
        self.ax1.legend(loc="best", fontsize=7, ncol=2)
        self.ax1.grid(True, alpha=0.3)

        # Ax2: CIDS obs vs pred
        for j, lab in enumerate(labels):
            self.ax2.plot(wl, cids_obs[:, j], linestyle="--", label=f"CIDS obs ({lab})")
            self.ax2.plot(wl, cids_pred[:, j], linestyle="-", label=f"CIDS pred ({lab})")

        self.ax2.set_ylabel("CIDS (dimensionless)")
        self.ax2.set_title("CIDS validation: observed (dashed) vs predicted (solid)")
        self.ax2.legend(loc="best", fontsize=7, ncol=2)
        self.ax2.grid(True, alpha=0.3)

        # Ax3: MSPE
        self.ax3.plot(wl, mspe, label="MSPE(λ)")
        self.ax3.set_xlabel("Wavelength (nm)")
        self.ax3.set_ylabel("Mean square prediction error")
        self.ax3.set_title("MSPE spectrum (avg across groups)")
        self.ax3.legend(loc="best", fontsize=9)
        self.ax3.grid(True, alpha=0.3)

        self.fig.tight_layout()
        self.canvas.draw()

    def save_output(self):
        if not self.outputs:
            messagebox.showinfo("Nothing to save", "Run 'Solve + Plot' first.")
            return

        path = filedialog.asksaveasfilename(
            title="Save output Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not path:
            return

        try:
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                self.outputs["out_df"].to_excel(writer, index=False, sheet_name="CIDS_Decomposition")
                self.outputs["deltaA_only_df"].to_excel(writer, index=False, sheet_name="DeltaA_only")
        except Exception as e:
            messagebox.showerror("Save failed", f"Could not save the Excel file.\n\n{e}")
            return

        self._set_status(f"Saved output: {path.split('/')[-1]}")
        messagebox.showinfo("Saved", "Output Excel saved successfully.")

    def run(self):
        self.mainloop()


if __name__ == "__main__":
    App().run()
