"""
CIDS correction, validation, and chirality decomposition
-------------------------------------------------------
GUI UPDATED: plots ΔA, ΔS, and ΔE_total vs wavelength in 3 separate graphs (3 subplots).

Header format supported:
  CIDS-<sampleID>-<conc>
  CD-<sampleID>-<conc>
  ST-<sampleID>-<conc>

Equations:
  denom = 10^ST - 1   (ST clamped to >= 0.01)

  (A) ΔA:
    ΔA = (CIDS + CD/denom) / (l_T + 1/denom)

  (B) ΔS:
    ΔS = (l_T*CD - CIDS) / ( K * ( l_T + 1/denom ) )

  (C) ΔE_total:
    ΔE_total = ΔA + ΔS

Outputs (Excel):
  Sheet 1: CIDS_Decomposition (full table incl ΔA, ΔS, ΔE_total)
  Sheet 2: DeltaA_only
  Sheet 3: DeltaS_only
  Sheet 4: DeltaE_only

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
    s = str(s).strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^0-9A-Za-z._\-]+", "_", s)
    return s


def _parse_kind_sample_conc(colname: str):
    """
    Parse: KIND-<sampleID>-<conc>
    sampleID may contain dashes; split from the RIGHT to get concentration.
    """
    m = re.match(r"^\s*(CIDS|CD|ST)\s*[-_]\s*(.+)\s*$", colname, flags=re.IGNORECASE)
    if not m:
        return None
    kind = m.group(1).upper()
    rest = m.group(2).strip()

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

    keys = sorted(set((sid, c) for (k, sid, c) in found.keys()))
    groups = []
    for sid, cval in keys:
        if (("CIDS", sid, cval) in found) and (("CD", sid, cval) in found) and (("ST", sid, cval) in found):
            cids_col, lab1 = found[("CIDS", sid, cval)]
            cd_col, lab2 = found[("CD", sid, cval)]
            st_col, lab3 = found[("ST", sid, cval)]
            conc_label = lab1 or lab2 or lab3 or f"{cval:g}"
            groups.append({
                "sample_id": sid,
                "conc_val": cval,
                "conc_label": conc_label,
                "key_label": f"{sid} @ {conc_label}",
                "CIDS": cids_col,
                "CD": cd_col,
                "ST": st_col,
            })

    groups.sort(key=lambda g: (g["sample_id"].lower(), g["conc_val"]))
    return wl_col, groups


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1200x880")

        self.df = None
        self.wl_col = None
        self.groups = None
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

        row1 = ttk.Frame(par)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="l_T (path length l):").pack(side=tk.LEFT)
        self.lt_entry = ttk.Entry(row1, width=10)
        self.lt_entry.insert(0, "0.83")
        self.lt_entry.pack(side=tk.LEFT, padx=(6, 14))

        ttk.Label(row1, text="K (constant in ΔS):").pack(side=tk.LEFT)
        self.k_entry = ttk.Entry(row1, width=10)
        self.k_entry.insert(0, "1.0")
        self.k_entry.pack(side=tk.LEFT, padx=(6, 0))

        info = ttk.Frame(par)
        info.pack(fill=tk.X, pady=(10, 0))
        self.group_info_var = tk.StringVar(value="Detected groups: (none yet)")
        ttk.Label(info, textvariable=self.group_info_var).pack(anchor="w")

        plot_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        plot_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.fig = Figure(figsize=(9, 6), dpi=100)
        # Three plots: ΔA, ΔS, ΔE_total
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
            self.group_info_var.set("Detected groups: 0 (need CIDS-<sample>-<conc>, CD-<sample>-<conc>, ST-<sample>-<conc>)")
            messagebox.showerror(
                "No matched groups found",
                "I couldn't find any complete triplets:\n"
                "  CIDS-<sampleID>-<conc>, CD-<sampleID>-<conc>, ST-<sampleID>-<conc>"
            )
            return

        preview = [g["key_label"] for g in groups[:8]]
        more = "" if len(groups) <= 8 else f" … (+{len(groups)-8} more)"
        self.group_info_var.set(f"Detected groups: {len(groups)} | e.g., " + "; ".join(preview) + more)
        self._set_status(f"Loaded {path.split('/')[-1]} | WL='{wl_col}' | groups={len(groups)}")

    def _read_params(self):
        if not self.var_cd_ok.get():
            messagebox.showerror("Confirmation required", "Terminate: CD must be straylight-corrected ΔE spectra.")
            return None
        if not self.var_cids_ok.get():
            messagebox.showerror("Confirmation required", "Terminate: CIDS must be fully K-factor-normalized.")
            return None

        try:
            lt = float(self.lt_entry.get().strip())
            if not np.isfinite(lt) or lt <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Invalid l_T", "Please enter a valid positive number for l_T.")
            return None

        try:
            K = float(self.k_entry.get().strip())
            if not np.isfinite(K) or K == 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Invalid K", "Please enter a valid nonzero number for K.")
            return None

        if self.df is None or not self.groups:
            messagebox.showerror("Missing groups", "Load a valid Excel file first.")
            return None

        return lt, K

    def solve_and_plot(self):
        params = self._read_params()
        if params is None:
            return
        lt, K = params

        wl = pd.to_numeric(self.df[self.wl_col], errors="coerce").to_numpy(dtype=float)
        M = wl.shape[0]
        G = len(self.groups)

        cd = np.full((M, G), np.nan, dtype=float)
        cids = np.full((M, G), np.nan, dtype=float)
        st = np.full((M, G), np.nan, dtype=float)
        labels = []
        tags = []

        for j, g in enumerate(self.groups):
            labels.append(g["key_label"])
            tag = f"{_slug(g['sample_id'])}_{g['conc_val']:g}"
            tags.append(tag)

            cd[:, j] = pd.to_numeric(self.df[g["CD"]], errors="coerce").to_numpy(dtype=float)
            cids[:, j] = pd.to_numeric(self.df[g["CIDS"]], errors="coerce").to_numpy(dtype=float)
            st[:, j] = pd.to_numeric(self.df[g["ST"]], errors="coerce").to_numpy(dtype=float)

        # Clamp ST
        st_clamped = np.where(np.isfinite(st), st, np.nan)
        st_clamped = np.where(st_clamped < MIN_ST, MIN_ST, st_clamped)

        denom = np.power(10.0, st_clamped) - 1.0  # 10^ST - 1

        # ΔA
        numerator_A = cids + (cd / denom)
        x_A = lt + (1.0 / denom)
        deltaA = np.where(np.isfinite(x_A) & (x_A != 0.0), numerator_A / x_A, np.nan)

        # Predict CIDS + MSPE (still computed for QC, not plotted)
        cids_pred = (lt) * deltaA - (cd - deltaA) / denom
        resid = cids_pred - cids
        mspe = np.nanmean(resid * resid, axis=1)

        # ΔS
        denom_B = K * (lt + (1.0 / denom))
        deltaS = np.where(np.isfinite(denom_B) & (denom_B != 0.0), (lt * cd - cids) / denom_B, np.nan)

        # ΔE_total
        deltaE = deltaA + deltaS

        # Outputs
        out = pd.DataFrame({self.wl_col: wl, "MSPE_CIDS": mspe})
        deltaA_only = pd.DataFrame({self.wl_col: wl})
        deltaS_only = pd.DataFrame({self.wl_col: wl})
        deltaE_only = pd.DataFrame({self.wl_col: wl})

        for j, g in enumerate(self.groups):
            tag = tags[j]
            out[f"SampleID_{tag}"] = g["sample_id"]
            out[f"Conc_{tag}"] = g["conc_val"]

            out[f"CD_{tag}"] = cd[:, j]
            out[f"CIDS_obs_{tag}"] = cids[:, j]
            out[f"ST_{tag}"] = st_clamped[:, j]
            out[f"denom_(10^ST-1)_{tag}"] = denom[:, j]

            out[f"DeltaA_CC_{tag}"] = deltaA[:, j]
            out[f"DeltaS_{tag}"] = deltaS[:, j]
            out[f"DeltaE_total_{tag}"] = deltaE[:, j]

            out[f"CIDS_pred_{tag}"] = cids_pred[:, j]

            deltaA_only[f"DeltaA_CC_{tag}"] = deltaA[:, j]
            deltaS_only[f"DeltaS_{tag}"] = deltaS[:, j]
            deltaE_only[f"DeltaE_total_{tag}"] = deltaE[:, j]

        for df_ in (out, deltaA_only, deltaS_only, deltaE_only):
            df_["l_T"] = lt
            df_["K"] = K

        self.outputs = {
            "wl": wl,
            "labels": labels,
            "tags": tags,
            "deltaA": deltaA,
            "deltaS": deltaS,
            "deltaE": deltaE,
            "mspe": mspe,
            "out_df": out,
            "deltaA_only_df": deltaA_only,
            "deltaS_only_df": deltaS_only,
            "deltaE_only_df": deltaE_only,
        }

        self._plot_deltaA_deltaS_deltaE()
        self._set_status(f"Computed ΔA, ΔS, and ΔE_total for {G} groups. Ready to save Excel.")

    def _plot_deltaA_deltaS_deltaE(self):
        """Plots ΔA, ΔS, ΔE_total vs wavelength on three separate subplots."""
        if not self.outputs:
            return

        wl = self.outputs["wl"]
        labels = self.outputs["labels"]
        dA = self.outputs["deltaA"]
        dS = self.outputs["deltaS"]
        dE = self.outputs["deltaE"]

        self.ax1.clear()
        self.ax2.clear()
        self.ax3.clear()

        # ΔA
        for j, lab in enumerate(labels):
            self.ax1.plot(wl, dA[:, j], label=lab)
        self.ax1.set_ylabel("ΔA (dimensionless)")
        self.ax1.set_title("Differential absorption extinction ΔA(λ)")
        self.ax1.legend(loc="best", fontsize=7, ncol=2)
        self.ax1.grid(True, alpha=0.3)

        # ΔS
        for j, lab in enumerate(labels):
            self.ax2.plot(wl, dS[:, j], label=lab)
        self.ax2.set_ylabel("ΔS (dimensionless)")
        self.ax2.set_title("Differential scattering extinction ΔS(λ)")
        self.ax2.legend(loc="best", fontsize=7, ncol=2)
        self.ax2.grid(True, alpha=0.3)

        # ΔE_total
        for j, lab in enumerate(labels):
            self.ax3.plot(wl, dE[:, j], label=lab)
        self.ax3.set_xlabel("Wavelength (nm)")
        self.ax3.set_ylabel("ΔE_total (dimensionless)")
        self.ax3.set_title("Differential total extinction ΔE_total(λ) = ΔA(λ) + ΔS(λ)")
        self.ax3.legend(loc="best", fontsize=7, ncol=2)
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
                self.outputs["deltaS_only_df"].to_excel(writer, index=False, sheet_name="DeltaS_only")
                self.outputs["deltaE_only_df"].to_excel(writer, index=False, sheet_name="DeltaE_only")
        except Exception as e:
            messagebox.showerror("Save failed", f"Could not save the Excel file.\n\n{e}")
            return

        self._set_status(f"Saved output: {path.split('/')[-1]}")
        messagebox.showinfo("Saved", "Output Excel saved successfully.")

    def run(self):
        self.mainloop()


if __name__ == "__main__":
    App().run()