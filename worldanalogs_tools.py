
# worldanalogs_tools.py — .xls via pyexcel-xls (works with Pandas 2.x), .xlsx via openpyxl
from __future__ import annotations
import os, re
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional, Any
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

def _safe_read_excel(path: str) -> Dict[str, pd.DataFrame]:
    """
    Robustly read all sheets from an Excel workbook into a dict of DataFrames.
    - .xls  -> use pyexcel-xls (works even when Pandas 2.x + xlrd conflict)
    - .xlsx -> use pandas + openpyxl
    """
    ext = os.path.splitext(path)[1].lower()
    sheets: Dict[str, pd.DataFrame] = {}

    if ext == ".xls":
        try:
            from pyexcel_xls import get_data
        except Exception as e:
            raise RuntimeError("pyexcel-xls is required to open .xls files. Please ensure 'pyexcel-xls' is installed.") from e

        raw = get_data(path)  # dict: sheet_name -> list of rows
        for name, rows in raw.items():
            if not rows:
                continue
            # Find first non-empty row as header
            header_idx = None
            for i, r in enumerate(rows):
                # consider a row non-empty if any cell has non-blank content
                if any((c is not None and str(c).strip() != "") for c in r):
                    header_idx = i
                    break
            if header_idx is None:
                continue
            header = [str(c).strip() for c in rows[header_idx]]
            body = rows[header_idx+1:]
            # Normalize row lengths to header length
            norm = []
            for r in body:
                r = list(r)
                if len(r) < len(header):
                    r = r + [None] * (len(header) - len(r))
                norm.append(r[:len(header)])
            df = pd.DataFrame(norm, columns=header)
            df.columns = [str(c).strip() for c in df.columns]
            sheets[name.strip()] = df
        if not sheets:
            raise RuntimeError("No readable sheets found in .xls file. Ensure the first non-empty row in each sheet contains headers.")
        return sheets

    else:
        # .xlsx path
        try:
            xls = pd.ExcelFile(path, engine="openpyxl")
        except Exception as e:
            raise RuntimeError(f"Failed to open '{path}' with openpyxl: {e}")
        for name in xls.sheet_names:
            df = xls.parse(name)
            df.columns = [str(c).strip() for c in df.columns]
            sheets[name.strip()] = df
        return sheets

@dataclass
class WorldAnalogs:
    sheets: Dict[str, pd.DataFrame]

    @classmethod
    def load(cls, path: str) -> "WorldAnalogs":
        sheets = _safe_read_excel(path)
        return cls(sheets=sheets)

    def sheet(self, name: str) -> pd.DataFrame:
        keys = {k.lower(): k for k in self.sheets.keys()}
        key = keys.get(name.lower())
        if key is None:
            raise KeyError(f"Sheet '{name}' not found. Available: {list(self.sheets.keys())}")
        return self.sheets[key]

    @property
    def geology(self) -> pd.DataFrame:
        return self.sheet("Geology")

    @property
    def oil(self) -> pd.DataFrame:
        return self.sheet("Oil")

    @property
    def gas(self) -> pd.DataFrame:
        return self.sheet("Gas")

    @property
    def boe(self) -> pd.DataFrame:
        return self.sheet("BOE")

    def ancillary(self) -> Optional[pd.DataFrame]:
        try:
            return self.sheet("Ancillary")
        except KeyError:
            return None

    
    @staticmethod
    def _find_col(df: pd.DataFrame, patterns: List[str]) -> Optional[str]:
        """
        Robust column finder:
        - Case-insensitive
        - Ignores spaces/underscores/dashes and punctuation
        - Falls back to regex search (case-insensitive)
        """
        def norm(s: str) -> str:
            return re.sub(r"\W+", "", str(s).lower())  # remove non-alphanumerics

        cols = list(df.columns)
        norms = {c: norm(c) for c in cols}

        # First pass: exact normalized match
        for pat in patterns:
            npat = norm(pat)
            for c in cols:
                if norms[c] == npat:
                    return c

        # Second pass: regex search on raw column names
        for pat in patterns:
            try:
                rx = re.compile(pat, re.I)
            except re.error:
                continue
            for c in cols:
                if rx.search(str(c)):
                    return c
        return None
    

    def find_column(self, sheet: str, patterns: List[str]) -> str:
        df = self.sheet(sheet)
        col = self._find_col(df, patterns)
        if not col:
            raise KeyError(f"Column matching {patterns} not found in sheet '{sheet}'.")
        return col

    def list_classification_vars(self) -> List[str]:
        candidates = [
            "AU_Code", "AU Name", "AU_Name", "TPS_Code", "TPS_Name",
            "Province Code", "Province Name",
            "Structural Setting", "Crustal System", "Architecture",
            "Trap System (Major)", "Depositional System",
            "Source Rock Depositional Environment", "Kerogen Type",
            "Source Type", "Source Rock Qualifier", "Status",
            "General Reservoir Rock Age", "Reservoir Rock Lithology",
            "Reservoir Rock Depositional Environment",
            "Seal Rock Lithology", "Trap Type", "General Source Rock Age",
            "Migration Distance"
        ]
        g = self.geology
        present = [c for c in candidates if c in g.columns]
        for c in g.columns:
            if c not in present and g[c].dtype == "object":
                present.append(c)
        return present

    def filter_analogs(self, filters: Dict[str, List[Any]]) -> pd.DataFrame:
        g = self.geology.copy()
        for col, allowed in filters.items():
            if col not in g.columns:
                match = self._find_col(g, [col])
                if match:
                    col = match
                else:
                    raise KeyError(f"Filter column '{col}' not found in Geology sheet.")
            allowed_norm = {str(v).strip().lower() for v in allowed}
            g = g[g[col].astype(str).str.strip().str.lower().isin(allowed_norm)]
        return g

    def extend_selection(self, selected_au_codes: List[Any]) -> Dict[str, pd.DataFrame]:
        key_geo = self._find_col(self.geology, ["AU_Code"])
        if not key_geo:
            raise KeyError("AU_Code not found in Geology sheet.")
        au_set = set(map(str, selected_au_codes))
        out = {}
        for name, df in self.sheets.items():
            key = self._find_col(df, ["AU_Code"])
            if key:
                sub = df[df[key].astype(str).isin(au_set)].copy()
                out[name] = sub
        return out

    def resolve_utility_column(self, which_sheet: str, logical_name: str) -> str:
        mapping = {
            "number_density_gt5":       [r"^Number\s*/\s*1000\s*km2\s*for\s*>\s*5"],
            "number_density_gt50":      [r"^Number\s*/\s*1000\s*km2\s*for\s*>\s*50"],
            "discovered_pct_num_gt5":   [r"^Discovered\s*%\s*by\s*Number\s*for\s*>\s*5"],
            "discovered_pct_num_gt50":  [r"^Discovered\s*%\s*by\s*Number\s*for\s*>\s*50"],
            "discovered_pct_vol_gt5":   [r"^Discovered\s*%\s*by\s*Volume\s*for\s*>\s*5"],
            "discovered_pct_vol_gt50":  [r"^Discovered\s*%\s*by\s*Volume\s*for\s*>\s*50"],
            "median_gt5":               [r"^Median\s*of\s*>\s*5"],
            "median_gt50":              [r"^Median\s*of\s*>\s*50"],
            "maximum_gt5":              [r"^Maximum\s*of\s*>\s*5"],
            "maximum_gt50":             [r"^Maximum\s*of\s*>\s*50"],
        }
        pats = mapping.get(logical_name)
        if not pats:
            raise KeyError(f"Unknown logical variable '{logical_name}'.")
        return self.find_column(which_sheet, pats)

    def _au_name_column(self, df: pd.DataFrame) -> Optional[str]:
        return self._find_col(df, ["AU Name", "AU_Name", "AU name", "Assessment Unit Name", "AUName"])

    def analog_plot(self, sheet_name: str, value_col_logical: str, maturity_mode: str = "volume_gt50",
                    selected_au_codes: Optional[List[Any]] = None, fig_size: Tuple[int, int] = (10, 6)) -> plt.Figure:
        df = self.sheet(sheet_name).copy()
        if selected_au_codes is not None:
            key = self._find_col(df, ["AU_Code"])
            if not key:
                raise KeyError(f"AU_Code not found in sheet '{sheet_name}'.")
            df = df[df[key].astype(str).isin(set(map(str, selected_au_codes)))].copy()

        val_col = self.resolve_utility_column(sheet_name, value_col_logical)
        maturity_map = {
            "number_gt5":  "discovered_pct_num_gt5",
            "number_gt50": "discovered_pct_num_gt50",
            "volume_gt5":  "discovered_pct_vol_gt5",
            "volume_gt50": "discovered_pct_vol_gt50",
        }
        mat_logical = maturity_map.get(maturity_mode.lower())
        if not mat_logical:
            raise KeyError("maturity_mode must be one of: number_gt5, number_gt50, volume_gt5, volume_gt50")
        mat_col = self.resolve_utility_column(sheet_name, mat_logical)

        name_col = self._au_name_column(df) or df.columns[0]
        sub = df[[name_col, val_col, mat_col]].dropna().copy()
        sub = sub.sort_values(val_col).reset_index(drop=True)

        vals = sub[val_col].astype(float).values
        mats = sub[mat_col].astype(float).clip(0, 100).values / 100.0
        names = sub[name_col].astype(str).values

        fig = plt.figure(figsize=fig_size)
        ax = fig.add_subplot(111)
        for x, m, nm in zip(vals, mats, names):
            ax.vlines(x, 0, 1, linewidth=1.0, alpha=0.5)
            ax.vlines(x, 0, m, linewidth=4.0, alpha=0.9)
            ax.text(x, 1.02, nm, rotation=90, va="bottom", ha="center", fontsize=8)
        ax.set_ylim(0, 1.1)
        ax.set_xlabel(val_col)
        ax.set_ylabel("Maturity (fraction discovered)")
        ax.set_title(f"Analog Plot • {sheet_name} • {val_col}  • maturity={mat_col}")
        return fig

    def analog_histogram(self, sheet_name: str, value_col_logical: str,
                         selected_au_codes: Optional[List[Any]] = None, bins: Optional[int] = None,
                         fig_size: Tuple[int, int] = (10, 6)) -> plt.Figure:
        df = self.sheet(sheet_name).copy()
        if selected_au_codes is not None:
            key = self._find_col(df, ["AU_Code"])
            if not key:
                raise KeyError(f"AU_Code not found in sheet '{sheet_name}'.")
            df = df[df[key].astype(str).isin(set(map(str, selected_au_codes)))].copy()

        val_col = self.resolve_utility_column(sheet_name, value_col_logical)
        vals = df[val_col].dropna().astype(float).values

        if bins is None:
            if len(vals) > 1:
                iqr = np.subtract(*np.percentile(vals, [75, 25]))
                bin_width = 2 * iqr * (len(vals) ** (-1/3)) if iqr > 0 else 0
                if bin_width > 0:
                    bins = max(8, int(np.ceil((vals.max() - vals.min()) / bin_width)))
                else:
                    bins = 20
            else:
                bins = 10

        fig = plt.figure(figsize=fig_size)
        ax = fig.add_subplot(111)
        ax.hist(vals, bins=bins)
        ax.set_xlabel(val_col)
        ax.set_ylabel("Frequency")
        ax.set_title(f"Analog Histogram • {sheet_name} • {val_col}  (n={len(vals)})")
        return fig

    def export_selection_csvs(self, selected_au_codes: List[Any], out_dir: str) -> List[str]:
        import os
        os.makedirs(out_dir, exist_ok=True)
        files = []
        au_set = set(map(str, selected_au_codes))
        for name, df in self.sheets.items():
            key = self._find_col(df, ["AU_Code"])
            if key:
                sub = df[df[key].astype(str).isin(au_set)].copy()
                out_path = os.path.join(out_dir, f"{name}_selection.csv")
                sub.to_csv(out_path, index=False)
                files.append(out_path)
        return files
