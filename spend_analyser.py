"""
Procurement Spend Analyzer (OOP Version)
Automated tool for analyzing supplier spend data and generating insights
Author: Jean Kocman
"""

import pandas as pd
import matplotlib.pyplot as plt
import argparse
import glob
import csv
import os
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any, NoReturn


logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

COLUMN_MAP = {
    "supplier": ["supplier", "vendor", "supplier name", "nominated supplier", "counterpart"],
    "spend": ["spend", "amount", "value", "cost"],
    "category": ["category", "group", "type"],
    "year": ["year", "fiscal year", "period"],
}

YEAR_ALLOWED_HELP = (
    "Column 'Year' must be in format YYYY (e.g., 2024) or YYYY-MM-DD (e.g., 2024-03-15). "
    "Other formats are not supported; please correct the input data accordingly."
)


def die(msg: str) -> NoReturn:
    logger.error(msg)
    raise SystemExit(1)


def choose_file() -> str:
    """Let user choose a sample_data file from current directory by index and return its path."""
    files: list[str] = []
    for ext in ["csv", "xls", "xlsx"]:
        files.extend(glob.glob(f"sample_data*.{ext}"))

    if not files:
        die("No sample_data file found (csv/xls/xlsx)")

    print("\nAvailable files:")
    for i, f in enumerate(files, start=1):
        stat = os.stat(f)
        modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        print(f"{i}. {Path(f).name} (modified: {modified})")

    choice = input("\nSelect a file by number: ")
    try:
        idx = int(choice)
        if 1 <= idx <= len(files):
            return files[idx - 1]
        die("Invalid choice")
    except ValueError:
        die("Invalid input")


def detect_sep(path: str | Path) -> str:
    """Try to sniff CSV delimiter. Falls back to ',' if detection fails."""
    # Fallback encoding: CSV z evropských ERP často není čisté UTF-8
    for enc in ("utf-8", "cp1250", "latin2"):
        try:
            with open(path, newline="", encoding=enc) as f:
                sample = f.read(4096)
            dialect = csv.Sniffer().sniff(sample, delimiters=",;|	")
            return dialect.delimiter
        except UnicodeDecodeError:
            continue
        except Exception:
            return ","
    return ","


def parse_spend(series: pd.Series) -> pd.Series:
    s = series.astype("string").str.strip()
    s = s.str.replace(r"[^0-9,.\-]", "", regex=True)

    has_comma = s.str.contains(",", na=False)
    has_dot = s.str.contains(r"\.", na=False)
    both = has_comma & has_dot

    if both.any():
        last_comma = s.str.rfind(",")
        last_dot = s.str.rfind(".")

        # US: 1,234.56  -> remove commas
        us = both & (last_dot > last_comma)
        s = s.mask(us, s.str.replace(",", "", regex=False))

        # EU: 1.234,56  -> remove dots, replace comma -> dot
        eu = both & ~us
        s = s.mask(eu, s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))

    # Only comma: 123,45 -> 123.45
    only_comma = has_comma & ~has_dot
    s = s.mask(only_comma, s.str.replace(",", ".", regex=False))

    return pd.to_numeric(s, errors="coerce")


class SpendAnalyzer:
    def __init__(self, file_path: str | Path) -> None:
        """Initialize spend analyser with path to input file."""
        self.file_path: Path = Path(file_path)
        self.df: pd.DataFrame = pd.DataFrame()
        self.results: Optional[Dict[str, Any]] = None

    def load_data(self, csv_sep: Optional[str] = None) -> None:
        """Load data from CSV or Excel (all sheets combined)."""
        try:
            suffix = self.file_path.suffix.lower()

            if suffix == ".csv":
                sep = csv_sep or detect_sep(self.file_path)
                # encoding_errors='replace' pomůže, když se objeví „divné“ znaky
                try:
                    self.df = pd.read_csv(self.file_path, sep=sep, encoding_errors="replace")
                except TypeError:
                    self.df = pd.read_csv(self.file_path, sep=sep)

            elif suffix in [".xls", ".xlsx"]:
                with pd.ExcelFile(self.file_path) as xls:
                    frames = [pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names]
                self.df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

            else:
                raise ValueError(f"Unsupported file format: {suffix}")

            if self.df.empty or len(self.df.columns) == 0:
                die("The file is empty or has no columns")

            logger.info(f"Loaded {len(self.df)} records from {self.file_path}")

        except Exception as e:
            die(f"Error loading data: {e}")

        self.standardize_columns()

    def standardize_columns(self) -> None:
        """Rename columns to unified internal names (supplier, spend, category, year)."""
        if self.df.empty or len(self.df.columns) == 0:
            die("No data loaded, cannot standardize columns. Run load_data() first.")

        self.df.columns = self.df.columns.str.strip().str.lower()

        rename_map: dict[str, str] = {}
        for unified_name, possible_names in COLUMN_MAP.items():
            for name in possible_names:
                name_l = name.lower()
                if name_l in self.df.columns:
                    rename_map[name_l] = unified_name
                    break

        if rename_map:
            self.df = self.df.rename(columns=rename_map)

        required = ["supplier", "spend"]
        missing_required = [col for col in required if col not in self.df.columns]
        if missing_required:
            die(f"Missing required columns: {missing_required}")

        optional = ["category", "year"]
        missing_optional = [col for col in optional if col not in self.df.columns]
        if missing_optional:
            logger.warning(f"Missing optional columns: {missing_optional}. Some analyses will be skipped.")

        logger.info("Column names standardized")

    def clean_data(self) -> None:
        """Clean and standardize data formats for analysis."""
        if self.df.empty or "spend" not in self.df.columns:
            die("No data loaded or missing 'spend' column. Run load_data() first.")

        df = self.df.copy()

        # spend -> numeric (robust)
        df["spend"] = parse_spend(df["spend"])

        # year -> YYYY or YYYY-MM-DD (optional)
        if "year" in df.columns:
            y = df["year"]
            if pd.api.types.is_numeric_dtype(y):
                y_num = pd.to_numeric(y, errors="coerce")
                # 2024.0 OK, 2024.5 -> NA
                y_int = y_num.where((y_num % 1) == 0)
                df["year"] = y_int.astype("Int64")
            else:
                ys = y.astype("string").str.strip()
                ys = ys.replace(["", "nan", "none", "nat"], pd.NA)

                out = pd.Series(pd.NA, index=df.index, dtype="Int64")

                is_year = ys.str.fullmatch(r"\d{4}", na=False)
                out.loc[is_year] = pd.to_numeric(ys.loc[is_year], errors="coerce").astype("Int64")

                rem = out.isna() & ys.notna()
                if rem.any():
                    try:
                        dt = pd.to_datetime(ys.loc[rem], format="%Y-%m-%d", errors="raise")
                        out.loc[rem] = pd.Series(dt.dt.year.to_numpy(), index=ys.loc[rem].index).astype("Int64")
                    except Exception:
                        examples = ys.loc[rem].head(5).tolist()
                        die(f"{YEAR_ALLOWED_HELP} Found: {examples}")

                df["year"] = out

            bad = df["year"].notna() & ((df["year"] < 2000) | (df["year"] > 2100))
            if bad.any():
                examples = df.loc[bad, "year"].head(5).tolist()
                die(f"{YEAR_ALLOWED_HELP} Suspicious range 2000–2100, e.g.: {examples}")

        # Clean text fields (do not convert NA -> 'nan')
        if "supplier" in df.columns:
            s = df["supplier"].astype("string").str.strip()
            df["supplier"] = s.where(s.notna(), pd.NA).str.title()
        if "category" in df.columns:
            c = df["category"].astype("string").str.strip()
            df["category"] = c.where(c.notna(), pd.NA).str.title()

        # Drop rows without spend value
        df = df.dropna(subset=["spend"])

        # comment in case you wan to inlude debit notes 
        df = df.loc[df["spend"] >= 0].copy()

        # Remove duplicates
        df = df.drop_duplicates()

        if df.empty:
            die("No valid rows after cleaning")

        self.df = df
        logger.info("Data cleaned and standardized")

    def analyze(self) -> Dict[str, Any]:
        """Perform spend analysis and store results dictionary."""
        if self.df.empty or "supplier" not in self.df.columns or "spend" not in self.df.columns:
            die("No data to analyze. Run load_data() and clean_data() first.")

        top_suppliers = (
            self.df.groupby("supplier", dropna=True)["spend"].sum().nlargest(10).sort_values(ascending=True)
        )

        spend_by_category = (
            self.df.groupby("category", dropna=True)["spend"].sum().sort_values(ascending=True)
            if "category" in self.df.columns
            else pd.Series(dtype=float)
        )

        yoy_trend = (
            self.df.groupby("year", dropna=True)["spend"].sum().sort_index()
            if "year" in self.df.columns
            else pd.Series(dtype=float)
        )

        self.results = {
            "top_suppliers": top_suppliers,
            "spend_by_category": spend_by_category,
            "yoy_trend": yoy_trend,
            "total_spend": float(self.df["spend"].sum()),
            "avg_spend": float(self.df["spend"].mean()),
            "supplier_count": int(self.df["supplier"].nunique(dropna=True)),
        }

        logger.info("Analysis complete")
        return self.results

    def report(self) -> None:
        """Print a readable text summary report."""
        if self.results is None:
            die("No results to report")

        r = self.results

        print("=" * 60)
        print("PROCUREMENT SPEND ANALYSIS REPORT")
        print("=" * 60)
        print(f"Total Spend: €{r['total_spend']:,.2f}")
        print(f"Average Transaction: €{r['avg_spend']:,.2f}")
        print(f"Number of Suppliers: {r['supplier_count']}")

        print("Top 3 Suppliers:")
        for supplier, spend in r["top_suppliers"].tail(3).items():
            print(f"  • {supplier}: €{spend:,.2f}")

        if not r["spend_by_category"].empty:
            print("Top 3 Categories:")
            for category, spend in r["spend_by_category"].tail(3).items():
                print(f"  • {category}: €{spend:,.2f}")

        if not r["yoy_trend"].empty and len(r["yoy_trend"].index) >= 2:
            latest = r["yoy_trend"].index[-1]
            previous = r["yoy_trend"].index[-2]

            prev_val = float(r["yoy_trend"].loc[previous])
            last_val = float(r["yoy_trend"].loc[latest])

            if prev_val == 0:
                print(f"YoY Growth ({previous} → {latest}): n/a (previous year spend = 0)")
            else:
                growth = ((last_val / prev_val) - 1) * 100
                print(f"YoY Growth ({previous} → {latest}): {growth:+.1f}%")

        print("=" * 60 + "")

    def visualize(self, output_path: str = "spend_analysis.png") -> None:
        """Generate and save visual charts from analysis results."""
        if self.results is None:
            die("No results to visualize")

        r = self.results

        fig, axes = plt.subplots(1, 3, figsize=(18, 5))
        fig.suptitle("Procurement Spend Analysis Dashboard", fontsize=16, fontweight="bold")

        # Top suppliers
        if not r["top_suppliers"].empty:
            r["top_suppliers"].plot(kind="barh", ax=axes[0], color="steelblue")
            axes[0].set_title("Top 10 Suppliers by Spend")
        else:
            axes[0].set_visible(False)

        # Spend by category
        if not r["spend_by_category"].empty:
            r["spend_by_category"].plot(kind="barh", ax=axes[1], color="coral")
            axes[1].set_title("Spend by Category")
        else:
            axes[1].set_visible(False)

        # YoY trend
        if not r["yoy_trend"].empty:
            r["yoy_trend"].plot(kind="line", ax=axes[2], marker="o", color="green", linewidth=2)
            axes[2].set_title("Year-over-Year Spend Trend")
            axes[2].grid(True, alpha=0.3)
        else:
            axes[2].set_visible(False)

        plt.tight_layout()
        plt.savefig(output_path, dpi=300, bbox_inches="tight")
        plt.close(fig)
        logger.info(f"Visualization saved to {output_path}")

    def export(self, output_path: str = "cleaned_data.xlsx") -> None:
        """Export cleaned data to Excel or CSV."""
        if self.df.empty:
            die("No data to export (empty dataframe).")

        out = Path(output_path)
        suffix = out.suffix.lower()

        if suffix == ".csv":
            try:
                self.df.to_csv(out, index=False)
                logger.info(f"Cleaned data exported to {out}")
                return
            except Exception as e:
                die(f"CSV export failed: {e}")

        if suffix == ".xlsx":
            try:
                self.df.to_excel(out, index=False, engine="openpyxl")
                logger.info(f"Cleaned data exported to {out}")
                return
            except Exception as e:
                fallback = out.with_suffix(".csv")
                try:
                    self.df.to_csv(fallback, index=False)
                    logger.warning(f"Could not write Excel ({e}). Saved CSV fallback to {fallback}")
                    return
                except Exception as e2:
                    die(f"Export failed: {e2}")

        die(f"Unsupported export format: '{suffix}'. Use .xlsx or .csv.")

def main() -> None:
    parser = argparse.ArgumentParser(description="Procurement Spend Analyzer")
    parser.add_argument(
        "file",
        nargs="?",
        default=None,
        help="Path to input data file (CSV/XLS/XLSX). If not provided, sample_data.* will be used.",
    )
    parser.add_argument("--sep", "-s", default=None, help="CSV separator (e.g. ',' or ';').")
    parser.add_argument(
        "--export-format",
        "-e",
        choices=["xlsx", "csv"],
        default="xlsx",
        help="Export format for cleaned data. Fallback to CSV if Excel writer missing.",
    )

    args = parser.parse_args()

    file_path = args.file or choose_file()

    analyzer = SpendAnalyzer(file_path)
    analyzer.load_data(csv_sep=args.sep)
    analyzer.clean_data()
    analyzer.analyze()
    analyzer.report()
    analyzer.visualize()
    analyzer.export(output_path=f"cleaned_data.{args.export_format}")

if __name__ == "__main__":
    main()
