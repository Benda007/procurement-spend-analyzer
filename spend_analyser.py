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
from typing import Optional, Dict, Any, NoReturn, cast



logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

COLUMN_MAP = {
    "supplier": ["supplier", "vendor", "supplier name", "nominated supplier", "counterpart"],
    "spend": ["spend", "amount", "value", "cost"],
    "category": ["category", "group", "type"],
    "year": ["year", "fiscal year", "period"]
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
    files = []
    for ext in ["csv", "xls", "xlsx"]:
        files.extend(glob.glob(f"sample_data*.{ext}"))

    if not files:
        die("No sample_data file found (csv/xls/xlsx)")


    print("\nAvailable files:")
    for i, f in enumerate(files, start=1):
        stat = os.stat(f)
        created = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        print(f"{i}. {Path(f).name} (created: {created})")

    choice = input("\nSelect a file by number: ")
    try:
        choice = int(choice)
        if 1 <= choice <= len(files):
            return files[choice - 1]
        else:
            die("Invalid choice")
    except ValueError:
        die("Invalid input")

def detect_sep(path):
    with open(path, newline='', encoding='utf-8') as f:
        sample = f.read(2048)
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;|\t")
        return dialect.delimiter
    except Exception:
        return ","


class SpendAnalyzer:
    def __init__(self, file_path: str | Path) -> None:
        """Initialize spend analyser with path to input file."""
        self.file_path: Path = Path(file_path)
        self.df: pd.DataFrame = pd.DataFrame()
        self.results: Optional[Dict[str, Any]] = None
        
    def load_data(self, csv_sep: Optional[str] = None) -> None:
        """Load data from CSV or Excel (all sheets combined)."""
        try:
            suffix = Path(self.file_path).suffix.lower()

            if suffix == ".csv":
                if csv_sep:
                    sep = csv_sep
                else:
                    sep = detect_sep(self.file_path)
                self.df = pd.read_csv(self.file_path, sep=sep)
            elif suffix in [".xls", ".xlsx"]:
                xls = pd.ExcelFile(self.file_path)
                frames = [pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names]
                if frames:
                    self.df = pd.concat(frames, ignore_index=True)
                else:
                    self.df = pd.DataFrame()
            else:
                raise ValueError(f"Unsupported file format: {suffix}")

            if not isinstance(self.df, pd.DataFrame):
                die("Loaded object is not a DataFrame")
                
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
        # Normalize column names: strip and lowercase
        self.df.columns = self.df.columns.str.strip().str.lower()

        # Build rename map using lowercase names from COLUMN_MAP
        rename_map = {}
        for unified_name, possible_names in COLUMN_MAP.items():
            for name in possible_names:
                name_l = name.lower()
                if name_l in self.df.columns:
                    rename_map[name_l] = unified_name
                    break

        # Apply renaming (only keys present)
        if rename_map:
            self.df = self.df.rename(columns=rename_map)

        # Required columns
        required = ["supplier", "spend"]
        missing_required = [col for col in required if col not in self.df.columns]
        if missing_required:
            die(f"Missing required columns: {missing_required}")

        # Optional columns
        optional = ["category", "year"]
        missing_optional = [col for col in optional if col not in self.df.columns]
        if missing_optional:
            logger.warning(f"Missing optional columns: {missing_optional}. Some analyses will be skipped.")
        logger.info("Column names standardized")

    def clean_data(self) -> None:
        """Clean and standardize data formats for analysis.

        - Converts 'spend' column to numeric.
        - Normalizes supplier and category text fields.
        - Drops rows with missing or invalid spend.
        - Removes duplicates and negative spend values.
        - Enforces Year format: YYYY or ISO YYYY-MM-DD (optional column).
        """
        if self.df.empty or "spend" not in self.df.columns:
            die("No data loaded or missing 'spend' column. Run load_data() first.")

        df = self.df  # local alias

        # spend -> numeric
        # TBD - Oprava: před to_numeric() „očistit“ měny, mezery a sjednotit desetinný oddělovač.
        df["spend"] = pd.to_numeric(df["spend"], errors="coerce")



        # year -> supports only YYYY or ISO YYYY-MM-DD (optional)
        if "year" in df.columns:
            logger.info(f"year dtype before parse: {df['year'].dtype}")

            y = cast(pd.Series, df["year"])

            # numeric years like 2026
            if pd.api.types.is_numeric_dtype(y):
                df["year"] = pd.to_numeric(y, errors="coerce").astype("Int64")
            else:
                ys = cast(pd.Series, y.astype("string").str.strip())

                # allow empty values -> keep as NA
                na_like = ys.str.lower().isin(["", "nan", "none", "nat"])
                ys = ys.mask(na_like, "")

                # (a) pure year YYYY
                is_year = ys.str.fullmatch(r"\d{4}")
                out = pd.Series(pd.NA, index=df.index, dtype="Int64")

                nums = cast(pd.Series, pd.to_numeric(ys[is_year], errors="coerce"))
                out.loc[is_year] = nums.astype("Int64")

                # (b) remaining non-empty must be ISO date YYYY-MM-DD
                rem = out.isna() & ys.ne("")
                if rem.any():
                    try:
                        dt = pd.to_datetime(ys[rem], format="%Y-%m-%d", errors="raise")
                        out.loc[rem] = pd.Series(pd.DatetimeIndex(dt).year, index=ys[rem].index).astype("Int64")
                    except Exception:
                        examples = ys[rem].head(5).tolist()
                        die(f"{YEAR_ALLOWED_HELP} Found: {examples}")

                    df["year"] = out

                # sanity check range
                bad = df["year"].notna() & ((df["year"] < 2000) | (df["year"] > 2100))
                if bad.any():
                    die(f"{YEAR_ALLOWED_HELP} Suspicious range 2000–2100, e.g.: {df.loc[bad, 'year'].head(5).tolist()}")

                df["year"] = out

            # sanity check range
            bad = df["year"].notna() & ((df["year"] < 2000) | (df["year"] > 2100))
            if bad.any():
                die(f"{YEAR_ALLOWED_HELP} Suspicious range 2000–2100, e.g.: {df.loc[bad, 'year'].head(5).tolist()}")

        # Clean text fields
        if "supplier" in df.columns:
            df["supplier"] = df["supplier"].astype(str).str.strip().str.title()
        if "category" in df.columns:
            df["category"] = df["category"].astype(str).str.strip().str.title()

        # Drop rows without spend value
        df = df.dropna(subset=["spend"])

        # Remove duplicates
        df = df.drop_duplicates()

        # Remove negative spend
        df = cast(pd.DataFrame, df.loc[df["spend"] >= 0, :].copy())

        if df.empty:
            die("No valid rows after cleaning")

        self.df = df
        logger.info("Data cleaned and standardized")
    
    
    def analyze(self) -> Dict[str, Any]:
        """Perform spend analysis and store results dictionary."""
        if self.df.empty or "supplier" not in self.df.columns or "spend" not in self.df.columns:
            die("No data to analyze. Run load_data() and clean_data() first.")
        top_suppliers = self.df.groupby("supplier")["spend"].sum().nlargest(10).sort_values(ascending=True)

        if "category" in self.df.columns:
            spend_by_category = self.df.groupby("category")["spend"].sum().sort_values(ascending=True)
        else:
            spend_by_category = pd.Series(dtype=float)

        if "year" in self.df.columns:
            yoy_trend = self.df.groupby("year")["spend"].sum().sort_index()
        else:
            yoy_trend = pd.Series(dtype=float)

        self.results = {
            "top_suppliers": top_suppliers,
            "spend_by_category": spend_by_category,
            "yoy_trend": yoy_trend,
            "total_spend": float(self.df["spend"].sum()),
            "avg_spend": float(self.df["spend"].mean()),
            "supplier_count": int(self.df["supplier"].nunique())
        }
        # print("✓ Analysis complete")
        logger.info("Analysis complete")
        return self.results

    def report(self) -> None:
        """Print a readable text summary report."""
        if self.results is None:
            die("No results to report")

        r = self.results

        print("\n" + "=" * 60)
        print("PROCUREMENT SPEND ANALYSIS REPORT")
        print("=" * 60)
        print(f"\nTotal Spend: €{r['total_spend']:,.2f}")
        print(f"Average Transaction: €{r['avg_spend']:,.2f}")
        print(f"Number of Suppliers: {r['supplier_count']}")

        print("\nTop 3 Suppliers:")
        for supplier, spend in r["top_suppliers"].tail(3).items():
            print(f"  • {supplier}: €{spend:,.2f}")

        if not r["spend_by_category"].empty:
            print("\nTop 3 Categories:")
            for category, spend in r["spend_by_category"].tail(3).items():
                print(f"  • {category}: €{spend:,.2f}")

        if not r["yoy_trend"].empty and len(r["yoy_trend"].index) >= 2:
            latest = r["yoy_trend"].index[-1]
            previous = r["yoy_trend"].index[-2]
            growth = ((r["yoy_trend"][latest] / r["yoy_trend"][previous]) - 1) * 100
            print(f"\nYoY Growth ({previous} → {latest}): {growth:+.1f}%")

        print("=" * 60 + "\n")

    def visualize(self, output_path: str ="spend_analysis.png") -> None:
        """Generate and save visual charts from analysis results.
        Panels: 
        - Top suppliers by spend. 
        - Spend by category (if available).
        - Year-over-year spend trend (if available)."""
        if self.results is None:
            die("No results to visualize")

        r = self.results

        # Prepare subplots; if some series empty, leave blank axes
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
        """Export cleaned data to Excel or CSV.

        - If output_path ends with .csv -> write CSV
        - If output_path ends with .xlsx -> write Excel (fallback to CSV if Excel writer missing)
        """
        if self.df.empty:
            die("No data to export (empty dataframe).")

        out = Path(output_path)
        suffix = out.suffix.lower()

        # 1) CSV requested -> write CSV directly
        if suffix == ".csv":
            try:
                self.df.to_csv(out, index=False)
                logger.info(f"Cleaned data exported to {out}")
                return
            except Exception as e:
                die(f"CSV export failed: {e}")

        # 2) Excel requested -> try Excel, fallback to CSV
        if suffix in (".xlsx", ".xls"):
            try:
                self.df.to_excel(out, index=False)
                logger.info(f"Cleaned data exported to {out}")
                return
            except Exception as e:
                fallback = out.with_suffix(".csv")
                try:
                    self.df.to_csv(fallback, index=False)
                    logger.warning(
                        f"Could not write Excel ({e}). Saved CSV fallback to {fallback}"
                    )
                    return
                except Exception as e2:
                    die(f"Export failed: {e2}")

        # 3) Unknown extension -> fail fast with clear message
        die(f"Unsupported export format: '{suffix}'. Use .xlsx or .csv.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Procurement Spend Analyzer")
    parser.add_argument(
        "file", nargs="?", default=None,
        help="Path to input data file (CSV/XLS/XLSX). If not provided, sample_data.* will be used."
    )
    parser.add_argument("--sep", "-s", default=None, help="CSV separator (e.g. ',' or ';').")
    parser.add_argument(
        "--export-format", "-e", choices=["xlsx", "csv"], default="xlsx",
        help="Export format for cleaned data. Fallback to CSV if Excel writer missing."
    )

    args = parser.parse_args()

    # Determine input file path
    if args.file:
        file_path = args.file
    else:
        file_path = choose_file()

    # Create analyzer and run pipeline, pass CSV separator if provided
    analyzer = SpendAnalyzer(file_path)
    analyzer.load_data(csv_sep=args.sep)
    analyzer.clean_data()
    analyzer.analyze()
    analyzer.report()
    analyzer.visualize()
    analyzer.export(output_path=f"cleaned_data.{args.export_format}")