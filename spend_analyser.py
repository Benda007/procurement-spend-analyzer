"""
Procurement Spend Analyzer (OOP Version)
Automated tool for analyzing supplier spend data and generating insights
Author: Jean Kocman
"""

import pandas as pd
import matplotlib.pyplot as plt
import argparse
import sys
import glob
import os
from pathlib import Path
from datetime import datetime

COLUMN_MAP = {
    "supplier": ["supplier", "vendor", "supplier name", "nominated supplier", "counterpart"],
    "spend": ["spend", "amount", "value", "cost"],
    "category": ["category", "group", "type"],
    "year": ["year", "fiscal year", "period"]
}


def choose_file():
    """Let user choose a sample_data file by index."""
    files = []
    for ext in ["csv", "xls", "xlsx"]:
        files.extend(glob.glob(f"sample_data.{ext}"))

    if not files:
        print("✗ No sample_data file found (csv/xls/xlsx)")
        sys.exit(1)

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
            print("✗ Invalid choice")
            sys.exit(1)
    except ValueError:
        print("✗ Invalid input")
        sys.exit(1)


class SpendAnalyzer:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.results = None

    def load_data(self, csv_sep=None):
        """Load data from CSV or Excel (all sheets combined)."""
        try:
            suffix = Path(self.file_path).suffix.lower()

            if suffix == ".csv":
                # If your CSV uses semicolons, pass csv_sep=";" when calling load_data
                if csv_sep:
                    self.df = pd.read_csv(self.file_path, sep=csv_sep)
                else:
                    self.df = pd.read_csv(self.file_path)
            elif suffix in [".xls", ".xlsx"]:
                xls = pd.ExcelFile(self.file_path)
                frames = [pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names]
                if frames:
                    self.df = pd.concat(frames, ignore_index=True)
                else:
                    self.df = pd.DataFrame()
            else:
                raise ValueError(f"Unsupported file format: {suffix}")

            # Basic validations
            if not isinstance(self.df, pd.DataFrame):
                print("✗ Loaded object is not a DataFrame")
                sys.exit(1)

            if self.df.empty or len(self.df.columns) == 0:
                print("✗ The file is empty or has no columns")
                sys.exit(1)

            print(f"✓ Loaded {len(self.df)} records from {self.file_path}")

        except Exception as e:
            print(f"✗ Error loading data: {e}")
            sys.exit(1)

        # Standardize column names immediately after loading
        self.standardize_columns()

    def standardize_columns(self):
        """Rename columns to unified internal names (supplier, spend, category, year)."""
        if self.df is None:
            print("✗ No data loaded, cannot standardize columns")
            sys.exit(1)

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
            print(f"✗ Missing required columns: {missing_required}")
            sys.exit(1)

        # Optional columns
        optional = ["category", "year"]
        missing_optional = [col for col in optional if col not in self.df.columns]
        if missing_optional:
            print(f"⚠ Warning: Missing optional columns: {missing_optional}. Some analyses will be skipped.")

        print("✓ Column names standardized")

    def clean_data(self):
        """Clean and standardize data formats."""
        if self.df is None:
            print("✗ No data loaded, cannot clean")
            sys.exit(1)

        # Convert spend to numeric
        self.df["spend"] = pd.to_numeric(self.df["spend"], errors="coerce")

        # If year exists, try to coerce to integer year
        if "year" in self.df.columns:
            try:
                self.df["year"] = pd.to_datetime(self.df["year"], errors="coerce").dt.year
            except Exception:
                # fallback: try numeric conversion
                self.df["year"] = pd.to_numeric(self.df["year"], errors="coerce")

        # Clean text fields
        if "supplier" in self.df.columns:
            self.df["supplier"] = self.df["supplier"].astype(str).str.strip().str.title()
        if "category" in self.df.columns:
            self.df["category"] = self.df["category"].astype(str).str.strip().str.title()

        # Drop rows without spend value
        self.df = self.df.dropna(subset=["spend"])

        # Remove duplicates
        self.df = self.df.drop_duplicates()

        # Remove negative or zero spend if that is desired (keep zero if you want)
        self.df = self.df[self.df["spend"] >= 0]

        if self.df.empty:
            print("✗ No valid rows after cleaning")
            sys.exit(1)

        print("✓ Data cleaned and standardized")

    def analyze(self):
        """Perform spend analysis calculations."""
        if self.df is None:
            print("✗ No data loaded, cannot analyze")
            sys.exit(1)

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
        print("✓ Analysis complete")

    def report(self):
        """Generate text summary report."""
        if self.results is None:
            print("✗ No results to report")
            sys.exit(1)

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

    def visualize(self, output_path="spend_analysis.png"):
        """Generate visual charts from analysis results."""
        if self.results is None:
            print("✗ No results to visualize")
            sys.exit(1)

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
        print(f"✓ Visualization saved to {output_path}")

    def export(self, output_path="cleaned_data.xlsx"):
        """Export cleaned data to Excel, fallback to CSV if openpyxl missing."""
        if self.df is None:
            print("✗ No data to export")
            sys.exit(1)

        # Try Excel first
        try:
            # pandas will use openpyxl for .xlsx; if missing, this will raise
            self.df.to_excel(output_path, index=False)
            print(f"✓ Cleaned data exported to {output_path}")
        except Exception as e:
            # Fallback to CSV
            fallback = Path(output_path).with_suffix(".csv")
            try:
                self.df.to_csv(fallback, index=False)
                print(f"⚠ Could not write Excel ({e}). Saved CSV fallback to {fallback}")
            except Exception as e2:
                print(f"✗ Export failed: {e2}")
                sys.exit(1)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Procurement Spend Analyzer")
    parser.add_argument(
        "file", nargs="?", default=None,
        help="Path to input data file (CSV/XLS/XLSX). If not provided, sample_data.* will be used."
    )
    args = parser.parse_args()

    # If user provided a file, use it. Otherwise fall back to sample_data.*
    if args.file:
        file_path = args.file
    else:
        file_path = choose_file()

    analyzer = SpendAnalyzer(file_path)
    analyzer.load_data()
    analyzer.clean_data()
    analyzer.analyze()
    analyzer.report()
    analyzer.visualize()
    analyzer.export()