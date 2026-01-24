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
from pathlib import Path
from datetime import datetime

COLUMN_MAP = {
    "supplier": ["Supplier", "Vendor", "Supplier Name", "Nominated Supplier", "Counterpart"],
    "spend": ["Spend", "Amount", "Value", "Cost"],
    "category": ["Category", "Group", "Type"],
    "year": ["Year", "Fiscal Year", "Period"]
}

def find_sample_file():
    """Find a file named 'sample_data' with any supported extension."""
    for ext in ["csv", "xls", "xlsx"]:
        files = glob.glob(f"sample_data.{ext}")
        if files:
            return files[0]
    print("✗ No sample_data file found (csv/xls/xlsx)")
    sys.exit(1)


def find_column(df, possible_names):
    """Find the first matching column name from a list of possible names."""
    for name in possible_names:
        if name in df.columns:
            return name
    raise KeyError(f"None of the expected columns found: {possible_names}")

class SpendAnalyzer:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.results = None

    def load_data(self):
        """Load data from CSV or Excel (all sheets combined)."""
        try:
            suffix = Path(self.file_path).suffix.lower()

            if suffix == ".csv":
                self.df = pd.read_csv(self.file_path)

            elif suffix in [".xls", ".xlsx"]:
                xls = pd.ExcelFile(self.file_path)
                frames = []
                for sheet in xls.sheet_names:
                    frames.append(pd.read_excel(xls, sheet_name=sheet))
                self.df = pd.concat(frames, ignore_index=True)

            else:
                raise ValueError(f"Unsupported file format: {suffix}")

            print(f"✓ Loaded {len(self.df)} records from {self.file_path}")

        except Exception as e:
            print(f"✗ Error loading data: {e}")
            sys.exit(1)

        # Standardize column names immediately after loading
        self.standardize_columns()

    def standardize_columns(self):
        """Rename columns to unified internal names (supplier, spend, category, year)."""
        rename_map = {}

        for unified_name, possible_names in COLUMN_MAP.items():
            for name in possible_names:
                if name in self.df.columns:
                    rename_map[name] = unified_name
                    break

        self.df = self.df.rename(columns=rename_map)

        missing = [col for col in COLUMN_MAP.keys() if col not in self.df.columns]
        if missing:
            print(f"✗ Missing required columns after standardization: {missing}")
            sys.exit(1)

        print("✓ Column names standardized")

    def clean_data(self):
        """Clean and standardize data formats."""
        # Convert spend to numeric
        self.df["spend"] = pd.to_numeric(self.df["spend"], errors="coerce")

        # Drop rows without spend value
        self.df = self.df.dropna(subset=["spend"])

        print("✓ Data cleaned and standardized")

    def analyze(self):
        """Perform spend analysis calculations."""
        top_suppliers = self.df.groupby("supplier")["spend"].sum().nlargest(10).sort_values(ascending=True)
        spend_by_category = self.df.groupby("category")["spend"].sum().sort_values(ascending=True)
        yoy_trend = self.df.groupby("year")["spend"].sum().sort_index()

        self.results = {
            "top_suppliers": top_suppliers,
            "spend_by_category": spend_by_category,
            "yoy_trend": yoy_trend,
            "total_spend": self.df["spend"].sum(),
            "avg_spend": self.df["spend"].mean(),
            "supplier_count": self.df["supplier"].nunique()
        }

        print("✓ Analysis complete")

    def report(self):
        """Generate text summary report."""
        r = self.results

        print("\n" + "="*60)
        print("PROCUREMENT SPEND ANALYSIS REPORT")
        print("="*60)
        print(f"\nTotal Spend: €{r['total_spend']:,.2f}")
        print(f"Average Transaction: €{r['avg_spend']:,.2f}")
        print(f"Number of Suppliers: {r['supplier_count']}")

        print("\nTop 3 Suppliers:")
        for supplier, spend in r["top_suppliers"].tail(3).items():
            print(f"  • {supplier}: €{spend:,.2f}")

        print("\nTop 3 Categories:")
        for category, spend in r["spend_by_category"].tail(3).items():
            print(f"  • {category}: €{spend:,.2f}")

        years = r["yoy_trend"].index
        if len(years) >= 2:
            latest = years[-1]
            previous = years[-2]
            growth = ((r["yoy_trend"][latest] / r["yoy_trend"][previous]) - 1) * 100
            print(f"\nYoY Growth ({previous} → {latest}): {growth:+.1f}%")

        print("="*60 + "\n")

    def visualize(self, output_path="spend_analysis.png"):
        """Generate visual charts from analysis results."""
        r = self.results

        fig, axes = plt.subplots(1, 3, figsize=(18, 5))
        fig.suptitle("Procurement Spend Analysis Dashboard", fontsize=16, fontweight="bold")

        r["top_suppliers"].plot(kind="barh", ax=axes[0], color="steelblue")
        axes[0].set_title("Top 10 Suppliers by Spend")

        r["spend_by_category"].plot(kind="barh", ax=axes[1], color="coral")
        axes[1].set_title("Spend by Category")

        r["yoy_trend"].plot(kind="line", ax=axes[2], marker="o", color="green", linewidth=2)
        axes[2].set_title("Year-over-Year Spend Trend")
        axes[2].grid(True, alpha=0.3)

        plt.tight_layout()
        plt.savefig(output_path, dpi=300, bbox_inches="tight")
        print(f"✓ Visualization saved to {output_path}")

    def export(self, output_path="cleaned_data.xlsx"):
        """Export cleaned data to Excel."""
        self.df.to_excel(output_path, index=False)
        print(f"✓ Cleaned data exported to {output_path}")


if __name__ == "__main__":
    file_path = find_sample_file()
    analyzer = SpendAnalyzer(file_path)
    analyzer.load_data()
    analyzer.clean_data()
    analyzer.analyze()
    analyzer.report()
    analyzer.visualize()
    analyzer.export()
