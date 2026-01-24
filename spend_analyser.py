"""
Procurement Spend Analyzer (OOP Version)
Automated tool for analyzing supplier spend data and generating insights
Author: Jean Kocman
"""

import pandas as pd
import matplotlib.pyplot as plt
import sys
from pathlib import Path

COLUMN_MAP = {
    "supplier": ["Supplier", "Vendor", "Supplier Name", "Nominated Supplier", "Counterpart"],
    "spend": ["Spend", "Amount", "Value", "Cost"],
    "category": ["Category", "Group", "Type"],
    "year": ["Year", "Fiscal Year", "Period"]
}

def find_column(df, possible_names):
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
        """Load supplier spend data from CSV / Excel file"""
        try:
            suffix = Path(self.file_path).suffix.lower()
            if suffix == ".csv":
                self.df = pd.read_csv(self.file_path)
            elif suffix in [".xls", ".xlsx"]:
                self.df = pd.read_excel(self.file_path)
            else:
                raise ValueError(f"Unsupported file format: {suffix}")
            print(f"✓ Loaded {len(self.df)} records from {self.file_path}")
        except Exception as e:
            print(f"✗ Error loading data: {e}")
            sys.exit(1)

    def clean_data(self):
        """Clean and standardize data formats"""
        spend_col = find_column(self.df, COLUMN_MAP["spend"])
        # Convert spend column to numeric
        self.df[spend_col] = pd.to_numeric(self.df[spend_col], errors="coerce")
        # Drop rows with missing spend
        self.df = self.df.dropna(subset=[spend_col])
        print("✓ Data cleaned and standardized")

    def analyze(self):
        """Perform spend analysis calculations"""
        supplier_col = find_column(self.df, COLUMN_MAP["supplier"])
        spend_col = find_column(self.df, COLUMN_MAP["spend"])
        category_col = find_column(self.df, COLUMN_MAP["category"])
        year_col = find_column(self.df, COLUMN_MAP["year"])

        top_suppliers = self.df.groupby(supplier_col)[spend_col].sum().nlargest(10).sort_values(ascending=True)
        spend_by_category = self.df.groupby(category_col)[spend_col].sum().sort_values(ascending=True)
        yoy_trend = self.df.groupby(year_col)[spend_col].sum().sort_index()

        total_spend = self.df[spend_col].sum()
        avg_spend = self.df[spend_col].mean()
        supplier_count = self.df[supplier_col].nunique()

        self.results = {
            'top_suppliers': top_suppliers,
            'spend_by_category': spend_by_category,
            'yoy_trend': yoy_trend,
            'total_spend': total_spend,
            'avg_spend': avg_spend,
            'supplier_count': supplier_count
        }
        print("✓ Analysis complete")

    def report(self):
        """Generate text summary report"""
        r = self.results
        print("\n" + "="*60)
        print("PROCUREMENT SPEND ANALYSIS REPORT")
        print("="*60)
        print(f"\nTotal Spend: €{r['total_spend']:,.2f}")
        print(f"Average Transaction: €{r['avg_spend']:,.2f}")
        print(f"Number of Suppliers: {r['supplier_count']}")

        print(f"\nTop 3 Suppliers:")
        for supplier, spend in r['top_suppliers'].tail(3).items():
            print(f"  • {supplier}: €{spend:,.2f}")

        print(f"\nTop 3 Categories:")
        for category, spend in r['spend_by_category'].tail(3).items():
            print(f"  • {category}: €{spend:,.2f}")

        years = r['yoy_trend'].index
        if len(years) >= 2:
            latest_year = years[-1]
            previous_year = years[-2]
            growth = ((r['yoy_trend'][latest_year] / r['yoy_trend'][previous_year]) - 1) * 100
            print(f"\nYoY Growth ({previous_year} → {latest_year}): {growth:+.1f}%")

        print("="*60 + "\n")

    def visualize(self, output_path="spend_analysis.png"):
        """Generate visual charts from analysis results"""
        r = self.results
        fig, axes = plt.subplots(1, 3, figsize=(18, 5))
        fig.suptitle('Procurement Spend Analysis Dashboard', fontsize=16, fontweight='bold')

        r['top_suppliers'].plot(kind='barh', ax=axes[0], color='steelblue')
        axes[0].set_title('Top 10 Suppliers by Spend')
        axes[0].set_xlabel('Total Spend (€)')
        axes[0].set_ylabel('Supplier')

        r['spend_by_category'].plot(kind='barh', ax=axes[1], color='coral')
        axes[1].set_title('Spend by Category')
        axes[1].set_xlabel('Total Spend (€)')
        axes[1].set_ylabel('Category')

        r['yoy_trend'].plot(kind='line', ax=axes[2], marker='o', color='green', linewidth=2)
        axes[2].set_title('Year-over-Year Spend Trend')
        axes[2].set_xlabel('Year')
        axes[2].set_ylabel('Total Spend (€)')
        axes[2].grid(True, alpha=0.3)

        plt.tight_layout()
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        print(f"✓ Visualization saved to {output_path}")

    def export(self, output_path="cleaned_data.xlsx"):
        """Export cleaned data to Excel"""
        self.df.to_excel(output_path, index=False)
        print(f"✓ Cleaned data exported to {output_path}")


if __name__ == "__main__":
    analyzer = SpendAnalyzer("sample_data.xlsx")
    analyzer.load_data()
    analyzer.clean_data()
    analyzer.analyze()
    analyzer.report()
    analyzer.visualize()
    analyzer.export()