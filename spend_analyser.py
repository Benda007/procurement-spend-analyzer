"""
Procurement Spend Analyzer
Automated tool for analyzing supplier spend data and generating insights
Author: Jean Kocman
"""

import pandas as pd
import matplotlib.pyplot as plt
import tabulate
import sys
from pathlib import Path

def load_data(file_path):
    """Load supplier spend data from CSV file"""
    try:
        df = pd.read_csv(file_path)
        print(f"✓ Loaded {len(df)} records from {file_path}")
        return df
    except FileNotFoundError:
        print(f"✗ Error: File {file_path} not found")
        sys.exit(1)
    except Exception as e:
        print(f"✗ Error loading data: {e}")
        sys.exit(1)

def analyze_spend(df):
    """Perform spend analysis calculations"""
    
    # Top 10 suppliers by total spend
    top_suppliers = df.groupby('Supplier')['Spend'].sum().nlargest(10).sort_values(ascending=True)
    
    # Spend by category
    spend_by_category = df.groupby('Category')['Spend'].sum().sort_values(ascending=True)
    
    # Year-over-year trend
    yoy_trend = df.groupby('Year')['Spend'].sum().sort_index()
    
    # Summary statistics
    total_spend = df['Spend'].sum()
    avg_spend = df['Spend'].mean()
    supplier_count = df['Supplier'].nunique()
    
    return {
        'top_suppliers': top_suppliers,
        'spend_by_category': spend_by_category,
        'yoy_trend': yoy_trend,
        'total_spend': total_spend,
        'avg_spend': avg_spend,
        'supplier_count': supplier_count
    }

def generate_visualizations(analysis_results, output_path='spend_analysis.png'):
    """Generate visual charts from analysis results"""
    
    fig, axes = plt.subplots(1, 3, figsize=(18, 5))
    fig.suptitle('Procurement Spend Analysis Dashboard', fontsize=16, fontweight='bold')
    
    # Chart 1: Top 10 Suppliers
    analysis_results['top_suppliers'].plot(
        kind='barh',
        ax=axes[0],
        color='steelblue'
    )
    axes[0].set_title('Top 10 Suppliers by Spend')
    axes[0].set_xlabel('Total Spend (€)')
    axes[0].set_ylabel('Supplier')
    
    # Chart 2: Spend by Category
    analysis_results['spend_by_category'].plot(
        kind='barh',
        ax=axes[1],
        color='coral'
    )
    axes[1].set_title('Spend by Category')
    axes[1].set_xlabel('Total Spend (€)')
    axes[1].set_ylabel('Category')
    
    # Chart 3: Year-over-Year Trend
    analysis_results['yoy_trend'].plot(
        kind='line',
        ax=axes[2],
        marker='o',
        color='green',
        linewidth=2
    )
    axes[2].set_title('Year-over-Year Spend Trend')
    axes[2].set_xlabel('Year')
    axes[2].set_ylabel('Total Spend (€)')
    axes[2].grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    print(f"✓ Visualization saved to {output_path}")

def generate_report(analysis_results):
    """Generate text summary report"""
    
    print("\n" + "="*60)
    print("PROCUREMENT SPEND ANALYSIS REPORT")
    print("="*60)
    print(f"\nTotal Spend: €{analysis_results['total_spend']:,.2f}")
    print(f"Average Transaction: €{analysis_results['avg_spend']:,.2f}")
    print(f"Number of Suppliers: {analysis_results['supplier_count']}")
    
    print(f"\nTop 3 Suppliers:")
    for supplier, spend in analysis_results['top_suppliers'].tail(3).items():
        print(f"  • {supplier}: €{spend:,.2f}")
    
    print(f"\nTop 3 Categories:")
    for category, spend in analysis_results['spend_by_category'].tail(3).items():
        print(f"  • {category}: €{spend:,.2f}")
    
    # Calculate YoY growth
    years = analysis_results['yoy_trend'].index
    if len(years) >= 2:
        latest_year = years[-1]
        previous_year = years[-2]
        growth = ((analysis_results['yoy_trend'][latest_year] / 
                  analysis_results['yoy_trend'][previous_year]) - 1) * 100
        print(f"\nYoY Growth ({previous_year} → {latest_year}): {growth:+.1f}%")
    
    print("="*60 + "\n")

def main():
    """Main execution function"""
    
    print("\n🚀 Procurement Spend Analyzer")
    print("-" * 40)
    
    # Load data
    data_file = 'sample_data.csv'
    df = load_data(data_file)
    
    # Perform analysis
    print("\n📊 Analyzing spend data...")
    results = analyze_spend(df)
    
    # Generate report
    generate_report(results)
    
    # Generate visualizations
    print("📈 Generating visualizations...")
    generate_visualizations(results)
    
    print("✅ Analysis complete!\n")

if __name__ == "__main__":
    main()
