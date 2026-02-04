Excel Data Validation & Cleaning Tool
Automated Python script that validates and cleans product data across Excel worksheets, ensuring data quality for analytics pipelines.

Features
Dynamic Data Detection: Finds actual used ranges in Excel sheets (ignores empty rows/columns)

Multi-Sheet Validation: Compares "Sheet1" (reference data) against "Data" sheet using green-highlighted columns

Smart Cleaning: Removes quotes, standardizes commas, trims special characters, fixes numeric formats

Range Validation: Checks numeric values with extensions (e.g., "10 KG", "5 L") against allowed min/max ranges

Error Reporting: Flags duplicates, invalid values, formulas, out-of-range numbers; adds comments per row

Summary Dashboard: Creates "Validation_Summary" sheet with error counts and percentages

Demo Results
text
Processed: 1,247 cells across 156 rows
Errors Found: 23 (1.84% error rate)
- Invalid values: 12
- Range violations: 5  
- Formatting issues: 6
Time saved: ~2 hours manual validation
Prerequisites
bash
pip install openpyxl
Quick Start
bash
# Place your Excel file in the same directory
python excel_validator.py

# Input: Outdoor-recreation_Scope-Rings-and-Adaptors_reverse_PDW_[by_Sarang-P]_1763041008_ce28de14.xlsx
# Output: validated_report_no_price_range.xlsx
How It Works
text
Sheet1 (Reference)    →    Green Columns    →    Data Sheet (Validation)
   |                       in Data Sheet           |
   ↓                                             ↓
Allowed Values +      ←   Dynamic Comparison   →  Error Comments + Summary
Range Info (KG,L,etc)                          (Per Row + Dashboard)
Input Excel Requirements:

Sheet1: Reference values (exact matches)

Data: Sheet to validate (green fill = columns to check)

Comments column (auto-created if missing)

Sample Error Comments
text
Row 5: Size: Added missing extension "KG", Price: Value "ABC" not allowed
Row 12: Weight: Numeric value 150 exceeds allowed range [10, 100]
Row 23: Color: Duplicated values in cell
Customization
python
# Edit these in the script:
input_file = 'your_file.xlsx'
output_file = 'validated_output.xlsx'

# Add new validation rules in run_validation_all()
# Extend numeric_extension_info for custom units
Project Metrics
Metric	Value
Lines of Code	180+
Libraries Used	openpyxl, re, collections
Validation Rules	12+
Error Types Tracked	8
Built for Data Analyst workflows - Deployable in ETL pipelines, FinTech data processing, e-commerce inventory validation.

By Sarang P |Data Analyst
