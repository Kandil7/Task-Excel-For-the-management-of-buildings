#!/usr/bin/env python3
"""
Legacy Excel Generator for Building Management (Pandas version)
Note: This is the legacy version. Consider using excel_generate_v2.py for better features.
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
import numpy as np


def generate_excel_pandas(output_path: Path) -> None:
    """
    Generate Excel file using pandas and xlsxwriter.

    Args:
        output_path: Path where Excel file will be saved
    """
    # Create an Excel writer object
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

    # --- Sheet 1: بيانات العقارات ---
    # Create a DataFrame for properties
    properties_data = {
        'اسم العمارة': ['العمارة 1', 'العمارة 1', 'العمارة 1', 'العمارة 1', 'العمارة 1',
                        'العمارة 2', 'العمارة 2', 'العمارة 2', 'العمارة 2', 'العمارة 2'],
        'رقم الوحدة': [1, 2, 3, 4, 5, 1, 2, 3, 4, 5],
        'نوع الوحدة': ['شقة', 'شقة', 'شقة', 'استوديو', 'استوديو',
                       'شقة', 'شقة', 'شقة', 'استوديو', 'استوديو'],
        'الإيجار الشهري': [5000, 5500, 6000, 3000, 3200, 6500, 7000, 7200, 3500, 3700],
        'الإيجار اليومي': [200, 220, 240, 150, 160, 250, 270, 280, 180, 190],
        'تكاليف الصيانة الشهرية': [300, 350, 400, 200, 220, 450, 500, 520, 250, 270]
    }
    properties_df = pd.DataFrame(properties_data)

    # Write to Excel
    properties_df.to_excel(writer, sheet_name='بيانات العقارات', index=False)

    # --- Sheet 2: بيانات المستأجرين ---
    # Create a DataFrame for tenants
    tenants_data = {
        'اسم المستأجر': ['أحمد علي', 'محمد سعيد', 'خالد عمر', 'سامي حسن', 'يوسف أحمد'],
        'رقم الوحدة': [1, 2, 3, 4, 5],
        'تاريخ بدء العقد': ['2024-01-01', '2024-02-01', '2024-03-01', '2024-04-01', '2024-05-01'],
        'قيمة الإيجار': [5000, 5500, 6000, 3000, 3200],
        'حالة الدفعات': ['مدفوع', 'مدفوع', 'لم يدفع', 'مدفوع', 'لم يدفع']
    }
    tenants_df = pd.DataFrame(tenants_data)

    # Write to Excel
    tenants_df.to_excel(writer, sheet_name='بيانات المستأجرين', index=False)

    # --- Sheet 3: تقارير الإيرادات ---
    # Create a DataFrame for revenues report
    monthly_revenue_building1 = sum(properties_df['الإيجار الشهري'][:5])
    monthly_revenue_building2 = sum(properties_df['الإيجار الشهري'][5:])
    weekly_revenue_building1 = sum(properties_df['الإيجار اليومي'][:5]) * 7
    weekly_revenue_building2 = sum(properties_df['الإيجار اليومي'][5:]) * 7

    revenues_data = {
        'اسم العمارة': ['العمارة 1', 'العمارة 1', 'العمارة 2', 'العمارة 2'],
        'الإيرادات الشهرية': [monthly_revenue_building1, 0, monthly_revenue_building2, 0],
        'الإيرادات الأسبوعية': [weekly_revenue_building1, 0, weekly_revenue_building2, 0],
        'التاريخ': ['2024-07', '2024-08', '2024-07', '2024-08']
    }
    revenues_df = pd.DataFrame(revenues_data)

    # Write to Excel
    revenues_df.to_excel(writer, sheet_name='تقارير الإيرادات', index=False)

    # --- Sheet 4: المصروفات ---
    # Create a DataFrame for expenses
    monthly_expenses_building1 = sum(properties_df['تكاليف الصيانة الشهرية'][:5])
    monthly_expenses_building2 = sum(properties_df['تكاليف الصيانة الشهرية'][5:])
    weekly_expenses_building1 = monthly_expenses_building1 / 4
    weekly_expenses_building2 = monthly_expenses_building2 / 4

    expenses_data = {
        'اسم العمارة': ['العمارة 1', 'العمارة 1', 'العمارة 2', 'العمارة 2'],
        'المصروفات الشهرية': [monthly_expenses_building1, 0, monthly_expenses_building2, 0],
        'المصروفات الأسبوعية': [weekly_expenses_building1, 0, weekly_expenses_building2, 0],
        'التاريخ': ['2024-07', '2024-08', '2024-07', '2024-08']
    }
    expenses_df = pd.DataFrame(expenses_data)

    # Write to Excel
    expenses_df.to_excel(writer, sheet_name='المصروفات', index=False)

    # --- Sheet 5: إضافة بيانات جديدة ---
    # This sheet will have placeholders for adding new data
    new_data_df = pd.DataFrame({
        'نوع البيانات': ['عمارة جديدة', 'وحدة جديدة', 'مستأجر جديد'],
        'الوصف': ['إدخال بيانات عمارة جديدة', 'إدخال بيانات وحدة جديدة', 'إدخال بيانات مستأجر جديد']
    })

    new_data_df.to_excel(writer, sheet_name='إضافة بيانات جديدة', index=False)

    # Save the workbook (using close() instead of deprecated _save())
    writer.close()
    print(f"✓ Excel file generated successfully: {output_path}")


def main() -> int:
    """
    Main entry point for the legacy script.

    Returns:
        Exit code (0 for success, non-zero for errors)
    """
    parser = argparse.ArgumentParser(
        description='Generate Excel spreadsheet for building management (Legacy/Pandas version)',
        epilog='Note: Consider using excel_generate_v2.py for more features and better configuration.'
    )

    parser.add_argument(
        '-o', '--output',
        type=Path,
        default=Path('نموذج إدارة العمارات والشقق.xlsx'),
        help='Output Excel file path (default: نموذج إدارة العمارات والشقق.xlsx)'
    )

    args = parser.parse_args()

    try:
        generate_excel_pandas(args.output)
        return 0
    except PermissionError:
        print(
            f"Error: Unable to write to {args.output}. "
            f"Please close the file if it's open and try again.",
            file=sys.stderr
        )
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
