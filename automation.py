import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
import sys

def compare_worksheets(file_path, sheet1, sheet2, output_path, key_column='ID'):
    """
    Compare two worksheets within the same Excel file and create a report of mismatches.
    
    Args:
        file_path: Path to Excel file containing both worksheets
        sheet1: Name of first worksheet
        sheet2: Name of second worksheet
        output_path: Path to save the comparison report
        key_column: Column name to use as unique identifier (default: 'ID')
    """
    
    try:
        # Read worksheets
        df1 = pd.read_excel(file_path, sheet_name=sheet1)
        df2 = pd.read_excel(file_path, sheet_name=sheet2)
        
        print(f"✓ Loaded '{sheet1}': {len(df1)} rows")
        print(f"✓ Loaded '{sheet2}': {len(df2)} rows")
        
        # Ensure key column exists
        if key_column not in df1.columns or key_column not in df2.columns:
            print(f"✗ Error: '{key_column}' column not found in one or both worksheets")
            return
        
        # Set key column as index
        df1_indexed = df1.set_index(key_column)
        df2_indexed = df2.set_index(key_column)
        
        # Initialize results list
        mismatches = []
        
        # Get all unique keys from both worksheets
        all_keys = set(df1_indexed.index) | set(df2_indexed.index)
        
        # Compare each record
        for key in sorted(all_keys):
            in_sheet1 = key in df1_indexed.index
            in_sheet2 = key in df2_indexed.index
            
            if in_sheet1 and in_sheet2:
                # Record exists in both sheets - check for differences
                row1 = df1_indexed.loc[key]
                row2 = df2_indexed.loc[key]
                
                # Get all columns from both sheets
                all_cols = set(df1.columns) | set(df2.columns)
                
                for col in all_cols:
                    if col == key_column:
                        continue
                    
                    val1 = row1.get(col) if in_sheet1 else 'N/A'
                    val2 = row2.get(col) if in_sheet2 else 'N/A'
                    
                    # Convert NaN to string for comparison
                    val1_str = str(val1) if pd.notna(val1) else 'MISSING'
                    val2_str = str(val2) if pd.notna(val2) else 'MISSING'
                    
                    if val1_str != val2_str:
                        mismatches.append({
                            'ID': key,
                            'Field': col,
                            f'{sheet1}_Value': val1_str,
                            f'{sheet2}_Value': val2_str,
                            'Status': f"MISMATCH - {sheet1}: {val1_str}, {sheet2}: {val2_str}"
                        })
            
            elif in_sheet1 and not in_sheet2:
                # Record only in sheet 1
                row1 = df1_indexed.loc[key]
                for col in df1.columns:
                    if col != key_column:
                        val1 = row1.get(col)
                        val1_str = str(val1) if pd.notna(val1) else 'MISSING'
                        mismatches.append({
                            'ID': key,
                            'Field': col,
                            f'{sheet1}_Value': val1_str,
                            f'{sheet2}_Value': 'RECORD_NOT_FOUND',
                            'Status': f'RECORD_ONLY_IN_{sheet1.upper()}'
                        })
            
            else:
                # Record only in sheet 2
                row2 = df2_indexed.loc[key]
                for col in df2.columns:
                    if col != key_column:
                        val2 = row2.get(col)
                        val2_str = str(val2) if pd.notna(val2) else 'MISSING'
                        mismatches.append({
                            'ID': key,
                            'Field': col,
                            f'{sheet1}_Value': 'RECORD_NOT_FOUND',
                            f'{sheet2}_Value': val2_str,
                            'Status': f'RECORD_ONLY_IN_{sheet2.upper()}'
                        })
        
        # Create results DataFrame
        results_df = pd.DataFrame(mismatches)
        
        if len(results_df) == 0:
            print("\n✓ No mismatches found! Both worksheets are identical.")
        else:
            print(f"\n✓ Found {len(results_df)} mismatches")
            print("\nMismatch Summary:")
            print(results_df.to_string(index=False))
        
        # Save to Excel with formatting
        results_df.to_excel(output_path, index=False, sheet_name='Mismatches')
        
        # Apply formatting to output file
        wb = openpyxl.load_workbook(output_path)
        ws = wb.active
        
        # Header formatting
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        
        # Color coding for status
        mismatch_fill = PatternFill(start_color='FFE699', end_color='FFE699', fill_type='solid')
        sheet1_only_fill = PatternFill(start_color='C6E0B4', end_color='C6E0B4', fill_type='solid')
        sheet2_only_fill = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            status_cell = row[4]  # Status column (5th column)
            status = status_cell.value
            
            if status:
                if 'MISMATCH' in status:
                    status_cell.fill = mismatch_fill
                elif sheet1.upper() in status:
                    status_cell.fill = sheet1_only_fill
                elif sheet2.upper() in status:
                    status_cell.fill = sheet2_only_fill
        
        # Adjust column widths
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 40
        
        wb.save(output_path)
        print(f"\n✓ Report saved to: {output_path}")
        
    except FileNotFoundError as e:
        print(f"✗ Error: File not found - {e}")
    except ValueError as e:
        print(f"✗ Error: Worksheet not found - {e}")
    except Exception as e:
        print(f"✗ Error: {e}")

def list_worksheets(file_path):
    """
    List all worksheets available in an Excel file.
    
    Args:
        file_path: Path to Excel file
    """
    try:
        xls = pd.ExcelFile(file_path)
        print(f"Available worksheets in '{file_path}':")
        for idx, sheet in enumerate(xls.sheet_names, 1):
            print(f"  {idx}. {sheet}")
        return xls.sheet_names
    except Exception as e:
        print(f"✗ Error reading file: {e}")
        return []

# Example usage
if __name__ == '__main__':
    # Modify these to your actual file and worksheet names
    file_path = 'politicians_data.xlsx'
    sheet1_name = 'Sheet1'
    sheet2_name = 'Sheet2'
    output = 'worksheet_comparison_report.xlsx'
    
    # Optional: List available worksheets first
    # list_worksheets(file_path)
    
    # Use 'ID' as the key column (adjust if your column name is different)
    compare_worksheets(file_path, sheet1_name, sheet2_name, output, key_column='ID')
    
    # Uncomment below to use command line arguments
    # if len(sys.argv) > 3:
    #     compare_worksheets(sys.argv[1], sys.argv[2], sys.argv[3], 
    #                        sys.argv[4] if len(sys.argv) > 4 else 'comparison_report.xlsx')
    # else:
    #     print("Usage: python script.py <file.xlsx> <sheet1_name> <sheet2_name> [output.xlsx]")