#!/usr/bin/env python3
"""
Convert XML files from downloaded_tables directory to Excel format
"""

import xml.etree.ElementTree as ET
import pandas as pd
import os
from pathlib import Path
import sys

def parse_xml_to_dataframe(xml_file_path):
    """Parse XML file and convert to pandas DataFrame"""
    try:
        # Parse the XML file
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        
        # Define namespace if present
        namespace = {'': 'http://schemas.datacontract.org/2004/07/E1212_ServiceAPI.Models'}
        
        # Find DataList element
        data_list = root.find('.//DataList', namespace)
        if data_list is None:
            # Try without namespace
            data_list = root.find('.//DataList')
        
        if data_list is None:
            return None, "No DataList found in XML"
        
        # Extract data from TN_DT elements
        records = []
        tn_dt_elements = data_list.findall('.//TN_DT', namespace)
        if not tn_dt_elements:
            # Try without namespace
            tn_dt_elements = data_list.findall('.//TN_DT')
        
        for tn_dt in tn_dt_elements:
            record = {}
            for child in tn_dt:
                # Remove namespace from tag name for cleaner column names
                tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                
                # Handle nil values
                if child.text is not None:
                    record[tag_name] = child.text
                else:
                    record[tag_name] = None
            
            records.append(record)
        
        if not records:
            return None, "No data records found"
        
        # Create DataFrame
        df = pd.DataFrame(records)
        
        # Convert numeric columns
        numeric_columns = ['DTVAL_CO', 'Period', 'CODE', 'CODE1', 'CODE2']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return df, f"Successfully parsed {len(records)} records"
        
    except Exception as e:
        return None, f"Error parsing XML: {str(e)}"

def create_pivot_table(df):
    """Create a pivot table if the data structure supports it"""
    try:
        # Check if we have the required columns for pivoting
        required_cols = ['Period', 'DTVAL_CO']
        if not all(col in df.columns for col in required_cols):
            return df, "No pivot - missing required columns"
        
        # Identify columns to use as row identifiers
        id_cols = []
        potential_id_cols = ['CODE', 'SCR_MN', 'SCR_ENG', 'SCR_MN1', 'SCR_ENG1']
        
        for col in potential_id_cols:
            if col in df.columns:
                id_cols.append(col)
        
        if not id_cols:
            return df, "No pivot - no identifier columns found"
        
        # Check if we have multiple periods to justify pivoting
        unique_periods = df['Period'].nunique()
        if unique_periods <= 1:
            return df, "No pivot - only one period found"
        
        # Create pivot table
        pivot_df = df.pivot_table(
            index=id_cols,
            columns='Period',
            values='DTVAL_CO',
            aggfunc='first'
        )
        
        # Reset index to make identifier columns regular columns
        pivot_df = pivot_df.reset_index()
        
        # Sort columns: put identifier columns first, then years in ascending order
        year_cols = [col for col in pivot_df.columns if col not in id_cols]
        year_cols = sorted(year_cols, key=lambda x: int(str(x)) if str(x).replace('-', '').isdigit() else float('inf'))
        
        # Reorder columns
        pivot_df = pivot_df[id_cols + year_cols]
        
        return pivot_df, f"Pivoted data: {len(pivot_df)} categories across {len(year_cols)} periods"
        
    except Exception as e:
        return df, f"Pivot failed: {str(e)}, using original format"

def convert_xml_to_excel(xml_file_path, excel_file_path):
    """Convert a single XML file to Excel format"""
    
    # Parse XML to DataFrame
    df, parse_message = parse_xml_to_dataframe(xml_file_path)
    
    if df is None:
        return False, parse_message
    
    # Try to create pivot table for better presentation
    pivot_df, pivot_message = create_pivot_table(df)
    
    try:
        # Save to Excel
        pivot_df.to_excel(excel_file_path, index=False, engine='openpyxl')
        
        # Add metadata sheet with original structure info
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
            # Create metadata
            metadata = pd.DataFrame({
                'Property': ['Source File', 'Total Records', 'Conversion Status', 'Processing Notes'],
                'Value': [xml_file_path.name, len(df), 'Success', pivot_message]
            })
            metadata.to_excel(writer, sheet_name='Metadata', index=False)
            
            # Also save original data structure if pivoted
            if len(pivot_df.columns) != len(df.columns):
                df.to_excel(writer, sheet_name='Original_Data', index=False)
        
        return True, f"{parse_message}. {pivot_message}"
        
    except Exception as e:
        return False, f"Error saving Excel: {str(e)}"

def main():
    """Main function to convert all XML files to Excel"""
    
    # Define paths
    xml_dir = Path("downloaded_tables")
    
    # Create output directory
    excel_dir = xml_dir / "excel_files"
    excel_dir.mkdir(exist_ok=True)
    
    # Find all XML files
    xml_files = list(xml_dir.glob("*.xml"))
    
    if not xml_files:
        print("No XML files found in downloaded_tables directory")
        return
    
    print(f"Found {len(xml_files)} XML files to convert")
    print(f"Output directory: {excel_dir}")
    print("-" * 50)
    
    successful_conversions = 0
    failed_conversions = 0
    
    for i, xml_file in enumerate(xml_files, 1):
        # Create Excel file path
        excel_file = excel_dir / f"{xml_file.stem}.xlsx"
        
        # Convert XML to Excel
        success, message = convert_xml_to_excel(xml_file, excel_file)
        
        if success:
            successful_conversions += 1
            print(f"[{i:2d}/{len(xml_files)}] ✓ {xml_file.name}")
            print(f"         → {excel_file.name}")
            print(f"         → {message}")
        else:
            failed_conversions += 1
            print(f"[{i:2d}/{len(xml_files)}] ✗ {xml_file.name}")
            print(f"         → Error: {message}")
        
        print()
    
    print("=" * 50)
    print(f"Conversion Summary:")
    print(f"✓ Successfully converted: {successful_conversions}")
    print(f"✗ Failed conversions: {failed_conversions}")
    print(f"📁 Excel files saved in: {excel_dir}")
    
    if successful_conversions > 0:
        print(f"\nExcel files include:")
        print(f"- Main data sheet (pivoted when possible)")
        print(f"- Metadata sheet with conversion info")
        print(f"- Original data sheet (when data was pivoted)")

if __name__ == "__main__":
    main() 