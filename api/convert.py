from flask import Flask, request, jsonify
from flask_cors import CORS
import xml.etree.ElementTree as ET
import pandas as pd
import io
import base64
import json

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

def parse_xml_to_dataframe(xml_content):
    """Parse XML content and convert to pandas DataFrame"""
    try:
        # Parse the XML content
        root = ET.fromstring(xml_content)
        
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

@app.route('/api/convert', methods=['POST'])
def convert_xml_to_xlsx():
    """API endpoint to convert XML to XLSX"""
    try:
        # Get XML content from request
        xml_content = None
        
        # Try to get XML from different possible sources
        if request.is_json:
            # JSON request
            data = request.get_json()
            xml_content = data.get('xml_content') or data.get('xml')
        elif request.files and 'xml' in request.files:
            # File upload
            xml_file = request.files['xml']
            xml_content = xml_file.read().decode('utf-8')
        elif request.form:
            # Form data
            xml_content = request.form.get('xml_content') or request.form.get('xml')
        elif request.data:
            # Raw data (assume it's XML)
            xml_content = request.data.decode('utf-8')
        
        if not xml_content:
            return jsonify({'error': 'No XML content provided. Send XML data in JSON body, form data, or as file upload.'}), 400
        
        # Parse XML to DataFrame
        df, parse_message = parse_xml_to_dataframe(xml_content)
        
        if df is None:
            return jsonify({'error': parse_message}), 400
        
        # Try to create pivot table for better presentation
        pivot_df, pivot_message = create_pivot_table(df)
        
        # Create Excel file in memory
        excel_buffer = io.BytesIO()
        
        # Save to Excel
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            pivot_df.to_excel(writer, sheet_name='Data', index=False)
            
            # Create metadata
            metadata = pd.DataFrame({
                'Property': ['Total Records', 'Conversion Status', 'Processing Notes'],
                'Value': [len(df), 'Success', pivot_message]
            })
            metadata.to_excel(writer, sheet_name='Metadata', index=False)
            
            # Also save original data structure if pivoted
            if len(pivot_df.columns) != len(df.columns):
                df.to_excel(writer, sheet_name='Original_Data', index=False)
        
        excel_buffer.seek(0)
        excel_data = excel_buffer.getvalue()
        
        # Encode Excel file as base64 for JSON response
        excel_base64 = base64.b64encode(excel_data).decode('utf-8')
        
        # Return success response with Excel file
        return jsonify({
            'success': True,
            'message': parse_message,
            'processing_notes': pivot_message,
            'records_count': len(df),
            'excel_file': excel_base64,
            'filename': 'converted_data.xlsx'
        })
        
    except Exception as e:
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500

@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({'status': 'healthy', 'message': 'XML to XLSX converter API is running'})

@app.route('/', methods=['GET'])
def home():
    """Home endpoint with API documentation"""
    return jsonify({
        'service': 'XML to XLSX Converter API',
        'version': '1.0.0',
        'endpoints': {
            'POST /api/convert': 'Convert XML to XLSX format',
            'GET /api/health': 'Health check endpoint',
            'GET /': 'This documentation'
        },
        'usage': {
            'endpoint': '/api/convert',
            'method': 'POST',
            'content_types': [
                'application/json with xml_content field',
                'multipart/form-data with xml file or xml_content field',
                'text/xml or application/xml (raw XML data)'
            ],
            'response': 'JSON with base64 encoded Excel file'
        }
    })

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)