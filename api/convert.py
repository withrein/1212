import xml.etree.ElementTree as ET
import pandas as pd
import io
import base64
import json

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

def handler(request, response):
    """n8n handler function to convert XML to XLSX"""
    try:
        # Set CORS headers for browser requests
        response.headers['Access-Control-Allow-Origin'] = '*'
        response.headers['Access-Control-Allow-Methods'] = 'POST, GET, OPTIONS'
        response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        
        # Handle preflight OPTIONS request
        if request.method == 'OPTIONS':
            response.status_code = 200
            response.body = ''
            return
        
        # Only accept POST requests for conversion
        if request.method != 'POST':
            response.status_code = 405
            response.headers['Content-Type'] = 'application/json'
            response.body = json.dumps({'error': 'Method not allowed. Use POST.'})
            return
        
        # Get XML content from request
        xml_content = None
        
        # Try to get XML from different possible sources
        if hasattr(request, 'body') and request.body:
            # If body is already a string (XML content)
            if isinstance(request.body, str):
                xml_content = request.body
            elif isinstance(request.body, bytes):
                xml_content = request.body.decode('utf-8')
            else:
                # Try to parse as JSON in case XML is in a field
                try:
                    body_data = json.loads(request.body)
                    xml_content = body_data.get('xml_content') or body_data.get('xml')
                except:
                    xml_content = str(request.body)
        
        # Try to get from form data if available
        if not xml_content and hasattr(request, 'form'):
            xml_content = request.form.get('xml_content') or request.form.get('xml')
        
        # Try to get from query parameters
        if not xml_content and hasattr(request, 'args'):
            xml_content = request.args.get('xml_content') or request.args.get('xml')
        
        if not xml_content:
            response.status_code = 400
            response.headers['Content-Type'] = 'application/json'
            response.body = json.dumps({'error': 'No XML content provided. Send XML data in body, form, or as xml_content parameter.'})
            return
        
        # Parse XML to DataFrame
        df, parse_message = parse_xml_to_dataframe(xml_content)
        
        if df is None:
            response.status_code = 400
            response.headers['Content-Type'] = 'application/json'
            response.body = json.dumps({'error': parse_message})
            return
        
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
        response.status_code = 200
        response.headers['Content-Type'] = 'application/json'
        response.body = json.dumps({
            'success': True,
            'message': parse_message,
            'processing_notes': pivot_message,
            'records_count': len(df),
            'excel_file': excel_base64,
            'filename': 'converted_data.xlsx'
        })
        
    except Exception as e:
        response.status_code = 500
        response.headers['Content-Type'] = 'application/json'
        response.body = json.dumps({'error': f'Conversion failed: {str(e)}'})

# Health check function (optional)
def health_handler(request, response):
    """Health check handler for n8n"""
    response.status_code = 200
    response.headers['Content-Type'] = 'application/json'
    response.body = json.dumps({'status': 'healthy', 'message': 'XML to XLSX converter is running'})