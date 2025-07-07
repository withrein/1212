# XML to XLSX Converter API

A Flask API service that converts XML data to Excel (XLSX) format with automatic pivot table generation.

## Features

- 🔄 Convert XML to XLSX format
- 📊 Automatic pivot table creation for time-series data
- 🌐 CORS-enabled for web applications
- 📁 Multiple input formats supported (JSON, form data, file upload, raw XML)
- 🔍 Metadata sheet with conversion statistics
- 💾 Base64 encoded Excel files in JSON response

## API Endpoints

### `POST /api/convert`

Converts XML data to XLSX format.

**Input formats:**

- JSON: `{"xml_content": "your-xml-here"}`
- Form data: `xml_content` or `xml` field
- File upload: Upload XML file with key `xml`
- Raw XML: Send XML directly in request body

**Response:**

```json
{
  "success": true,
  "message": "Successfully parsed X records",
  "processing_notes": "Pivoted data: X categories across Y periods",
  "records_count": 123,
  "excel_file": "base64-encoded-excel-data",
  "filename": "converted_data.xlsx"
}
```

### `GET /api/health`

Health check endpoint.

### `GET /`

API documentation and service information.

## Deployment on Render

### Quick Deploy

1. Fork this repository
2. Connect your GitHub account to Render
3. Create a new Web Service
4. Select this repository
5. Configure the following settings:

**Build Settings:**

- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn api.convert:app`

**Environment:**

- Runtime: Python 3.11+

### Manual Deployment Steps

1. **Create a Render account** at [render.com](https://render.com)

2. **Connect your repository:**

   - Click "New +" → "Web Service"
   - Connect your GitHub repository
   - Select this repository

3. **Configure the service:**

   - **Name:** `xml-to-xlsx-converter`
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn api.convert:app`
   - **Instance Type:** Free (or paid for production)

4. **Deploy:**

   - Click "Create Web Service"
   - Wait for deployment to complete

5. **Test your API:**
   - Your API will be available at: `https://your-service-name.onrender.com`
   - Test with: `GET https://your-service-name.onrender.com/api/health`

## Local Development

1. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

2. **Run the application:**

   ```bash
   python api/convert.py
   ```

3. **Test locally:**
   ```bash
   curl http://localhost:5000/api/health
   ```

## Usage Examples

### Using cURL

```bash
# JSON request
curl -X POST https://your-app.onrender.com/api/convert \
  -H "Content-Type: application/json" \
  -d '{"xml_content": "<your-xml-data>"}'

# Form data
curl -X POST https://your-app.onrender.com/api/convert \
  -F "xml_content=<your-xml-data>"

# File upload
curl -X POST https://your-app.onrender.com/api/convert \
  -F "xml=@your-file.xml"
```

### Using Python

```python
import requests
import base64

# Send XML data
response = requests.post(
    'https://your-app.onrender.com/api/convert',
    json={'xml_content': your_xml_data}
)

if response.json()['success']:
    # Decode the Excel file
    excel_data = base64.b64decode(response.json()['excel_file'])

    # Save to file
    with open('converted_data.xlsx', 'wb') as f:
        f.write(excel_data)
```

## Dependencies

- Flask 2.3.3
- Flask-CORS 4.0.0
- pandas 2.1.1
- openpyxl 3.1.2
- gunicorn 21.2.0

## Supported XML Structure

The API expects XML with the following structure:

- `DataList` element containing data records
- `TN_DT` elements as individual data records
- Automatic namespace detection
- Numeric field detection for `DTVAL_CO`, `Period`, `CODE`, etc.

## Features

- **Automatic Pivoting:** Creates pivot tables for time-series data
- **Multiple Sheets:** Data, Metadata, and Original Data (if pivoted)
- **Error Handling:** Comprehensive error messages and status codes
- **CORS Support:** Ready for web application integration
