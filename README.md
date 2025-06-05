# Statement Analyzer

A Python-based web application for analyzing bank statements and searching through transactions.

## Features

- Upload bank statements in PDF or CSV format
- Extract and parse transaction data
- Search through transactions by phrase
- Filter transactions by date range
- Modern, responsive web interface
- Drag-and-drop file upload

## Requirements

- Python 3.8 or higher
- Flask
- PyPDF2
- pandas
- Other dependencies listed in requirements.txt

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd statement-analyzer
```

2. Create a virtual environment and activate it:
```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

3. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Start the application:
```bash
python app.py
```

2. Open your web browser and navigate to:
```
http://localhost:5000
```

3. Upload your bank statement (PDF or CSV format) using the upload interface

4. Use the search interface to find specific transactions:
   - Enter a search phrase
   - Optionally set a date range
   - Click "Search" to find matching transactions

## Supported File Formats

- PDF bank statements
- CSV bank statements

## Notes

- The application currently supports basic date formats (DD/MM/YYYY)
- File size limit is set to 16MB
- Uploaded files are stored temporarily in the 'uploads' directory

## Contributing

Feel free to submit issues and enhancement requests! 