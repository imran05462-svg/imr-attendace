# Teacher Attendance Processor

A Python application that processes attendance logs and generates comprehensive attendance reports in Excel and PDF formats.

## Features

- **Multiple Input Formats**: Supports CSV and Excel file formats
- **Flexible Report Layouts**: Choose from three different report layouts:
  - **Compact Dashboard**: Stacked In/Out format, optimized for printing
  - **Single Page**: Entire month on one sheet
  - **Vertical Report**: Individual teacher mode for detailed views
- **Dual Output Formats**: Generate reports as Excel spreadsheets or PDF documents
- **User-Friendly GUI**: Built with Tkinter for easy interaction
- **Smart Data Parsing**: Automatically detects and maps various column naming conventions

## Requirements

- Python 3.7+
- pandas
- openpyxl
- fpdf2

## Installation

1. Clone or download the project
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Quick Setup

To quickly install all required dependencies, run:

```bash
pip install -r requirements.txt
```

This will install:
- pandas
- openpyxl
- fpdf2

## Usage

### Running the Application

```bash
python app.py
```

This launches the GUI dashboard where you can:

1. **Select an attendance log file** by clicking the "Browse" button
2. **Choose your desired report layout**:
   - Compact Dashboard (Stacked In/Out) - Best for prints
   - Single Page (Entire Month on One Sheet)
   - Vertical Report (Individual Teacher Mode)
3. **Generate Your Report**:
   - Click "Generate Excel" for Excel output
   - Click "Generate PDF" for PDF output

### Input File Format

Your attendance log file should contain the following columns:
- **DateTime** (or similar): Timestamp of attendance record
- **EnNo**, **Employee No**, **ID**, **Code**, or **No**: Employee/Teacher identifier
- **Name** (or similar): Employee/Teacher name
- **IOMd**, **Mode**, or **Direction**: In/Out indicator (optional)

## Project Structure

- `app.py` - Main GUI application
- `processor.py` - Core processing logic for attendance data
- `requirements.txt` - Python dependencies
- `log_input.csv` - Sample attendance log (CSV format)

## How It Works

1. The application reads the attendance log file (CSV or Excel)
2. Data is parsed and normalized to handle various column naming conventions
3. Attendance records are processed and aggregated by employee and date
4. Reports are generated in the selected format and layout
5. Output files are saved in the project directory

## Output Files

- Excel reports are saved as `.xlsx` files
- PDF reports are saved as `.pdf` files

Files are generated with timestamps to avoid overwriting previous reports.

## License

This project is provided as-is for educational and organizational use.
