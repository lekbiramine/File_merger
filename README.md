# File Merger

A Python tool that automatically merges CSV and Excel files from an input directory, cleans and normalizes the data, and generates a comprehensive master report.

## Features

- **Automatic File Processing**: Reads all CSV and Excel files from the input directory
- **Data Cleaning**: Normalizes headers, removes duplicates, and handles missing values
- **Report Generation**: Creates a master Excel report with cleaned data and department summaries
- **Email Notifications**: Automatically sends the generated report via email (optional)
- **Logging**: Comprehensive logging for tracking processing activities

## Requirements

- Python 3.x
- See `requirements.txt` for dependencies

## Installation

1. Clone or download this repository
2. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   ```

3. Activate the virtual environment:
   - Windows: `venv\Scripts\activate`
   - Linux/Mac: `source venv/bin/activate`

4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Setup

1. Create an `.env` file in the project root with the following variables:
   ```
   SENDER=your_email@gmail.com
   PASSWORD=your_app_password
   RECIEVER=recipient_email@gmail.com
   ```

   Note: For Gmail, you'll need to use an app-specific password instead of your regular password.

2. Place your CSV and Excel files in the `input/` directory

## Usage

Run the script:
```bash
python main.py
```

The script will:
1. Process all CSV and Excel files from the `input/` directory
2. Clean and merge the data
3. Generate a master report in the `output/` directory
4. Send the report via email (if configured)

## Expected Data Format

The tool expects the following columns in your files:
- `name` - Name field
- `department` - Department field
- `amount` - Numeric amount field
- `date` - Date field

## Project Structure

```
File_Merger/
├── input/              # Place your input files here
├── output/             # Generated reports are saved here
├── logs/               # Application logs
├── main.py             # Main script
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

## Output

The generated report (`master_report.xlsx`) contains two sheets:
- **Cleaned Data**: All merged and cleaned records sorted by date
- **Summary**: Total amounts grouped by department

## Logging

All processing activities are logged to `logs/automation.log` for troubleshooting and audit purposes.

## License

This project is open source and available for personal use.
