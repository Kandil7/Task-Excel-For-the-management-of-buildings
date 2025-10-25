# Excel لإدارة العمارات والشقق الخاصة
# Excel for Building and Apartment Management

A Python-based tool for generating Excel spreadsheets to manage buildings, apartments, tenants, rental payments, and expenses. Designed for property managers to track and organize their real estate portfolios efficiently.

## Features

- **Property Management**: Track buildings and individual units with detailed information
- **Tenant Database**: Maintain comprehensive tenant records including contact information and lease details
- **Financial Tracking**: Monitor rental income, payment status, and expenses
- **Automated Reporting**: Generate financial reports for revenues and expenses
- **Conditional Formatting**: Visual indicators for unpaid rents and other important data
- **Flexible Configuration**: Use JSON configuration files for easy data management
- **Multi-language Support**: Full Arabic language support with bilingual documentation

## Quick Start

### Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

### Installation

1. Clone the repository:
```bash
git clone https://github.com/Kandil7/Task-Excel-For-the-management-of-buildings.git
cd Task-Excel-For-the-management-of-buildings
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Create your configuration file:
```bash
cp config.example.json config.json
```

4. Edit `config.json` with your building and tenant data.

### Basic Usage

#### Using the Modern Version (Recommended)

Generate an Excel file using the default config:
```bash
python excel_generate_v2.py
```

Specify a custom config file:
```bash
python excel_generate_v2.py -c mydata.json
```

Specify output filename:
```bash
python excel_generate_v2.py -o my_report.xlsx
```

Use verbose mode for detailed output:
```bash
python excel_generate_v2.py -v
```

#### Using the Legacy Version

The legacy pandas-based version is still available:
```bash
python excel_generate.py
```

Specify output filename:
```bash
python excel_generate.py -o output.xlsx
```

## Configuration

The configuration file (`config.json`) should follow this structure:

```json
{
  "output_filename": "إدارة عمارات سكنية.xlsx",
  "buildings": [
    {
      "name": "عمارة أ",
      "units": 5,
      "notes": ""
    }
  ],
  "units": [
    {
      "unit_no": "101",
      "building": "عمارة أ",
      "type": "استديو",
      "rent": 2000,
      "status": "مُؤجّرة",
      "notes": ""
    }
  ],
  "tenants": [
    {
      "unit_no": "101",
      "name": "محمد أحمد",
      "id": "1234567890",
      "mobile": "05xxxxxxxx",
      "start_date": "2024-01-01",
      "end_date": "2024-12-31",
      "rent": 2000,
      "email": "mohamed@example.com",
      "notes": ""
    }
  ],
  "rents_paid": [
    {
      "unit_no": "101",
      "month": "يناير",
      "year": 2024,
      "amount": 2000,
      "date": "2024-01-05",
      "method": "تحويل بنكي",
      "status": "مدفوع",
      "notes": ""
    }
  ],
  "expenses": [
    {
      "building": "عمارة أ",
      "date": "2024-01-01",
      "type": "فاتورة كهرباء",
      "amount": 500,
      "category": "فواتير",
      "notes": ""
    }
  ]
}
```

See `config.example.json` for a complete example.

## Generated Excel Structure

The generated Excel file contains the following sheets:

1. **العمارات (Buildings)**: List of all buildings with unit counts
2. **الوحدات (Units)**: Detailed information about each unit
3. **المستأجرين (Tenants)**: Tenant database with contact and lease information
4. **الإيجارات (Rents)**: Rental payment tracking with conditional formatting for unpaid rents
5. **المصروفات (Expenses)**: Expense tracking by building and category

## Command-Line Options

### excel_generate_v2.py

```
usage: excel_generate_v2.py [-h] [-c CONFIG] [-o OUTPUT] [-v]

Generate Excel spreadsheet for building management

optional arguments:
  -h, --help            show help message and exit
  -c CONFIG, --config CONFIG
                        Path to configuration JSON file (default: config.json)
  -o OUTPUT, --output OUTPUT
                        Output Excel file path (default: from config or "output.xlsx")
  -v, --verbose         Enable verbose output
```

### excel_generate.py (Legacy)

```
usage: excel_generate.py [-h] [-o OUTPUT]

Generate Excel spreadsheet for building management (Legacy/Pandas version)

optional arguments:
  -h, --help            show help message and exit
  -o OUTPUT, --output OUTPUT
                        Output Excel file path (default: نموذج إدارة العمارات والشقق.xlsx)
```

## Project Structure

```
.
├── excel_generate_v2.py      # Modern version (recommended)
├── excel_generate.py          # Legacy pandas version
├── config.example.json        # Example configuration file
├── requirements.txt           # Python dependencies
├── readme.md                  # This file
├── LICENSE                    # MIT License
└── .gitignore                 # Git ignore rules
```

## Development

### Running Tests

```bash
python -m pytest tests/
```

### Code Quality

The code follows Python best practices:
- Type hints for better code clarity
- Comprehensive error handling
- Detailed docstrings
- PEP 8 style guidelines

## Differences Between Versions

| Feature | excel_generate.py | excel_generate_v2.py |
|---------|------------------|----------------------|
| Library | pandas + xlsxwriter | openpyxl |
| Configuration | Hardcoded | JSON file |
| Styling | Basic | Advanced |
| Error Handling | Basic | Comprehensive |
| Type Hints | Partial | Full |
| CLI Arguments | Limited | Full featured |
| **Status** | Legacy | **Recommended** |

## Troubleshooting

### Permission Error

If you get a permission error when generating the Excel file:
- Make sure the output file is not currently open in Excel or another program
- Check that you have write permissions in the output directory

### Configuration File Not Found

If you get a "config file not found" error:
- Make sure you've created a `config.json` file
- Use `config.example.json` as a template
- Or specify a different config file with `-c path/to/config.json`

### Invalid JSON Error

If you get a JSON parsing error:
- Validate your JSON file using a JSON validator
- Check for missing commas, brackets, or quotes
- Ensure date formats are YYYY-MM-DD

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

See [CONTRIBUTING.md](CONTRIBUTING.md) for more details.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

Mohamed Kandil

## Acknowledgments

- Built with Python and openpyxl
- Designed for Arabic-speaking property managers
- Inspired by the need for simple, effective property management tools

## Support

If you encounter any issues or have questions:
1. Check the [Troubleshooting](#troubleshooting) section
2. Review the example configuration in `config.example.json`
3. Open an issue on GitHub with details about your problem

---

**Note**: For the best experience and latest features, use `excel_generate_v2.py` instead of the legacy version.
