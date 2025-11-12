# Excel Comparison Tool - Python Version

Compare two Excel files (KMS_DML and Documentation) to verify data consistency.

## Features

- Reads Excel files using openpyxl
- Compares references and versions between two files
- Generates detailed comparison reports
- Identifies matches, mismatches, and missing records
- **Windows executable available** - no Python installation required!

## Installation

### Option 1: Run as Python Script

1. Install Python 3.8 or higher
2. Install dependencies:
```bash
pip install -r requirements.txt
```

### Option 2: Use Windows Executable (No Python Required)

1. Download `ExcelComparator.exe` from the `dist` folder
2. Run directly from command line or double-click

## Building the Executable

To build the Windows executable yourself:

```bash
# Run the build script
build.bat
```

Or manually:
```bash
# Install dependencies
pip install -r requirements.txt

# Build with PyInstaller
pyinstaller --clean ExcelComparator.spec
```

The executable will be created in the `dist` folder.

## Usage

### Compare Excel Files (Python Script)

```bash
python main.py <kms-dml-file> <documentation-file>
```

Example:
```bash
python main.py ../test_data/KMS_DML.xlsx ../test_data/Documentation.xlsx
```

### Compare Excel Files (Executable)

```bash
ExcelComparator.exe <kms-dml-file> <documentation-file>
```

Example:
```bash
ExcelComparator.exe ..\test_data\KMS_DML.xlsx ..\test_data\Documentation.xlsx
```

Or use the provided batch file:
```bash
run_example.bat
```

### Generate Test Files

```bash
python generate_test_files.py
```

This will create sample Excel files in the `test_data` folder.

## Comparison Logic

The tool compares:
- **KMS_DML Column A** ↔ **Documentation Column C** (Reference codes)
- **KMS_DML Column Q** ↔ **Documentation Column H** (Version information)

## Output

The comparison report includes:
- ✅ **Matches**: Records with matching reference and version
- ⚠️ **Mismatches**: Records with matching reference but different versions
- ➖ **Only in KMS_DML**: Records missing from Documentation
- ➕ **Only in Documentation**: Records missing from KMS_DML

## Project Structure

```
python/
├── excel_reader.py          # Excel file reader class
├── comparator.py            # Comparison logic and result formatting
├── main.py                  # Main entry point
├── generate_test_files.py   # Test data generator
├── requirements.txt         # Python dependencies
├── ExcelComparator.spec     # PyInstaller configuration
├── build.bat                # Build script for Windows executable
├── run_example.bat          # Example batch file to run the tool
├── README.md                # This file
└── dist/                    # Built executable (after running build.bat)
    └── ExcelComparator.exe
```

## Distribution

To distribute the tool to users without Python:
1. Run `build.bat` to create the executable
2. Share the `dist\ExcelComparator.exe` file
3. Users can run it directly without installing Python!
