"""
Main script to compare KMS_DML and Documentation Excel files
"""
import sys
from excel_reader import ExcelReader
from comparator import ExcelComparator


def main():
    """Main function to run the Excel comparison"""
    # Get file paths from command line arguments
    if len(sys.argv) < 3:
        print('Error: Please provide both file paths as arguments')
        print('Usage: ./Program main.py <kms-dml-file> <documentation-file>')
        print('Example: ./Program main.py ../test_data/KMS_DML.xlsx ../test_data/Documentation.xlsx')
        sys.exit(1)

    kms_dml_file_path = sys.argv[1]
    documentation_file_path = sys.argv[2]

    try:
        print('📂 Loading Excel files...\n')

        # Read KMS_DML file
        kms_dml_reader = ExcelReader(kms_dml_file_path)
        kms_dml_reader.load()
        print(f'✓ Loaded KMS_DML file: {kms_dml_file_path}')
        print(f'  Sheets: {", ".join(kms_dml_reader.get_sheet_names())}')

        # Read Documentation file
        doc_reader = ExcelReader(documentation_file_path)
        doc_reader.load()
        print(f'✓ Loaded Documentation file: {documentation_file_path}')
        print(f'  Sheets: {", ".join(doc_reader.get_sheet_names())}')

        # Extract data from KMS_DML sheet (columns A and Q)
        print('\n📊 Extracting data...')
        kms_dml_data = kms_dml_reader.extract_columns('KMS_DML', ['A', 'Q'])
        print(f'  Extracted {len(kms_dml_data)} rows from KMS_DML (columns A, Q)')

        # Extract data from Documentation sheet (columns C and H)
        doc_data = doc_reader.extract_columns('Documentation', ['C', 'H'])
        print(f'  Extracted {len(doc_data)} rows from Documentation (columns C, H)')

        # Perform comparison
        print('\n🔍 Comparing data...')
        comparator = ExcelComparator()
        result = comparator.compare(kms_dml_data, doc_data)

        # Print results
        comparator.print_results(result)

    except Exception as error:
        print(f'Error: {error}')
        sys.exit(1)


if __name__ == '__main__':
    main()
