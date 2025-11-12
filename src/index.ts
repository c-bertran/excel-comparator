import { ExcelReader } from '@/excel';
import { ExcelComparator } from './comparator';

(async () => {
  try {
    // Get file paths from command line arguments
    const kmsDmlFilePath = process.argv[2];
    const documentationFilePath = process.argv[3];
    
    if (!kmsDmlFilePath || !documentationFilePath) {
      console.error('Error: Please provide both file paths as arguments');
      console.log('Usage: npm run start <kms-dml-file> <documentation-file>');
      console.log('Example: npm run start test_data/KMS_DML.xlsx test_data/Documentation.xlsx');
      process.exit(1);
    }

    console.log('📂 Loading Excel files...\n');

    // Read KMS_DML file
    const kmsDmlReader = new ExcelReader(kmsDmlFilePath);
    await kmsDmlReader.load();
    console.log(`✓ Loaded KMS_DML file: ${kmsDmlFilePath}`);
    console.log(`  Sheets: ${kmsDmlReader.getSheetNames().join(', ')}`);

    // Read Documentation file
    const docReader = new ExcelReader(documentationFilePath);
    await docReader.load();
    console.log(`✓ Loaded Documentation file: ${documentationFilePath}`);
    console.log(`  Sheets: ${docReader.getSheetNames().join(', ')}`);

    // Extract data from KMS_DML sheet (columns A and Q)
    console.log('\n📊 Extracting data...');
    const kmsDmlData = kmsDmlReader.extractColumns('KMS_DML', ['A', 'Q']);
    console.log(`  Extracted ${kmsDmlData.length} rows from KMS_DML (columns A, Q)`);

    // Extract data from Documentation sheet (columns C and H)
    const docData = docReader.extractColumns('Documentation', ['C', 'H']);
    console.log(`  Extracted ${docData.length} rows from Documentation (columns C, H)`);

    // Perform comparison
    console.log('\n🔍 Comparing data...');
    const comparator = new ExcelComparator();
    const result = comparator.compare(kmsDmlData, docData);

    // Print results
    comparator.printResults(result);

  } catch (error) {
    console.error('Error:', error);
    process.exit(1);
  }
})();
