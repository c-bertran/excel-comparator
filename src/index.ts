import { ExcelReader } from '@/excel';
import { ExcelComparator } from './comparator';
import { ArgumentParser } from './argParser';
import * as fs from 'fs';
import * as path from 'path';

function validateFile(filePath: string): void {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
  
  const ext = path.extname(filePath).toLowerCase();
  if (!['.xlsx', '.xlsm', '.xls'].includes(ext)) {
    throw new Error(`Unsupported file format: ${ext}. Supported formats: .xlsx, .xlsm, .xls`);
  }
}

(async () => {
  try {
    // Parse command-line arguments
    const parser = new ArgumentParser(process.argv);
    const args = parser.parse();

    // Show help if requested
    if (args.showHelp) {
      ArgumentParser.printHelp();
      process.exit(0);
    }

    const [file1, file2] = args.files;

    // Validate files
    validateFile(file1.filePath);
    validateFile(file2.filePath);

    console.log('📂 Loading Excel files...\n');

    // Read first file
    const reader1 = new ExcelReader(file1.filePath);
    await reader1.load();
    console.log(`✓ Loaded file 1: ${file1.filePath}`);
    console.log(`  Sheets: ${reader1.getSheetNames().join(', ')}`);
    console.log(`  Target: Sheet "${file1.sheetName}", Column ${file1.columnLetter}`);

    // Read second file
    const reader2 = new ExcelReader(file2.filePath);
    await reader2.load();
    console.log(`✓ Loaded file 2: ${file2.filePath}`);
    console.log(`  Sheets: ${reader2.getSheetNames().join(', ')}`);
    console.log(`  Target: Sheet "${file2.sheetName}", Column ${file2.columnLetter}`);

    // Extract data from first file
    console.log('\n📊 Extracting data...');
    const data1 = reader1.extractColumns(file1.sheetName, [file1.columnLetter]);
    console.log(`  Extracted ${data1.length} rows from file 1 (${file1.sheetName}, column ${file1.columnLetter})`);

    // Extract data from second file
    const data2 = reader2.extractColumns(file2.sheetName, [file2.columnLetter]);
    console.log(`  Extracted ${data2.length} rows from file 2 (${file2.sheetName}, column ${file2.columnLetter})`);

    // Perform comparison
    console.log('\n🔍 Comparing data...');
    const comparator = new ExcelComparator();

    // Create display names for the files
    const file1DisplayName = `${path.basename(file1.filePath)} [${file1.sheetName}:${file1.columnLetter}]`;
    const file2DisplayName = `${path.basename(file2.filePath)} [${file2.sheetName}:${file2.columnLetter}]`;

    // Pass data directly with the column letters - no transformation needed
    const result = comparator.compare(
      data1, 
      data2, 
      file1.columnLetter, 
      file2.columnLetter,
      file1DisplayName, 
      file2DisplayName
    );

    // Print results
    comparator.printResults(result);

  } catch (error) {
    if (error instanceof Error) {
      console.error('Error:', error.message);
    } else {
      console.error('Error:', error);
    }
    console.log('\nUse -h or --help for usage information');
    process.exit(1);
  }
})();
