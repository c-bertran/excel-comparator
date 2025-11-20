/**
 * Command-line argument parser for Excel comparison tool
 */

export interface FileArgument {
  filePath: string;
  sheetName: string;
  columnLetter: string;
}

export interface ParsedArgs {
  files: [FileArgument, FileArgument];
  showHelp: boolean;
}

export class ArgumentParser {
  private args: string[];

  constructor(args: string[]) {
    // Remove first two elements (node and script path)
    this.args = args.slice(2);
  }

  /**
   * Parse command-line arguments
   */
  parse(): ParsedArgs {
    // Check for help flag
    if (this.args.length === 0 || this.args.includes('--help') || this.args.includes('-h')) {
      return {
        files: [] as any,
        showHelp: true
      };
    }

    // Parse file arguments
    const files = this.parseFileArguments();

    if (files.length !== 2) {
      throw new Error('Exactly 2 files must be specified with -f flag');
    }

    return {
      files: files as [FileArgument, FileArgument],
      showHelp: false
    };
  }

  /**
   * Parse -f file arguments
   */
  private parseFileArguments(): FileArgument[] {
    const files: FileArgument[] = [];
    let i = 0;

    while (i < this.args.length) {
      if (this.args[i] === '-f') {
        // Expect 3 arguments after -f: filePath, sheetName, columnLetter
        if (i + 3 >= this.args.length) {
          throw new Error(`Invalid -f argument at position ${i}. Expected: -f <file-path> <sheet-name> <column-letter>`);
        }

        const filePath = this.args[i + 1];
        const sheetName = this.args[i + 2];
        const columnLetter = this.args[i + 3];

        if (!filePath || !sheetName || !columnLetter) {
          throw new Error(`Missing arguments for -f flag at position ${i}`);
        }

        // Validate column letter
        if (!this.isValidColumnLetter(columnLetter)) {
          throw new Error(`Invalid column letter: ${columnLetter}. Must be A-Z or AA-ZZ format`);
        }

        files.push({
          filePath,
          sheetName,
          columnLetter: columnLetter.toUpperCase()
        });

        i += 4; // Move past -f and its 3 arguments
      } else {
        throw new Error(`Unknown argument: ${this.args[i]}`);
      }
    }

    return files;
  }

  /**
   * Validate column letter format (A-Z or AA-ZZ)
   */
  private isValidColumnLetter(col: string): boolean {
    return /^[A-Za-z]{1,2}$/.test(col);
  }

  /**
   * Print help message
   */
  static printHelp(): void {
    console.log(`
Excel Comparator Tool - Compare references and versions between two Excel files

USAGE:
  comparator -f <file1> <sheet1> <column1> -f <file2> <sheet2> <column2>

OPTIONS:
  -h, --help              Show this help message
  -f <file> <sheet> <col> Specify an Excel file, sheet name, and column letter

ARGUMENTS:
  <file>    Path to Excel file (.xlsx, .xlsm, .xls)
  <sheet>   Name of the sheet/tab within the Excel file
  <col>     Column letter to extract (A-Z or AA-ZZ)

EXAMPLES:
  # Compare column A from Sheet1 in file1.xlsx with column B from Sheet2 in file2.xlsx
  comparator -f file1.xlsx Sheet1 A -f file2.xlsx Sheet2 B

  # Compare KMS_DML files
  comparator -f KMS_DML.xlsm KMS_DML A -f Documentation.xlsx Documentation C

  # Real-world example with macro-enabled file
  comparator -f RNC_T_A442203en.xlsm Documentation C -f DML_final.xlsx KMS_DML A

COMPARISON LOGIC:
  - First file (-f) is treated as the reference (KMS_DML-like)
  - Second file (-f) is treated as the target (Documentation-like)
  - The tool compares the specified columns from each sheet
  - Outputs matches, mismatches, and missing records

SUPPORTED FORMATS:
  .xlsx  - Excel Workbook
  .xlsm  - Excel Macro-Enabled Workbook
  .xls   - Excel 97-2003 Workbook

EXIT CODES:
  0  - Success
  1  - Error (invalid arguments, file not found, etc.)
`);
  }
}
