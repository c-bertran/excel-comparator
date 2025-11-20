/**
 * Comparator class to compare data between KMS_DML and Documentation
 */

export interface ComparisonResult {
  matches: MatchedRecord[];
  mismatches: MismatchedRecord[];
  onlyInFile1: UnmatchedRecord[];
  onlyInFile2: UnmatchedRecord[];
  summary: ComparisonSummary;
  file1Name: string;
  file2Name: string;
}

export interface MatchedRecord {
  reference: string;
  file1Version: any;
  file2Version: any;
  file1Row: number;
  file2Row: number;
}

export interface MismatchedRecord {
  reference: string;
  file1Version: any;
  file2Version: any;
  file1Row: number;
  file2Row: number;
}

export interface UnmatchedRecord {
  reference: string;
  version: any;
  row: number;
}

export interface ComparisonSummary {
  totalFile1: number;
  totalFile2: number;
  matches: number;
  mismatches: number;
  onlyInFile1: number;
  onlyInFile2: number;
}

export class ExcelComparator {
  /**
   * Compare data from two files
   * @param file1Data - Data from first file
   * @param file2Data - Data from second file
   * @param file1Column - Column letter from first file
   * @param file2Column - Column letter from second file
   * @param file1Name - Display name for first file
   * @param file2Name - Display name for second file
   * @returns Comparison results
   */
  compare(
    file1Data: Array<Record<string, any>>,
    file2Data: Array<Record<string, any>>,
    file1Column: string,
    file2Column: string,
    file1Name: string = 'File 1',
    file2Name: string = 'File 2'
  ): ComparisonResult {
    const matches: MatchedRecord[] = [];
    const mismatches: MismatchedRecord[] = [];
    const onlyInFile1: UnmatchedRecord[] = [];
    const onlyInFile2: UnmatchedRecord[] = [];

    // Create maps for easier lookup
    const file2Map = new Map<string, { version: any; row: number }>();
    
    // Populate file2 map using the specified column
    file2Data.forEach((row) => {
      const reference = this.normalizeReference(row[file2Column]);
      if (reference) {
        file2Map.set(reference, {
          version: row[file2Column],
          row: row['_rowNumber']
        });
      }
    });

    // Track which file2 records were matched
    const matchedFile2Refs = new Set<string>();

    // Compare file1 records using the specified column
    file1Data.forEach((file1Row) => {
      const reference = this.normalizeReference(file1Row[file1Column]);
      
      if (!reference) return;

      const file1Version = file1Row[file1Column];
      const file1RowNumber = file1Row['_rowNumber'];

      const file2Record = file2Map.get(reference);

      if (file2Record) {
        // Found matching reference
        matchedFile2Refs.add(reference);

        if (this.versionsMatch(file1Version, file2Record.version)) {
          // Versions match
          matches.push({
            reference,
            file1Version,
            file2Version: file2Record.version,
            file1Row: file1RowNumber,
            file2Row: file2Record.row
          });
        } else {
          // Versions don't match
          mismatches.push({
            reference,
            file1Version,
            file2Version: file2Record.version,
            file1Row: file1RowNumber,
            file2Row: file2Record.row
          });
        }
      } else {
        // Only in file1
        onlyInFile1.push({
          reference,
          version: file1Version,
          row: file1RowNumber
        });
      }
    });

    // Find records only in file2
    file2Data.forEach((file2Row) => {
      const reference = this.normalizeReference(file2Row[file2Column]);
      
      if (!reference) return;

      if (!matchedFile2Refs.has(reference)) {
        onlyInFile2.push({
          reference,
          version: file2Row[file2Column],
          row: file2Row['_rowNumber']
        });
      }
    });

    const summary: ComparisonSummary = {
      totalFile1: file1Data.length,
      totalFile2: file2Data.length,
      matches: matches.length,
      mismatches: mismatches.length,
      onlyInFile1: onlyInFile1.length,
      onlyInFile2: onlyInFile2.length
    };

    return {
      matches,
      mismatches,
      onlyInFile1,
      onlyInFile2,
      summary,
      file1Name,
      file2Name
    };
  }

  /**
   * Normalize reference strings for comparison (trim, uppercase)
   */
  private normalizeReference(ref: any): string | null {
    if (!ref) return null;
    
    const refStr = String(ref).trim();
    return refStr.length > 0 ? refStr : null;
  }

  /**
   * Check if two versions match (handles null, undefined, and string comparison)
   */
  private versionsMatch(version1: any, version2: any): boolean {
    // Convert to strings and trim
    const v1 = String(version1 || '').trim();
    const v2 = String(version2 || '').trim();
    
    // Case-insensitive comparison
    return v1.toLowerCase() === v2.toLowerCase();
  }

  /**
   * Print comparison results to console in a formatted way
   */
  printResults(result: ComparisonResult): void {
    console.log('\n' + '='.repeat(80));
    console.log('COMPARISON RESULTS');
    console.log('='.repeat(80));

    // Summary
    console.log('\n📊 SUMMARY:');
    console.log(`   Total records in ${result.file1Name}: ${result.summary.totalFile1}`);
    console.log(`   Total records in ${result.file2Name}: ${result.summary.totalFile2}`);
    console.log(`   ✅ Matches: ${result.summary.matches}`);
    console.log(`   ⚠️  Mismatches: ${result.summary.mismatches}`);
    console.log(`   ➖ Only in ${result.file1Name}: ${result.summary.onlyInFile1}`);
    console.log(`   ➕ Only in ${result.file2Name}: ${result.summary.onlyInFile2}`);

    // Matches
    if (result.matches.length > 0) {
      console.log('\n' + '─'.repeat(80));
      console.log('✅ MATCHES (Reference and Version match):');
      console.log('─'.repeat(80));
      result.matches.forEach((match) => {
        console.log(`   ${match.reference}`);
        console.log(`      Version: ${match.file1Version}`);
        console.log(`      ${result.file1Name} row: ${match.file1Row}, ${result.file2Name} row: ${match.file2Row}`);
      });
    }

    // Mismatches
    if (result.mismatches.length > 0) {
      console.log('\n' + '─'.repeat(80));
      console.log('⚠️  MISMATCHES (Reference found but Version differs):');
      console.log('─'.repeat(80));
      result.mismatches.forEach((mismatch) => {
        console.log(`   ${mismatch.reference}`);
        console.log(`      ${result.file1Name} version: "${mismatch.file1Version}" (row ${mismatch.file1Row})`);
        console.log(`      ${result.file2Name} version: "${mismatch.file2Version}" (row ${mismatch.file2Row})`);
      });
    }

    // Only in File 1
    if (result.onlyInFile1.length > 0) {
      console.log('\n' + '─'.repeat(80));
      console.log(`➖ ONLY IN ${result.file1Name.toUpperCase()} (not found in ${result.file2Name}):`);
      console.log('─'.repeat(80));
      result.onlyInFile1.forEach((record) => {
        console.log(`   ${record.reference} (version: ${record.version}, row: ${record.row})`);
      });
    }

    // Only in File 2
    if (result.onlyInFile2.length > 0) {
      console.log('\n' + '─'.repeat(80));
      console.log(`➕ ONLY IN ${result.file2Name.toUpperCase()} (not found in ${result.file1Name}):`);
      console.log('─'.repeat(80));
      result.onlyInFile2.forEach((record) => {
        console.log(`   ${record.reference} (version: ${record.version}, row: ${record.row})`);
      });
    }

    console.log('\n' + '='.repeat(80));
  }
}
