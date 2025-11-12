/**
 * Comparator class to compare data between KMS_DML and Documentation
 */

export interface ComparisonResult {
  matches: MatchedRecord[];
  mismatches: MismatchedRecord[];
  onlyInKmsDml: UnmatchedRecord[];
  onlyInDocumentation: UnmatchedRecord[];
  summary: ComparisonSummary;
}

export interface MatchedRecord {
  reference: string;
  kmsDmlVersion: any;
  documentationVersion: any;
  kmsDmlRow: number;
  documentationRow: number;
}

export interface MismatchedRecord {
  reference: string;
  kmsDmlVersion: any;
  documentationVersion: any;
  kmsDmlRow: number;
  documentationRow: number;
}

export interface UnmatchedRecord {
  reference: string;
  version: any;
  row: number;
}

export interface ComparisonSummary {
  totalKmsDml: number;
  totalDocumentation: number;
  matches: number;
  mismatches: number;
  onlyInKmsDml: number;
  onlyInDocumentation: number;
}

export class ExcelComparator {
  /**
   * Compare KMS_DML data with Documentation data
   * @param kmsDmlData - Data from KMS_DML (columns A and Q)
   * @param docData - Data from Documentation (columns C and H)
   * @returns Comparison results
   */
  compare(
    kmsDmlData: Array<Record<string, any>>,
    docData: Array<Record<string, any>>
  ): ComparisonResult {
    const matches: MatchedRecord[] = [];
    const mismatches: MismatchedRecord[] = [];
    const onlyInKmsDml: UnmatchedRecord[] = [];
    const onlyInDocumentation: UnmatchedRecord[] = [];

    // Create maps for easier lookup
    const docMap = new Map<string, { version: any; row: number }>();
    
    // Populate Documentation map (using column C as key, column H as version)
    docData.forEach((row) => {
      const reference = this.normalizeReference(row['C']);
      if (reference) {
        docMap.set(reference, {
          version: row['H'],
          row: row['_rowNumber']
        });
      }
    });

    // Track which documentation records were matched
    const matchedDocRefs = new Set<string>();

    // Compare KMS_DML records (column A as reference, column Q as version)
    kmsDmlData.forEach((kmsRow) => {
      const reference = this.normalizeReference(kmsRow['A']);
      
      if (!reference) return;

      const kmsDmlVersion = kmsRow['Q'];
      const kmsDmlRowNumber = kmsRow['_rowNumber'];

      const docRecord = docMap.get(reference);

      if (docRecord) {
        // Found matching reference
        matchedDocRefs.add(reference);

        if (this.versionsMatch(kmsDmlVersion, docRecord.version)) {
          // Versions match
          matches.push({
            reference,
            kmsDmlVersion,
            documentationVersion: docRecord.version,
            kmsDmlRow: kmsDmlRowNumber,
            documentationRow: docRecord.row
          });
        } else {
          // Versions don't match
          mismatches.push({
            reference,
            kmsDmlVersion,
            documentationVersion: docRecord.version,
            kmsDmlRow: kmsDmlRowNumber,
            documentationRow: docRecord.row
          });
        }
      } else {
        // Only in KMS_DML
        onlyInKmsDml.push({
          reference,
          version: kmsDmlVersion,
          row: kmsDmlRowNumber
        });
      }
    });

    // Find records only in Documentation
    docData.forEach((docRow) => {
      const reference = this.normalizeReference(docRow['C']);
      
      if (!reference) return;

      if (!matchedDocRefs.has(reference)) {
        onlyInDocumentation.push({
          reference,
          version: docRow['H'],
          row: docRow['_rowNumber']
        });
      }
    });

    const summary: ComparisonSummary = {
      totalKmsDml: kmsDmlData.length,
      totalDocumentation: docData.length,
      matches: matches.length,
      mismatches: mismatches.length,
      onlyInKmsDml: onlyInKmsDml.length,
      onlyInDocumentation: onlyInDocumentation.length
    };

    return {
      matches,
      mismatches,
      onlyInKmsDml,
      onlyInDocumentation,
      summary
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
    console.log(`   Total records in KMS_DML: ${result.summary.totalKmsDml}`);
    console.log(`   Total records in Documentation: ${result.summary.totalDocumentation}`);
    console.log(`   ✅ Matches: ${result.summary.matches}`);
    console.log(`   ⚠️  Mismatches: ${result.summary.mismatches}`);
    console.log(`   ➖ Only in KMS_DML: ${result.summary.onlyInKmsDml}`);
    console.log(`   ➕ Only in Documentation: ${result.summary.onlyInDocumentation}`);

    // Matches
    if (result.matches.length > 0) {
      console.log('\n' + '─'.repeat(80));
      console.log('✅ MATCHES (Reference and Version match):');
      console.log('─'.repeat(80));
      result.matches.forEach((match) => {
        console.log(`   ${match.reference}`);
        console.log(`      Version: ${match.kmsDmlVersion}`);
        console.log(`      KMS_DML row: ${match.kmsDmlRow}, Documentation row: ${match.documentationRow}`);
      });
    }

    // Mismatches
    if (result.mismatches.length > 0) {
      console.log('\n' + '─'.repeat(80));
      console.log('⚠️  MISMATCHES (Reference found but Version differs):');
      console.log('─'.repeat(80));
      result.mismatches.forEach((mismatch) => {
        console.log(`   ${mismatch.reference}`);
        console.log(`      KMS_DML version: "${mismatch.kmsDmlVersion}" (row ${mismatch.kmsDmlRow})`);
        console.log(`      Documentation version: "${mismatch.documentationVersion}" (row ${mismatch.documentationRow})`);
      });
    }

    // Only in KMS_DML
    if (result.onlyInKmsDml.length > 0) {
      console.log('\n' + '─'.repeat(80));
      console.log('➖ ONLY IN KMS_DML (not found in Documentation):');
      console.log('─'.repeat(80));
      result.onlyInKmsDml.forEach((record) => {
        console.log(`   ${record.reference} (version: ${record.version}, row: ${record.row})`);
      });
    }

    // Only in Documentation
    if (result.onlyInDocumentation.length > 0) {
      console.log('\n' + '─'.repeat(80));
      console.log('➕ ONLY IN DOCUMENTATION (not found in KMS_DML):');
      console.log('─'.repeat(80));
      result.onlyInDocumentation.forEach((record) => {
        console.log(`   ${record.reference} (version: ${record.version}, row: ${record.row})`);
      });
    }

    console.log('\n' + '='.repeat(80));
  }
}
