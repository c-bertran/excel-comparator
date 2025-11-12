"""
Comparator class to compare data between KMS_DML and Documentation
"""
from typing import List, Dict, Any
from dataclasses import dataclass


@dataclass
class MatchedRecord:
    """Record that matches in both files"""
    reference: str
    kms_dml_version: Any
    documentation_version: Any
    kms_dml_row: int
    documentation_row: int


@dataclass
class MismatchedRecord:
    """Record with matching reference but different versions"""
    reference: str
    kms_dml_version: Any
    documentation_version: Any
    kms_dml_row: int
    documentation_row: int


@dataclass
class UnmatchedRecord:
    """Record found in only one file"""
    reference: str
    version: Any
    row: int


@dataclass
class ComparisonSummary:
    """Summary statistics of the comparison"""
    total_kms_dml: int
    total_documentation: int
    matches: int
    mismatches: int
    only_in_kms_dml: int
    only_in_documentation: int


@dataclass
class ComparisonResult:
    """Complete comparison results"""
    matches: List[MatchedRecord]
    mismatches: List[MismatchedRecord]
    only_in_kms_dml: List[UnmatchedRecord]
    only_in_documentation: List[UnmatchedRecord]
    summary: ComparisonSummary


class ExcelComparator:
    """Comparator class to compare KMS_DML and Documentation data"""

    def compare(
        self,
        kms_dml_data: List[Dict[str, Any]],
        doc_data: List[Dict[str, Any]]
    ) -> ComparisonResult:
        """
        Compare KMS_DML data with Documentation data
        
        Args:
            kms_dml_data: Data from KMS_DML (columns A and Q)
            doc_data: Data from Documentation (columns C and H)
            
        Returns:
            Comparison results
        """
        matches = []
        mismatches = []
        only_in_kms_dml = []
        only_in_documentation = []

        # Create map for easier lookup (Documentation)
        doc_map = {}
        for row in doc_data:
            reference = self._normalize_reference(row.get('C'))
            if reference:
                doc_map[reference] = {
                    'version': row.get('H'),
                    'row': row.get('_rowNumber')
                }

        # Track which documentation records were matched
        matched_doc_refs = set()

        # Compare KMS_DML records (column A as reference, column Q as version)
        for kms_row in kms_dml_data:
            reference = self._normalize_reference(kms_row.get('A'))
            
            if not reference:
                continue

            kms_dml_version = kms_row.get('Q')
            kms_dml_row_number = kms_row.get('_rowNumber')

            doc_record = doc_map.get(reference)

            if doc_record:
                # Found matching reference
                matched_doc_refs.add(reference)

                if self._versions_match(kms_dml_version, doc_record['version']):
                    # Versions match
                    matches.append(MatchedRecord(
                        reference=reference,
                        kms_dml_version=kms_dml_version,
                        documentation_version=doc_record['version'],
                        kms_dml_row=kms_dml_row_number,
                        documentation_row=doc_record['row']
                    ))
                else:
                    # Versions don't match
                    mismatches.append(MismatchedRecord(
                        reference=reference,
                        kms_dml_version=kms_dml_version,
                        documentation_version=doc_record['version'],
                        kms_dml_row=kms_dml_row_number,
                        documentation_row=doc_record['row']
                    ))
            else:
                # Only in KMS_DML
                only_in_kms_dml.append(UnmatchedRecord(
                    reference=reference,
                    version=kms_dml_version,
                    row=kms_dml_row_number
                ))

        # Find records only in Documentation
        for doc_row in doc_data:
            reference = self._normalize_reference(doc_row.get('C'))
            
            if not reference:
                continue

            if reference not in matched_doc_refs:
                only_in_documentation.append(UnmatchedRecord(
                    reference=reference,
                    version=doc_row.get('H'),
                    row=doc_row.get('_rowNumber')
                ))

        summary = ComparisonSummary(
            total_kms_dml=len(kms_dml_data),
            total_documentation=len(doc_data),
            matches=len(matches),
            mismatches=len(mismatches),
            only_in_kms_dml=len(only_in_kms_dml),
            only_in_documentation=len(only_in_documentation)
        )

        return ComparisonResult(
            matches=matches,
            mismatches=mismatches,
            only_in_kms_dml=only_in_kms_dml,
            only_in_documentation=only_in_documentation,
            summary=summary
        )

    @staticmethod
    def _normalize_reference(ref: Any) -> str | None:
        """
        Normalize reference strings for comparison (trim)
        
        Args:
            ref: Reference value to normalize
            
        Returns:
            Normalized reference or None
        """
        if ref is None:
            return None
        
        ref_str = str(ref).strip()
        return ref_str if len(ref_str) > 0 else None

    @staticmethod
    def _versions_match(version1: Any, version2: Any) -> bool:
        """
        Check if two versions match (handles null, undefined, and string comparison)
        
        Args:
            version1: First version to compare
            version2: Second version to compare
            
        Returns:
            True if versions match, False otherwise
        """
        # Convert to strings and trim
        v1 = str(version1 or '').strip()
        v2 = str(version2 or '').strip()
        
        # Case-insensitive comparison
        return v1.lower() == v2.lower()

    def print_results(self, result: ComparisonResult) -> None:
        """
        Print comparison results to console in a formatted way
        
        Args:
            result: Comparison results to print
        """
        print('\n' + '=' * 80)
        print('COMPARISON RESULTS')
        print('=' * 80)

        # Summary
        print('\n📊 SUMMARY:')
        print(f'   Total records in KMS_DML: {result.summary.total_kms_dml}')
        print(f'   Total records in Documentation: {result.summary.total_documentation}')
        print(f'   ✅ Matches: {result.summary.matches}')
        print(f'   ⚠️  Mismatches: {result.summary.mismatches}')
        print(f'   ➖ Only in KMS_DML: {result.summary.only_in_kms_dml}')
        print(f'   ➕ Only in Documentation: {result.summary.only_in_documentation}')

        # Matches
        if result.matches:
            print('\n' + '─' * 80)
            print('✅ MATCHES (Reference and Version match):')
            print('─' * 80)
            for match in result.matches:
                print(f'   {match.reference}')
                print(f'      Version: {match.kms_dml_version}')
                print(f'      KMS_DML row: {match.kms_dml_row}, Documentation row: {match.documentation_row}')

        # Mismatches
        if result.mismatches:
            print('\n' + '─' * 80)
            print('⚠️  MISMATCHES (Reference found but Version differs):')
            print('─' * 80)
            for mismatch in result.mismatches:
                print(f'   {mismatch.reference}')
                print(f'      KMS_DML version: "{mismatch.kms_dml_version}" (row {mismatch.kms_dml_row})')
                print(f'      Documentation version: "{mismatch.documentation_version}" (row {mismatch.documentation_row})')

        # Only in KMS_DML
        if result.only_in_kms_dml:
            print('\n' + '─' * 80)
            print('➖ ONLY IN KMS_DML (not found in Documentation):')
            print('─' * 80)
            for record in result.only_in_kms_dml:
                print(f'   {record.reference} (version: {record.version}, row: {record.row})')

        # Only in Documentation
        if result.only_in_documentation:
            print('\n' + '─' * 80)
            print('➕ ONLY IN DOCUMENTATION (not found in KMS_DML):')
            print('─' * 80)
            for record in result.only_in_documentation:
                print(f'   {record.reference} (version: {record.version}, row: {record.row})')

        print('\n' + '=' * 80)
