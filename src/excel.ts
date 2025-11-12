import * as ExcelJS from 'exceljs';

/**
 * Excel reader class to simplify reading and extracting data from Excel files
 */
export class ExcelReader {
	private workbook: ExcelJS.Workbook;
	private filePath: string;

	constructor(filePath: string) {
		this.filePath = filePath;
		this.workbook = new ExcelJS.Workbook();
	}

	/**
	 * Load the Excel file
	 */
	async load(): Promise<void> {
		await this.workbook.xlsx.readFile(this.filePath);
	}

	/**
	 * Get a worksheet by name
	 */
	getWorksheet(sheetName: string): ExcelJS.Worksheet | undefined {
		return this.workbook.getWorksheet(sheetName);
	}

	/**
	 * Extract data from specific columns in a worksheet
	 * @param sheetName - Name of the sheet
	 * @param columns - Column letters to extract (e.g., ['A', 'Q'])
	 * @param startRow - Starting row (default: 2, skip header)
	 * @returns Array of row data with specified columns
	 */
	extractColumns(
		sheetName: string,
		columns: string[],
		startRow: number = 2
	): Array<Record<string, any>> {
		const worksheet = this.getWorksheet(sheetName);
		if (!worksheet) {
			throw new Error(`Worksheet "${sheetName}" not found`);
		}

		const data: Array<Record<string, any>> = [];
		
		worksheet.eachRow((row, rowNumber) => {
			// Skip rows before startRow
			if (rowNumber < startRow) return;

			const rowData: Record<string, any> = {};
			columns.forEach(col => {
				const cell = row.getCell(col);
				rowData[col] = cell.value;
			});

			// Only add row if at least one column has data
			if (Object.values(rowData).some(val => val !== null && val !== undefined)) {
				rowData['_rowNumber'] = rowNumber;
				data.push(rowData);
			}
		});

		return data;
	}

	/**
	 * Get all sheet names in the workbook
	 */
	getSheetNames(): string[] {
		return this.workbook.worksheets.map(sheet => sheet.name);
	}

	/**
	 * Get row count for a specific sheet
	 */
	getRowCount(sheetName: string): number {
		const worksheet = this.getWorksheet(sheetName);
		if (!worksheet) {
			throw new Error(`Worksheet "${sheetName}" not found`);
		}
		return worksheet.rowCount;
	}
}
