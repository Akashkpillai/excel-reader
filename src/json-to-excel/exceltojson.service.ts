import { Injectable } from '@nestjs/common';
import * as ExcelJS from 'exceljs';

@Injectable()
export class ExcelToJsonService {
    async excelToJson1(filePath: string): Promise<any[]> {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const sheetNo = workbook.worksheets.length;
        console.log('lenth', sheetNo);
        const name = workbook.worksheets[0].name;
        console.log('name', name);
        const worksheet = workbook.worksheets[0]; // First sheet

        // Extract multi-level headers
        const headers = this.extractNestedHeaders(worksheet).flat();

        // Extract data rows
        return this.extractRows(worksheet, headers);
    }
    async excelToJson(filePath: string): Promise<any[]> {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);

        const sheetCount = workbook.worksheets.length;
        console.log(`✅ Total Sheets: ${sheetCount}`);

        const allSheetsData: any[] = [];

        // Loop through all sheets
        for (const worksheet of workbook.worksheets) {
            console.log(`📄 Processing Sheet: ${worksheet.name}`);

            // Extract headers
            const headers = this.extractNestedHeaders(worksheet).flat();

            // Extract data rows
            const sheetData = this.extractRows(worksheet, headers);

            // Store with sheet name for reference
            allSheetsData.push({
                sheetName: worksheet.name,
                data: sheetData,
            });
        }

        return allSheetsData;
    }

    /**
     * Extracts multi-level headers dynamically, detecting parent-child relationships.
     */
    private extractNestedHeaders(worksheet: ExcelJS.Worksheet): string[] {
        const maxHeaderRow = this.detectHeaderRows(worksheet); // Get max header row index
        const headersMap: string[][] = [];

        // Iterate only over header rows
        for (let rowIndex = 1; rowIndex <= maxHeaderRow; rowIndex++) {
            const row = worksheet.getRow(rowIndex);
            row.eachCell((cell, colIndex) => {
                if (!headersMap[colIndex]) headersMap[colIndex] = [];
                headersMap[colIndex][rowIndex - 1] = (cell.text || '').trim(); // Store header
            });
        }

        // Flatten multi-level headers correctly
        return headersMap.map((colHeaders) => {
            const uniqueHeaders = [...new Set(colHeaders.filter((h) => h))]; // Remove duplicates
            return uniqueHeaders.join('.'); // Join with dot notation
        });
    }

    /**
     * Detects the number of header rows dynamically.
     * Assumes data rows start after headers.
     */
    private detectHeaderRows1(worksheet: ExcelJS.Worksheet): number {
        let maxHeaderRow = 1;
        let consecutiveTextRows = 0;

        worksheet.eachRow((row, rowIndex) => {
            const values = row.values as any[]; // Ensure it's an array
            const nonEmptyCells = values.filter(
                (val) => val !== undefined && val !== null && val !== ''
            ).length;

            // Skip empty rows
            if (nonEmptyCells === 0) return;

            // Count how many cells contain pure text
            const textCells = values.filter(
                (val) => typeof val === 'string' && isNaN(Number(val))
            ).length;

            // Count how many cells contain numbers (likely data)
            const numericCells = values.filter(
                (val) => typeof val === 'number' || (!isNaN(Number(val)) && val !== '')
            ).length;

            // If row has mostly text, consider it a header row
            if (textCells / nonEmptyCells >= 0.8 && numericCells === 0) {
                maxHeaderRow = rowIndex;
                consecutiveTextRows++;
            } else if (consecutiveTextRows > 0) {
                // If we see a data row after headers, stop checking further
                return;
            }
        });

        console.log('✅ Corrected Header Row Count:', maxHeaderRow);
        return maxHeaderRow;
    }
    private detectHeaderRows(worksheet: ExcelJS.Worksheet): number {
        let lastHeaderRow = 1;
        let columnCount = 0;

        // Get merged cell ranges
        const mergedRanges = worksheet.model.merges || [];

        worksheet.eachRow((row, rowIndex) => {
            const values = row.values as any[];
            const nonEmptyCells = values.filter(
                (val) => val !== undefined && val !== null && val !== ''
            ).length;

            if (nonEmptyCells === 0) return; // Skip empty rows

            if (columnCount === 0) columnCount = nonEmptyCells; // Set expected column count

            // Check if row has merged cells
            const hasMergedCells = mergedRanges.some((range) => {
                const [start, end] = range.split(':').map((ref) => worksheet.getCell(ref).row);
                return rowIndex >= Number(start) && rowIndex <= Number(end);
            });

            if (hasMergedCells) {
                lastHeaderRow = rowIndex; // Continue checking merged rows
            } else {
                return;
            }
        });

        console.log(`✅ Headers Cover ${lastHeaderRow} Rows`);
        return lastHeaderRow;
    }

    /**
     * Extracts row data dynamically and maps it into a structured JSON format.
     */
    private extractRows1(worksheet: ExcelJS.Worksheet, headers: string[]): any[] {
        const rows: any[] = [];
        const maxHeaderRow = this.detectHeaderRows(worksheet);

        worksheet.eachRow((row, rowIndex) => {
            if (rowIndex <= maxHeaderRow) return; // Skip header rows

            const rowData: any = {};
            row.eachCell((cell, colIndex) => {
                const header = headers[colIndex - 1]; // Match column to correct header
                if (header) {
                    this.setNestedValue(rowData, header, cell.value);
                }
            });

            rows.push(rowData);
        });

        return rows;
    }
    private extractRows(worksheet: ExcelJS.Worksheet, headers: string[]): any[] {
        const rows: any[] = [];
        const maxHeaderRow = this.detectHeaderRows(worksheet);

        worksheet.eachRow((row, rowIndex) => {
            if (rowIndex <= maxHeaderRow) return; // Skip header rows

            const rowData: any = {};

            row.eachCell((cell, colIndex) => {
                const header = headers[colIndex - 1]; // Get corresponding header
                if (!header) return;

                let value = cell.value;

                try {
                    // Attempt to parse JSON strings
                    if (typeof value === 'string') {
                        const parsed = JSON.parse(value);
                        value = this.convertNumericObjectsToArray(parsed);
                    }
                } catch (error) {
                    // Keep original value if parsing fails
                }

                this.setNestedValue(rowData, header, value);
            });

            rows.push(rowData);
        });

        return rows;
    }

    /**
     * Converts numeric-keyed objects into arrays recursively
     */
    private convertNumericObjectsToArray(value: any): any {
        if (Array.isArray(value)) {
            return value.map(this.convertNumericObjectsToArray.bind(this));
        } else if (value !== null && typeof value === 'object') {
            const keys = Object.keys(value);
            const allKeysAreNumbers = keys.every((key) => /^\d+$/.test(key)); // Check if all keys are numbers

            if (allKeysAreNumbers) {
                // Convert object with numeric keys into an array
                return keys
                    .sort((a, b) => Number(a) - Number(b)) // Ensure correct order
                    .map((key) => this.convertNumericObjectsToArray(value[key]));
            } else {
                // Recursively process objects
                return Object.fromEntries(
                    keys.map((key) => [key, this.convertNumericObjectsToArray(value[key])])
                );
            }
        }
        return value;
    }

    private isNumericObject(obj: any): boolean {
        if (typeof obj !== 'object' || obj === null) return false;

        return Object.keys(obj).every((key) => !isNaN(Number(key)));
    }

    /**
     * Sets a nested value in an object dynamically based on dot notation.
     */
    private setNestedValue(obj: any, path: string, value: any) {
        const keys = path.split('.');
        let current = obj;

        for (let i = 0; i < keys.length - 1; i++) {
            const key = keys[i];

            // ✅ Ensure the current key is an object, not a string
            if (typeof current[key] === 'string') {
                console.error(`❌ Cannot assign properties to a string at key: ${key}`);
                return; // Skip assignment if it's a string
            }

            // ✅ Create an object if it doesn't exist
            if (!current[key]) {
                current[key] = {};
            }

            current = current[key];
        }

        // ✅ Ensure we're not assigning to a string
        if (typeof current === 'string') {
            console.error(`❌ Cannot assign to a string at path: ${path}`);
            return;
        }

        // Set the final value
        current[keys[keys.length - 1]] = value;
    }

    private setNestedValue1(obj: any, path: string, value: any) {
        const keys = path.split('.');
        let current = obj;

        keys.forEach((key, index) => {
            if (index === keys.length - 1) {
                current[key] = value;
            } else {
                current[key] = current[key] || {};
                current = current[key];
            }
        });
    }
}
