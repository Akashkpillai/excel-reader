import { Injectable } from '@nestjs/common';
import * as ExcelJS from 'exceljs';

@Injectable()
export class ExcelService {
    /**
     * Converts JSON data into an Excel file.
     * @param jsonData - Array of JSON objects to be converted.
     * @param filePath - Path where the generated Excel file will be saved.
     * @param frozenColumns - Optional list of column names to freeze.
     * @returns The file path of the generated Excel file.
     * @throws Error if JSON data is invalid or empty.
     */
    async jsonToExcel(
        jsonData: any[],
        filePath: string,
        frozenColumns: string[] = []
    ): Promise<string> {
        try {
            if (!this.isUniformStructure(jsonData)) {
                throw new Error('Invalid or Not Same Array JSON data');
            }
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Data');

            const headers = this.extractHeaders(jsonData);
            const structuredHeaders = this.structureHeaders(headers);

            if (!structuredHeaders || Object.keys(structuredHeaders).length === 0) {
                throw new Error('No valid headers found in JSON data');
            }

            const columnMapping = this.createHeaderRow(worksheet, structuredHeaders);
            const fullPathMapping = this.generateFullPathMapping(structuredHeaders, '');

            this.fillDataRows(worksheet, jsonData, headers);

            if (frozenColumns.length > 0) {
                this.applyFreezing(worksheet, columnMapping, fullPathMapping, frozenColumns);
            }

            await workbook.xlsx.writeFile(filePath);
            return filePath;
        } catch (error) {
            console.error('Error generating Excel file:', error);
            throw new Error('Failed to generate Excel file');
        }
    }

    /**
     * Extracts all unique headers (including nested properties) from the JSON data.
     * @param data - The JSON data array.
     * @param parentKey - The parent key for nested properties (used for recursion).
     * @returns A set of unique header paths.
     */
    private extractHeaders(data: any[], parentKey = ''): Set<string> {
        const headers = new Set<string>();
        data.forEach((item) => {
            Object.entries(item).forEach(([key, value]) => {
                const newKey = parentKey ? `${parentKey}.${key}` : key;
                if (typeof value === 'object' && value !== null) {
                    const subHeaders = this.extractHeaders([value], newKey);
                    subHeaders.forEach((sub) => headers.add(sub));
                } else {
                    headers.add(newKey);
                }
            });
        });
        return headers;
    }

    /**
     * Converts a flat header set into a structured hierarchy.
     * @param headers - The set of flat headers.
     * @returns A hierarchical object representing the header structure.
     */
    private structureHeaders(headers: Set<string>): any {
        const structuredHeaders: any = {};
        headers.forEach((path) => {
            const parts = path.split('.');
            let currentLevel = structuredHeaders;
            parts.forEach((part, index) => {
                if (!currentLevel[part]) currentLevel[part] = {};
                if (index === parts.length - 1) currentLevel[part] = null;
                else currentLevel = currentLevel[part];
            });
        });
        return structuredHeaders;
    }

    /**
     * Computes the maximum depth of the header structure.
     * @param headers - The structured header object.
     * @param depth - The current depth (used for recursion).
     * @returns The maximum depth found.
     */
    private getMaxDepth(headers: any, depth = 1): number {
        let maxDepth = depth;
        Object.values(headers).forEach((subHeaders) => {
            if (subHeaders !== null) {
                maxDepth = Math.max(maxDepth, this.getMaxDepth(subHeaders, depth + 1));
            }
        });
        return maxDepth;
    }

    /**
     * Creates the header row with hierarchical merging.
     * @param worksheet - The Excel worksheet.
     * @param headers - The structured header object.
     * @param rowIndex - The row index for the header (default: 1).
     * @param colIndex - The starting column index (default: 1).
     * @param maxDepth - The maximum depth of the headers.
     * @returns A mapping of column names to column indices.
     */
    private createHeaderRow(
        worksheet: ExcelJS.Worksheet,
        headers: any,
        rowIndex = 1,
        colIndex = 1,
        maxDepth = this.getMaxDepth(headers)
    ): Record<string, number> {
        let colSpan = 0;
        const columnMapping: Record<string, number> = {};

        Object.keys(headers).forEach((key) => {
            const startCol = colIndex + colSpan;
            let endCol = startCol;

            if (headers[key] !== null) {
                const subHeaderMapping = this.createHeaderRow(
                    worksheet,
                    headers[key],
                    rowIndex + 1,
                    startCol,
                    maxDepth
                );
                endCol = Math.max(...Object.values(subHeaderMapping));
                Object.assign(columnMapping, subHeaderMapping);
            } else {
                worksheet.mergeCells(rowIndex, startCol, maxDepth, startCol);
            }

            if (headers[key] !== null) {
                worksheet.mergeCells(rowIndex, startCol, rowIndex, endCol);
            }

            const cell = worksheet.getCell(rowIndex, startCol);
            cell.value = key;
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.font = { bold: true, color: { argb: '000000' }, size: 12 };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'bdb4b3' } };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
            };

            columnMapping[key] = startCol;
            colSpan += endCol - startCol + 1;
        });

        return columnMapping;
    }

    /**
     * Generates a mapping of full header paths to their respective column indices.
     * This is useful for determining which column an attribute belongs to, including nested attributes.
     * @param headers - The structured header object.
     * @param parentKey - The parent key used for recursion (default: '').
     * @returns A record mapping full header paths to column indices.
     */
    private generateFullPathMapping(headers: any, parentKey = ''): Record<string, number> {
        const fullPathMapping: Record<string, number> = {};
        let columnIndex = 1;

        const traverseHeaders = (currentHeaders: any, currentParentKey: string) => {
            const parentIndex = columnIndex;

            Object.keys(currentHeaders).forEach((key) => {
                const newKey = currentParentKey ? `${currentParentKey}.${key}` : key;

                if (currentHeaders[key] !== null && typeof currentHeaders[key] === 'object') {
                    fullPathMapping[newKey] = parentIndex;
                    traverseHeaders(currentHeaders[key], newKey);
                } else {
                    fullPathMapping[newKey] = columnIndex;
                    columnIndex++;
                }
            });
        };

        traverseHeaders(headers, parentKey);
        return fullPathMapping;
    }

    /**
     * Populates the worksheet with JSON data.
     * @param worksheet - The Excel worksheet.
     * @param data - The JSON data array.
     * @param headers - The set of headers.
     */
    private fillDataRows(worksheet: ExcelJS.Worksheet, data: any[], headers: Set<string>) {
        const flatHeaders = Array.from(headers);
        const rows = data.map((item) => flatHeaders.map((key) => this.getNestedValue(item, key)));
        worksheet.addRows(rows);
    }

    /**
     * Retrieves a nested value from an object using dot notation.
     * @param obj - The object to retrieve data from.
     * @param path - The dot-separated path string.
     * @returns The extracted value or an empty string if not found.
     */
    private getNestedValue(obj: any, path: string) {
        try {
            return path
                .split('.')
                .reduce((acc, part) => (acc && acc[part] !== undefined ? acc[part] : ''), obj);
        } catch (error) {
            console.error('Error extracting nested value:', error);
            return '';
        }
    }

    /**
     * Freezes specified columns in the worksheet.
     * @param worksheet - The Excel worksheet.
     * @param columnMapping - Mapping of column names to indices.
     * @param fullPathMapping - Mapping of full header paths to indices.
     * @param frozenColumns - List of columns to freeze.
     */
    private applyFreezing(
        worksheet: ExcelJS.Worksheet,
        columnMapping: Record<string, number>,
        fullPathMapping: Record<string, number>,
        frozenColumns: string[]
    ) {
        if (frozenColumns.length === 0) return;

        let maxFreezeCol = 0;
        const columnsToFreeze = new Set<number>();

        frozenColumns.forEach((columnName) => {
            const colIndex = columnMapping[columnName];

            if (colIndex !== undefined) {
                columnsToFreeze.add(colIndex);
                let lastCol = colIndex;

                Object.keys(fullPathMapping).forEach((key) => {
                    if (key.includes(columnName + '.')) {
                        lastCol = Math.max(lastCol, fullPathMapping[key]);
                    }
                });

                maxFreezeCol = Math.max(maxFreezeCol, lastCol);
            }
        });

        worksheet.views = [{ state: 'frozen', xSplit: maxFreezeCol, ySplit: 1 }];
    }

    private isUniformStructure(jsonData: any[]): boolean {
        if (!Array.isArray(jsonData) || jsonData.length === 0) {
            throw new Error('Invalid or empty JSON data');
        }

        const referenceStructure = this.getStructure(jsonData[0]); // Get structure from the first object

        return jsonData.every(
            (item) => JSON.stringify(this.getStructure(item)) === JSON.stringify(referenceStructure)
        );
    }

    private getStructure(obj: any): any {
        if (Array.isArray(obj)) {
            return obj.length > 0 ? [this.getStructure(obj[0])] : [];
        } else if (obj !== null && typeof obj === 'object') {
            return Object.fromEntries(
                Object.entries(obj).map(([key, value]) => [key, this.getStructure(value)])
            );
        }
        return typeof obj;
    }
}
