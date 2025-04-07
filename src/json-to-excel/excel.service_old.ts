import { Injectable } from '@nestjs/common';
import * as ExcelJS from 'exceljs';

@Injectable()
export class ExcelService {
    // Convert JSON to Excel
    async jsonToExcel(jsonData: any[], filePath: string): Promise<string> {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Data');

        const headers = this.extractHeaders(jsonData);
        const structuredHeaders = this.structureHeaders(headers);
        this.createHeaderRow(worksheet, structuredHeaders);
        this.fillDataRows(worksheet, jsonData, headers);

        await workbook.xlsx.writeFile(filePath);
        return filePath;
    }

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

    // Structure headers for merging
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
    private getMaxDepth(headers: any, depth = 1): number {
        let maxDepth = depth;
        Object.values(headers).forEach((subHeaders) => {
            if (subHeaders !== null) {
                maxDepth = Math.max(maxDepth, this.getMaxDepth(subHeaders, depth + 1));
            }
        });
        return maxDepth;
    }

    // Create header row with merged columns
    private createHeaderRow(
        worksheet: ExcelJS.Worksheet,
        headers: any,
        rowIndex = 1,
        colIndex = 1,
        maxDepth = this.getMaxDepth(headers)
    ): number {
        let colSpan = 0;
        Object.keys(headers).forEach((key) => {
            const startCol = colIndex + colSpan;
            let endCol = startCol;
            let depth = rowIndex;

            if (headers[key] !== null) {
                // Recursively create sub-headers and get the last column
                endCol = this.createHeaderRow(
                    worksheet,
                    headers[key],
                    rowIndex + 1,
                    startCol,
                    maxDepth
                );
            } else {
                // If no sub-headers, merge downward to max depth
                worksheet.mergeCells(rowIndex, startCol, maxDepth, startCol);
                // eslint-disable-next-line @typescript-eslint/no-unused-vars
                depth = maxDepth;
            }

            // Merge parent headers over their sub-headers
            if (headers[key] !== null) {
                worksheet.mergeCells(rowIndex, startCol, rowIndex, endCol);
            }

            const cell = worksheet.getCell(rowIndex, startCol);
            cell.value = key;

            // Apply styling
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.font = { bold: true, color: { argb: '000000' }, size: 12 };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'bdb4b3' },
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
            };

            colSpan += endCol - startCol + 1;
        });

        return colIndex + colSpan - 1;
    }

    // Fill data rows
    private fillDataRows(worksheet: ExcelJS.Worksheet, data: any[], headers: Set<string>) {
        const flatHeaders = Array.from(headers);
        const rows = data.map((item) => flatHeaders.map((key) => this.getNestedValue(item, key)));
        worksheet.addRows(rows);
    }

    // Get nested value
    private getNestedValue(obj: any, path: string) {
        return path.split('.').reduce((acc, part) => (acc && acc[part] ? acc[part] : ''), obj);
    }
}

export class TimezoneService {
    async convertUtcToTimezone(
        utcDate: string,
        timezone: string,
        locale: string = 'en-US',
        options?: Intl.DateTimeFormatOptions
    ): Promise<string> {
        const date = new Date(utcDate);

        // Default options (can be overridden by user)
        const defaultOptions: Intl.DateTimeFormatOptions = {
            timeZone: timezone,
            year: options?.year !== undefined ? options.year : 'numeric',
            month: options?.month !== undefined ? options.month : '2-digit',
            day: options?.day !== undefined ? options.day : '2-digit',
            hour: options?.hour !== undefined ? options.hour : '2-digit',
            minute: options?.minute !== undefined ? options.minute : '2-digit',
            second: options?.second !== undefined ? options.second : '2-digit',
            hour12: options?.hour12 !== undefined ? options.hour12 : false,
        };

        return new Intl.DateTimeFormat(locale, defaultOptions).format(date);
    }
}

export class HtmlTableService {
    async jsonToHtml(jsonData: any[]): Promise<string> {
        if (!jsonData || jsonData.length === 0) {
            return '<p>No data available</p>';
        }
        if (!this.isUniformStructure(jsonData)) {
            throw new Error('Invalid or Not Same Array JSON data');
        }

        const headers = this.extractHeaders(jsonData);
        const structuredHeaders = this.structureHeaders(headers);
        const maxDepth = this.getMaxDepth(structuredHeaders);

        let html = '<table border="1" style="border-collapse: collapse; width: 100%;">';
        html += this.createHeaderRows(
            structuredHeaders,
            maxDepth,
            'black',
            'gray',
            '2px solid black'
        );
        html += this.createBodyRows(jsonData, headers, 'center');
        html += '</table>';
        console.log(html);
        const data = html.toString();
        return data;
    }

    private extractHeaders(data: any[], parentKey = ''): Set<string> {
        const headers = new Set<string>();

        data.forEach((item) => {
            Object.entries(item).forEach(([key, value]) => {
                const newKey = parentKey ? `${parentKey}.${key}` : key;
                if (typeof value === 'object' && value !== null) {
                    this.extractHeaders([value], newKey).forEach((sub) => headers.add(sub));
                } else {
                    headers.add(newKey);
                }
            });
        });

        return headers;
    }

    private structureHeaders(headers: Set<string>): Record<string, any> {
        const structuredHeaders: Record<string, any> = {};

        headers.forEach((path) => {
            const parts = path.split('.');
            let currentLevel = structuredHeaders;

            parts.forEach((part, index) => {
                if (!currentLevel[part]) {
                    currentLevel[part] = {};
                }
                if (index === parts.length - 1) {
                    currentLevel[part] = null;
                } else {
                    currentLevel = currentLevel[part];
                }
            });
        });

        return structuredHeaders;
    }

    private getMaxDepth(headers: Record<string, any>, depth = 1): number {
        let maxDepth = depth;

        Object.values(headers).forEach((subHeaders) => {
            if (subHeaders !== null) {
                maxDepth = Math.max(maxDepth, this.getMaxDepth(subHeaders, depth + 1));
            }
        });

        return maxDepth;
    }

    // private createHeaderRows(headers: Record<string, any>, maxDepth: number): string {
    //     const rows: string[][] = Array.from({ length: maxDepth }, () => []);

    //     const processHeader = (subHeaders: Record<string, any>, depth: number): number => {
    //         let span = 0;

    //         Object.keys(subHeaders).forEach((key) => {
    //             const colSpan = subHeaders[key] !== null ? this.getColumnSpan(subHeaders[key]) : 1;
    //             const rowSpan = subHeaders[key] === null ? maxDepth - depth : 1;

    //             rows[depth].push(`<th colspan="${colSpan}" rowspan="${rowSpan}">${key}</th>`);

    //             if (subHeaders[key] !== null) {
    //                 span += processHeader(subHeaders[key], depth + 1);
    //             } else {
    //                 span += 1;
    //             }
    //         });

    //         return span;
    //     };

    //     processHeader(headers, 0);

    //     return '<thead>' + rows.map((row) => `<tr>${row.join('')}</tr>`).join('') + '</thead>';
    // }
    private createHeaderRows(
        headers: Record<string, any>,
        maxDepth: number,
        textColor: string,
        backgroundColor: string,
        borderStyle: string
    ): string {
        const rows: string[][] = Array.from({ length: maxDepth }, () => []);

        const processHeader = (subHeaders: Record<string, any>, depth: number): number => {
            let span = 0;

            Object.keys(subHeaders).forEach((key) => {
                const colSpan = subHeaders[key] !== null ? this.getColumnSpan(subHeaders[key]) : 1;
                const rowSpan = subHeaders[key] === null ? maxDepth - depth : 1;

                // Apply dynamic colors
                rows[depth].push(
                    `<th colspan="${colSpan}" rowspan="${rowSpan}" 
                     style="color: ${textColor}; background-color: ${backgroundColor}; 
                            border: ${borderStyle}; padding: 8px; text-align: center;">
                     ${key}
                 </th>`
                );

                if (subHeaders[key] !== null) {
                    span += processHeader(subHeaders[key], depth + 1);
                } else {
                    span += 1;
                }
            });

            return span;
        };

        processHeader(headers, 0);

        return '<thead>' + rows.map((row) => `<tr>${row.join('')}</tr>`).join('') + '</thead>';
    }

    private getColumnSpan(headers: Record<string, any>): number {
        let span = 0;
        Object.values(headers).forEach((subHeaders) => {
            span += subHeaders === null ? 1 : this.getColumnSpan(subHeaders);
        });
        return span;
    }

    private createBodyRows(data: any[], headers: Set<string>, textAlign: string): string {
        const flatHeaders = Array.from(headers);
        let rows = '<tbody>';

        data.forEach((item) => {
            rows += '<tr>';
            flatHeaders.forEach((key) => {
                rows += `<td style="text-align: ${textAlign}; padding: 8px; border: 1px solid black;">${this.getNestedValue(item, key)}</td>`;
            });
            rows += '</tr>';
        });

        rows += '</tbody>';
        return rows;
    }

    private getNestedValue(obj: any, path: string): string {
        return path
            .split('.')
            .reduce((acc, part) => (acc && acc[part] !== undefined ? acc[part] : ''), obj);
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
