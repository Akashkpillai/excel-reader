import { Injectable } from '@nestjs/common';
import * as pdfParse from 'pdf-parse';
import * as pdfTableExtractor from 'pdf-table-extractor';
import PDF2Pic from 'pdf2pic';
import { existsSync, readdirSync } from 'fs';
import { PdfReader, TableParser } from 'pdfreader';
import * as ExcelJS from 'exceljs';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import fs from 'fs';
//import pdfParse from 'pdf-parse';

@Injectable()
export class PdfService {
    // Extracts text and counts paragraphs
    async extractText(filePath: string): Promise<{ text: string; paragraphCount: number }> {
        const dataBuffer = await pdfParse(filePath);
        const text = dataBuffer.text;
        const paragraphs = text.split(/\n\s*\n/).filter((p) => p.trim().length > 0);
        return { text, paragraphCount: paragraphs.length };
    }

    // Detects tables in the PDF
    async extractTables(
        filePath: string
    ): Promise<{ hasTables: boolean; tableCount: number; tables: any[] }> {
        return new Promise((resolve, reject) => {
            pdfTableExtractor(
                filePath,
                (result) => {
                    const tablePages = result.pageTables
                        .filter((page) => page.tables.length > 0)
                        .map((page) => ({
                            pageNumber: page.page,
                            tables: page.tables, // Each table as an array of rows/columns
                        }));

                    resolve({
                        hasTables: tablePages.length > 0,
                        tableCount: tablePages.length,
                        tables: tablePages,
                    });
                },
                (error) => reject(error)
            );
        });
    }
    async extractTables4(filePath) {
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        const fs = require('fs');
        const dataBuffer = fs.readFileSync(filePath);

        try {
            const data = await pdfParse(dataBuffer);
            console.log(data.text); // Extracted text from PDF
        } catch (err) {
            console.error('Error extracting PDF text:', err);
        }
    }
    async extractTables5(filePath: string): Promise<{ hasTables: boolean; tables: string[][] }> {
        return new Promise((resolve, reject) => {
            // eslint-disable-next-line @typescript-eslint/no-var-requires
            const fs = require('fs');
            const dataBuffer = fs.readFileSync(filePath);

            pdfParse(dataBuffer)
                .then((data) => {
                    const text = data.text;
                    const tables = this.parseTextToTable(text);

                    resolve({
                        hasTables: tables.length > 0,
                        tables,
                    });
                })
                .catch((err) => {
                    reject(`Error extracting PDF text: ${err}`);
                });
        });
    }

    // Function to parse extracted text into a table format
    private parseTextToTable(text: string): string[][] {
        const rows = text
            .split('\n')
            .map((row) => row.trim())
            .filter((row) => row.length > 0);

        // Convert text rows into a table format (simple column splitting)
        const tableData: string[][] = rows.map((row) => row.split(/\s{2,}/)); // Split by multiple spaces
        console.log('table', tableData);
        return tableData;
    }
    async extractTables1(filePath: string): Promise<any[]> {
        return new Promise((resolve, reject) => {
            // eslint-disable-next-line @typescript-eslint/no-var-requires
            const PDFParser = require('pdf2json'); // ✅ Correct way to import

            const pdfParser = new PDFParser();

            pdfParser.on('pdfParser_dataError', (errData: any) => {
                reject(errData.parserError);
            });

            pdfParser.on('pdfParser_dataReady', (pdfData: any) => {
                if (!pdfData.formImage || !pdfData.formImage.Pages) {
                    return reject(new Error('Invalid PDF structure'));
                }

                const extractedData: any[] = [];

                pdfData.formImage.Pages.forEach((page: any, pageIndex: number) => {
                    const rows: Record<number, string[]> = {};

                    page.Texts.forEach((text: any) => {
                        const y = Math.floor(text.y * 10); // Normalize position
                        const decodedText = decodeURIComponent(text.R[0].T); // Extract text

                        rows[y] = rows[y] || [];
                        rows[y].push(decodedText);
                    });

                    // Convert rows object into an ordered array
                    const table = Object.keys(rows)
                        .sort((a, b) => parseFloat(a) - parseFloat(b))
                        .map((y) => rows[y]);

                    extractedData.push({
                        pageNumber: pageIndex + 1,
                        tables: table,
                    });
                });

                resolve(extractedData);
            });

            pdfParser.loadPDF(filePath);
        });
    }
    // Extracts images by converting PDF to images and counting non-empty images
    async extractImages(filePath: string): Promise<{ hasImages: boolean; imageCount: number }> {
        const outputDir = './pdf-images';
        const converter = PDF2Pic.fromPath(filePath, {
            density: 100,
            saveFilename: 'output',
            savePath: outputDir,
            format: 'png',
        });

        // Convert PDF pages to images
        await converter.bulk(-1);

        // Count the number of generated images
        if (existsSync(outputDir)) {
            const images = readdirSync(outputDir).filter((file) => file.endsWith('.png'));
            return { hasImages: images.length > 0, imageCount: images.length };
        }

        return { hasImages: false, imageCount: 0 };
    }
    private processTables(tables: any[]): any[] {
        return tables.map((page) => ({
            pageNumber: page.pageNumber,
            tables: page.tables.map((table) => {
                const headers = table[0].map((h, i) => h.trim() || `Column${i + 1}`);
                return table
                    .slice(1)
                    .map((row) =>
                        Object.fromEntries(row.map((value, index) => [headers[index], value]))
                    );
            }),
        }));
    }
    async extractTables2(
        filePath: string
    ): Promise<{ hasTables: boolean; tableCount: number; tables: any[] }> {
        return new Promise((resolve, reject) => {
            const tables: any[] = [];
            const rows: Record<number, { x: number; text: string }[]> = {};
            const pdfReader = new PdfReader();

            pdfReader.parseFileItems(filePath, (err, item) => {
                if (err) {
                    return reject(err);
                }

                if (!item) {
                    // After parsing is complete, reconstruct table
                    const table = Object.keys(rows)
                        .sort((a, b) => parseFloat(a) - parseFloat(b)) // Sort rows by Y position
                        .map((y) => {
                            // Sort words in a row by X position and join adjacent characters
                            const words = [];
                            let currentWord = '';
                            let prevX = null;

                            rows[y]
                                .sort((a, b) => a.x - b.x) // Sort words by X position
                                .forEach(({ x, text }) => {
                                    if (prevX !== null && x - prevX > 0.5) {
                                        words.push(currentWord.trim()); // Push completed word
                                        currentWord = text;
                                    } else {
                                        currentWord += text; // Merge characters into a word
                                    }
                                    prevX = x;
                                });

                            words.push(currentWord.trim()); // Push last word in row
                            return words;
                        });

                    if (table.length > 0) {
                        tables.push({ pageNumber: 1, tables: table });
                    }

                    return resolve({
                        hasTables: tables.length > 0,
                        tableCount: tables.length,
                        tables: tables,
                    });
                }

                if (item.text) {
                    const y = Math.floor(item['y'] * 10); // Normalize Y position
                    rows[y] = rows[y] || [];
                    rows[y].push({ x: item['x'], text: item.text }); // Store X position and text
                }
            });
        });
    }
    async extractTables3(filePath: string): Promise<Record<number, string>> {
        return new Promise((resolve, reject) => {
            const filename = filePath;
            const nbCols = 5;
            const cellPadding = 20;
            const columnQuantitizer = (item) => (parseFloat(item.x) >= 10 ? 1 : 0);

            const padColumns = (array, nb) =>
                // eslint-disable-next-line prefer-spread
                Array.apply(null, { length: nb }).map((val, i) => array[i] || []);

            const mergeCells = (cells) => (cells || []).map((cell) => cell.text).join('');

            const formatMergedCell = (mergedCell) =>
                mergedCell.substr(0, cellPadding).padEnd(cellPadding, ' ');

            const renderMatrix = (matrix) =>
                (matrix || [])
                    .map(
                        (row) =>
                            '| ' +
                            padColumns(row, nbCols)
                                .map(mergeCells)
                                .map(formatMergedCell)
                                .join(' | ') +
                            ' |'
                    )
                    .join('\n');

            let table = new TableParser();
            const pageData: Record<number, string> = {};
            let currentPage = 1;

            new PdfReader().parseFileItems(filename, function (err, item) {
                if (err) {
                    console.error('Error processing PDF:', err);
                    reject(err);
                } else if (!item || item.page) {
                    // When a new page starts or end of the file is reached
                    if (table.getMatrix().length > 0) {
                        pageData[currentPage] = renderMatrix(table.getMatrix());
                    }
                    if (!item) {
                        resolve(pageData); // Resolve when the PDF parsing is complete
                    } else {
                        currentPage = item.page;
                        table = new TableParser(); // Reset table for next page
                    }
                } else if (item.text) {
                    const x: number = item['x'];
                    const y: number = item['y'];
                    table.processItem(
                        {
                            x,
                            y,
                            sw: 0,
                            w: 0,
                            A: '',
                            clr: 0,
                            R: [],
                            text: item.text,
                        },
                        columnQuantitizer(item)
                    );
                }
            });
        });
    }
    async extractPdfData(filePath: string): Promise<any> {
        return new Promise((resolve, reject) => {
            const reader = new PdfReader();
            const pages: Record<number, string[]> = {};
            let currentPage = 1;

            reader.parseFileItems(filePath, (err, item) => {
                if (err) {
                    console.error('Error reading PDF:', err);
                    reject(err);
                } else if (!item) {
                    resolve(pages); // Resolve when parsing is complete
                } else if (item.page) {
                    currentPage = item.page;
                    pages[currentPage] = []; // Initialize new page
                } else if (item.text) {
                    pages[currentPage].push(item.text);
                }
            });
        });
    }
    async extractTablesToExcel(filePath: string, outputExcel: string): Promise<void> {
        return new Promise((resolve, reject) => {
            pdfTableExtractor(
                filePath,
                async (result) => {
                    const workbook = new ExcelJS.Workbook();
                    const tablePages = result.pageTables.filter((page) => page.tables.length > 0);

                    if (tablePages.length === 0) {
                        console.log('No tables found in the PDF.');
                        return resolve();
                    }

                    // Loop through each page and create a sheet
                    tablePages.forEach((page) => {
                        const sheet = workbook.addWorksheet(`Page ${page.page}`);

                        console.log(`Processing Page ${page.page}...`);

                        page.tables.forEach((table, tableIndex) => {
                            console.log(`Table ${tableIndex + 1}:`, table);

                            // Ensure table is structured as rows
                            const rowSize = table.length; // Adjust based on your PDF table structure
                            const structuredRows = [];

                            for (let i = 0; i < table.length; i += rowSize) {
                                const row = table.slice(i, i + rowSize);
                                structuredRows.push(row);
                            }

                            structuredRows.forEach((row) => {
                                if (Array.isArray(row)) {
                                    sheet.addRow(
                                        row.filter(
                                            (cell) => typeof cell === 'string' //&& cell.trim() !== ''
                                        )
                                    );
                                }
                            });

                            sheet.addRow([]); // Add an empty row to separate tables
                        });
                    });

                    // Write to an Excel file
                    await workbook.xlsx.writeFile(outputExcel);
                    console.log(`Excel file saved as ${outputExcel}`);
                    resolve();
                },
                (error) => reject(error)
            );
        });
    }

    // async extractTablesFromPDF(

    //     pdfPath
    // ): Promise<{ hasTables: boolean; tableCount: number; tables: any[] }> {
    //     try {
    //         // eslint-disable-next-line @typescript-eslint/no-var-requires
    //         const fs = require('fs');
    //         // eslint-disable-next-line @typescript-eslint/no-var-requires
    //         //const getDocument = require('pdfjs-dist');
    //         // Read the PDF file as a buffer
    //         const pdfBuffer = fs.readFileSync(pdfPath);
    //         const pdfData = await getDocument({ data: pdfBuffer }).promise;

    //         const extractedPages = [];

    //         for (let pageNum = 1; pageNum <= pdfData.numPages; pageNum++) {
    //             const page = await pdfData.getPage(pageNum);
    //             const textContent = await page.getTextContent();

    //             // Extract text and positions
    //             const items = textContent.items.map((item) => {
    //                 const textItem = item as any as { str: string; transform: number[] };
    //                 return {
    //                     text: textItem.str,
    //                     x: textItem.transform[4], // X-coordinate
    //                     y: textItem.transform[5], // Y-coordinate
    //                 };
    //             });

    //             // Sort by Y-coordinate to group rows together
    //             items.sort((a, b) => b.y - a.y || a.x - b.x);

    //             // Convert sorted text into a table
    //             const table = [];
    //             let currentRow = [];
    //             let lastY = null;

    //             items.forEach((item) => {
    //                 if (lastY !== null && Math.abs(item.y - lastY) > 5) {
    //                     table.push(currentRow);
    //                     currentRow = [];
    //                 }
    //                 currentRow.push(item.text);
    //                 lastY = item.y;
    //             });

    //             if (currentRow.length) table.push(currentRow);

    //             extractedPages.push({ pageNumber: pageNum, table });
    //         }

    //         return {
    //             hasTables: extractedPages.length > 0,
    //             tableCount: extractedPages.length,
    //             tables: extractedPages,
    //         };
    //     } catch (error) {
    //         console.error('Error extracting tables:', error);
    //     }
    // }

    // Combines all detections into a single function
    async extractTables6(filePath: string): Promise<{ hasTables: boolean; tables: string[][] }> {
        return new Promise((resolve, reject) => {
            // eslint-disable-next-line @typescript-eslint/no-var-requires
            const fs = require('fs');
            const dataBuffer = fs.readFileSync(filePath);

            pdfParse(dataBuffer)
                .then((data) => {
                    const text = data.text;
                    console.log(text);
                    //const tables = this.formatTableData(this.parseTextToTable2(text));
                    const tables = this.parseTextToTable2(text);
                    resolve({
                        hasTables: tables.length > 0,
                        tables,
                    });
                })
                .catch((err) => {
                    reject(`Error extracting PDF text: ${err}`);
                });
        });
    }

    // Function to parse extracted text into a table format
    private parseTextToTable2(text: string): string[][] {
        const rows = text
            .split('\n')
            .map((row) => row.trim())
            .filter((row) => row.length > 0);
        console.log('raww', rows);
        // Convert each row into properly formatted data
        return rows.map((row) => this.splitRow(row));
    }

    // Function to split rows correctly
    private splitRow(row: string): string[] {
        const data = row
            .replace(/([A-Za-z]+)(\d+)/g, '$1 $2') // Add space between words and numbers
            .replace(/(\d+)([A-Za-z]+)/g, '$1 $2') // Add space between numbers and words
            .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2') // Fix uppercase words merging
            .replace(/([a-z])([A-Z])/g, '$1 $2') // Split words where lowercase meets uppercase
            .split(/\s{2,}|\t|,/) // Split by multiple spaces, tabs, or commas
            .map((cell) => cell.trim()); // Trim spaces
        console.log('data', data);

        return data;
    }
    async analyzePdf(filePath: string) {
        //const textData = await this.extractText(filePath);
        console.log('textData');
        const tableData = await this.extractTables3(filePath);
        //const tableData = await this.extractTablesToExcel(filePath, './uploads/output1.xlsx');
        console.log('tableData', tableData);
        //const result = await this.processTables(tableData.tables);
        //console.log('result', result);
        // const imageData = await this.extractImages(filePath);
        // console.log('imageData', imageData);

        return {
            // text: textData.text,
            // paragraphCount: textData.paragraphCount,
            hasTables: tableData,
            tableCount: tableData,
            //hasImages: imageData.hasImages,
            // imageCount: imageData.imageCount,
        };
    }
}
