import {
    Controller,
    Post,
    Body,
    Res,
    UploadedFile,
    UseInterceptors,
    Query,
    Get,
} from '@nestjs/common';
import { Response } from 'express';
import { FileInterceptor } from '@nestjs/platform-express';
import {
    ApiBody,
    ApiConsumes,
    ApiOperation,
    ApiQuery,
    ApiResponse,
    ApiTags,
} from '@nestjs/swagger';
import { ExcelService } from './excel.service';
import * as fs from 'fs';
import { HtmlTableService, TimezoneService } from './excel.service_old';
import { ExcelToJsonService } from './exceltojson.service';

class JsonInputDto {
    example: any[];
}

@ApiTags('excel')
@Controller('excel')
export class ExcelController {
    constructor(
        private readonly excelService: ExcelService,
        private readonly htmlService: HtmlTableService,
        private readonly timezoneService: TimezoneService,
        private readonly excelToJsonService: ExcelToJsonService
    ) {}

    @Post('export')
    @ApiOperation({ summary: 'Convert JSON to Excel' })
    @ApiResponse({ status: 200, description: 'Returns an Excel file' })
    @ApiBody({
        description: 'JSON data to be converted',
        type: JsonInputDto,
        examples: {
            example1: {
                value: [
                    {
                        id: 1,
                        name: 'John Doe',
                        contact: {
                            email: 'john@example.com',
                            phone: '123-456-7890',
                            social: {
                                twitter: '@johndoe',
                                linkedin: 'linkedin.com/in/johndoe',
                            },
                        },
                        address: {
                            home: {
                                city: 'New York',
                                country: 'USA',
                            },
                            work: {
                                city: 'San Francisco',
                                country: 'USA',
                            },
                        },
                    },
                    {
                        id: 2,
                        name: 'Jane Doe',
                        contact: {
                            email: 'jane@example.com',
                            phone: '987-654-3210',
                            social: {
                                twitter: '@janedoe',
                                linkedin: 'linkedin.com/in/janedoe',
                            },
                        },
                        address: {
                            home: {
                                city: 'London',
                                country: 'UK',
                            },
                            work: {
                                city: 'Berlin',
                                country: 'Germany',
                            },
                        },
                    },
                ],
            },
        },
    })
    @ApiQuery({
        name: 'frozenColumns',
        required: false, // Mark as optional in Swagger
        description: 'Comma-separated list of columns to freeze (e.g., "name,contact,email")',
        example: 'name,contact.email',
    })
    async convertJsonToExcel(
        @Body() jsonData: any[],
        @Res() res: Response,
        @Query('frozenColumns') frozenColumnsQuery?: string
    ) {
        const frozenColumns = frozenColumnsQuery ? frozenColumnsQuery.split(',') : [];
        const filePath = './data.xlsx';

        await this.excelService.jsonToExcel(jsonData, filePath, frozenColumns);
        res.download(filePath, 'data.xlsx', () => fs.unlinkSync(filePath));
    }
    @Post('export-html')
    @ApiOperation({ summary: 'Convert JSON to HTML' })
    @ApiResponse({ status: 200, description: 'Returns an HTML file' })
    @ApiBody({
        description: 'JSON data to be converted',
        type: Object,
        examples: {
            example1: {
                value: [
                    {
                        id: 1,
                        name: 'John Doe',
                        contact: {
                            email: 'john@example.com',
                            phone: '123-456-7890',
                            social: {
                                twitter: '@johndoe',
                                linkedin: 'linkedin.com/in/johndoe',
                            },
                        },
                        address: {
                            home: {
                                city: 'New York',
                                country: 'USA',
                            },
                            work: {
                                city: 'San Francisco',
                                country: 'USA',
                            },
                        },
                    },
                    {
                        id: 2,
                        name: 'Jane Doe',
                        contact: {
                            email: 'jane@example.com',
                            phone: '987-654-3210',
                            social: {
                                twitter: '@janedoe',
                                linkedin: 'linkedin.com/in/janedoe',
                            },
                        },
                        address: {
                            home: {
                                city: 'London',
                                country: 'UK',
                            },
                            work: {
                                city: 'Berlin',
                                country: 'Germany',
                            },
                        },
                    },
                ],
            },
        },
    })
    async convertJsonToHtml(@Body() jsonData: any[]) {
        const htmlContent = this.htmlService.jsonToHtml(jsonData);

        console.log('Generated HTML:', htmlContent); // Debugging

        if (!htmlContent) {
            //return res.status(500).send('Error generating HTML');
        }

        //res.setHeader('Content-Type', 'text/html');
        return htmlContent;
    }

    // @Post('import')
    // @ApiOperation({ summary: 'Convert Excel to JSON' })
    // @ApiConsumes('multipart/form-data')
    // @ApiResponse({ status: 200, description: 'Returns JSON data extracted from Excel' })
    // @UseInterceptors(FileInterceptor('file'))
    // async convertExcelToJson(@UploadedFile() file: Express.Multer.File) {
    //     const filePath = `./uploads/${file.originalname}`;
    //     fs.writeFileSync(filePath, file.buffer);
    //     const jsonData = await this.excelToJsonService.excelToJson(filePath);
    //     fs.unlinkSync(filePath);
    //     return jsonData;
    // }
    @Post('import')
    @ApiOperation({ summary: 'Upload Excel and Convert to JSON' })
    @ApiConsumes('multipart/form-data') // Important for Swagger file uploads
    @ApiBody({
        schema: {
            type: 'object',
            properties: {
                file: {
                    type: 'string',
                    format: 'binary', // Tells Swagger to expect a file
                },
            },
        },
    })
    @ApiResponse({ status: 200, description: 'Returns JSON data extracted from Excel' })
    @UseInterceptors(FileInterceptor('file'))
    async convertExcelToJson(@UploadedFile() file: Express.Multer.File) {
        const filePath = `./uploads/${file.originalname}`;
        fs.writeFileSync(filePath, file.buffer);

        const jsonData = await this.excelToJsonService.excelToJson(filePath);
        console.log(JSON.stringify(jsonData, null, 2));
        fs.unlinkSync(filePath); // Remove file after processing
        return jsonData;
    }

    @Get('test')
    @ApiOperation({ summary: 'Testing' })
    @ApiResponse({ status: 200, description: 'Returns JSON data extracted from Excel' })
    async test(@Query('utcDate') utcDate: string, @Query('timezone') timezone: string) {
        return this.timezoneService.convertUtcToTimezone(utcDate, timezone, 'en-US', {
            hour12: true,
        });
    }
}
