import { Controller, Post, UploadedFile, UseInterceptors } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { PdfService } from './pdfreader.service';
import * as fs from 'fs';
import { ApiBody, ApiConsumes, ApiOperation, ApiResponse } from '@nestjs/swagger';

@Controller('pdf')
export class PdfController {
    constructor(private readonly pdfService: PdfService) {}

    @Post('import')
    @ApiOperation({ summary: 'Upload Pdf and Analysis Data' })
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
    @ApiResponse({ status: 200, description: 'Returns PDF analysis Data' })
    @UseInterceptors(FileInterceptor('file'))
    async analyzePdf(@UploadedFile() file: Express.Multer.File) {
        const filePath = `./uploads/${file.originalname}`;
        fs.writeFileSync(filePath, file.buffer);

        //const jsonData = await this.pdfService.analyzePdf(filePath);
        const analysis = await this.pdfService.analyzePdf(filePath);
        // return { message: 'PDF analysis completed', analysis };
        console.log('data', analysis);
        fs.unlinkSync(filePath); // Remove file after processing
        return { message: 'PDF analysis completed', analysis };
    }
}
