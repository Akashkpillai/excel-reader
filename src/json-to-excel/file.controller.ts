// import { Controller, Post, UploadedFile, UseInterceptors } from '@nestjs/common';
// import { ApiOperation, ApiResponse, ApiConsumes, ApiBody } from '@nestjs/swagger';
// import { FileInterceptor } from '@nestjs/platform-express';
// import * as fs from 'fs';
// import { ExcelToJsonService } from './exceltojson.service';

// @Controller('file')
// export class FileController {
//     constructor(private readonly excelToJsonService: ExcelToJsonService) {}

//     @Post('import')
//     @ApiOperation({ summary: 'Upload Excel and Convert to JSON' })
//     @ApiConsumes('multipart/form-data') // Important for Swagger file uploads
//     @ApiBody({
//         schema: {
//             type: 'object',
//             properties: {
//                 file: {
//                     type: 'string',
//                     format: 'binary', // Tells Swagger to expect a file
//                 },
//             },
//         },
//     })
//     @ApiResponse({ status: 200, description: 'Returns JSON data extracted from Excel' })
//     @UseInterceptors(FileInterceptor('file'))
//     async convertExcelToJson(@UploadedFile() file: Express.Multer.File) {
//         const filePath = `./uploads/${file.originalname}`;
//         fs.writeFileSync(filePath, file.buffer);

//         const jsonData = await this.excelToJsonService.excelToJson(filePath);

//         fs.unlinkSync(filePath); // Remove file after processing
//         return jsonData;
//     }
// }
