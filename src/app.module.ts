import { Module } from '@nestjs/common';
import { ExcelController } from './json-to-excel/excel.controller';
import { ExcelService } from './json-to-excel/excel.service';
import { HtmlTableService, TimezoneService } from './json-to-excel/excel.service_old';
import { ExcelToJsonService } from './json-to-excel/exceltojson.service';
import { PdfController } from './pdfreader/pdfreader.controller';
import { PdfService } from './pdfreader/pdfreader.service';

@Module({
    imports: [],
    controllers: [ExcelController, PdfController],
    providers: [ExcelService, HtmlTableService, TimezoneService, ExcelToJsonService, PdfService],
})
export class AppModule {}
