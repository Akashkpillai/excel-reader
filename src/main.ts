import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { DocumentBuilder, SwaggerModule } from '@nestjs/swagger';

async function bootstrap() {
    const app = await NestFactory.create(AppModule);

    // Swagger Configuration
    const config = new DocumentBuilder()
        .setTitle('JSON ↔ Excel API')
        .setDescription('API for converting JSON to Excel and vice versa')
        .setVersion('1.0')
        .addTag('excel')
        .build();

    const document = SwaggerModule.createDocument(app, config);
    SwaggerModule.setup('api', app, document);

    await app.listen(3050);
}
bootstrap();
