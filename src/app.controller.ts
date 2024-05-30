import {
  Body,
  Controller,
  Get,
  ParseIntPipe,
  Post,
  UploadedFiles,
  UseInterceptors,
} from '@nestjs/common';
import { FileFieldsInterceptor } from '@nestjs/platform-express';
import { AppService } from './app.service';
import { parse } from 'csv-parse/sync';
import * as moment from 'moment';
import 'moment/locale/vi';
moment.locale('vi');
import 'moment-timezone';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Get()
  getHello(): string {
    return this.appService.getHello();
  }

  @Post()
  @UseInterceptors(
    FileFieldsInterceptor([
      { name: 'rawData', maxCount: 1 },
      { name: 'idMapping', maxCount: 1 },
    ]),
  )
  async generateTimesheet(
    @UploadedFiles()
    files: {
      rawData?: Express.Multer.File[];
      idMapping?: Express.Multer.File[];
    },
    @Body('month', ParseIntPipe) month: number,
    @Body('year', ParseIntPipe) year: number,
  ) {
    const rawDataFile = files.rawData[0];
    const idMappingFile = files.idMapping[0];
    const rawData: any[] = parse(rawDataFile.buffer, {
      // columns: true,
      fromLine: 2,
      skipEmptyLines: true,
      encoding: 'utf-8',
      delimiter: ',',
      trim: true,
    });
    const idMapping: any[] = parse(idMappingFile.buffer, {
      // columns: true,
      fromLine: 2,
      skipEmptyLines: true,
      encoding: 'utf-8',
      delimiter: ',',
      trim: true,
    });
    console.log({ idMapping });
    console.log({ rawData });

    const employees = this.appService.processData(
      { month, year },
      rawData,
      idMapping,
    );

    return this.appService.generateSheet({ month, year }, employees);
  }
}
