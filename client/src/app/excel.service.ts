import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
import { DatePipe } from '../../node_modules/@angular/common';
@Injectable({
  providedIn: 'root'
})
export class ExcelService {
  constructor(private datePipe: DatePipe) {
  }
  exportAsExcelFile(json: any, excelFileName: string, headersArray: any): void {

    const data = json;
    // Create workbook and worksheet
    const workbook = new Workbook();
    // Add sheet in workbook
    const worksheet = workbook.addWorksheet(excelFileName);

    // Add Data and Conditional Formatting
    data.forEach((element) => {
      const eachRow = [];
      headersArray.forEach((header, index) => {
        eachRow.push(element[header.value]);
        worksheet.getColumn(index + 1).width = header.width;
      });
      if (element.header) {
        // Add Header Row
        const headerRow = worksheet.addRow(eachRow);
        // Cell Style
        headerRow.eachCell((cell, number) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' },
            bgColor: { argb: 'FF0000FF' }
          };
          cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
        });
      }
    });

    worksheet.addRow([]);

    // Generate Excel File with given name
    workbook.xlsx.writeBuffer().then((excelData) => {
      const blob = new Blob([excelData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, excelFileName);
    });
  }
}
