import {Component} from '@angular/core';
import * as Excel from 'exceljs';
//import * as Excel from 'exceljs/dist/exceljs.min.js';
import * as FileSaver from 'file-saver';

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'PruebasExcel';

  public generateExcel() {
    const workbook = new Excel.Workbook();
    workbook.creator = 'Me';
    workbook.lastModifiedBy = 'Her';
    workbook.created = new Date(1985, 8, 30);
    workbook.modified = new Date();
    workbook.lastPrinted = new Date(2016, 9, 27);

    const worksheet = workbook.addWorksheet('My Sheet', {properties: {tabColor: {argb: 'FFC0000'}}});

    // Title
    const title = 'Excel file example';

    // Add Row
    worksheet.addRow([title]);

    // Set Row 2 to Comic Sans.
    worksheet.getRow(1).font = {name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true};

    // add a table to a sheet
    worksheet.addTable({
      name: 'MyTable',
      ref: 'C3',
      headerRow: true,
      totalsRow: true,
      style: {
        theme: 'TableStyleLight1',
        showRowStripes: true,
      },
      columns: [
        {name: 'Date', totalsRowLabel: 'Totals:', filterButton: true},
        {name: 'Amount', totalsRowFunction: 'sum', filterButton: false},
      ],
      rows: [
        [new Date('2019-07-20'), 70.10],
        [new Date('2019-07-21'), 70.60],
        [new Date('2019-07-22'), 70.10],
      ],
    });

    // Generate Excel File
    workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], {type: EXCEL_TYPE});
      // Given name
      FileSaver.saveAs(blob, 'download.xlsx');
    });
  }
}
