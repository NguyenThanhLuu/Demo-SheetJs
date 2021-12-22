import { Injectable } from "@angular/core";
import * as FileSaver from "file-saver";
import * as XLSX from "xlsx";

const EXCEL_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8";
const EXCEL_EXTENSION = ".xlsx";

@Injectable()
export class ExcelService {
  constructor() {}

  public exportAsExcelFile(
    dataInput: any[],
    excelSheetName: string,
    excelFileName: string
  ): void {
    /* Make the worksheet */
    let ws = XLSX.utils.json_to_sheet(dataInput);
    // Merge cell
    let merge = [
      { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } }, // merge a2:a3 (hàng 1 cột 0 và hàng 2 cột 0)
      { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } },
      { s: { r: 0, c: 1 }, e: { r: 0, c: 3 } },
    ];
    ws["!merges"] = merge;
    // Width Cell
    let wscols = [{ width: 40 }, { width: 7 }, { width: 10 }, { width: 20 }]; // độ rộng cột tính từ cột đầu tiên sang phải
    ws["!cols"] = wscols;
    let wsrows = [{ hpt: 16 }, { hpx: 16 }]; // pt là point
    ws["!rows"] = wsrows;
    // Color fill cell

    /* add to workbook */
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, excelSheetName);

    /* generate an XLSX file */
    XLSX.writeFile(wb, excelFileName + new Date().getTime() + EXCEL_EXTENSION);
  }
}

// ĐÂY CŨNG TẠO MỘT WB ĐƯỢC
// public exportAsExcelFile(dataInput: any[], excelFileName: string): void {
//     const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(dataInput);
//     // Merge cell
//     const merge = [
//       { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } }, // merge a2:a3 (hàng 1 cột 0 và hàng 2 cột 0)
//       { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } },
//       { s: { r: 0, c: 1 }, e: { r: 0, c: 3 } },
//     ];
//     worksheet["!merges"] = merge;
//     // Fill cell
//     console.log("worksheet", worksheet);
//     const workbook: XLSX.WorkBook = {
//       Sheets: { data: worksheet },
//       SheetNames: ["data"],
//     };
//     const excelBuffer: any = XLSX.write(workbook, {
//       bookType: "xlsx",
//       type: "array",
//     });
//     this.saveAsExcelFile(excelBuffer, excelFileName);
//   }

//   private saveAsExcelFile(buffer: any, fileName: string): void {
//     const data: Blob = new Blob([buffer], {
//       type: EXCEL_TYPE,
//     });
//     FileSaver.saveAs(
//       data,
//       fileName + "_export_" + new Date().getTime() + EXCEL_EXTENSION
//     );
//   }
