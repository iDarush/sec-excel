"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const xlsx_1 = require("xlsx");
class ExcelFile {
    /**
     * Create an instance of ExcelFile
     * @param filePath Path to .xlsx file
     */
    constructor(filePath) {
        /**
         * Get worksheet cell object
         * @param sheetName Workbook sheet
         * @param row Row number (starts from 1)
         * @param column Column name (A, B, etc)
         */
        this.getCell = (sheetName, column, row) => {
            const index = this.workbook.SheetNames.find((s) => s === sheetName);
            if (!index) {
                throw new Error(`Worksheet ${sheetName} not found within ${this.workbook.SheetNames.join(', ')}`);
            }
            column = column || '';
            const worksheet = this.workbook.Sheets[index];
            const address = `${column.toUpperCase()}${row}`;
            const cell = worksheet[address];
            return cell;
        };
        /**
         * Get worksheet cell value
         * @param sheetName Workbook sheet
         * @param row Row number (starts from 1)
         * @param column Column name (A, B, etc)
         */
        this.getCellValue = (sheetName, column, row) => {
            const cell = this.getCell(sheetName, column, row);
            return cell ? cell.w || cell.v : undefined;
        };
        this.workbook = xlsx_1.readFile(filePath);
    }
}
exports.ExcelFile = ExcelFile;
//# sourceMappingURL=index.js.map