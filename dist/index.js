"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelFile = void 0;
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
        this.getCell = (sheetName, row, column) => {
            sheetName = this.workbook.SheetNames.find((s) => s === sheetName);
            if (!sheetName) {
                throw new Error(`Worksheet ${sheetName} not found.`);
            }
            column = column || '';
            const worksheet = this.workbook.Sheets[sheetName];
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
        this.getCellValue = (sheetName, row, column) => {
            const cell = this.getCell(sheetName, row, column);
            return cell ? cell.w || cell.v : undefined;
        };
        this.workbook = xlsx_1.readFile(filePath);
    }
}
exports.ExcelFile = ExcelFile;
//# sourceMappingURL=index.js.map