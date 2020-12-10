import { readFile, WorkBook, WorkSheet, CellObject } from "xlsx";

class ExcelFile {
    private workbook: WorkBook;

    /**
     * Create an instance of ExcelFile
     * @param filePath Path to .xlsx file
     */
    constructor(filePath: string) {
        this.workbook = readFile(filePath);
    }

    /**
     * Get worksheet cell object
     * @param sheetName Workbook sheet
     * @param row Row number (starts from 1)
     * @param column Column name (A, B, etc)
     */
    getCell = (sheetName: string, column: string, row: number) => {
        sheetName = this.workbook.SheetNames.find((s) => s === sheetName);
        if (!sheetName) {
            throw new Error(`Worksheet ${sheetName} not found.`);
        }

        column = column || '';

        const worksheet: WorkSheet = this.workbook.Sheets[sheetName];
        const address = `${column.toUpperCase()}${row}`;
        const cell = worksheet[address] as CellObject;

        return cell;
    };

    /**
     * Get worksheet cell value
     * @param sheetName Workbook sheet
     * @param row Row number (starts from 1)
     * @param column Column name (A, B, etc)
     */
    getCellValue = (sheetName: string, column: string, row: number) => {
        const cell = this.getCell(sheetName, column, row);
        return cell ? cell.w || cell.v : undefined;
    };
}

export { ExcelFile };
