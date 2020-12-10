const path = require('path');
const { ExcelFile } = require('../dist');

const PATH = path.join(__dirname, 'book.xlsx');

describe('Read from file', () => {
    const file = new ExcelFile(PATH);

    test('Text value must be correct', () => {
        const value = file.getCellValue('Page', 2, 'A');
        expect(value).toBe('Text');
    });

    test('Common value must be correct', () => {
        const value = file.getCellValue('Page', 2, 'D');
        expect(value).toBe('Common');
    });

    test('Formula value must be correct', () => {
        const value = file.getCellValue('Page', 2, 'E');
        expect(value).toBe('TextCommon');
    });

    test('Number value must be correct', () => {
        const value = file.getCellValue('Page', 2, 'B');
        expect(value).toBe('2.00');
    });

    test('Date value must be correct', () => {
        const value = file.getCellValue('Page', 2, 'C');
        expect(value).toBe('Saturday, February 01, 2020');
    });
});
