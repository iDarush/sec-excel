const path = require('path');
const { ExcelFile } = require('../dist');

const PATH = path.join(__dirname, 'book.xlsx');

describe('Read from file', () => {
    const file = new ExcelFile(PATH);

    test('Text value must be correct', () => {
        const value = file.getCellValue('Page', 'A', 2);
        expect(value).toBe('Text');
    });

    test('Common value must be correct', () => {
        const value = file.getCellValue('Page', 'D', 2);
        expect(value).toBe('Common');
    });

    test('Formula value must be correct', () => {
        const value = file.getCellValue('Page', 'E', 2);
        expect(value).toBe('TextCommon');
    });

    test('Number value must be correct', () => {
        const value = file.getCellValue('Page', 'B', 2);
        expect(value).toBe('2.00');
    });

    test('Date value must be correct', () => {
        const value = file.getCellValue('Page', 'C', 2);
        expect(value).toBe('Saturday, February 01, 2020');
    });
});
