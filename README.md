## Installation
```bash
npm i sec-excel --save
```

## Usage
```javascript
const { ExcelFile } = require('sec-excel');

...

const file = new ExcelFile(pathToFile);
const cellValue = file.getCellValue('Page 1', 'D', 2);

...
```