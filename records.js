// data to be converted
const data = [
    {
        crno: 274800,
        name: 'Gurkeerat',
        age: 4,
        sex: 'M',
        topics: 'English',
        work: 'Alphabets'
    },
    {
        crno: 200547,
        name: 'Gurjashan',
        age: 5,
        sex: 'M',
        topics: 'Maths',
        work: 'Tables 2 - 10'
    },
    {
        crno: 154876,
        name: 'Rayna',
        age: 5,
        sex: 'F',
        topics: '',
        work: 'Reasoning'
    },
    {
        crno: 226334,
        name: 'Inaya',
        age: 3,
        sex: 'F',
        topics: '',
        work: 'Colouring'
    },
    {
        crno: 284692,
        name: 'Ayush',
        age: 10,
        sex: 'M',
        topics: 'English',
        work: 'Prepositions'
    }
];

// importing exceljs
const Excel = require('exceljs');
const path = require('path');

// 
let area = 'academics';
// date
let date = '18.05.2021';
// need to create a workbook object. Almost everything in ExcelJS is based off of the workbook object.
let workbook = new Excel.Workbook();

// adding a worksheet
let worksheet = workbook.addWorksheet(`${date}`);

let keys = Object.keys(data[0]);
console.log(keys);

let columnList = [];
keys.forEach((key) => {columnList.push({header: key.toUpperCase(), key: key})});

console.log(columnList);

// defining columns
worksheet.columns = columnList;
// worksheet.columns = [
//     { header: 'CR No', key: 'crno' },
//     { header: 'Name', key: 'name' },
//     { header: 'Age', key: 'age' },
//     { header: 'Sex', key: 'sex' },
//     { header: 'Assignment', key: 'topics' },
//     { header: '', key: 'work' }
// ];

// force the columns to be at least as long as their header row.
// Have to take this approach because ExcelJS doesn't have an autofit property.
worksheet.columns.forEach(column => {
    column.width = column.header.length < 12 ? 12 : column.header.length;
});

// worksheet.columns[worksheet.columns.length].width = 20;

// Make the header bold.
// Note: in Excel the rows are 1 based, meaning the first row is 1 instead of 0.
worksheet.getRow(1).font = { bold: true };

// Dump all the data into Excel
data.forEach((e, index) => {
    // row 1 is the header.
    // const rowIndex = index + 2

    // By using destructuring we can easily dump all of the data into the row without doing much
    // We can add formulas pretty easily by providing the formula property.
    worksheet.addRow({...e});
});

worksheet.addRow();
worksheet.addRow(['',`Total = ${data.length}`]);

// loop through all of the rows and set the outline style.
worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
    const columns = ['A','B', 'C', 'D', 'E', 'F'];

    columns.forEach((v) => {
        worksheet.getCell(`${v}1`).alignment = {horizontal: 'center'}
    });

    worksheet.getCell(`A${rowNumber}`).alignment = {horizontal: 'center'}
    worksheet.getCell(`C${rowNumber}`).alignment = {horizontal: 'center'}
    worksheet.getCell(`D${rowNumber}`).alignment = {horizontal: 'center'}
});

// Create a freeze pane, which means we'll always see the header as we scroll around.
worksheet.views = [
    { state: 'frozen', xSplit: 0, ySplit: 1, activeCell: 'B2' }
];

// Keep in mind that reading and writing is promise based.
workbook.xlsx.writeFile(path.join(__dirname, `/reports/${area}_${date}.xlsx`));
