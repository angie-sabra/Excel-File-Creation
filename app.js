// const ExcelJs = require('exceljs/dist/es5');
const prompt = require('prompt-sync')({ sigint: true });
const ExcelJS = require('exceljs');
let workbook = new ExcelJS.Workbook();
let worksheet = workbook.addWorksheet('Names');
let namesList = [];
const { addNames, readFromExcel, addWorksheet } = require('./services/excelManipulation.service');
worksheet.columns = [
    { header: 'FirstName', key: 'firstName', width: 32 },
    { header: 'LastName', key: 'lastName', width: 32 },
    { header: 'Age', key: 'age', width: 10 }
];

worksheet.getRow(1).font = { bold: true };



readFromExcel(workbook);