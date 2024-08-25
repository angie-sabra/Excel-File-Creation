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


// const express = require('express');

// const app = express();

// app.get('/users-list', (req, res) => {
//   // Get complete list of users
//   const usersList = [];

//   // Send the usersList as a response to the client
//   res.send(usersList);
// });