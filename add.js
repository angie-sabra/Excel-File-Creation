
const prompt = require('prompt-sync')({ sigint: true });
const ExcelJS = require('exceljs');
let workbook = new ExcelJS.Workbook();
let worksheet = workbook.addWorksheet('Names');
let namesList = [];

try {

    workbook.xlsx.readFile('namesList.xlsx')

        .then(() => {
            worksheet = workbook.getWorksheet('Names');
            if (!worksheet) {
                worksheet = workbook.addWorksheet('Names');
            }
            addNames();
        })

        .catch((error) => {
            console.log('file does not exist')
            addNames();
        });


} catch (error) {
    console.log(error)
}

worksheet.columns = [
    { header: 'FirstName', key: 'firstName', width: 32 },
    { header: 'LastName', key: 'lastName', width: 32 },
    { header: 'Age', key: 'age', width: 10 }
];



worksheet.getRow(1).font = { bold: true };


function addNames() {
    while (true) {
        let firstName = prompt('Enter the first name / type "exit" to exit: ');

        if (firstName.toLowerCase() === 'exit') {
            break;
        }

        let lastName = prompt('Enter the last name: ');
        let age = prompt('Enter the age :')
        let fullName = { firstName: firstName, lastName: lastName, age: age };
        namesList.push(fullName);
    }


    namesList.forEach(person => {
        console.log(person)
        worksheet.addRow([person.firstName, person.lastName, person.age]);
    });


    workbook.xlsx.writeFile('namesList.xlsx')
        .then(() => {
            console.log('List saved to namesList.xlsx');
        })
        .catch(err => {
            console.error('Error writing to file:', err);
        });

    // console.log('List:', namesList);
}
