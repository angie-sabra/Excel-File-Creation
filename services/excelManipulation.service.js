const ExcelJS = require('exceljs');

let workbook = new ExcelJS.Workbook();
let worksheet;
let namesList = [];

async function initialize() {
    try {
        await workbook.xlsx.readFile('namesList.xlsx');
        worksheet = workbook.getWorksheet('Names');
        if (!worksheet) {
            worksheet = workbook.addWorksheet('Names');
        }
    } catch (error) {
        console.log('File does not exist, creating a new one.');
        worksheet = workbook.addWorksheet('Names');
    }

    worksheet.columns = [
        { header: 'FirstName', key: 'firstName', width: 32 },
        { header: 'LastName', key: 'lastName', width: 32 },
        { header: 'Age', key: 'age', width: 10 }
    ];

    worksheet.getRow(1).font = { bold: true };

    loadUsersFromWorksheet();
}

function loadUsersFromWorksheet() {
    namesList = [];
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber === 1) return; 
        const user = {
            firstName: row.getCell('A').value,
            lastName: row.getCell('B').value,
            age: row.getCell('C').value
        };
        namesList.push(user);
    });
}

function addUser(firstName, lastName, age) {
    const newUser = { firstName, lastName, age };
    namesList.push(newUser);
    worksheet.addRow(newUser);

    return workbook.xlsx.writeFile('namesList.xlsx')
        .then(() => newUser)
        .catch(err => {
            throw new Error('Error saving the user to the file: ' + err.message);
        });
}

function updateUser(firstName, newDetails) {
    let user = namesList.find(u => u.firstName === firstName);

    if (!user) {
        throw new Error('User not found');
    }

    user.lastName = newDetails.lastName || user.lastName;
    user.age = newDetails.age || user.age;

    worksheet.eachRow((row, rowNumber) => {
        if (row.getCell('A').value === firstName) {
            row.getCell('B').value = user.lastName;
            row.getCell('C').value = user.age;
        }
    });

    return workbook.xlsx.writeFile('namesList.xlsx')
        .then(() => user)
        .catch(err => {
            throw new Error('Error updating the user in the file: ' + err.message);
        });
}

function deleteUser(firstName) {
    const userIndex = namesList.findIndex(u => u.firstName === firstName);

    if (userIndex === -1) {
        throw new Error('User not found');
    }

    namesList.splice(userIndex, 1);
    worksheet.spliceRows(userIndex + 2, 1); 

    return workbook.xlsx.writeFile('namesList.xlsx')
        .catch(err => {
            throw new Error('Error deleting the user from the file: ' + err.message);
        });
}

module.exports = {
    initialize,
    loadUsersFromWorksheet,
    addUser,
    updateUser,
    deleteUser,
    getUsers: () => namesList  
};





// function addNames(prompt, worksheet, namesList, workbook) {
//     while (true) {
//         let firstName = prompt('Enter the first name / type "exit" to exit: ');
//         if (firstName.toLowerCase() === 'exit') {
//             break;
//         }
//         let lastName = prompt('Enter the last name: ');
//         let age = prompt('Enter the age: ');
//         let fullName = { firstName: firstName, lastName: lastName, age: age };
//         namesList.push(fullName);
//     }

//     namesList.forEach(person => {
//         console.log(person);
//         worksheet.addRow([person.firstName, person.lastName, person.age]);
//     });

//     workbook.xlsx.writeFile('namesList.xlsx')
//         .then(() => {
//             console.log('List saved to namesList.xlsx');
//         })
//         .catch(err => {
//             console.error('Error writing to file:', err);
//         });
// }

// function readFromExcel(workbook) {
//     return workbook.xlsx.readFile('namesList.xlsx')
//         .then(() => {
//             console.log('I read the file !!!');
//             const worksheet = workbook.getWorksheet('Names');

//             if (worksheet) {
//                 const expectedHeaders = ['FirstName', 'LastName', 'Age'];
//                 const headerRow = worksheet.getRow(1);
//                 const headers = {};
//                 const actualHeaders = [];
                
//                 headerRow.eachCell((cell, colNumber) => {
//                     const header = cell.value;
//                     headers[header] = colNumber;
//                     actualHeaders.push(header);
//                 });

//                 const additionalHeaders = actualHeaders.filter(header => !expectedHeaders.includes(header));

//                 if (additionalHeaders.length > 0) {
//                     console.log('Additional columns found:', additionalHeaders);
//                 } else {
//                     console.log('No additional columns found.');
//                 }

//                 const userList = [];
//                 worksheet.eachRow((row, rowNumber) => {
//                     if (rowNumber === 1) return;

//                     let userObject = {};
//                     if (headers['FirstName']) {
//                         userObject.firstName = row.getCell(headers['FirstName']).value;
//                     }
//                     if (headers['LastName']) {
//                         userObject.lastName = row.getCell(headers['LastName']).value;
//                     }
//                     if (headers['Age']) {
//                         userObject.age = row.getCell(headers['Age']).value;
//                     }

//                     if (Object.keys(userObject).length > 0) {
//                         userList.push(userObject);
//                     }
//                 });

//                 console.log(userList);
//                 return userList;
//             } else {
//                 console.log('Worksheet does not exist.');
//                 return [];
//             }
//         })
//         .catch((error) => {
//             console.log('File does not exist.', error);
//             return [];
//         });
// }

// function addWorksheet(workbook) {
//     let worksheet = workbook.getWorksheet('Names');
//     if (!worksheet) {
//         worksheet = workbook.addWorksheet('Names');
//     }
//     return worksheet;
// }

// module.exports = {
//     addNames,
//     readFromExcel,
//     addWorksheet
// };
