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
}


function readFromExcel(workbook) {
    return workbook.xlsx.readFile('namesList.xlsx')
        .then(() => {
            console.log('I read the file !!!');
            const worksheet = workbook.getWorksheet('Names');
            const userList = [];
            // for(let i=2 ; i<worksheet.length ; i++){
            //     console.log("hiii");
            // }




            
          
            if (worksheet) {
                worksheet.eachRow((row, rowNumber) => {
                    console.log("i am iterating");
                    if (rowNumber === 1) return;
                    let userObject = {
                        firstName: row.getCell('A').value, 
                        lastName: row.getCell('B').value, 
                        age: row.getCell('C').value        
                    };
                    
                    userList.push(userObject);
                });

                console.log(userList);
                return userList; 
            } else {
                console.log(' does not exist.');
                return ;
            }
        })
        .catch((error) => {
            console.log('File does not exist.', error);
            return ;
        });
}

function addWorksheet(workbook) {
    let worksheet = workbook.getWorksheet('Names');
    if (!worksheet) {
        worksheet = workbook.addWorksheet('Names');
    }
    return worksheet;
}

module.exports = {
    addNames,
    readFromExcel,
    addWorksheet
};