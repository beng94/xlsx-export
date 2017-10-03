const xlsx = require('xlsx');
const csv = require('fast-csv');
const fs = require('fs');

const path = './live_teams.xlsx';

const workbook = xlsx.readFile(path);
const studentData = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

const exportData = {};
studentData.forEach((student) => {
    let team = exportData[student.team];
    if(team) {
        team.emails.push(student.email);
    } else {
        exportData[student.team] = {
            name: student.team,
            emails: [ student.email]
        };
    }
});

console.log(JSON.stringify(exportData, null, 4));

let csvData = Object.keys(exportData).map((key) => {
    const data = exportData[key];
    return {
        name: data.name,
        email_1: data.emails[0],
        email_2: data.emails[1],
        email_3: data.emails[2]
    };
});

console.log(JSON.stringify(csvData, null, 4));

const fileStream = fs.createWriteStream('out.csv');
csv
.write(csvData, { headers: true })
.pipe(fileStream);