const xlsx = require('xlsx');


function excelDateToJSDate(serial) {
    const excelEpoch = new Date(1899, 11, 30);
    const daysInMillis = serial * 86400000; 
    return new Date(excelEpoch.getTime() + daysInMillis);
}

function calculateTimeGaps(serialDates) {
    const THIRTY_MINUTES_IN_MILLIS = 30 * 60 * 1000; 
    let totalGap = 0;

    for (let i = 1; i < serialDates.length; i++) {
        const date1 = excelDateToJSDate(serialDates[i - 1]);
        const date2 = excelDateToJSDate(serialDates[i]);
        const gapInMillis = date2 - date1;

        console.log(gapInMillis);

        if (gapInMillis >= THIRTY_MINUTES_IN_MILLIS) {
            totalGap += gapInMillis;
        }
    }
    return totalGap / (1000 * 60); 
}

function readTimestampsFromExcel(filePath, sheetName) {
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[sheetName];

    const timestamps = xlsx.utils.sheet_to_json(sheet, { header: 1 }).map(row => row[0]);

    return timestamps.filter(Boolean); 
}

const filePath = 'timestamps.xlsx';
const sheetName = 'Sheet1'; 

const timestamps = readTimestampsFromExcel(filePath, sheetName);
const totalGap = calculateTimeGaps(timestamps);

console.log(`Total time gap >= 30 mins: ${totalGap} minutes`);
