// Requiring the module 
const reader = require('xlsx')

// Reading our test file 
const file = reader.readFile('Assignment_Timecard.xlsx')

// for storing data
let data = []
let PositionId = [];
let PositionStatus = [];
let Time = [];
let TimeOut = [];
let TimeCardHours = [];
let PCycleStart = [];
let PCycleEnd = [];
let EmployeeName = [];
let FileNumber = [];

const sheets = file.SheetNames
// access the data from excel sheet and store in above arrays
for (let i = 0; i < sheets.length; i++) {

    const temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
    temp.forEach((res) => {
        PositionId.push(res['Position ID']);
        PositionStatus.push(res['Position Status']);
        Time.push(res['Time']);
        TimeOut.push(res['Time Out']);
        TimeCardHours.push(res['Timecard Hours (as Time)']);
        PCycleStart.push(res['Pay Cycle Start Date']);
        PCycleEnd.push(res['Pay Cycle End Date']);
        EmployeeName.push(res['Employee Name']);
        FileNumber.push(res['File Number']);
    })
}
// convert excel time data into standred time and date using excelTimestampToDate function
for (let i = 0; i < Time.length; i++) {
    Time[i] = excelTimestampToDate(Time[i]);
    TimeOut[i] = excelTimestampToDate(TimeOut[i]);
    PCycleStart[i] = excelTimestampToDate(PCycleStart[i]);
    PCycleEnd[i] = excelTimestampToDate(PCycleEnd[i]);
}


function SolutionOfThirdQuestion() {

    let i = 1;
    while (i < TimeCardHours.length) {

        let str1 = new String(TimeCardHours[i]);
        let str2 = new String(TimeCardHours[i + 1])
        let hour1 = (+str1.slice(0, 1)); // converting in ineteger from string using + operator
        let minutes1 = (+str1.slice(2, 4));
        let hour2 = (+str2.slice(0, 1));
        let minutes2 = (+str2.slice(2, 4));

        let finalhour = hour1 + hour2;
        let finalminutes = minutes1 + minutes2;


        let sumofbothhours = +(finalhour + '.' + finalminutes)
        // if total hour will be greater than 14 then print employee details Using hours in morning to afternoon then afternoon to evening
        if (sumofbothhours > 14) {
            console.log('-----------------------------------------------------------------------');
            console.log("Employee Name is: ", EmployeeName[i]);
            console.log("Employee Position Id is: ", PositionId[i],);
            console.log("Employee Position Status is: ", PositionStatus[i],);
            console.log("Morning to afternoon Timecard hours: ", hour1+":"+minutes1)
            console.log("afternoon to evening Timecard hours: ", hour2+":"+minutes2)
        }
        i += 2; //jump to next day
    }
}

function SolutionOfSecondQuestion() {

    let i = 1;

    while (i < TimeCardHours.length) {

        let str1 = new String(TimeCardHours[i]);
        let str2 = new String(TimeCardHours[i + 1])
        let hour1 = (+str1.slice(0, 1));
        let minutes1 = (+str1.slice(2, 4));
        let hour2 = (+str2.slice(0, 1));
        let minutes2 = (+str2.slice(2, 4));

        if (PositionId[i] != PositionId[i + 1]) {
            // not full day only half day
            let totaltime = +(hour1 + "." + minutes1);
            if (totaltime > 1 && totaltime < 10) {
                console.log("Employee Name is:", EmployeeName[i]);
                console.log("Employee Position Id is:", PositionId[i], );
                console.log("Employee Position Status is:", PositionStatus[i],);
                console.log("Morning to afternoon Timecard hours:", hour1+":"+minutes1)
            }
            i++;
        }
        else {
            // check employee with greater than 1 hour but less than 10 hours
            let finalhour = hour1 + hour2;
            let finalminutes = minutes1 + minutes2;

            let sumofbothhours = +(finalhour + '.' + finalminutes)
            if (sumofbothhours > 1 && sumofbothhours < 10) {
                console.log('-----------------------------------------------------------------------');
                console.log("Employee Name is: ", EmployeeName[i]);
                console.log("Employee Position Id is: ", PositionId[i], );
                console.log("Employee Position Status is: ", PositionStatus[i], );
                console.log("Morning to afternoon Timecard hours: ", hour1+":"+minutes1)
                console.log("afternoon to evening Timecard hours: ", hour2+":"+minutes2)
            }
            i += 2; // jump to next day
        }
    }
}

function SolutionOfFirstQuestion() {
    let indx = 1;
    // if date after seven day is currdate + 7 then its answer
    while (indx < Time.length - 13) {
        let firstdate = Time[indx].getDate();
        let seventhdate = Time[indx + 13].getDate();

        if (firstdate + 6 == seventhdate) {
            console.log("---------------------------------------------------------------------------")
            console.log("From", Time[indx].toDateString(), "to", Time[indx+13].toDateString())
            console.log("Employee Name is:", EmployeeName[indx]);
            console.log("Employee Position Id is:", PositionId[indx]);
            console.log("Employee Position Status is:", PositionStatus[indx]);
            indx += 6;
        }
        else
        indx++;
    }
}

function excelTimestampToDate(excelTimestamp) {
    // Excel timestamp starts from January 1, 1900 (Windows) or January 1, 1904 (Mac)
    const excelEpoch = new Date('1899-12-30T00:00:00Z');

    // Convert days to milliseconds
    const milliseconds = excelTimestamp * 24 * 60 * 60 * 1000;

    // Create a new Date object by adding milliseconds to the Excel epoch
    const date = new Date(excelEpoch.getTime() + milliseconds);

    return date;
}

console.log("\nEmployees who have worked for 7 consecutive days:\n");
SolutionOfFirstQuestion();
console.log('----------------------------------------------------------------------');
console.log("\nEmployees who have worked for more than 14 hours in a single shift:");
SolutionOfThirdQuestion();
console.log('----------------------------------------------------------------------');
console.log("Employees with less than 10 hours between shifts but greater than 1 hour:");
SolutionOfSecondQuestion();