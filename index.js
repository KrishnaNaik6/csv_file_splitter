const fs = require('fs');
const parse = require('papaparse');
const path = require('path')
const XLSX = require('xlsx')

const PATH = './csv/trudata_fo_1min_2024-03-18.csv';
const output_path = `./xlsx/${path.basename(PATH).replace('.csv', '.xlsx')}`


// Function to read CSV file and return an array of arrays (header and data)
function readCSVFile(PATH) {
    try {
        // Read the CSV file
        const fileContent = fs.readFileSync(PATH, 'utf-8');

        // Parse the CSV data
        const { data, meta } = parse.parse(fileContent, {
            header: true // Treat the first row as header
        });

        // Convert the parsed data into an array of arrays
        const dataArray = [meta.fields]; // Header as the first array
        data.forEach(row => {
            dataArray.push(Object.values(row)); // Push data rows as arrays
        });


        return dataArray;
    } catch (error) {
        console.error('Error:', error);
        return null;
    }
}

// =================================append data to xlsx file==========================
function append_data(newdata, sheet) {
    // Load the existing workbook
    const workbook = XLSX.readFile(output_path);

    // Get the first sheet in the workbook
    const sheetName = sheet;
    const worksheet = workbook.Sheets[sheetName];

    // Create new data to append
    const newData = newdata

    // Append the new data to the existing sheet
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const newRowStartIndex = range.e.r + 1;
    newData.forEach((row, rowIndex) => {
        row.forEach((cellValue, cellIndex) => {
            console.log("row" , newRowStartIndex + rowIndex, "COLUM", cellIndex)
            const cellAddress = XLSX.utils.encode_cell({ r: newRowStartIndex + rowIndex, c: cellIndex });
            worksheet[cellAddress] = { t: 's', v: cellValue }; // 's' indicates string type for cell value
        });
    });

    console.log(`range of column c: ${range.e.c}, r : ${newRowStartIndex + newData.length - 1}`,)
    // Update the range of the worksheet
    worksheet['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: range.e.c, r: newRowStartIndex + newData.length - 1 } });

    // Write the modified data back to the workbook
    XLSX.writeFile(workbook, output_path, { bookType: 'xlsx', type: 'buffer' }, function (err) {
        if (err) {
            console.error(err);
        } else {
            console.log("New data appended successfully!");
        }
    });

}

function newsheet(webbook, name, data=[]){
    const worksheet = XLSX.utils.aoa_to_sheet(data)
    XLSX.utils.book_append_sheet(webbook, worksheet, name);
}



// Function to create a workbook with two worksheets and save it to a specified directory
function createWorkbookAndSave(data, output_path) {
    let min_row = 104857
    let start = 0
    let end = min_row
    let hed = data[0]
    let range = data.length / min_row
    let fsliced = data.slice(start, end)
    console.log("range", Math.ceil(range))

    try {
        var sheet_name = 'Sheet1'
        // Create a new workbook
        let newWorkbook = XLSX.utils.book_new();
        console.log("Created new work-book")

        // =========================Convert data1 to worksheet
        newsheet(newWorkbook, sheet_name, fsliced)
        newsheet(newWorkbook, 'Sheet2')
        start = end
        end +=min_row
        console.log("Created sheet", sheet_name)
        
        XLSX.writeFile(newWorkbook, output_path, { bookType: 'xlsx', type: 'buffer', encoding: 'utf-8' });
        console.log(`Workbook saved to ${output_path}`);


        for ( let i =0; i<Math.ceil(range); i++){
            if(start>data.length){
                console.log("breaking")
                break
            }

            if(start === 1048570){
                console.log("above the sheets min_row")
                sheet_name = 'Sheet2'
                // newsheet(newWorkbook, sheet_name)
                let sliced = data.slice(start, end)
                var newdata = sliced.unshift[hed]
                append_data(newdata, sheet_name)
                start = end
                end += min_row
            }
            else{
                console.log(start, end)
                let sliced = data.slice(start, end)
                append_data(sliced, sheet_name)
                start = end
                end += min_row
            }
        }
        
    } catch (error) {
        console.error('Error:', error);
    }
}


// // read the csv file
const data = readCSVFile(PATH)
// const hed = data[0]

// const data1 = data.slice(0, 10)
// const data2 = data.slice(data.length - 10, data.length)
// data2.unshift(hed);


// create workbook and save
createWorkbookAndSave(data, output_path);