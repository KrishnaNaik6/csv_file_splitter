const fs = require('fs');
const parse = require('papaparse');
const path = require('path')
const XLSX = require('xlsx')

const PATH = './csv/trudata_fo_1min_2024-03-18.csv';
const output_path = `./xlsx/new_${path.basename(PATH).replace('.csv', '.xlsx')}`

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

let d = readCSVFile(PATH)

console.log(d.length)
let size = d.length
const pageSize = 904857
const totalPages = Math.ceil(size / pageSize)
console.log(`total paf=ges:${totalPages}`)
let newWorkbook = XLSX.utils.book_new();


function addDataToWorkbook(newWorkbook, data, index) {
    try {
        const worksheet = XLSX.utils.aoa_to_sheet(data)
        XLSX.utils.book_append_sheet(newWorkbook, worksheet, 'page-' + index);
    } catch (e) {
        console.log(e)
    }
}

for (i = 0; i < totalPages; i++) {
    console.log(`Creating Page${i + 1} sindex:${i * pageSize} to ${pageSize}`)
    addDataToWorkbook(newWorkbook, d.slice(i * pageSize, (i * pageSize) + pageSize), i + 1)
}
XLSX.writeFile(newWorkbook, output_path, { bookType: 'xlsx', type: 'buffer', encoding: 'utf-8' });
