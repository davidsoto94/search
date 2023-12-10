import ExcelJS from 'exceljs';

async function main() {
    //get param number from command line
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('LMSDcompleta.xlsx');
    const spectersWorkbook = new ExcelJS.Workbook();
    await spectersWorkbook.xlsx.readFile('listaEspectros.xlsx');
    const worksheet = workbook.getWorksheet(1);
    const spectersWorksheet = spectersWorkbook.getWorksheet(1);
    const resultArray = [];
    
    for (let x = 1; x <= spectersWorksheet.rowCount; x++) {
        const param = Math.floor(spectersWorksheet.getRow(x).getCell('A').value);
        resultArray.push(["FFFFFF",spectersWorksheet.getRow(x).getCell('A').value]);
        for (let i = 1; i <= worksheet.rowCount; i++) {
            const row = worksheet.getRow(i);
            //check if the value of the cell CH, CI or CJ is equal to the param
            if (['CH', 'CI'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param)|| Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = [
                    'CC', 'CD', 'CE', 'CF','CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM'];
                // Iterate over the cell names
                cellNames.forEach(cellName => {
                    let cellValue = row.getCell(cellName).value;
                    if (typeof cellValue == 'object') {
                        row.getCell(cellName).value = cellValue.result;
                    }
                });
                //insert the row from vell CC to cell CM in the resultArray all in one line 
                resultArray.push(["D0CECE",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
    
            //check if the value of the cell CW, CX or CY is equal to the param
            if (['CW', 'CX'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param) || Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = ['CR', 'CS', 'CT', 'CU','CV', 'CW', 'CX', 'CY', 'CZ', 'DA', 'DB'];
    
                cellNames.forEach(cellName => {
                    let cellValue = row.getCell(cellName).value;
                    if (typeof cellValue == 'object') {
                        row.getCell(cellName).value = cellValue.result;
                    }
                });
    
                resultArray.push(["C6E0B4",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
            if (['DL', 'DM'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param) || Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = ['DG', 'DH', 'DI', 'DJ', 'DK', 'DL', 'DM', 'DN', 'DO', 'DP', 'DQ'];
                // Iterate over the cell names
                for (let cellName of cellNames) {
                    // Get the cell value
                    let cellValue = row.getCell(cellName).value;
    
                    // Check if the cell value is an object
                    if (typeof cellValue == 'object') {
                        // If it is, update the cell value with the result property of the object
                        row.getCell(cellName).value = cellValue.result;
                    }
                }
                //insert the row from vell DG to cell DS in the resultArray all in one line 
                resultArray.push(["FFE699",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
    
             //check if the value of the cell CW, CX or CY is equal to the param
             if (['DZ', 'EA'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param) || Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = ['DU', 'DV', 'DW', 'DX','DY', 'DZ', 'EA', 'EB', 'EC', 'ED', 'EE'];
    
                cellNames.forEach(cellName => {
                    let cellValue = row.getCell(cellName).value;
                    if (typeof cellValue == 'object') {
                        row.getCell(cellName).value = cellValue.result;
                    }
                });
    
                resultArray.push(["BDD7EE",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
    
            if (['EO', 'EP'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param) || Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = ['EJ', 'EK', 'EL', 'EM','EN', 'EO', 'EP', 'EQ', 'ER', 'ES', 'ET'];
    
                cellNames.forEach(cellName => {
                    let cellValue = row.getCell(cellName).value;
                    if (typeof cellValue == 'object') {
                        row.getCell(cellName).value = cellValue.result;
                    }
                });
    
                resultArray.push(["FFE699",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
    
            if (['FD', 'FE'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param) || Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = ['EY', 'EZ', 'FA', 'FB', 'FC', 'FD', 'FE', 'FF', 'FG', 'FH', 'FI'];
    
                cellNames.forEach(cellName => {
                    let cellValue = row.getCell(cellName).value;
                    if (typeof cellValue == 'object') {
                        row.getCell(cellName).value = cellValue.result;
                    }
                }); 
                resultArray.push(["FFFF99",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
    
        }
    }


    //create a new workbook and worksheet
    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet('Resultados');
    //add the resultArray to the new worksheet
    for (let i = 0; i < resultArray.length; i++){
        let row = newWorksheet.addRow(resultArray[i]);
        row.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: resultArray[i][0] }
        };
    }
    //save the new workbook
    newWorkbook.xlsx.writeFile('resultados.xlsx');
}

main()