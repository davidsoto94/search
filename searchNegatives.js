import ExcelJS from 'exceljs';

async function main() {
    //get param number from command line
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('LMSDcompleta.xlsx');
    const spectersWorkbook = new ExcelJS.Workbook();
    await spectersWorkbook.xlsx.readFile('listaEspectros.xlsx');
    const worksheet = workbook.getWorksheet(2);
    const spectersWorksheet = spectersWorkbook.getWorksheet(2);
    const resultArray = [];
    for (let x = 1; x <= spectersWorksheet.rowCount; x++) {
        const param = Math.floor(spectersWorksheet.getRow(x).getCell('A').value);
        if(param==0){
            continue
        }
        resultArray.push(["FFFFFF",spectersWorksheet.getRow(x).getCell('A').value]);
        for (let i = 1; i <= worksheet.rowCount; i++) {
            const row = worksheet.getRow(i);
            //check if the value of the cell CH, CI or CJ is equal to the param
            if (['G'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param)|| Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = [
                    'B', 'C','D', 'E', 'F','G', 'H', 'I', 'J'];
                // Iterate over the cell names
                cellNames.forEach(cellName => {
                    let cellValue = row.getCell(cellName).value;
                    if (typeof cellValue == 'object' && cellValue!=null) {
                        // If it is, update the cell value with the result property of the object
                        row.getCell(cellName).value = cellValue.result;
                    }
                    else if(typeof cellValue == Number){
                        console.log(cellName)
                        console.log(cellValue)
                        row.getCell(cellName).value = cellValue;
                    }
                });
                //insert the row from vell CC to cell CM in the resultArray all in one line 
                resultArray.push(["D0CECE",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
    
            //check if the value of the cell CW, CX or CY is equal to the param
            if (['T'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param) || Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = ['O','P','Q', 'R', 'S', 'T','U', 'V', 'W'];
    
                cellNames.forEach(cellName => {
                    let cellValue = row.getCell(cellName).value;
                    if (typeof cellValue == 'object' && cellValue!=null) {
                        console.log(cellName)
                        console.log(cellValue)
                        // If it is, update the cell value with the result property of the object
                        row.getCell(cellName).value = cellValue.result;
                    }
                    else if(typeof cellValue == Number){
                        row.getCell(cellName).value = cellValue;
                    }
                });
    
                resultArray.push(["C6E0B4",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
            if (['AF', 'AG'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param) || Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = ['AB', 'AC', 'AD', 'AE','AF', 'AG', 'AH', 'AI', 'AJ'];
                // Iterate over the cell names
                for (let cellName of cellNames) {
                    // Get the cell value
                    let cellValue = row.getCell(cellName).value;
    
                    // Check if the cell value is an object
                    if (typeof cellValue == 'object' && cellValue!=null) {
                        console.log(cellName)
                        console.log(cellValue)
                        // If it is, update the cell value with the result property of the object
                        row.getCell(cellName).value = cellValue.result;
                    }
                    else if(typeof cellValue == Number){
                        row.getCell(cellName).value = cellValue;
                    }
                }
                //insert the row from vell DG to cell DS in the resultArray all in one line 
                resultArray.push(["FFE699",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
    
             //check if the value of the cell CW, CX or CY is equal to the param
             if (['AR', 'AS'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param) || Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = ['AN', 'AO', 'AP', 'AQ','AR','AS', 'AT', 'AU', 'AV'];
    
                cellNames.forEach(cellName => {
                    let cellValue = row.getCell(cellName).value;
                    if (typeof cellValue == 'object' && cellValue!=null) {
                        console.log(cellName)
                        console.log(cellValue)
                        // If it is, update the cell value with the result property of the object
                        row.getCell(cellName).value = cellValue.result;
                    }
                    else if(typeof cellValue == Number){
                        row.getCell(cellName).value = cellValue;
                    }
                });
    
                resultArray.push(["BDD7EE",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
    
            if (['BE', 'BF'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param) || Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = ['BA', 'BB', 'BC', 'BD', 'BE','BF','BG','BH', 'BI'];
    
                cellNames.forEach(cellName => {
                    let cellValue = row.getCell(cellName).value;
                    if (typeof cellValue == 'object' && cellValue!=null) {
                        console.log(cellName)
                        console.log(cellValue)
                        // If it is, update the cell value with the result property of the object
                        row.getCell(cellName).value = cellValue.result;
                    }
                    else if(typeof cellValue == Number){
                        row.getCell(cellName).value = cellValue;
                    }
                });
    
                resultArray.push(["FFE699",...cellNames.map(cellName => row.getCell(cellName).value)]);
            }
    
            if (['BR', 'BS'].some(cell => Math.floor(row.getCell(cell).value?.result) == Number(param) || Math.floor(row.getCell(cell).value) == Number(param))) {
                const cellNames = ['BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV'];
    
                cellNames.forEach(cellName => {
                    let cellValue = row.getCell(cellName).value;
                    if (typeof cellValue == 'object' && cellValue!=null) {
                        console.log(cellName)
                        console.log(cellValue)
                        // If it is, update the cell value with the result property of the object
                        row.getCell(cellName).value = cellValue.result;
                    }
                    else if(typeof cellValue == Number){
                        row.getCell(cellName).value = cellValue;
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