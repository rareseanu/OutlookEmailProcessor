export class ExcelFile {

    // Loads the excel file from the given path.
    constructor(filePath) {
        this.filePath = filePath;
    }

    loadExcelFile() {
        return fetch(this.filePath)
        .then((data) => {
            if(!data.ok) {
                throw data;
            }
            var test = new ExcelJS.Workbook();
            this.workbook = new ExcelJS.Workbook();
            return data.arrayBuffer();
        })
        .then(array => this.workbook.xlsx.load(array))
        .then(() => {
            this.workbook.eachSheet((sheet, id) => {
                sheet.eachRow((row, rowIndex) => {
                    console.log(row.values, rowIndex)
                })
            });
        })
        .catch(error => {
            // TODO: Logging system (file & user interface.)
            console.log(error);
        });
    }

    getCellObject(worksheetNo, rowNo, columnNo) {
        let worksheet = this.workbook.worksheets[worksheetNo];
        console.log(this.workbook.worksheets[0]);
        let row = worksheet.getRow(rowNo);
        return row.getCell(columnNo);
    }

    async writeToCell(worksheetNo, rowNo, columnNo, value) {
        let cell = this.getCellObject(worksheetNo, rowNo, columnNo);
        cell.value = value;
    
        return await this.workbook.xlsx.writeBuffer();
    }
    
}