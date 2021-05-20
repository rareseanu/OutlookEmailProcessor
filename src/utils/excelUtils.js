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
        let row = worksheet.getRow(rowNo);
        return row.getCell(columnNo);
    }

    colFirstEmpty(worksheetNo, startPos, colNo) {
        let worksheet = this.workbook.worksheets[worksheetNo];
        let col = worksheet.getColumn(colNo);
        let resRowNumber = 1;
        col.eachCell({ includeEmpty: true }, function(cell, rowNumber) {
            resRowNumber = rowNumber;
            console.log("Cell " + col + " on Row " + rowNumber + " = " + cell.value);
        });

        return this.getCellObject(worksheetNo, resRowNumber + 1, colNo);
    }

    findValueColumn(worksheetNo, colNo, value) {
        let worksheet = this.workbook.worksheets[worksheetNo];
        let col = worksheet.getColumn(colNo);
        let celltemp = null;
        col.eachCell({ includeEmpty: true }, function(cell, rowNumber) {
            if(cell.value == value) {
                console.log("GASIT" + cell.value);
                celltemp = cell;
            }
        });
        return celltemp;
    }

    getDateFromCell(cell) {
        if(cell != null) {
            var parts = cell.value.split("-");
            var dt = new Date(parseInt(parts[2], 10),
                  parseInt(parts[1], 10) - 1,
                  parseInt(parts[0], 10));
            return dt;
        }
        return null;
    }

    dateToString(date) {
        let month = date.getMonth() + 1;
        return date.getDate()+ '-' + month + '-' + date.getFullYear();
    }

    addDays(date, days) {
        var result = new Date(date);
        result.setDate(result.getDate() + days);
        return result;
      }

    async addEmail(worksheetNo, colNo, email, dates) {
        let emailFoundCell = this.findValueColumn(worksheetNo, colNo, email);
        let currentDateCell = this.getCellObject(0, 1, 2);
        let startDate = dates[0];
        let endDate = dates[1];

        if(emailFoundCell == null) {
            emailFoundCell = this.colFirstEmpty(worksheetNo, 2, 1);
            emailFoundCell.value = email;
        }

        if(emailFoundCell != null) {
            if(this.getCellObject(0, 1, 2).value == null) {
                currentDateCell.value = this.dateToString(startDate);
                console.log(currentDateCell.value);
            }
            let currentDate = this.getDateFromCell(currentDateCell);
            if(currentDate >= startDate && currentDate <= endDate) {
                this.getCellObject(0, emailFoundCell.row, currentDateCell.col).value = 'x';
            }

            while(currentDate < endDate) {
                currentDateCell = this.getCellObject(0, 1, currentDateCell.col + 1);

                if(currentDateCell.value == null) {
                    currentDateCell.value = 
                        this.dateToString(this.addDays(this.getDateFromCell(this.getCellObject(0, 1, currentDateCell.col - 1)), 1));
                }
                currentDate = this.getDateFromCell(currentDateCell);
                console.log(currentDateCell.value);
                if(currentDate >= startDate) {
                    this.getCellObject(0, emailFoundCell.row, currentDateCell.col).value = 'x';
                }
            }
        }
        return await this.workbook.xlsx.writeBuffer();
    }

    rowFirstEmpty(worksheetNo, startPos, rowNo) {
        let worksheet = this.workbook.worksheets[worksheetNo];
        let row = worksheet.getRow(rowNo);
        let celltemp = null;
        let foundFirst = false;
        row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
            if(colNumber >= startPos && cell.value == null && !foundFirst) {
                celltemp = cell;
                foundFirst = true;
            }
        });
        return celltemp;
    }

    async writeToCell(worksheetNo, rowNo, columnNo, value) {
        let cell = this.getCellObject(worksheetNo, rowNo, columnNo);
        cell.value = value;
    
        return await this.workbook.xlsx.writeBuffer();
    }
    
}