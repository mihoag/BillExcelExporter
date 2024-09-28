const ExcelJS = require('exceljs');
const ConfigStyle = require('./ConfigStyle');
const DataCellTable = require('./DataCellTable');

class ExporterService {
    constructor(filename, mau, jtt, jct, jft ) {
        this.filename = filename;
        this.mau = mau;
        this.jttMap = jtt; // Assume jtt is a single object in an array
        this.listJctMap = jct; // jct is an array of objects
        this.jftMap = jft; // Assume jft is a single object in an array
        this.workbook = new ExcelJS.Workbook();
        this.totalPerProductPositions = [];
    }

    sortJct()
    {
        
    }

    async exportToExcel(res) {
        const filePath = `Templates/${this.mau}.xlsx`;
        await this.workbook.xlsx.readFile(filePath);
        const sheet = this.workbook.getWorksheet(1);

        this.fillGeneralData(sheet, this.jttMap);

        const configEntry = ConfigStyle.getConfig()[this.mau];
      
        let startRowNum = configEntry.headerRow;
        
        //Set stt for Jct
        this.setSttForJct();
        
        this.parseNumberInJct(configEntry);
    
        this.jftMap = this.resetJfTt();

        
        sheet.spliceRows(startRowNum + 1, 0, ...Array(this.listJctMap.length).fill([]));

        this.updateFormula(sheet, startRowNum,this.listJctMap.length);
    
        this.fillOrderData(sheet, configEntry, startRowNum);

        this.setFormularOrderData(sheet, configEntry);

    
        this.setTotalFomular(sheet, configEntry, startRowNum + 1 + this.listJctMap.length);

        /*
        await this.calculateAfterFormula(sheet, configEntry);
        */
        
    
        if(!this.filename.endsWith(".xlsx"))
        {
            this.filename = this.filename + ".xlsx";
        }
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(this.filename)}`);
        await this.workbook.xlsx.write(res);
        res.end();
    }

    updateFormula(sheet, startIndex, sizeOrders)
    {
        for (let rowIndex = startIndex + 1 + sizeOrders; rowIndex <= sheet.rowCount; rowIndex++) {
            const row = sheet.getRow(rowIndex);
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                if (cell.formula) {
                    // Adjust the formula references
                    const updatedFormula = this.adjustFormulaReferences(cell.formula, sizeOrders);
                    cell.value = { formula : updatedFormula};
                }
            });
        }
    }

    adjustFormulaReferences(formula,numRows) {
        // Basic example of updating formula row references
        // This will need to be more complex depending on formula syntax
        return formula.replace(/(?<=[A-Z])\d+/g, (row) => {
            const rowNumber = parseInt(row, 10);
            return  rowNumber + numRows;
        });
    }

    captureInitialColumnWidths(sheet) {
        const columnWidths = new Map();
        sheet.columns.forEach((col, index) => {
            columnWidths.set(index, col.width);
        });
        return columnWidths;
    }

    restoreColumnWidths(sheet, columnWidths) {
        columnWidths.forEach((width, index) => {
            sheet.getColumn(index + 1).width = width;
        });
    }

    fillGeneralData(sheet, data) {
        for (var key of data.keys()) {
            const { rowIndex, colIndex } = this.convertCellPosition(key);
            const cell = sheet.getRow(rowIndex).getCell(colIndex);
            cell.value = data.get(key);
        }
    }

    setSttForJct() {
        let count = 1;
        for(var i = 0 ; i < this.listJctMap.length ; i++) {
            for (const key of this.listJctMap[i].keys()) {
                if (this.listJctMap[i].get(key) === 'stt') {
                    this.listJctMap[i].set(key,count++);
                }
            }
        }
    }

    parseNumberInJct(configEntry) {

        for(var i = 0 ; i < this.listJctMap.length ; i++) {
            for (const key of this.listJctMap[i].keys()) {
                if (configEntry.format_number.includes(key)) {
                
                    this.listJctMap[i].set(key,parseFloat(this.listJctMap[i].get(key)));
                }
            }
        }
    }

    resetJfTt() {
        const newJftMap = {};
        for (const key of this.jftMap.keys()) {
            const colLetter = key.replace(/\d+/g, '');
            const rowIndex = key.replace(/\D+/g, '',10);
            const newRowIndex = parseInt(rowIndex) + this.listJctMap.length;
            const newPosition = `${colLetter}${newRowIndex}`;
            newJftMap[newPosition] = this.jftMap.get(key);
        }
        return newJftMap;
    }

    setConfigOrderCell(cell, position, configEntry) {
        const columnLetters = position.replace(/\d/g, '');
        const font = {
            name: 'Times New Roman',
            size: 11,
            color: { argb: 'FF000000' } // Default to black
        };
    
        // Set alignments
        if (configEntry.alignment[columnLetters]) {
            const alignmentValue = configEntry.alignment[columnLetters];
            if (alignmentValue === 'right') {
                cell.alignment = { wrapText: true, 
                    vertical: 'middle' , horizontal: 'right' };
            } else if (alignmentValue === 'left') {
                cell.alignment = {wrapText: true, 
                    vertical: 'middle' , horizontal: 'left' };
            }
        } else {
            cell.alignment = {wrapText: true, 
                vertical: 'middle' , horizontal: 'center' };
        }
    
        // Set color
        if (configEntry.fontColor[columnLetters]) {
            const colorValue = configEntry.fontColor[columnLetters];
            font.color = { argb: colorValue };
        }
    
        // Set number format
        if (configEntry.format_number.includes(columnLetters)) {
            cell.numFmt = '#,##0';
        }
        // Apply the font to the cell
        cell.font = font;
    }


    fillOrderData(sheet, configEntry, startRowNum) {
        startRowNum++;
        for (const map of this.listJctMap) {
            this.totalPerProductPositions.push(startRowNum);
            let maxRowHeight = 0;  // Variable to store the maximum row height

            for (const key of map.keys()) {
                const position = `${key}${startRowNum}`;
                const { rowIndex, colIndex } = this.convertCellPosition(position);
                const row = sheet.getRow(rowIndex);
                const cell = row.getCell(colIndex);
                cell.value = map.get(key);
                cell.style = {...DataCellTable};
                this.setConfigOrderCell(cell,position,configEntry);
                // Calculate the height for this cell
                const cellHeight = this.calculateCellHeight(sheet, cell, colIndex);
                if (cellHeight > maxRowHeight) {
                   maxRowHeight = cellHeight;  // Update maxRowHeight if this cell's height is greater
                }
            }
            const row = sheet.getRow(startRowNum);
            row.height = maxRowHeight;
            startRowNum++;
        }
    }

    calculateCellHeight(sheet, cell, colIndex) {
        const columnWidth = sheet.getColumn(colIndex).width;
        const fontSize = 11; // Assume a default font size, adjust as needed
        const charWidth = 1.35; // Approximate width of a character at default font size
        
        const content = cell.value || '';
        const numLines = Math.ceil(content.length / (columnWidth / charWidth));
        const rowHeight = numLines * fontSize * 1.2; // 1.2 is a scaling factor for line spacing
    
        return rowHeight;
    }

    setFormularOrderData(sheet, configEntry) {
        for (const rowNum of this.totalPerProductPositions) {
            const colIndex = this.convertColumnLetter(configEntry.formulaMultiply[2]);
            const cell = sheet.getRow(rowNum).getCell(colIndex);

            const col1 = `${configEntry.formulaMultiply[0]}${rowNum}`;
            const col2 = `${configEntry.formulaMultiply[1]}${rowNum}`;
            cell.value = { formula: `${col1}*${col2}` };
        }
    }

    setTotalFomular(sheet, configEntry, startRowNum) {
        const colIndex = this.convertColumnLetter(configEntry.formulaMultiply[2]);
        const cell = sheet.getRow(startRowNum).getCell(colIndex);

        const param1 = `${configEntry.formulaMultiply[2]}${this.totalPerProductPositions[0]}`;
        const param2 = `${configEntry.formulaMultiply[2]}${this.totalPerProductPositions[this.totalPerProductPositions.length - 1]}`;
        cell.value = { formula: `SUM(${param1}:${param2})` };
    }

    convertCellPosition(position) {
        const colIndex = this.convertColumnLetter(position.replace(/\d+/g, ''));
        const rowIndex = position.replace(/\D+/g, '',10);
        return { rowIndex, colIndex };
    }

    convertColumnLetter(letter) {
        return letter.charCodeAt(0) - 'A'.charCodeAt(0)+1;
    }
}

module.exports = ExporterService;
