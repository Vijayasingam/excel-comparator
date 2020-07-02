import * as XLSX from 'xlsx';
import merge from 'lodash.merge';
class ExcelParser {
    _config: any;
    getCellRow = (cell: string) => Number(cell.replace(/[A-z]/gi, ''));
    getCellColumn = (cell: string) => cell.replace(/[0-9]/g, '').toUpperCase();
    getSheetCellValue = (sheetCell: { t: string; v: string; w: string; }) => {
        if (!sheetCell) {
            return undefined;
        }
        if (sheetCell.t === 'z' && this._config.sheetStubs) {
            return null;
        }
        return (sheetCell.t === 'n' || sheetCell.t === 'd') ? sheetCell.v : (sheetCell.w && sheetCell.w.trim && sheetCell.w.trim()) || sheetCell.w;
    };
    parseSheet = (sheetData: { constructor: StringConstructor; name: any; columnToKey: any; range: any; header: { rows: any; rowToKeys: any; }; appendData: any; }, workbook: { Sheets: { [x: string]: any; }; }) => {
        const sheetName = (sheetData.constructor == String) ? sheetData : sheetData.name;
        const sheet = workbook.Sheets[sheetName];
        const columnToKey = sheetData.columnToKey;
        const headerRows = (sheetData.header && sheetData.header.rows);
        const headerRowToKeys = (sheetData.header && sheetData.header.rowToKeys);

        let rows: any[] = [];
        for (let cell in sheet) {

            // !ref is not a data to be retrieved || this cell doesn't have a value
            if (cell == '!ref' || (sheet[cell].v === undefined && !(this._config.sheetStubs && sheet[cell].t === 'z'))) {
                continue;
            }

            const row = this.getCellRow(cell);
            const column = this.getCellColumn(cell);
            // Is a Header row
            if (headerRows && row <= headerRows) {
                continue;
            }

            // This column is not this._configured to be retrieved
            if (columnToKey && !(columnToKey[column] || columnToKey['*'])) {
                continue;
            }

            const rowData = rows[row] = rows[row] || {};
            let columnData = (columnToKey && (columnToKey[column] || columnToKey['*'])) ?
                columnToKey[column] || columnToKey['*'] :
                (headerRowToKeys) ?
                `{{${column}${headerRowToKeys}}}` :
                column;

            let dataVariables = columnData.match(/{{([^}}]+)}}/g);
            if (dataVariables) {
                dataVariables.forEach((dataVariable: string) => {
                    let dataVariableRef = dataVariable.replace(/[\{\}]*/gi, '');
                    let variableValue;
                    switch (dataVariableRef) {
                        case 'columnHeader':
                            dataVariableRef = (headerRows) ? `${column}${headerRows}` : `${column + 1}`;
                            break;
                        default:
                            variableValue = this.getSheetCellValue(sheet[dataVariableRef]);
                    }
                    columnData = columnData.replace(dataVariable, variableValue);
                });
            }

            if (columnData === '') {
                continue;
            }

            rowData[columnData] = this.getSheetCellValue(sheet[cell]);
            
            if (sheetData.appendData) {
                merge(true, rowData, sheetData.appendData);
            }
        }

        // removing first row i.e. 0th rows because first cell itself starts from A1
        rows.shift();

        // Cleaning empty if required
        if (!this._config.includeEmptyLines) {
            rows = rows.filter(v => v !== null && v !== undefined);
        }

        return rows;
    };

    convertExcelToJson = (config: any) => {
        this._config = config;
        let workbook: any = {};
        workbook = XLSX.readFile(this._config.sourceFile, {
            sheetStubs: true,
            cellDates: true
        });
        let parsedData: any = {};
        Object.keys(workbook.Sheets).forEach((sheet: any) => {
            sheet = (sheet.constructor == String) ? {
                name: sheet
            } : sheet;
            parsedData[sheet.name] = this.parseSheet(sheet, workbook);
        });
        return parsedData;
    };
}

export default ExcelParser;