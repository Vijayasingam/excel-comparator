import ExcelParser from "./src/ExcelParser";
import Comparator from "./src/Comparator";

var commandLineArgs = process.argv.slice(2);

if (commandLineArgs[0] && commandLineArgs[1]) {
    const excelParser = new ExcelParser();
    const comparator = new Comparator();
    const bookOne = excelParser.convertExcelToJson({
        sourceFile: commandLineArgs[0],
    })
    const bookTwo = excelParser.convertExcelToJson({
        sourceFile: commandLineArgs[1],
    })
    const differences = comparator.compare(bookOne, bookTwo);
    console.log(differences);
} else {
    console.log('Missing Files to be comapared');
}
