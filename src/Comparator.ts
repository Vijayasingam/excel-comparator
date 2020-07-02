class Comparator {
    compare(jsonObjectOne: any, jsonObjectTwo: any) {
        const differences: any = [];
        const sheetsInObjectOne = Object.keys(jsonObjectOne);
        sheetsInObjectOne.forEach((sheetName: string) => {
            const sheet: any[] = jsonObjectOne[sheetName];
            if (jsonObjectTwo[sheetName]) {
                if (sheet.length !== jsonObjectTwo[sheetName].length) {
                    console.log('Number of records in both the files are different');
                    console.log(`File one has ${sheet.length} rows but file two has ${jsonObjectTwo[sheetName].length} rows`);
                    console.log('Comparison might not be accurate.')
                }
                sheet.forEach((rows: any, index: number) => {
                    const cols = Object.keys(rows);
                    cols.forEach((col: string) => {
                        if (rows[col] !== jsonObjectTwo[sheetName][index][col]) {
                            const differentObject = {
                                one: rows,
                                two: jsonObjectTwo[sheetName][index],
                                valueInOne: rows[col],
                                valueInTwo: jsonObjectTwo[sheetName][index][col],
                                differenceInColumn: col
                            }
                            differences.push(differentObject);
                        }
                    })
                })
            } else {
                console.log('Sheet names in both the file are differnt');
                console.log('Comparison might not be accurate.')
            }
        });
        return differences.length > 0 ? differences : 'No Difference';
    }
}
export default Comparator;