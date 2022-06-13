const XLSX = require('xlsx');
const AEPS = "AEPS";
const FEE_CHG = "FEE CHG";

function readFile(file) {
    if (!file) {
      console.log("File cannot be null");
      return null;
    }

    const inputFile = XLSX.readFile(file);
    const workSheet = inputFile.Sheets[inputFile.SheetNames[0]];
    const jsonSheet = XLSX.utils.sheet_to_json(workSheet);

    if (!jsonSheet.length > 0) {
      console.log("No data in the input file")
      return null;
   }

   let jsonData = [];
   for (let i = 0; i < jsonSheet.length; i++) {
      const keys = Object.keys(jsonSheet[i]).values().next().value.split('|');
      const values = Object.values(jsonSheet[i]).values().next().value.split('|');
      jsonRow = Object.assign.apply(
        {},
        keys.map((v, i) => ({ [v]: values[i] }))
      );
      jsonData.push(jsonRow);
   }
    return jsonData
}

function parseAndEditData(jsonData) {
  const updatedData = jsonData.map((prop) => {
      if (prop.Description.match(FEE_CHG)) {
        prop.Flag = FEE_CHG;
      } else if (prop.Description.match(AEPS)) {
        prop.Flag = AEPS;
      } else prop.Flag = "";
      [prop["Description_Delimited"], prop[""]] = prop.Description.split("/");
    });

  return updatedData;
}

function writeJsonDataToExcel(outputfile) {
  const newWorkBook = XLSX.utils.book_new();
  const newSheet = XLSX.utils.json_to_sheet(jsonData);
  XLSX.utils.book_append_sheet(newWorkBook, newSheet, "Output Data");
  XLSX.writeFile(newWorkBook, outputfile);

}

const jsonData = readFile('./files/input.xlsx');
if(jsonData) {
  updatedData = parseAndEditData(jsonData)
  writeJsonDataToExcel("./files/output.xlsx")
}
