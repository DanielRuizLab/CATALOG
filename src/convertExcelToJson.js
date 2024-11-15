const xlsx = require('xlsx');
const fs = require('fs');

const workbook = xlsx.readFile(
    "C:\\Users\\AV-30580\\Desktop\\CATALOG\\Datexce\\Catálogo 03 de oct. 2024.xlsx"
);

const jsonData = {};
workbook.SheetNames.forEach((sheetName) => {
    const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    jsonData[sheetName] = sheetData;
});

fs.writeFileSync('./Datexce/catalogo.json', JSON.stringify(jsonData, null, 2));
console.log("Archivo JSON creado con éxito en './Datexce/catalogo.json'");
