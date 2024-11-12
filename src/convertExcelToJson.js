const xlsx = require('xlsx');
const fs = require('fs');

// Carga el archivo Excel
const workbook = xlsx.readFile(
    "C:\\Users\\AV-30580\\Desktop\\CATALOG\\Datexce\\Catálogo actualizado 05 de sep.xlsx"
);

// Convierte cada hoja en el archivo Excel a JSON
const jsonData = {};
workbook.SheetNames.forEach((sheetName) => {
    const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    jsonData[sheetName] = sheetData;
});

// Guarda el JSON en el proyecto
fs.writeFileSync('./Datexce/catalogo.json', JSON.stringify(jsonData, null, 2));
console.log("Archivo JSON creado con éxito en './Datexce/catalogo.json'");
