//DEPENDENCIAS NATIVAS
const path = require("path");
const fs = require("fs");

//DEPENCIAS DEL PARSER DE EXCEL
//LA IMPORATCION DESDE "excel" module no sirvio (version 1.0.1)
const parseXlsx = require(path.resolve(__dirname, "node_modules\\excel\\excelParser.js"));
//Y EL ARCHIVO excelParser.js se ajusto para que fuera commonjs en la importacion de dependencias y en su exportacion:
/* 
const fs = require('fs');
const Stream = require('stream');
const unzip = require('unzipper');
const xpath = require('xpath');
const XMLDOM = require('xmldom');
 */

/* 
module.exports = function parseXlsx(path, sheet) {
  if (!sheet) sheet = '1';
  return extractFiles(path, sheet).then(function(files) {
    return extractData(files);
  });
};
*/

//DEPENDENCIA DEL DOCXTEMPLATER
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
//DEPENDENCIAS DEL CONVERTIDOR PDF
const topdf = require('docx2pdf-converter')



//PROCESO PRINCIPAL

async function mainProcess() {
    //aqui va la ruta del archivo excel
    var excelName = "base.xlsx";
    const data = await parseXlsx(path.resolve(__dirname, 'templates\\'+excelName));
    const parsedData = parseData(data);
    console.log(parsedData)

    // CREAR ARCHIVOS DOCX Y PDF

    parsedData.forEach(rowObject => {
        //SE PUEDE ELEGIR O NO CREAR LOS PDF CON EL ULTIMO PARAMETRO
        crearDocx("plantilla diploma.docx",rowObject["Cédula"], rowObject, true)    
    });

}

mainProcess()
/*   .then()
  .catch(e => { console.error(e) })
*/

function parseData(dataArray) {
    const headers = dataArray.shift(); // extraer la primera fila (nombres de columnas)
    const result = dataArray.map((row) => {
        return headers.reduce((obj, header, index) => {
            obj[header] = row[index];
            return obj;
        }, {});
    });
    return result
}

function crearDocx(docxTemplateName, outputFileName, rowObject, onlyDocx) {

    // Load the docx file as binary content
    const content = fs.readFileSync(
        path.resolve(__dirname, "templates\\"+docxTemplateName),
        "binary"
    );

    // Unzip the content of the file
    const zip = new PizZip(content);

    // This will parse the template, and will throw an error if the template is
    // invalid, for example, if the template is "{user" (no closing tag)
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
    doc.render(rowObject);

    // Get the zip document and generate it as a nodebuffer
    const buf = doc.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });

    // buf is a nodejs Buffer, you can either write it to a
    // file or res.send it with express for example.
    var outputFilePath = path.resolve(__dirname,"Docx\\" +outputFileName+".docx");
    fs.writeFileSync(outputFilePath, buf);

    if(onlyDocx) return;

    topdf.convert(outputFilePath,path.resolve(__dirname, "Pdf\\" +outputFileName+".pdf"))

}