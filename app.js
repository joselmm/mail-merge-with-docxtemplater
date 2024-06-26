// PizZip is required because docx/pptx/xlsx files are all zipped files, and
// the PizZip library allows us to load the file in memory
var d=Date.now();
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");
const topdf = require('docx2pdf-converter')




// Load the docx file as binary content
const content = fs.readFileSync(
    path.resolve(__dirname, "templates\\Carta ejemplo.docx"),
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
doc.render({
    Nombre: "John",
    Apellido: "Doe",
    Direccion: "La villadeleiba",
    Correo: "jhondoe@example.com",
});

// Get the zip document and generate it as a nodebuffer
const buf = doc.getZip().generate({
    type: "nodebuffer",
    // compression: DEFLATE adds a compression step.
    // For a 50MB output document, expect 500ms additional CPU time
    compression: "DEFLATE",
});

// buf is a nodejs Buffer, you can either write it to a
// file or res.send it with express for example.
var outputFilePath = path.resolve(__dirname, "output.docx");
fs.writeFileSync(outputFilePath, buf);

topdf.convert(outputFilePath,path.resolve(__dirname,'output.pdf'))

console.log("duracion ejecucion: "+(Date.now()-d))


