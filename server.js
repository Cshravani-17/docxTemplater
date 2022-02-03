const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");

// Load the docx file as binary content
const content = fs.readFileSync(
    path.resolve(__dirname, "input.docx"),
    "binary"
);
const stri ="I undersigned,{nameOfSignee}, {{{designationOfSignee}}} certify that:.\n{{{Salutation}}} {{{First Name}}} {{{Last Name}}} is employed by our Company since {{{Hire Date}}}, currently employed by the establishment of {{{Address Establishment}}}.<.>\nTo date, {{{Salutation}}} {{{First Name}}}{{{Last Name}}} is employed under permanent contract as {{{Job Title Local}}}<.>\n{{{Salutation}}} {{{First Name}}} {{{Last Name}}} is not on probation period, nor notice period of resignation, nor dismissal.\nThis certificate is issued at the request of the person to serve and to assert that right."
const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

// Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
doc.render({
    body : stri
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    // compression: DEFLATE adds a compression step.
    // For a 50MB output document, expect 500ms additional CPU time
    compression: "DEFLATE",
});

// buf is a nodejs Buffer, you can either write it to a file or res.send it with express for example.
fs.writeFileSync(path.resolve(__dirname, "output2.docx"), buf);