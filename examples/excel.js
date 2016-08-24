//
// Requires
//
var sys = require('util');
var fs = require('fs');
var vm = require('vm');
var path = require('path');

var Enumerable = require('../index').Enumerable;
var Ltxml = require('../index').Ltxml;
var openXml = require('../index').openXml;

//
// Aliases
//
var XAttribute = Ltxml.XAttribute;
var XCData = Ltxml.XCData;
var XComment = Ltxml.XComment;
var XContainer = Ltxml.XContainer;
var XDeclaration = Ltxml.XDeclaration;
var XDocument = Ltxml.XDocument;
var XElement = Ltxml.XElement;
var XName = Ltxml.XName;
var XNamespace = Ltxml.XNamespace;
var XNode = Ltxml.XNode;
var XObject = Ltxml.XObject;
var XProcessingInstruction = Ltxml.XProcessingInstruction;
var XText = Ltxml.XText;
var XEntity = Ltxml.XEntity;
var cast = Ltxml.cast;
var castInt = Ltxml.castInt;

var S = openXml.S;
var R = openXml.R;

var templateFileName = 'sample.xlsx';
var templateExt = path.extname(templateFileName);
var templateBasename = path.basename(templateFileName, templateExt);

//
// Main
//
var beginTime = (new Date()).getTime();
var doc = new openXml.OpenXmlPackage(fs.readFileSync(templateFileName));
var workbookPart = doc.workbookPart();
var wbXDoc = workbookPart.getXDocument();

var sharedStringTablePart = workbookPart.sharedStringTablePart();
var sst = sharedStringTablePart.getXDocument().root;

var strings = sst.elements(S.si).select(function(si) {
    return si.descendants(S.t).aggregate("", function(text, t) {
        return text += t.value.replace(/\n/g, '\\n');
    });
}).toArray();

wbXDoc.root.element(S.sheets).elements(S.sheet).forEach(function(sheet) {
    console.log(sheet.attribute('name').value);

    var id = sheet.attribute(R.id).value;
    var worksheetPart = workbookPart.getPartById(id);
    var wsXDoc = worksheetPart.getXDocument();
    wsXDoc.descendants(S.row).forEach(function(row) {
        row.descendants(S.c).forEach(function(cell) {
            var msg = '';
            msg += 'r: ' + (cell.attribute('r') ? cell.attribute('r').value : '');
            msg += ', t: ' + (cell.attribute('t') ? cell.attribute('t').value : '');
            msg += ', v=' + (cell.element(S.v) && cell.element(S.v).value);
            msg += ' ' + (cell.attribute('t') && cell.attribute('t').value == 's' ? strings[cell.element(S.v).value] : '');
            console.log(msg);
        });
        console.log('-----');
    });
});

var theContent = doc.saveToBase64();
var buffer = new Buffer(theContent, 'base64');

var fileName = templateBasename + '_processed' + templateExt;
fs.writeFileSync(fileName, buffer);


var endTime = (new Date()).getTime();
var deltaTime = endTime - beginTime;

console.log('Finished writing a document in ' + (deltaTime / 1000).toString() + ' seconds.');
