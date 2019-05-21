var unoconv = require('unoconv');
var docxConverter = require('docx-pdf');
var fs = require('fs');*/

unoconv.convert('/OLAP.docx', 'pdf', {bin: 'unoconv'},function (err, result) {
	// result is returned as a Buffer
	fs.writeFile('converted.pdf', result);
});

//не unoconv, но работает :D
/*docxConverter('C:/Projects/testView/OLAP.docx','./output.pdf',function(err,result){
    if(err){
      console.log(err);
    }
    console.log('result'+result);
  });*/


  /*var converter = require('office-converter')();
  converter.generatePdf('./OLAP.docx', function(err, result) {
    // Process result if no error
      console.log('Output File located at ' + result.outputFile);
  });*/

/*const path = require('path');
const unoconv = require('awesome-unoconv');
 
const sourceFilePath = path.resolve('./OLAP.docx');
const outputFilePath = path.resolve('./myDoc.pdf');
 
unoconv
  .convert(sourceFilePath, outputFilePath)
  .catch(err => {
    console.log(err);
  });*/


  //const unoconv = require("unoconv-promise");

  /*unoconv
  .run({
    bin: 'C:/Program Files (x86)/unoconv-0.8.2/unoconv',
    file: "./OLAP.docx",
    output: "./temp.pdf",
    export: "PageRange=1-2"
  })
  .then(filePath => {
    console.log(filePath);
  })
  .catch(e => {
    throw e;
  });*/

  /*unoconv
  .run({
    file: "./OLAP.docx",
    output: "./temp.pdf",
    export: "PageRange=1-1"
  })
  .then(filePath => {
    console.log(filePath);
  })
  .catch(e => {
    throw e;
  });*/