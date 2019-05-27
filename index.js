var unoconv = require('unoconv2');
var docxConverter = require('docx-pdf');
var fs = require('fs');
var childProcess = require('child_process')
const Excel = require('exceljs');


var workbook = new Excel.Workbook();
workbook.xlsx.readFile('test.xlsx')
  .then(function() {
    workbook.worksheets.forEach(function(worksheet) {
      worksheet.pageSetup.fitToPage = true;
    });
    workbook.xlsx.writeFile('test2.xlsx')
      .then(function() {
        console.log('write done')
        unoconv.convert(workbook, 'pdf', {bin: 'unoconv.cmd'}, function (err, result) {
          fs.writeFileSync('convertedWithMutated.pdf', result);
        });
        
        /*unoconv.convert('test.xlsx', 'pdf', {bin: 'unoconv.cmd'}, function (err, result) {
          fs.writeFileSync('convertedWithoutMutated.pdf', result);
        });*/
      });
  });






/*var args = [
  '-fpdf', 'OLAP.docx'
];

childProcess.spawn('unoconv.cmd', args)*/

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


  /*const unoconv = require("unoconv-promise");

  unoconv
  .run({
    bin: "unoconv.cmd",
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


  //msoffice
  /*var msopdf = require('node-msoffice-pdf');


  msopdf(null, function(error, office) { 
 
    if (error) { 
      console.log("Init failed", error);
      return;
    }
 
 
   office.word({input: "OLAP.docx", output: "outfile.pdf"}, function(error, pdf) { 
      if (error) { 
           console.log("Woops", error);
       } else { 
           console.log("Saved to", pdf);
       }
   });
 
 
   office.excel({input: "infile.xlsx", output: "outfile.pdf"}, function(error, pdf) { 
       if (error) { 
           console.log("Woops", error);
       } else { 
           console.log("Saved to", pdf);
       }
   });
 
   office.close(null, function(error) { 
       if (error) { 
           console.log("Woops", error);
       } else { 
           console.log("Finished & closed");
       }
   });
});*/



/*var office2pdf = require(office2pdf),
  generatePdf = office2pdf.generatePdf;

generatePdf('OLAP.docx', function(err, result) {
  console.log(result);
});*/
