var express = require('express');
var excelbuilder = require('msexcel-builder');
var download= require('downloader');
var path = require('path');
var xlsx2 = require('node-xlsx');

var xlsx = require('xlsx');
var port =  8000;
var app = module.exports = express();
var underscore= require('underscore');
app.get('/', function(req, res) {
    res.sendFile(path.join(__dirname + '/index.html'));

})

var newExcel = []
var idStorage = [];

var readExcelFile = function(file) {

    var workbook = xlsx2.parse('MasterSheet.xlsx');
    var data = workbook[0].data;
    for (var x = 0; x < data.length; x++) {
        var person = data[x]
        if (idStorage.indexOf(+person[3]) === -1) {
            idStorage.push(+person[3]);
            newExcel.push(person);
        }
    }
    console.log('DONE');
    createExcelFile();

}

var createExcelFile = function() {
        var workbook = excelbuilder.createWorkbook('./','newData.xlsx');
        var patients = workbook.createSheet('PatientData',3000, 2000);
        for(var x=0; x<newExcel.length;x++){
           var patient=newExcel[x];
           // console.log('patient', patient)
           for(var i=1; i<patient.length+1;i++){
               console.log('length', i-1, patient[i-1])

            // console.log('patient:', patient[i-1])
            patients.set(i,x+1,patient[i-1]);
           }
        }
        console.log('done creating file');
workbook.save(function(err){
    if (err)
      throw err;
    else
      console.log('congratulations, your workbook created');
  });
};


readExcelFile();


app.listen(port);
console.log('Excel server is now running at port' + port);