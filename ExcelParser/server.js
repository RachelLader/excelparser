var express = require('express');
var excelbuilder = require('excel-builder');
var download= require('downloader');
var path = require('path');
var xlsx2 = require('node-xlsx');

var xlsx = require('xlsx');
var bodyParser = require('body-parser');
var port = process.env.PORT || 8000;
var app = module.exports = express();
var zip = new JSZip();
var underscore= require('underscore');
app.get('/', function(req, res) {
    res.sendFile(path.join(__dirname + '/index.html'));

})

app.use(bodyParser.json());
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
    console.log('DONE', newExcel);
    createExcelFile();

}

var createExcelFile = function() {
        var patientWB = excelbuilder.createWorkbook();
        var patients = patientWB.createWorksheet({ name: 'Patient Data' });

        patients.setData(newExcel); //<-- Here's the important part

        patientWB.addWorksheet(patients);

        var data = excelbuilder.createFile(patientWB);
        downloader('PatientData.xlsx', data);
    };


readExcelFile();


app.listen(port);
console.log('Excel server is now running at port' + port);