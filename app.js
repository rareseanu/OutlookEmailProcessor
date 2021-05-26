var ExcelFile = require("./src/utils/excelUtils.js")

const express = require('express')
const bodyParser = require('body-parser')
const ExcelJS = require('exceljs');
const app = express()
const fs = require('fs')
const https = require('https')

var jsonParser = bodyParser.json()
app.use(express.static('src'));
app.use(express.static('node_modules'));
app.use(express.static('assets'));

// Procesare pe server
// Trimitere doar parametrii de data & email
// Config in UI.

app.post('/update', jsonParser, (req, res) => {
    console.log(req.body.excel.data);
    fs.writeFile(__dirname + '/src/simulator/Book1.xlsx', Buffer.from(req.body.excel.data), (err) => {
        if (err) {
            console.log(err)
        }
    });
})

app.post('/updateExcel', jsonParser, (req, res) => {
    let email = req.body.email;
    let dates = req.body.dates;
    console.log(email);
    console.log(dates);
    let excel = new ExcelFile(__dirname + '/src/simulator/Book1.xlsx');
    excel.loadExcelFile()
    .then(() => {
      return excel.addEmail(0, 1, email, dates);
    })
    .then((data2) => {
        fs.writeFile(__dirname + '/src/simulator/Book1.xlsx', Buffer.from(data2), (err) => {
            if (err) {
                console.log(err)
            }
        });
    })
})

app.get('/link.html', (req, res) => {
    res.sendfile(__dirname + '/src/taskpane/link.html');
})

app.get('/simulator.html', (req, res) => {
    res.sendfile(__dirname + '/src/simulator/simulator.html');
})

app.get('/taskpane.html', (req, res) => {
    res.sendfile(__dirname + '/src/taskpane/taskpane.html');
})

https.createServer({
    key: fs.readFileSync('localhost.key'),
    cert: fs.readFileSync('localhost.crt')
  }, app)
  .listen(3000, function () {
    console.log('Example app listening on port 3000! Go to https://localhost:3000/')
  })

module.exports = app;