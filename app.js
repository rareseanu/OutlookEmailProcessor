const express = require('express')
const bodyParser = require('body-parser')
const ExcelJS = require('exceljs');
const app = express()
const fs = require('fs')
const port = 3000
const https = require('https')

var jsonParser = bodyParser.json()
app.use(express.static('src'));
app.use(express.static('node_modules'));
app.use(express.static('assets'));

app.post('/update', jsonParser, (req, res) => {
    console.log(req.body.excel.data);
    let workbook = new ExcelJS.Workbook();
    fs.writeFile(__dirname + '/src/simulator/Book1.xlsx', Buffer.from(req.body.excel.data), (err) => {
        if (err) {
            console.log(err)
        }
    });
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