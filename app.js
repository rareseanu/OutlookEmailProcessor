var ExcelFile = require("./src/utils/excelUtils.js")

const express = require('express')
const bodyParser = require('body-parser')
const app = express()
const fs = require('fs')
const https = require('https')

var jsonParser = bodyParser.json()
app.use(express.static('src'));
app.use(express.static('node_modules'));
app.use(express.static('assets'));

// Config file
const configPath = "config.json";
const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
const port = config.port;
const excelFilePath = config.excel;

app.post('/updateConfig', jsonParser, (req, res) => {
    let port = req.body.port;
    let domain = req.body.domain;
    let excel = req.body.excel;
    console.log(req.body);

    config.port = port;
    config.domain = domain;
    config.excel = excel;

    fs.writeFileSync(configPath, JSON.stringify(config));
})

app.post('/updateExcel', jsonParser, (req, res) => {
    let email = req.body.email;
    let dates = req.body.dates;
    console.log(email);
    console.log(dates);
    let excel = new ExcelFile(__dirname + excelFilePath);
    excel.loadExcelFile()
    .then(() => {
      return excel.addEmail(0, 1, email, dates);
    })
    .then((data2) => {
        fs.writeFile(__dirname + excelFilePath, Buffer.from(data2), (err) => {
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

app.get('/configuration.html', (req, res) => {
    res.sendfile(__dirname + '/src/configuration/configuration.html');
})

app.get('/config.json', (req, res) => {
    res.sendfile(__dirname + '/config.json');
})

https.createServer({
    key: fs.readFileSync('localhost.key'),
    cert: fs.readFileSync('localhost.crt')
  }, app)
  .listen(port, function () {
    console.log(`Email processor server listening on port ${port}!`)
  })

module.exports = app;