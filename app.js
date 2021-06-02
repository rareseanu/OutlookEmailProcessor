var ExcelFile = require("./src/utils/excelUtils.js")

const express = require('express')
const bodyParser = require('body-parser')
var devCerts = require('office-addin-dev-certs');
const app = express()
const fs = require('fs')
const https = require('https');

var jsonParser = bodyParser.json()
app.use(express.static('src'));
app.use(express.static('node_modules'));
app.use(express.static('assets'));

// Config file
const configPath = "config.json";
let config = JSON.parse(fs.readFileSync(configPath, 'utf8'));

app.post('/updateConfig', jsonParser, (req, res) => {
    let port = req.body.port;
    let domain = req.body.domain;
    let excel = req.body.excel;
    let environment = req.body.environment;

    config.port = port;
    config.domain = domain;
    config.excel = excel;
    config.environment = environment;

    fs.writeFileSync(configPath, JSON.stringify(config));
})

app.post('/updateExcel', jsonParser, (req, res) => {
    let email = req.body.email;
    let dates = req.body.dates;
    console.log(email);
    console.log(dates);
    let excel = new ExcelFile(__dirname + config.excel);
    excel.loadExcelFile()
        .then(() => {
            return excel.addEmail(0, 1, email, dates);
        })
        .then((data2) => {
            fs.writeFile(__dirname + config.excel, Buffer.from(data2), (err) => {
                if (err) {
                    console.log(err)
                }
            });
        })
})

app.get('/link.html', (req, res) => {
    res.sendFile(__dirname + '/src/taskpane/link.html');
})

app.get('/simulator.html', (req, res) => {
    res.sendFile(__dirname + '/src/simulator/simulator.html');
})

app.get('/taskpane.html', (req, res) => {
    res.sendFile(__dirname + '/src/taskpane/taskpane.html');
})

app.get('/configuration.html', (req, res) => {
    res.sendFile(__dirname + '/src/configuration/configuration.html');
})

app.get('/config.json', (req, res) => {
    res.sendFile(__dirname + '/config.json');
})

async function startServer(port) {
    if (config.environment === 'development') {
        const options = await devCerts.getHttpsServerOptions();
        https.createServer(options, app).listen(port, () => console.log(`Server running on ${port}`));
    }
    else {
        app.listen(config.port || 3000, () => console.log(`Server listening on port 3000`));
    }
}

startServer(config.port);

module.exports = app;