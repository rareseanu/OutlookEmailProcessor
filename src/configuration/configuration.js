let config;

async function getConfig(url) {
    const response = await fetch(url);

    return response.json();
}

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";

        getConfig("/config.json")
            .then(json => {
                config = json;
                document.getElementById("port").value = config.port;
                document.getElementById("domain").value = config.domain;
                document.getElementById("excel").value = config.excel;
            })
        document.getElementById("run").onclick = run;
    }
});

async function run() {
    let port = document.getElementById("port").value;
    let domain = document.getElementById("domain").value;
    let excel = document.getElementById("excel").value;

    var xhr = new XMLHttpRequest();
    xhr.open("POST", "https://localhost:3000/updateConfig", true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    let json = JSON.stringify({ port: port, domain: domain, excel: excel});
    xhr.send(json);
}