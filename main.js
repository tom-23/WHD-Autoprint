const express = require('express')
const csv = require('csv-parser')
const fs = require('fs')
const config = require('./config.json')
const fetch = require('node-fetch')
const NodeCache = require("node-cache");
const ThermalPrinter = require("node-thermal-printer").printer;
const Types = require("node-thermal-printer").types;
const Dymo = require('dymojs')
const cliProgress = require('cli-progress');
var progress = require('progress-stream');

const app = express()
const port = config.webPort

const ticketCache = new NodeCache();
const partsCache = new NodeCache();
const dymo = new Dymo();

var lastServerCheck = "n/a";
var lastLabelPrint = "n/a";
var lastTicketNumber = 999999;
var lastSubject = "AutoPrint";
var lastSerialNumber = "XXXXXXXXXXXX";
var lastOpenDate = "XX/XX/XX XX:XX";
var lastDetails = "No tickets have been submitted :(";

var checkingInterval = 30;

var reciptFooter;

app.set('view engine', 'ejs')
app.use(express.static('public'))

app.get('/', (req, res) => {
    res.render('index', { lastServerCheck, lastLabelPrint, lastSubject });
})

app.get('/oldindex', (req, res) => {
    res.render('oldindex', { lastServerCheck, lastLabelPrint, lastSubject });
})

app.get('/ui/ticketlabel', (req, res) => {
    res.render('ticketlabel');
})

app.get('/ui/ticketreceipt', (req, res) => {
    res.render('ticketreceipt');
})

app.get('/ui/partlabel', (req, res) => {
    res.render('partlabel');
})

app.get('/ui/pricelabel', (req, res) => {
    res.render('pricelabel');
})

app.get('/print/ticketlabel', (req, res) => {
    retriveTicketAndPrint(req.query.ticketNumber, false, true);
    res.redirect("/ui/ticketlabel");
})

app.get('/print/ticketreceipt', (req, res) => {
    retriveTicketAndPrint(req.query.ticketNumber, true, false);
    res.redirect("/ui/ticketreceipt");
})

app.get('/print/partlabel', (req, res) => {
    printPartLabel(req.query.partNumber)
    res.redirect("/ui/partlabel");
})

app.get('/print/pricelabel', (req, res) => {
    printPriceLabel(req.query.price, req.query.quantity)
    res.redirect("/ui/pricelabel");
})

app.get('/updaterestart', (req, res) => {
    runUpdateScript();
    res.redirect("/");
})

app.listen(port, '0.0.0.0', () => {
    console.log(`WHD-Autoprint app listening on port ${port}`)
    loadPartsCache(function () {
        loadRecpiptFooter();
        checkServer();
        setInterval(checkServer, checkingInterval * 1000)
    });

})


function loadRecpiptFooter() {
    try {
        var data = fs.readFileSync('recipt_footer.txt', 'utf8');
        reciptFooter = data.toString();
    } catch (e) {
        console.log('Error:', e.stack);
    }
}

async function printRecipt(ticketID, subject, detail, date) {
    let printer = new ThermalPrinter({
        type: Types.STAR,  // 'star' or 'epson'
        interface: "tcp://10.10.106.104",
        options: {
            timeout: 20000
        },
        width: 48,                         // Number of characters in one line - default: 48
        characterSet: 'SLOVENIA',          // Character set - default: SLOVENIA
        removeSpecialCharacters: false,    // Removes special characters - default: false
        lineCharacter: "-",                // Use custom character for drawing lines - default: -
    });

    let isConnected = await printer.isPrinterConnected();
    console.log("Printer connected:", isConnected);

    printer.newLine();
    printer.newLine();
    printer.newLine();
    printer.append(Buffer.from([0x1b, 0x34]));
    printer.append(Buffer.from([0x1b, 0x45]));
    printer.append(Buffer.from([0x1b, 0x69, 0x01, 0x01]));
    printer.append(Buffer.from([0x1b, 0x1d, 0x61, 0x01]));
    printer.println(String(ticketID));
    printer.newLine();
    printer.append(Buffer.from([0x1b, 0x69, 0x00, 0x00]));
    printer.append(Buffer.from([0x1b, 0x35]));
    printer.append(Buffer.from([0x1b, 0x45]));
    printer.println(subject);

    Buffer.from([0x1b, 0x46])

    printer.newLine();
    printer.newLine();


    printer.append(Buffer.from([0x1b, 0x1d, 0x61, 0x00]));

    printer.print(detail);
    printer.newLine();

    printer.newLine();
    printer.newLine();

    printer.println(date.toString());


    printer.newLine();
    printer.newLine();

    printer.alignCenter();
    await printer.printImage("istore.png");

    printer.newLine();
    printer.newLine();

    printer.append(Buffer.from([0x1b, 0x64, 0x02]));

    try {
        await printer.execute();
        console.log("Print success.");
    } catch (error) {
        console.error("Print error:", error);
        console.error("Trying again...");
        printRecipt(ticketID, subject, detail, date);
    }


}


function printLabel(ticketID, serialNumber, openDate) {
    var labelData = "";
    var fileName;
    if (openDate != "") {
        fileName = "./label_template_doa.label";
    } else if (serialNumber != "") {
        fileName = "./label_template_with_serialnumber.label";
    } else {
        fileName = "./label_template_without_serialnumber.label";
    }
    fs.readFile(fileName, 'utf8', function (err, data) {
        if (err) throw err;
        labelData = data.replace("TICKETNO", ticketID).replace("SERIALNO", serialNumber).replace("OPENDATE", openDate);
        dymo.renderLabel(labelData).then(imageData => {
            fs.writeFile("./public/last_label.png", imageData, 'base64', function (err) {
            });
        });
        dymo.print('DYMO LabelWriter 450', labelData);
        let date = new Date();
        lastLabelPrint = date.toUTCString();
        console.log("Label printed!\n")
    })

}

function printPartLabel(partNumber) {
    const partInfo = partsCache.get(partNumber);
    console.log(partInfo);

    var productName = partInfo["Product Name"]
    if (productName == undefined) {
        productName = partInfo["Part Type"]
    }

    const fileName = "./label_template_kgb.label";
    fs.readFile(fileName, 'utf8', function (err, data) {
        if (err) throw err;
        labelData = data.replace("000-0000", partInfo['Part Number']).replace("DEVICE_NAME", productName).replace("PART_DESC", partInfo['Part Description'].replace(partInfo['Product Name'], "")).replaceAll(",", "");
        dymo.renderLabel(labelData).then(imageData => {
            fs.writeFile("./public/last_label.png", imageData, 'base64', function (err) {
            });
        });
        dymo.print('DYMO LabelWriter 450', labelData);
        let date = new Date();
        lastLabelPrint = date.toUTCString();
        console.log("Label printed!\n")
    })
}

function printPriceLabel(price, quantity) {
    console.log("Price - " + price)
    console.log("Print Quanitity - " + quantity)
    const fileName = "./label_template_price.label";
    fs.readFile(fileName, 'utf8', function (err, data) {
        if (err) throw err;
        labelData = data.replace("00.00", price);
        dymo.renderLabel(labelData).then(imageData => {
            fs.writeFile("./public/last_label.png", imageData, 'base64', function (err) {
            });
        });
        dymo.print('DYMO LabelWriter 450', labelData, "", quantity);
        let date = new Date();
        lastLabelPrint = date.toUTCString();
        console.log("Label printed!\n")
    })
}

function printDOAWarning() {
    var fileName = "./label_doa_warning.label";
    fs.readFile(fileName, 'utf8', function (err, data) {
        if (err) throw err;
        labelData = data;
        dymo.renderLabel(labelData).then(imageData => {
            fs.writeFile("./public/last_label.png", imageData, 'base64', function (err) {
            });
        });
        dymo.print('DYMO LabelWriter 450', labelData);
        let date = new Date();
        lastLabelPrint = date.toUTCString();
        console.log("DOA Warning label printed!\n")
    })

}

function checkServer() {

    let date = new Date();
    const currentHour = date.getUTCHours();

    if (currentHour >= config.helpdesk.openingHour && currentHour < config.helpdesk.closingHour) {
        checkingInterval = config.helpdesk.checkingInterval;
    } else {
        console.log("Current hour: " + currentHour);
        console.log("Opening hour: " + config.helpdesk.openingHour + "\n");
        checkingInterval = 120; // Increase the checking interval cus why not

        if (ticketCache.keys().length != 0) {
            console.log("Clearing cache...");
            ticketCache.flushAll();
        }
        return;
    }

    if (ticketCache.keys().length == 0) {
        console.log("WARNING! No tickets cached :( Nothing will get printed until we cache all the current tickets.")
    } else {
        console.log("Checking for new tickets...");
    }
    var params = {
        username: config.helpdesk.username,
        password: config.helpdesk.password
    }

    var url = new URL("http://" + config.helpdesk.host + "/helpdesk/WebObjects/Helpdesk.woa/ra/Tickets/mine?qualifier=(statustype.statusTypeName %3D '" + config.helpdesk.searchStatusName + "')");

    Object.keys(params).forEach(key => url.searchParams.append(key, params[key]))

    fetch(url, { method: 'GET' })
        .then(res => res.json())
        .then(json => {
            if (ticketCache.keys().length == 0) {
                console.log("Caching current tickets...\n")
                for (var i = 0; i < json.length; i++) {
                    const ticketID = json[i].id;
                    const ticketSSubject = json[i].shortSubject;

                    console.log("Ticket ID: " + ticketID);
                    console.log("Ticket Subject: " + ticketSSubject)

                    if (ticketCache.set(ticketID, ticketSSubject)) {
                        console.log("Ticket cached!\n");
                    }
                }
                if (json.length == 0) {
                    console.log("Looks like helpdesk doesn't have any tickets to add. Since this is the case, TicketID 123456 will be used for cacheing purposes.")
                    if (ticketCache.set(123456, "TEMPTICKET")) {
                        console.log("Ticket cached!\n");
                    }
                }
            } else {
                for (var i = 0; i < json.length; i++) {
                    const ticketID = json[i].id;
                    const ticketSSubject = json[i].shortSubject;
                    if (ticketSSubject == null) {
                        ticketSSubject = "NOSUBJECTERR"
                    }

                    var retrivedSSubject = ticketCache.get(ticketID);

                    if (retrivedSSubject == undefined) {

                        console.log("New ticket found!");
                        console.log("Ticket ID: " + ticketID);
                        console.log("Ticket Subject: " + ticketSSubject);
                        if (ticketCache.set(ticketID, ticketSSubject)) {
                            console.log("Ticket cached!");
                            console.log("\n");
                        }
                        retriveTicketAndPrint(ticketID, true);
                    }
                }
            }

            let date = new Date();
            lastServerCheck = date.toUTCString();
        });
}

function retriveTicketAndPrint(ticketID, shouldPrintRecipt = true, shouldPrintLabel = true) {
    console.log("Retrieving ticket information...");

    var params = {
        username: config.helpdesk.username,
        password: config.helpdesk.password
    }

    var url = new URL("http://" + config.helpdesk.host + "/helpdesk/WebObjects/Helpdesk.woa/ra/Tickets/" + ticketID);

    Object.keys(params).forEach(key => url.searchParams.append(key, params[key]))

    fetch(url, { method: 'GET' })
        .then(res => res.json())
        .then(json => {

            const subject = json.subject;
            const detail = json.detail.replaceAll("<br/> ", "\n");
            const date = new Date(json.reportDateUtc);

            var serialNumber = subject;

            serialNumber = serialNumber.substr(0, 12);

            if (serialNumber.match("^[A-Za-z0-9]+$") == null) {
                serialNumber = "";
            }

            if (serialNumber.length < 12) {
                serialNumber = "";
            }

            serialNumber = serialNumber.toUpperCase();

            lastSubject = subject;
            lastSerialNumber = serialNumber;
            lastDetails = detail;
            lastTicketNumber = ticketID;

            console.log("Retrieved info!")
            console.log("Serial Number:" + serialNumber)

            if (shouldPrintLabel) {
                if (subject.startsWith("DOA")) {
                    printLabel(ticketID, subject, date.toLocaleDateString("en-UK"));
                    shouldPrintRecipt = false;
                } else {
                    printLabel(ticketID, serialNumber, "");
                }
            }



            if (!subject.startsWith("DOA") && subject.includes("DOA")) {
                printDOAWarning();
            }

            if (shouldPrintRecipt) {
                printRecipt(ticketID, subject, detail, date)
            }
        });
}

function loadPartsCache(callback) {
    console.log('Loading Parts DB...');
    const filename = 'iPhone_Parts.csv';
    const bar1 = new cliProgress.SingleBar({}, cliProgress.Presets.shades_classic);

    var stat = fs.statSync(filename);
    var str = progress({
        length: stat.size,
        time: 100 /* ms */
    });

    bar1.start(stat.size, 0);


    str.on('progress', function (progress) {
        bar1.update(progress.transferred);

        if (progress.remaining == 0) {
            bar1.stop();
            console.log('Parts DB Parsed and Cached!');
            callback();
        }

        /*
        {
            percentage: 9.05,
            transferred: 949624,
            length: 10485760,
            remaining: 9536136,
            eta: 42,
            runtime: 3,
            delta: 295396,
            speed: 949624
        }
        */
    });

    fs.createReadStream(filename)
        .pipe(str)
        .pipe(csv())
        .on('data', (row) => {
            partsCache.set(row['Part Number'], row);
        });
}

function runUpdateScript() {
    var spawn = require('child_process').spawn;
    spawn('/bin/bash', ['update_script.sh'], {
        stdio: 'ignore',
        detached: true
    }).unref();
}