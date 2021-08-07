const express = require('express')
const fs = require('fs')
const config = require('./config.json')
const fetch = require('node-fetch')
const NodeCache = require("node-cache");
const ThermalPrinter = require("node-thermal-printer").printer;
const Types = require("node-thermal-printer").types;
const Dymo = require('dymojs')
const app = express()
const port = config.webPort

const ticketCache = new NodeCache();
const dymo = new Dymo();

var lastServerCheck = "n/a";
var lastLabelPrint = "n/a";

var checkingInterval = 30;

var reciptFooter;

app.set('view engine', 'ejs')
app.use(express.static('public'))

app.get('/', (req, res) => {
    res.render('index', { lastServerCheck, lastLabelPrint});
})

app.listen(port, '0.0.0.0', () => {
    console.log(`WHD-Autoprint app listening on port ${port}`)
    checkServer();
    setInterval(checkServer,  checkingInterval * 1000)
})


function loadRecpiptFooter() {
    try {  
        var data = fs.readFileSync('recipt_footer.txt', 'utf8');
        reciptFooter = data.toString();    
    } catch(e) {
        console.log('Error:', e.stack);
    }
}
async function printRecipt (ticketID, subject, detail, date) {
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
    printer.println("iStore Ipswich");
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
  
    printer.newLine();
    printer.newLine();
  
    
    printer.append(Buffer.from([0x1b, 0x1d, 0x61, 0x00]));
  
    printer.drawLine();
    printer.print(detail);
    printer.newLine();
    printer.drawLine();

    printer.newLine();
    printer.newLine();

    printer.println(date.toString());

  
    printer.newLine();
    printer.newLine();

    printer.print(reciptFooter);

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
  

function printLabel(ticketID, ticketSSubject) {
    var labelData = "";
    fs.readFile("./label_template.label", 'utf8', function(err, data) {
        if (err) throw err;
        labelData = data.replace("TICKETNO", ticketID);
        dymo.renderLabel(labelData).then(imageData => {
            fs.writeFile("./public/last_label.png", imageData, 'base64', function(err) {
            });
        });
        dymo.print('DYMO LabelWriter 450', labelData);
        let date= new Date();
        lastLabelPrint = date.toUTCString();
        console.log("Label printed!\n")
    })

}



function checkServer() {
    
    let date= new Date();
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
                        retriveTicketAndPrint(ticketID);
                    }
                }
            }

            let date= new Date();
            lastServerCheck = date.toUTCString();
        });
}

function retriveTicketAndPrint(ticketID) {
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

            printLabel(ticketID, subject);
            printRecipt(ticketID, subject, detail, date)
        });
}

loadRecpiptFooter();