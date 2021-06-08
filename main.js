const express = require('express')
const fs = require('fs')
const config = require('./config.json')
const fetch = require('node-fetch')
const NodeCache = require("node-cache");
const Dymo = require('dymojs')
const app = express()
const port = config.webPort

const ticketCache = new NodeCache();
const dymo = new Dymo();

var lastServerCheck = "n/a";
var lastLabelPrint = "n/a";

var checkingInterval = 30;

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
                        console.log("Printing...");
                        printLabel(ticketID, ticketSSubject);
                    }
                }
            }

            let date= new Date();
            lastServerCheck = date.toUTCString();
        });
}