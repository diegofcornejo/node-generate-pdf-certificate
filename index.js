var express = require('express');
var app = express();
const readXlsxFile = require('read-excel-file/node');
const PDFDocument = require('pdfkit');
const fs = require('fs');

app.set('port', process.env.PORT || 4000);

var server = app.listen(app.get('port'), function () {
    console.log('Express server listening on port ' + server.address().port);

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    async function processArray(array) {
        const width = 792;
        //Folder
        var certs = `./certs`;
        if (!fs.existsSync(certs)) {
            fs.mkdirSync(certs);
        }
        for (const item of array) {
            var name = item[0] || '';
            console.log(item);
            await sleep(500);

            var doc = new PDFDocument({
                layout: 'landscape'
            });

            doc.image('./base.jpg', 0, 0, {
                width: width,
                height: 612
            });

            //Name
            doc.fontSize(24);
            doc.moveDown(5);
            doc.text(`${name}`, {
                align: 'center'
            });

            //Place
            doc.fontSize(18);
            doc.moveDown(0.25);
            doc.text(`${item[1]}`, {
                align: 'center'
            });

            //Course
            doc.fontSize(24);
            doc.moveDown(2);
            doc.text(`${item[2]}`, {
                align: 'center'
            });

            //Date
            doc.fontSize(12);
            doc.moveDown(2);
            doc.text(`${item[4]}`, {
                align: 'center'
            });

            //Folder
            var path = `./certs/${name}`;
            if (!fs.existsSync(path)) {
                fs.mkdirSync(path);
            }

            doc.pipe(fs.createWriteStream(`./certs/${item[0]}/${item[0]}_${item[2]}.pdf`));
            doc.end();
        };
    }

    readXlsxFile('./list.xlsx').then((rows) => {
        processArray(rows);
    })
});