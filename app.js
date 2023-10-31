const express = require('express');
const path = require('path');
const ExcelJS = require('exceljs');
const bodyParser = require('body-parser');
const app = express();
const nodemailer = require("nodemailer");
const fs = require("fs");
const PropertiesReader = require('properties-reader');
const prop = PropertiesReader('./app.properties');


app.use(express.json()) //Notice express.json middleware

// Parse application/x-www-form-urlencoded
app.use(bodyParser.urlencoded({ extended: false }));
const port = process.env.PORT || 3000;
// View engine setup
app.set('view engine', 'ejs');
// Serve static files (HTML, CSS, images, etc.)
app.use(express.static(path.join(__dirname, 'public')));


getProperty = (pty) => {return prop.get(pty);}

// Set up routes
app.get('/', (req, res) => {
    //res.sendFile(path.join(__dirname, 'index.html'));



    console.log(getProperty('server.port'))
    res.render('ind', { message: '' })
});


app.post('/', (req, res) => {
    const formData = req.body;
    console.log("hereee")
    console.log("COl2 " + formData.col8);
    saveToExcel(formData);
    res.render('ind', { message: 'Success Exported Data to Excel' });
});


app.get('/price', (req, res) => {

    console.log(findColumnToEnterData());
    res.render('priceDetails', { message: '' });
});

app.post('/price', (req, res) => {

    const formDataPrice = req.body;
    console.log(formDataPrice);
    copyToExcel(formDataPrice);

    res.render('priceDetails', { message: 'Success ! Price Details have been exported to Excel Sheet. Please verify.' });
});


app.get('/sendMail', (req, res)=>{

    let fa = getProperty('from.address')
    let sender = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: 'seethup@gmail.com',
            pass: 'wins zynw nddw vvct'
        }
    });
     
    let mail = {
        from: 'seethup@gmail.com',
        to:
            'nivi2310@gmail.com',
        subject: 'Sending Email using Node.js',
        text: 'That was easy!',
        html:
            '<h1>GeeksforGeeks</h1><p>Medical Camp excel</p>',
    attachments: [
        {
            filename: 'Medical Camp.xlsx',
            path: __dirname + '/Medical Camp.xlsx',
            cid: 'uniq-medicalcamp.xlsx'
        }
    ]
    };
     
    sender.sendMail(mail, function (error, info) {
        if (error) {
            console.log(error);
        } else {
            res.render('emailSuccess')
            console.log('Email sent successfully: '
                + info.response);
        }
    });    
});

app.listen(port, () => {
    console.log(`Server is running at http://localhost:${port}`);
});


function isEmpty(str) {
    return (!str || str.length === 0);
}



const formatDate = () => {
    let date = new Date();
    const day = date.toLocaleString('default', { day: '2-digit' });
    const month = date.toLocaleString('default', { month: 'short' });
    const year = date.toLocaleString('default', { year: 'numeric' });
    return day + '-' + month + '-' + year;
}

function saveToExcel(formData) {

    console.log(formData)

    const workbook = new ExcelJS.Workbook();
    workbook.xlsx.readFile('Medical Camp.xlsx')
        .then(() => {

            let dates = formatDate();
            //var ws1 = workbook.getWorksheet('TestSheet');
            //let copySHeet = workbook.addWorksheet(dates.toLocaleString());
            //copySHeet.model = ws1.model;
            //copySHeet.name = dates.toLocaleString();

            const worksheet = workbook.getWorksheet('QuantitySheet-November'); // Assuming you want to edit the first sheet

            // 1st Table
            const cellE12 = worksheet.getCell('E12');
            cellE12.value = formData.col1;

            const cellE13 = worksheet.getCell('E13');
            cellE13.value = formData.col2;

            const cellE14 = worksheet.getCell('E14');
            cellE14.value = formData.col3;

            const cellE15 = worksheet.getCell('E15');
            cellE15.value = formData.col4;

            //2nd Table
            const cellE18 = worksheet.getCell('E18');
            cellE18.value = formData.col5;

            const cellE19 = worksheet.getCell('E19');
            cellE19.value = formData.col6;

            const cellE20 = worksheet.getCell('E20');
            cellE20.value = formData.col7;

            const cellE21 = worksheet.getCell('E21');
            cellE21.value = formData.col8;

            const cellE22 = worksheet.getCell('E22');
            cellE22.value = formData.col9;

            //3rd Table

            const cellE27 = worksheet.getCell('E27');
            cellE27.value = formData.col10;

            const cellE28 = worksheet.getCell('E28');
            cellE28.value = formData.col11;

            const cellE29 = worksheet.getCell('E29');
            cellE29.value = formData.col12;

            const cellE30 = worksheet.getCell('E30');
            cellE30.value = formData.col13;

            const cellE31 = worksheet.getCell('E31');
            cellE31.value = formData.col14;

            const cellE32 = worksheet.getCell('E32');
            cellE32.value = formData.col15;

            worksheet.getCell('E33').value = formData.col16;

            //Table 4

            worksheet.getCell('E38').value = formData.col17;
            worksheet.getCell('E39').value = formData.col18;
            worksheet.getCell('E40').value = formData.col19;

            //Table5

            worksheet.getCell('J12').value = formData.col20;
            worksheet.getCell('J13').value = formData.col21;
            worksheet.getCell('J14').value = formData.col22;

            //Table 6
            worksheet.getCell('J18').value = formData.col23;
            worksheet.getCell('J19').value = formData.col24;
            worksheet.getCell('J20').value = formData.col25;
            worksheet.getCell('J21').value = formData.col26;
            worksheet.getCell('J22').value = formData.col27;
            worksheet.getCell('J23').value = formData.col28;

            //Table 7
            worksheet.getCell('J26').value = formData.col29;
            worksheet.getCell('J27').value = formData.col30;
            worksheet.getCell('J28').value = formData.col31;
            worksheet.getCell('J29').value = formData.col32;
            worksheet.getCell('J30').value = formData.col33;
            worksheet.getCell('J31').value = formData.col34;
            worksheet.getCell('J32').value = formData.col35;
            worksheet.getCell('J33').value = formData.col36;
            worksheet.getCell('J34').value = formData.col37;
            worksheet.getCell('J35').value = formData.col38;
            worksheet.getCell('J36').value = formData.col39;
            worksheet.getCell('J37').value = formData.col40;
            worksheet.getCell('J38').value = formData.col41;
            worksheet.getCell('J39').value = formData.col42;

            //Table 8

            worksheet.getCell('E44').value = formData.col43;
            worksheet.getCell('E45').value = formData.col44;
            worksheet.getCell('E46').value = formData.col45;
            worksheet.getCell('E47').value = formData.col46;
            worksheet.getCell('E48').value = formData.col46A;

            //Table 9
            worksheet.getCell('E50').value = formData.col47;

            //Table 10

            worksheet.getCell('E53').value = formData.col48;
            worksheet.getCell('E54').value = formData.col49;
            worksheet.getCell('E55').value = formData.col50;
            worksheet.getCell('E56').value = formData.col51;
            worksheet.getCell('E57').value = formData.col52;
            worksheet.getCell('E58').value = formData.col53;
            worksheet.getCell('E59').value = formData.col54;
            worksheet.getCell('E60').value = formData.col54A;

            //Table 11
            worksheet.getCell('E62').value = formData.col55;

            //Table 12
            worksheet.getCell('E65').value = formData.col56;

            //Table 13
            worksheet.getCell('E68').value = formData.col57;
            worksheet.getCell('E69').value = formData.col58;

            //Table 14
            worksheet.getCell('E72').value = formData.col59;

            //Table 15
            worksheet.getCell('J44').value = formData.col60;
            worksheet.getCell('J45').value = formData.col61;
            worksheet.getCell('J46').value = formData.col62;
            worksheet.getCell('J47').value = formData.col63;
            worksheet.getCell('J48').value = formData.col64;

            //Table 16
            worksheet.getCell('J51').value = formData.col65;
            worksheet.getCell('J52').value = formData.col66;

            //Table 17
            worksheet.getCell('J55').value = formData.col67;
            worksheet.getCell('J56').value = formData.col68;

            //Table 18
            worksheet.getCell('J60').value = formData.col69;
            worksheet.getCell('J61').value = formData.col70;
            worksheet.getCell('J62').value = formData.col71;
            worksheet.getCell('J63').value = formData.col72;
            worksheet.getCell('J64').value = formData.col73;
            worksheet.getCell('J65').value = formData.col74;
            worksheet.getCell('J66').value = formData.col75;

            //Table 19
            worksheet.getCell('J68').value = formData.col76;
            worksheet.getCell('J69').value = formData.col77;
            worksheet.getCell('J70').value = formData.col78;
            worksheet.getCell('J71').value = formData.col79;
            worksheet.getCell('J72').value = formData.col80;


            // Save the workbook
            return workbook.xlsx.writeFile('Medical Camp.xlsx');
        })
        .then(() => {
            console.log('Excel file updated successfully');
        })
        .catch(error => {
            console.error(error);
        });

}

const findColumnToEnterData = () => {

    var today = new Date();

    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    const d = new Date();
    let colName = monthNames[d.getMonth()] + ' ' + "01 " + d.getFullYear();
    //let colName = d.getFullYear()+'-'+d.getMonth().toLocaleString('en-US', {minimumIntegerDigits: 2, useGrouping:false})+'-'+'01';


    console.log("Date is == " + colName);

    return colName;
}


const searchForColumn = (worksheet, colName) => {
    //Iterate through rows and columns to find the value
    let cols = 1;

    worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        row.eachCell({ includeEmpty: false }, function (cell, colNumber) {

            let ba = cell.value;
            ba = ba.toString();
            // if(ba.includes("Nov 01 2023")){
            //     console.log(true)
            // }
            if (ba.indexOf(colName) != -1) {
                console.log(`Found at row ${rowNumber}, column ${colNumber}`);
                cols = colNumber;
            }
        });
    });

    return cols;
}

function copyToExcel(formData) {
    console.log(formData)

    const workbook = new ExcelJS.Workbook();
    workbook.xlsx.readFile('Medical Camp.xlsx')
        .then(() => {

            let dates = formatDate();

            

            const worksheet = workbook.getWorksheet('PriceSheet-November'); // Assuming you want to edit the first sheet

            let colName = findColumnToEnterData();
            console.log(colName);

            const colsNames = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA']


            let columnNumber = searchForColumn(worksheet, colName);

            console.log(columnNumber);
            let actualColumnName = colsNames[columnNumber - 1];




            // 1st Table
            const cellE12 = worksheet.getCell(actualColumnName + '5');
            cellE12.value = isEmpty(formData.fp1) ? 0 : parseFloat(formData.fp1);

            const cellE13 = worksheet.getCell(actualColumnName + '6');
            cellE13.value = isEmpty(formData.fp2) ? 0 : parseFloat(formData.fp2);

            const cellE14 = worksheet.getCell(actualColumnName + '7');
            cellE14.value = isEmpty(formData.fp3) ? 0 : parseFloat(formData.fp3);

            const cellE15 = worksheet.getCell(actualColumnName + '8');
            cellE15.value = isEmpty(formData.fp4) ? 0 : parseFloat(formData.fp4);

            //2nd Table
            const cellE18 = worksheet.getCell(actualColumnName + '9');
            cellE18.value = isEmpty(formData.fp5) ? 0 : parseFloat(formData.fp5);

            const cellE19 = worksheet.getCell(actualColumnName + '10');
            cellE19.value = isEmpty(formData.fp6) ? 0 : parseFloat(formData.fp6);

            const cellE20 = worksheet.getCell(actualColumnName + '11');
            cellE20.value = isEmpty(formData.fp7) ? 0 : parseFloat(formData.fp7);

            const cellE21 = worksheet.getCell(actualColumnName + '12');
            cellE21.value = isEmpty(formData.fp8) ? 0 : parseFloat(formData.fp8);

            const cellE22 = worksheet.getCell(actualColumnName + '13');
            cellE22.value = isEmpty(formData.fp9) ? 0 : parseFloat(formData.fp9);

            //3rd Table

            const cellE27 = worksheet.getCell(actualColumnName + '14');
            cellE27.value = isEmpty(formData.fp10) ? 0 : parseFloat(formData.fp10);

            const cellE28 = worksheet.getCell(actualColumnName + '15');
            cellE28.value = isEmpty(formData.fp11) ? 0 : parseFloat(formData.fp11);

            const cellE29 = worksheet.getCell(actualColumnName + '16');
            cellE29.value = isEmpty(formData.fp12) ? 0 : parseFloat(formData.fp12);

            const cellE30 = worksheet.getCell(actualColumnName + '17');
            cellE30.value = isEmpty(formData.fp13) ? 0 : parseFloat(formData.fp13);

            const cellE31 = worksheet.getCell(actualColumnName + '18');
            cellE31.value = isEmpty(formData.fp14) ? 0 : parseFloat(formData.fp14);

            const cellE32 = worksheet.getCell(actualColumnName + '19');
            cellE32.value = isEmpty(formData.fp15) ? 0 : parseFloat(formData.fp15);

            worksheet.getCell(actualColumnName + '20').value = isEmpty(formData.fp16) ? 0 : parseFloat(formData.fp16);

            //Table 4

            worksheet.getCell(actualColumnName + '21').value = isEmpty(formData.fp17) ? 0 : parseFloat(formData.fp17);
            worksheet.getCell(actualColumnName + '22').value = isEmpty(formData.fp18) ? 0 : parseFloat(formData.fp18);
            worksheet.getCell(actualColumnName + '23').value = isEmpty(formData.fp19) ? 0 : parseFloat(formData.fp19);

            //Table5

            worksheet.getCell(actualColumnName + '24').value = isEmpty(formData.fp20) ? 0 : parseFloat(formData.fp20);
            worksheet.getCell(actualColumnName + '25').value = isEmpty(formData.fp21) ? 0 : parseFloat(formData.fp21);
            worksheet.getCell(actualColumnName + '26').value = isEmpty(formData.fp22) ? 0 : parseFloat(formData.fp22);

            //Table 6
            worksheet.getCell(actualColumnName + '27').value = isEmpty(formData.fp23) ? 0 : parseFloat(formData.fp23);
            worksheet.getCell(actualColumnName + '28').value = isEmpty(formData.fp24) ? 0 : parseFloat(formData.fp24);
            worksheet.getCell(actualColumnName + '29').value = isEmpty(formData.fp25) ? 0 : parseFloat(formData.fp25);
            worksheet.getCell(actualColumnName + '30').value = isEmpty(formData.fp26) ? 0 : parseFloat(formData.fp26);
            worksheet.getCell(actualColumnName + '32').value = isEmpty(formData.fp27) ? 0 : parseFloat(formData.fp27);
            worksheet.getCell(actualColumnName + '33').value = isEmpty(formData.fp28) ? 0 : parseFloat(formData.fp28);

            //Table 7
            worksheet.getCell(actualColumnName + '34').value = isEmpty(formData.fp29) ? 0 : parseFloat(formData.fp29);
            worksheet.getCell(actualColumnName + '35').value = isEmpty(formData.fp30) ? 0 : parseFloat(formData.fp30);
            worksheet.getCell(actualColumnName + '36').value = isEmpty(formData.fp31) ? 0 : parseFloat(formData.fp31);
            worksheet.getCell(actualColumnName + '37').value = isEmpty(formData.fp32) ? 0 : parseFloat(formData.fp32);
            worksheet.getCell(actualColumnName + '38').value = isEmpty(formData.fp33) ? 0 : parseFloat(formData.fp33);
            worksheet.getCell(actualColumnName + '39').value = isEmpty(formData.fp34) ? 0 : parseFloat(formData.fp34);
            worksheet.getCell(actualColumnName + '40').value = isEmpty(formData.fp35) ? 0 : parseFloat(formData.fp35);
            worksheet.getCell(actualColumnName + '41').value = isEmpty(formData.fp36) ? 0 : parseFloat(formData.fp36);
            worksheet.getCell(actualColumnName + '42').value = isEmpty(formData.fp37) ? 0 : parseFloat(formData.fp37);
            worksheet.getCell(actualColumnName + '43').value = isEmpty(formData.fp38) ? 0 : parseFloat(formData.fp38);
            worksheet.getCell(actualColumnName + '44').value = isEmpty(formData.fp39) ? 0 : parseFloat(formData.fp39);
            worksheet.getCell(actualColumnName + '45').value = isEmpty(formData.fp40) ? 0 : parseFloat(formData.fp40);
            worksheet.getCell(actualColumnName + '46').value = isEmpty(formData.fp41) ? 0 : parseFloat(formData.fp41);
            worksheet.getCell(actualColumnName + '47').value = isEmpty(formData.fp42) ? 0 : parseFloat(formData.fp42);

            //Table 8

            worksheet.getCell(actualColumnName + '48').value = isEmpty(formData.fp43) ? 0 : parseFloat(formData.fp43);
            worksheet.getCell(actualColumnName + '49').value = isEmpty(formData.fp44) ? 0 : parseFloat(formData.fp44);
            worksheet.getCell(actualColumnName + '50').value = isEmpty(formData.fp45) ? 0 : parseFloat(formData.fp45);
            worksheet.getCell(actualColumnName + '51').value = isEmpty(formData.fp46) ? 0 : parseFloat(formData.fp46);
            worksheet.getCell(actualColumnName + '52').value = isEmpty(formData.fp46A) ? 0 : parseFloat(formData.fp46A);

            //Table 9
            worksheet.getCell(actualColumnName + '53').value = isEmpty(formData.fp47) ? 0 : parseFloat(formData.fp47);

            //Table 10

            worksheet.getCell(actualColumnName + '54').value = isEmpty(formData.fp48) ? 0 : parseFloat(formData.fp48);
            worksheet.getCell(actualColumnName + '55').value = isEmpty(formData.fp49) ? 0 : parseFloat(formData.fp49);
            worksheet.getCell(actualColumnName + '56').value = isEmpty(formData.fp50) ? 0 : parseFloat(formData.fp50);
            worksheet.getCell(actualColumnName + '57').value = isEmpty(formData.fp51) ? 0 : parseFloat(formData.fp51);
            worksheet.getCell(actualColumnName + '58').value = isEmpty(formData.fp52) ? 0 : parseFloat(formData.fp52);
            worksheet.getCell(actualColumnName + '60').value = isEmpty(formData.fp53) ? 0 : parseFloat(formData.fp53);
            worksheet.getCell(actualColumnName + '61').value = isEmpty(formData.fp54) ? 0 : parseFloat(formData.fp54);
            worksheet.getCell(actualColumnName + '62').value = isEmpty(formData.fp54A) ? 0 : parseFloat(formData.fp54A);

            //Table 11
            worksheet.getCell(actualColumnName + '63').value = isEmpty(formData.fp55) ? 0 : parseFloat(formData.fp55);

            //Table 12
            worksheet.getCell(actualColumnName + '64').value = isEmpty(formData.fp56) ? 0 : parseFloat(formData.fp56);

            //Table 13
            worksheet.getCell(actualColumnName + '65').value = isEmpty(formData.fp57) ? 0 : parseFloat(formData.fp57);
            worksheet.getCell(actualColumnName + '66').value = isEmpty(formData.fp58) ? 0 : parseFloat(formData.fp58);

            //Table 14
            worksheet.getCell(actualColumnName + '67').value = isEmpty(formData.fp59) ? 0 : parseFloat(formData.fp59);

            //Table 15
            worksheet.getCell(actualColumnName + '68').value = isEmpty(formData.fp60) ? 0 : parseFloat(formData.fp60);
            worksheet.getCell(actualColumnName + '69').value = isEmpty(formData.fp61) ? 0 : parseFloat(formData.fp61);
            worksheet.getCell(actualColumnName + '70').value = isEmpty(formData.fp62) ? 0 : parseFloat(formData.fp62);
            worksheet.getCell(actualColumnName + '71').value = isEmpty(formData.fp63) ? 0 : parseFloat(formData.fp63);
            worksheet.getCell(actualColumnName + '72').value = isEmpty(formData.fp64) ? 0 : parseFloat(formData.fp64);

            //Table 16
            worksheet.getCell(actualColumnName + '73').value = isEmpty(formData.fp65) ? 0 : parseFloat(formData.fp65);
            worksheet.getCell(actualColumnName + '74').value = isEmpty(formData.fp66) ? 0 : parseFloat(formData.fp66);

            //Table 17
            worksheet.getCell(actualColumnName + '75').value = isEmpty(formData.fp67) ? 0 : parseFloat(formData.fp67);
            worksheet.getCell(actualColumnName + '76').value = isEmpty(formData.fp68) ? 0 : parseFloat(formData.fp68);

            //Table 18
            worksheet.getCell(actualColumnName + '77').value = isEmpty(formData.fp69) ? 0 : parseFloat(formData.fp69);
            worksheet.getCell(actualColumnName + '78').value = isEmpty(formData.fp70) ? 0 : parseFloat(formData.fp70);
            worksheet.getCell(actualColumnName + '79').value = isEmpty(formData.fp71) ? 0 : parseFloat(formData.fp71);
            worksheet.getCell(actualColumnName + '80').value = isEmpty(formData.fp72) ? 0 : parseFloat(formData.fp72);
            worksheet.getCell(actualColumnName + '81').value = isEmpty(formData.fp73) ? 0 : parseFloat(formData.fp73);
            worksheet.getCell(actualColumnName + '82').value = isEmpty(formData.fp74) ? 0 : parseFloat(formData.fp74);
            worksheet.getCell(actualColumnName + '83').value = isEmpty(formData.fp75) ? 0 : parseFloat(formData.fp75);

            //Table 19
            worksheet.getCell(actualColumnName + '84').value = isEmpty(formData.fp76) ? 0 : parseFloat(formData.fp76);
            worksheet.getCell(actualColumnName + '85').value = isEmpty(formData.fp77) ? 0 : parseFloat(formData.fp77);
            worksheet.getCell(actualColumnName + '86').value = isEmpty(formData.fp78) ? 0 : parseFloat(formData.fp78);
            worksheet.getCell(actualColumnName + '87').value = isEmpty(formData.fp79) ? 0 : parseFloat(formData.fp79);
            worksheet.getCell(actualColumnName + '88').value = isEmpty(formData.fp80) ? 0 : parseFloat(formData.fp80);
            worksheet.getCell(actualColumnName + '89').value = isEmpty(formData.fp81) ? 0 : parseFloat(formData.fp81);

            worksheet.getCell(actualColumnName + '90').value = { formula: 'SUM(H5:H89)' };

            // Save the workbook
            return workbook.xlsx.writeFile('Medical Camp.xlsx');
        })
        .then(() => {
            console.log('Excel file updated successfully');
        })
        .catch(error => {
            console.error(error);
        });

}


