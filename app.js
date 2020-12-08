const GoogleSpreadsheet = require('google-spreadsheet');
const {promisify} = require('util');
const creds = require('./TML_GSheet-c7e2bdedafd6.json');
var request = require('request');
const express = require('express');
const bodyParser = require('body-parser');

const app = express()
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: true}) );
app.all("/*", function(req, res, next){
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, Content-Length, X-Requested-With');
    next();
  });

app.get('/',function(req,res){
    const returnObj = accessSpreadsheet()
    .then((result) => {
        res.send(result)
    });
})

app.get('/post/:id/:checked',function(req,res){
    const returnObj = appendSpreadsheet(req.params.id,req.params.checked)
    .then((result) => {
        res.send('Done')
    }); 
})

app.get('/protect',function(req,res){  
        var url = 'https://script.google.com/a/ragingriverict.com/macros/s/AKfycbxJ1u6SHrIiNOB-DRpgeJ9EGkJu0nCO3u972GGLz9Q/dev';
        request(url, function (error, response, body) {
            res.send(body)
    })
})
app.listen(3000,function(){
    console.log("Port 3000")
})



// function print(values){
//     console.log(`HCP ID: ${values.hcpid}`);
//     console.log(`HCP Name: ${values.hcpname}`);
//     console.log(`Exisiting Consent: ${values.existingconsent}`);
//     console.log(`Consent Date: ${values.consentdate}`);
//     console.log(`Is Approved: ${values.isapproved}`);

//     console.log(`------------`);

// }

async function accessSpreadsheet(){
    const doc = new GoogleSpreadsheet('1h0yG1eU3OhFEd4lhB2Oz4GHDahSZKPVgpDqCn7oRKeQ');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    const sheet = info.worksheets[2]
    //console.log(`Title: ${sheet.title}, Rows: ${sheet.rowCount}`);
    const cells = await promisify(sheet.getCells)({
        'min-row':7,
        'max-row':sheet.rowCount,
        'min-col':1,
        'max-col':5
    })
    var object={}
    var hcpid=''
    var hcpname=''
    var existingconsent=''
    var consentdate=''
    var isapproved=''
    for(const cell of cells){
        //console.log(cell)
       // console.log(`${cell.row}, ${cell.col}:${cell.value}`)
        if(cell.col == 1){
            hcpid = cell.value.trim()
            hcpname=''
            existingconsent=''
            consentdate=''
            isapproved=''
        }
        else if(cell.col == 2){
            hcpname = cell.value.trim()
        }
        else if(cell.col == 3){
            existingconsent = cell.value.trim()
        }
        else if(cell.col == 4){
            consentdate = cell.value.trim()
        }
        else if(cell.col == 5){
            isapproved = cell.value.trim().toLowerCase()
        }
        var obj={
            hcp_id:hcpid,
            hcp_name:hcpname,
            existing_consent:existingconsent,
            consent_date:consentdate,
            is_approved:isapproved,
        }
        var hcpidobj = object[hcpid] ||{}
        hcpidobj = obj
        object[hcpid] = hcpidobj
    
    }
    return object

    //console.log(object)
    // const rows = await promisify(sheet.getRows)({
    //     offset:0
    // })
    
    // console.log(rows)
    
    // rows.forEach(row=>{
    //     print(row)
    //     print(row)
    // })

    // const rows = await promisify(sheet.getRows)({
    //     offset:5,
    //     limit:10,
    //     orderby:'homestate'
    // })

    // const rows = await promisify(sheet.getRows)({
    //     query:'studentname=Clark'
    // })


    // rows.forEach(row=>{
    //     row.isgraduated=false
    //     row.save()
    // })

    // const insertRow ={
    //     studentname:'Clark',
    //     gender:'Male',
    //     major:'Computer Engineering',
    //     homestate:'PH',
    //     classlevel:'Graduated',
    //     extracurricularactivity:'Games',
    //     isgraduated:true,
    // }

    // await promisify(sheet.addRow)(insertRow)
}

async function appendSpreadsheet(id,checked){
    const doc = new GoogleSpreadsheet('1h0yG1eU3OhFEd4lhB2Oz4GHDahSZKPVgpDqCn7oRKeQ');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)();
    const sheet = info.worksheets[0]
    //console.log(`Title: ${sheet.title}, Rows: ${sheet.rowCount}`);
    const cells = await promisify(sheet.getCells)({
        'min-row':2,
        'max-row':sheet.rowCount,
        'min-col':1,
        'max-col':5
    }) 
    var count = 0
     for(const cell of cells){
        //console.log(cell)
         console.log(`${cell.row}, ${cell.col}:${cell.value} - ${count}`)
       if(id == cell.value.trim())
       { 
           cells[count+4].value= checked
           cells[count+4].save()
           console.log(`${count} - ${count+4} - ${cells[count+4].value}`)
           break
       }
    count =count + 1
    } 
}

// accessSpreadsheet();
//appendSpreadsheet('AM-00041',true);