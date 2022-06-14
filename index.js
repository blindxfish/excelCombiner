var XLSX = require("xlsx");
let filesArr = [];

//requiring path and fs modules
const path = require('path');
const fs = require('fs');
const async = require("async")

//joining path of directory 
const directoryPath = path.join(__dirname, 'files');
//passsing directoryPath and callback function
fs.readdir(directoryPath, function (err, files) {
    
    //handling error
    if (err) {
        return console.log('Unable to scan directory: ' + err);
    } 

    //listing all files using forEach
    files.forEach(function (file) {
        //Add the filenames in the array
        filesArr.push(file);
    });
    
    //Start the dataMiner
    for (var file in filesArr) {
        q.push(filesArr[file], function (err,file,res) {
            console.log("Started to read: " + file );       
        });
    }
});


// FUNCTION WHICH READS THE SPECIFIC ROWS FROM EVERY XCELL IN THE FOLDER: "/files"
function getdata(file){
    
    //start row 
    let crow = 11;

    //some specific data
    let plc = 'C4';
    
    //other rows
    let lac;
    let lacname;
    let module;
    let modulename;
    let compname;
    let layoutPosition;

    var workbook = XLSX.readFile('./files/'+file);

    //last row of each file 
    let maxRows = 1058

    //array we return for processing
    let rows = []
      

    //Customizable logic for extraction:
        
        //define sheet names or check
        workbook.SheetNames.forEach(sheet =>{
            if(sheet.startsWith('F')){
                crow = 11;
                let ws = workbook.Sheets[sheet]

                
               for(let i=0; i<=maxRows; i++){
                   
                   module = 'B'+crow.toString()
                   modulename = 'C'+crow.toString()
                   moduletext = 'H'+crow.toString()
                   compname = 'AE'+crow.toString()
                   productid = 'F'+crow.toString()
                   moduletype = 'E'+crow.toString()
                   layoutPosition = 'D'+crow.toString()
                   content = '' 

                if(ws['L'+crow.toString()] != undefined){
                    lac = 'L'+crow.toString()
                }
                
                if(ws['M'+crow.toString()] != undefined){
                    lacname = 'M'+crow.toString()
                }
                if(ws[modulename]!=undefined && ws[modulename].h != ' '){
// PLC
                    if(ws[plc]!=undefined){
                        content += ws[plc].h + ';'
                   }else{
                        content += ';'
                   }
// POSITION TYPE                   
                   if(ws[lacname]!=undefined){
                    content += ws[lacname].h + ';'
                    }else{
                            content += ';'
                    }
// LAC
                   if(ws[lac]!=undefined){
                    content += ws[lac].h + ';'
                    }else{
                            content += ';'
                    }
//MODULE NUM                      
                   if(ws[module]!=undefined){
                    content += ws[module].w + ';'
                    }else{
                            content += ';'
                    }
//MODULE NAME                    
                    content += ws[modulename].h + ';'
//COMPOINT                   
                    if(ws[compname]!=undefined){
                        content += ws[compname].h.replace(';',',') + ';'
                   }else{
                        content += ';'
                   }
//MODULE TEXT
                   if(ws[moduletext]!=undefined){
                       content += ws[moduletext].h + ';'
                    }else{
                        content += ';'
                   }
//PRODUCT ID      
                   if(ws[productid]!=undefined){ 
                    content += ws[productid].h + ';'
                    }else{
                        content += ';'
                   }
//MODULE TYPE
                   if(ws[moduletype]!=undefined){ 
                    content += ws[moduletype].w + ';'
                    }else{
                        content += ';'
                   }
//LAYOUT POSITION
                   if(ws[layoutPosition]!=undefined){ 
                    content += ws[layoutPosition].w + ';'
                    }else{
                        content += ';'
                   }


                   content += "\r\n"

                   rows.push(content)
                  
                }
                    
                    crow++

                }
            }
      
        })

    return rows
}

// CREATE THE QUEUE
var q = async.queue(function(file, callback) {

        dataArray = getdata(file)
        
        dataArray.forEach(el=>{
            fs.appendFileSync('file.csv', el, err => {
                if (err) {
                    console.error(err);
                }
                    console.log('Data extracted and saved!')
                });
        })
        
        callback(null,file);
}, 1);
