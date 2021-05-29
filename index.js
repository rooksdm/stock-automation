'use strict';
const excelToJson = require('convert-excel-to-json');
var json2xls = require('json2xls');
const fs = require('fs');
const Excel = require('exceljs');
const { isRegExp } = require('util');
const args = process.argv.slice(2)
console.log(args)
const path = args[0]
const parentPath = path.split('\\').reverse().splice(2).reverse().join('\\')
console.log(path)
console.log(parentPath)
const symbolsString = 'AP,CTVA,DD,DOW,IFF,KL,NTR,NUE,PPG,CHTR,CMCSA,T,TMUS,VZ,AMZN,BABA,BOOT,BWA,DBI,DIS,FOSL,GM,HD,LKQ,LOW,MCD,NFLX,PLCE,SBUX,TGT,ADM,CL,COST,GHC,GIS,K,KO,MO,NSRGY,PEP,PG,PM,TRMB,UL,WBA,WMT,BP,COP,CVX,ENB.TO,KMI,MPC,RDSB,VLO,XOM,AXP,BNS.TO,BRKB,HSBC,ICE,JPM,MA,MCO,PGR,PYPL,RY.TO,SCHW,SPGI,V,WFC,ABBV,ABT,AMGN,BMY,CVS,DHR,ISRG,JNJ,LLY,MDT,MRK,NVS,PFE,TMO,UNH,BA,CARR,CSX,DE,EMR,FDX,GD,ITW,LUV,MMM,MORN,NSC,OTIS,RTX,UAL,UNP,UPS,WEC,AAPL,ACN,ADBE,AKAM,AMAT,AMD,ANSS,ANTM,CSCO,FB,FTV,GLW,GOOG,HOLX,HPQ,IBM,INTC,MSFT,ORCL,QCOM,TXN,ZS,MLR,CNP,D,DNP,DUK,ED,EXC,LNT,NEE,SO,SRE'
const symbolsArray = symbolsString.toUpperCase().split(',')
const outArray = []
const dowTheory = excelToJson({
    sourceFile: path + '\\dowtheroy.xlsx',
    header:{
        rows: 1
    },
    columnToKey: {
        'A': '{{A1}}',
        'B': '{{B1}}',
        'C': '{{C1}}',
        'D': 'Mom Rank',
        'E': '{{E1}}',
        'F': '{{F1}}',
        'G': '{{G1}}',
        'H': '{{H1}}',
        'I': 'Perf Rank',
        'J': '{{J1}}',
        'K': '{{K1}}',
    }
})['Sheet1'];

const valueLine = excelToJson({
    sourceFile: path + '\\valueline.xlsx',
    columnToKey: {
        'A': 'Ticker',
        'B': 'Safety',
        'C': 'Performance',
        'D': 'Financial Strength',
        'E': 'Last Price',
        'F': 'Dividend Yield',
        'G': 'Trailing PE',
        'H': 'Target Price Range',
        'I': 'Beta',
        'J': 'Commentary'        
    }
})['Sheet1'];

symbolsArray.forEach(symbol => {
    let dt;
    if(symbol.includes('.TO')){
        let symbolWithoutTO = symbol.replace('.TO','')
        dt = dowTheory.find( ({ Ticker }) => Ticker === symbolWithoutTO )
    }else {
        dt = dowTheory.find( ({ Ticker }) => Ticker === symbol )
    }
    
    const vl = valueLine.find( ({ Ticker }) => Ticker === symbol )
    outArray.push({...dt,...vl})
});

var xls = json2xls(outArray);
//fs.writeFileSync(path + '\\stocks.xlsx', xls, 'binary');
const alphabet = ' abcdefghijklmnopqrstuvwxyz'.toUpperCase().split('');

async function updateExcel(){
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(parentPath + '\\old.xlsx');
    const worksheet = workbook.worksheets[0];
    const colMap = {}
    let newValues = {}
    worksheet.eachRow(function(row, rowNumber) {
        newValues = {}
        if(rowNumber === 1){
            //console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
            for (let i = 1; i < row.values.length; i++) {
                const element = row.values[i];
                colMap[element] = {column: alphabet[i], index: i}
            }
            //console.log(JSON.stringify(colMap))
            // console.log(JSON.stringify(outArray))
        }else {
            const ticker = row.values[colMap['Ticker'].index]
            newValues = outArray.find(obj => obj['Ticker'] === ticker)
            if(!newValues){
                console.log(`NewValues ${JSON.stringify(newValues)} ${rowNumber} issue ${row.values[colMap['Ticker'].index]}`)
            } else {
                for (let key in newValues){
                   // worksheet.getCell(`${colMap[key].column}${rowNumber}`).value = parseInt(newValues[key]) || newValues[key]
                }
            }
        }
    });
    await workbook.xlsx.writeFile('D:\\Projects\\rooksdm\\onedrive\\Rooksdm\\Projects - SB Financial - SB Financial\\Automation\\updated.xlsx');

}
updateExcel()

// console.log(JSON.stringify(outArray.find(obj => obj['Ticker'] === 'ZS')))