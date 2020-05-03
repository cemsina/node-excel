const Workbook = require('./workbook');
"use strict";

function createExcel(json){
    var wb = new Workbook(json.name,json.sheets);
    if(wb.isValid()){
        return {
            file: wb.save(),
        }
    }else{
        return {
            errorKey: wb.errorKey()
        }
    }
}

function toCell(row,col){
    const TABLE = "AZ";
    const start = TABLE.charCodeAt(0);
    const end = TABLE.charCodeAt(1);
    const diff = end - start;

    var temp = col;
    var columnStr = "";
    while(temp > 0){
        var cursor = temp % diff;
        columnStr += String.fromCharCode(start + cursor -1);
        temp -= cursor;
    }

    return columnStr + row;
}

module.exports = {
    createExcel: createExcel,
    toCell: toCell
};