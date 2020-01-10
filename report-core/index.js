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

module.exports = {
    createExcel: createExcel
};