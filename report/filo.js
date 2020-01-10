const {createExcel} = require('../report-core'); 
"use strict";

/**
    Hücre Parametreleri:

    type: "string" | "number",
    value: string | number,
    fontName: string,
    fontSize: number,
    bold: boolean,
    underline: boolean,
    italic: boolean,
    strike: boolean,
    outline: boolean,
    shadow: boolean,
    fontColor: "FFFFFF" (hex),
    fillColor: "FFFFFF" (hex),
    horizontal: "left" | "right" | "center",
    vertical: "left" | "right" | "center"
*/

/**
    Hücre Birleştirme:

    Key olarak "A1" denildiğinde sadece A1 hücresini belirtir.
    Key olarak "A1:A4" denildiğinde A1 den A4 e olan hücreleri birleştirir.
*/

function FiloRapor(){
    createExcel({
        name: "ornek_filo",
        sheets: [
            {
                name: "Sheet 1",
                cells: {
                    "A1:B1":{value: "Test"},
                    "A2":{value: 123,type:"number"},
                    "A3":{fontSize:22,value:"Cemsina"}
                }
            }
        ]
    });
}

module.exports = FiloRapor;