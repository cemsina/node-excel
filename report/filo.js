const { createExcel } = require("../report-core");
("use strict");
var Filo_Raporu = require("./../jsons/Filo_Raporu.json");
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
const row_count = Object.keys(Filo_Raporu.sheet_1.keys).length;

let content = {
  name: Filo_Raporu.name,
  sheets: [
    {
      name: Filo_Raporu.sheet_1.sheet_name,
      cells: {
        "A1:B1": { value: Filo_Raporu.sheet_1.header }
      }
    }
  ]
};

function FiloRapor() {
  Object.keys(Filo_Raporu.sheet_1.keys).map(key => {
    const indexStr = key.slice(-1);
    const index = Number(indexStr) + 1;
    content.sheets[0].cells[`A${index}`] = {
      bold: true,
      vertical: "center",
      fontSize: 10,
      value: Filo_Raporu.sheet_1.keys[`key_${indexStr}`]
    };
    content.sheets[0].cells[`B${index}`] = {
      vertical: "center",
      fontSize: 10,
      value: Filo_Raporu.sheet_1.values[`value_${indexStr}`]
    };
  });
  createExcel(content);
}

module.exports = FiloRapor;
