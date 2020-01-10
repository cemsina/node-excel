const { createExcel } = require("../report-core");
("use strict");
var Ihlal_Raporu = require("./../jsons/Ihlal_Raporu.json");
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

let content = {
  name: Ihlal_Raporu.name,
  sheets: [
    {
      name: Ihlal_Raporu.sheets[0].name,
      cells: {
        "A1:B1": { value: Ihlal_Raporu.sheets[0].name }
      }
    }
  ]
};

function IhlalRapor() {
  for (
    let i = 0;
    i < Object.keys(Ihlal_Raporu.sheets[0].subheaders).length;
    i++
  ) {
    content.sheets[0].cells[`A${i + 2}`] = {
      bold: true,
      vertical: "center",
      fontSize: 10,
      value: Ihlal_Raporu.sheets[0].subheaders[i]
    };
    content.sheets[0].cells[`B${i + 2}`] = {
      vertical: "center",
      fontSize: 10,
      value: Ihlal_Raporu.sheets[0].values[i]
    };
  }
  createExcel(content);
  console.log("TCL: content", content);
}

module.exports = IhlalRapor;
