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
    },
    {
      name: Ihlal_Raporu.sheets[1].name,
      cells: {
        F1: { value: "Tarih" },
        G1: { value: Ihlal_Raporu.sheets[1].date },
        "A3:G3": {
          value: Ihlal_Raporu.sheets[1].tables[0].name,
          bold: true,
          vertical: "center",
          fontSize: 10
        }
      }
    }
  ]
};

function getRow(number) {
  switch (number) {
    case value:
      break;

    default:
      break;
  }
}

function IhlalRapor() {
  table2Offset = 0;
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
  for (let i = 65; i < 72; i++) {
    let cell = String.fromCharCode(i) + 4;
    content.sheets[1].cells[cell] = {
      bold: true,
      vertical: "center",
      fontSize: 10,
      value: Ihlal_Raporu.sheets[1].tables[0].headers[i - 65]
    };
    for (
      let j = 0;
      j < Object.keys(Ihlal_Raporu.sheets[1].tables[0].items).length;
      j++
    ) {
      cell = String.fromCharCode(i) + (j + 5);
      content.sheets[1].cells[cell] = {
        vertical: "center",
        fontSize: 10,
        value: Ihlal_Raporu.sheets[1].tables[0].items[0].values[i - 65]
      };
      table2Offset = j + 8;
    }
  }
  cell = `A${table2Offset}:G${table2Offset}`;
  content.sheets[1].cells[cell] = {
    bold: true,
    vertical: "center",
    fontSize: 10,
    value: Ihlal_Raporu.sheets[1].tables[1].name
  };
  createExcel(content);
  console.log("TCL: content", content);
}

module.exports = IhlalRapor;
