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
  name: "Ihlal_Raporu",
  sheets: [
    {
      name: "Ihlal Raporu",
      cells: {
        "A1:B1": { value: "Ratio Mining", bold: true, horizontal: "center" }
      }
    },
    {
      name: "Günlük",
      cells: {
        F1: { value: "Tarih:" }
      }
    }
  ]
};

function writeCellBold(sheet, X, Y, value) {
  const ascii = Y + 64;
  yPrime = String.fromCharCode(ascii);
  const cell = `${yPrime}${X}`;
  content.sheets[sheet].cells[cell] = {
    bold: true,
    vertical: "center",
    fontSize: 10,
    value: value
  };
}
function writeCellHeader(sheet, X, Y, Z, T, value) {
  const ascii = Y + 64;
  yPrime = String.fromCharCode(ascii);
  const ascii_2 = T + 64;
  tPrime = String.fromCharCode(ascii_2);
  const cell = `${yPrime}${X}:${tPrime}${Z}`;
  content.sheets[sheet].cells[cell] = {
    bold: true,
    vertical: "center",
    fontSize: 10,
    value: value,
    horizontal: "center"
  };
}
function writeCell(sheet, X, Y, value) {
  const ascii = Y + 64;
  yPrime = String.fromCharCode(ascii);
  const cell = `${yPrime}${X}`;
  content.sheets[sheet].cells[cell] = {
    vertical: "center",
    fontSize: 10,
    value: value
  };
}

function IhlalRapor() {
  // Ihlaller Raporu
  writeCellBold(0, 2, 1, "Rapor:");
  writeCellBold(0, 3, 1, "Saha:");
  writeCellBold(0, 4, 1, "Oluşturma Tarihi:");
  writeCell(0, 2, 2, "Maden Ihlaller Raporu");
  writeCell(0, 3, 2, Ihlal_Raporu.saha);
  writeCell(0, 4, 2, Ihlal_Raporu.date);

  // Günlük
  const data = Ihlal_Raporu.reports[0].data;
  let currentRow = 3;
  for (let d = 0; d < data.length; d++) {
    const startTime = data[d].startTime;
    const endTime = data[d].endTime;
    writeCellHeader(
      1,
      currentRow,
      1,
      currentRow,
      7,
      `Vardiya ${d + 1}(${startTime} - ${endTime})`
    );
    currentRow++;
    const headers = [
      "Sürücüler",
      "Sürücü Tipi",
      "Kullandığı Araç",
      "Hız İhlali",
      "Bölge İhlali",
      "Yanlış Dökme",
      "Motor Zorlama Uyarısı"
    ];
    for (let i = 1; i <= 7; i++) {
      writeCellBold(1, currentRow, i, headers[i - 1]);
    }
    currentRow++;
    const entries = data[d].entries;
    let sumList = [0, 0, 0, 0, 0, 0];
    for (let i = 0; i < entries.length; i++) {
      let column = 1;
      let object = entries[i];
      for (var attributename in object) {
        writeCell(1, currentRow, column, object[attributename]);
        const value = Number(object[attributename]);
        if (attributename === "speedviolation") {
          if (object[attributename] !== "-") sumList[2] += value;
        } else if (attributename === "regionviolation") {
          if (object[attributename] !== "-") sumList[3] += value;
        } else if (attributename === "yanlisdokme") {
          if (object[attributename] !== "-") sumList[4] += value;
        } else if (attributename === "motorzorlama") {
          if (object[attributename] !== "-") sumList[5] += value;
        }
        column++;
      }
      currentRow++;
    }
    writeCellBold(1, currentRow, 1, "TOPLAM");
    for (let i = 0; i < sumList.length; i++) {
      if (sumList[i] === 0) {
        writeCell(1, currentRow, i + 2, "-");
      } else writeCell(1, currentRow, i + 2, sumList[i]);
    }
    currentRow += 2;
  }

  createExcel(content);
}

module.exports = IhlalRapor;
