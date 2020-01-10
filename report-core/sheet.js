const Cell = require('./cell');
const XLSX = require('xlsx-style');
"use strict";

function Sheet(name,cells_json){
    "use strict";
    var cells = Object.keys(cells_json).reduce((prev,cur)=>{
        prev.push(new Cell(cur,cells_json[cur]));
        return prev;
    },[]);

    var isValid = !cells.some(c=>!c.isValid());
    var errorKey = !isValid ? cells.find(c=>!c.isValid()).errorKey() : null;
    this.isValid = ()=>isValid;
    this.errorKey = ()=>errorKey;

    var ref = cells.reduce((prev,c)=>{
        return c.decoded().reduce((prv,cur)=>{
            if(!prv.min || !prv.max){
                prv.min = {r:cur.r,c:cur.c};
                prv.max = {r:cur.r,c:cur.c};
                return prv;
            }
            prv.min.r = Math.min(prv.min.r,cur.r);
            prv.min.c = Math.min(prv.min.c,cur.c);
            prv.max.r = Math.max(prv.max.r,cur.r);
            prv.max.c = Math.max(prv.max.c,cur.c);
            return prv;
        },prev);
    },{});

    this.ref = ()=>`${XLSX.utils.encode_cell(ref.min)}:${XLSX.utils.encode_cell(ref.max)}`;

    this.name = ()=>name;

    this.json = ()=>cells.reduce((prev,cur)=>{
        prev[cur.id()] = cur.json();
        var merges = cur.merges();
        if(merges)
            prev["!merges"].push(merges);
        return prev;
    },{"!merges":[],"!ref":this.ref()});

    this.cells = ()=>cells;

}


module.exports = Sheet;