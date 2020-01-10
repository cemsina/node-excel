const XLSX = require('xlsx-style');
"use strict";

const PARAM_FORMAT = {
    type: {target:"t", validation:s=>typeof s == "string"},
    value: {target: "v", validation:s=>typeof s == "string" || typeof s == "number"},
    fontName: {target:"s.font.name", validation:s=>typeof s == "string"},
    fontSize: {target:"s.font.sz", validation:s=>typeof s == "number"},
    bold: {target: "s.font.bold", validation:s=>typeof s == "boolean"},
    underline: {target: "s.font.underline", validation:s=>typeof s == "boolean"},
    italic: {target: "s.font.italic", validation:s=>typeof s == "boolean"},
    strike: {target: "s.font.strike", validation:s=>typeof s == "boolean"},
    outline: {target: "s.font.outline", validation:s=>typeof s == "boolean"},
    shadow: {target: "s.font.shadow", validation:s=>typeof s == "boolean"},
    fontColor: {target: "s.font.color.rgb", validation:s=>typeof s == "string"},
    fillColor: {target: "s.fill.fgColor.rgb", validation:s=>typeof s == "string"},
    horizontal: {target: "s.alignment.horizontal", validation:s=>s=="left" || s=="right" || s=="center"},
    vertical: {target: "s.alignment.vertical", validation:s=>s=="left" || s=="right" || s=="center"},
};

function set(ref, target,val){
    var arr = target.split(".");
    var last = arr.pop();
    var temp = arr.reduce((prev,cur)=>{
        prev[cur] = prev[cur] || {};
        return prev[cur];
    },ref);
    temp[last] = val;
}

function Cell(cell_range, params){
    "use strict";
    var json = {v:""};
    var isValid = true;
    var errorKey = null;
    var cells = cell_range.split(":");
    var decoded = cells.map(c=>XLSX.utils.decode_cell(c));

    this.decoded = ()=>decoded;
    this.id = ()=>cells[0];
    this.isValid = ()=>isValid;
    this.json = ()=>json;
    this.merges = ()=>cells.length==2 ? ({s:cells[0],e:cells[1]}) : null;
    this.errorKey = ()=>errorKey;

    (function init(){
        if(params.type)
            params.type = params.type[0];
        errorKey = Object.keys(params).find(p=>{
            if(PARAM_FORMAT[p] == undefined || !PARAM_FORMAT[p].validation(params[p]))
                return true;
            set(json,PARAM_FORMAT[p].target,params[p]);
        });
        isValid = !errorKey;
        return;
    })();
}


module.exports = Cell;