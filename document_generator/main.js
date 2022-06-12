/*********************************************************************
** ---------------------- Copyright notice ---------------------------
** This source code is part of the EVASoft project
** It is property of Alain Boute Ingenierie - www.abing.fr and is
** distributed under the GNU Public Licence version 2
** Commercial use is submited to licencing - contact eva@abing.fr
** -------------------------------------------------------------------
**        File : main.js
** Description : It's the main file of the documents generator
**      Author : Luc TORRES
**     Created : Apr 2022
*********************************************************************/
"use strict";

// includes
global.os = require('os');
global.execSync = require("child_process").execSync;
global.process = require('process');
global.fs = require("fs");
global.jsdom = require("jsdom");
global.path = require("path");
// global variable
// Get DOMParser, same API as in browser
const dom = new jsdom.JSDOM("");
const DOMParser = dom.window.DOMParser;
global.parser = new DOMParser;

require("./src/config");
require("./src/function");

// processing
cleanSpace();

// open template files
config.extention.forEach(ext => {
    zip("unzip", config.path.template[ext], config.path.unziped[ext]);
});

// edit template file
// open
let ods = {
    contentDocument: getDocument(config.path.unziped.ods + "content.xml"),
    settingsDocument: getDocument(config.path.unziped.ods + "settings.xml"),
    metaInfManifestDocument: getDocument(config.path.unziped.ods + "META-INF/manifest.xml"),
}

let sheetTableau = ods.contentDocument.querySelector("[table:name='Tableau']");
let settingsTableau = ods.settingsDocument.querySelector("[config:name='Tables'] > [config:name='Tableau']");
let sheetGraph = ods.contentDocument.querySelector("[table:name='Graph']");

let nbObjectGen = 0;

// add sheets
config.lsSheet.forEach((sheetConfig, i) => {
    // clone sheet template
    let newSheet = sheetTableau.cloneNode(true);
    let newSettings = settingsTableau.cloneNode(true);

    // rename sheet
    let sheetName = "T" + (i + 1);
    newSheet.setAttribute("table:name", sheetName);

    // insert settings
    newSettings.setAttribute("config:name", sheetName);
    settingsTableau.insertAdjacentElement('beforebegin', newSettings);

    // edit table
    sheetConfig.lsTable.forEach((tableConfig, i) => {
        let tableauConfig = config.spreadSheetTemplate.tableau;

        if (i !== 0) {
            // define where duplicate the first table
            if (tableConfig.xIndexSheet > 0) {
                // add new table to the right
                for (let j = 1; j <= tableConfig.nbCol; j++) {
                    let copyIndex;
                    if (j <= tableConfig.nbBodyCol) {
                        // add body col
                        copyIndex = tableauConfig.nbHeadCol + 1;
                    }
                    else {
                        // add foot col
                        copyIndex = sheetConfig.lsTable[i - 1].xIndexEndTable + j - tableConfig.nbCol;
                    }
                    addCol(newSheet, copyIndex, sheetConfig.lsTable[i - 1].xIndexEndTable + j);
                }

                // split col title
                let secondRow = newSheet.getElementsByTagName("table:table-row")[1];
                // get 2 last elements
                let lastElem = secondRow.children[secondRow.children.length - 1];
                let beforeLastElem = secondRow.children[secondRow.children.length - 2];
                let newLastElem = lastElem.cloneNode(true);
                let newBeforeLastElem = beforeLastElem.cloneNode(true);
                // set width
                lastElem.setAttribute("table:number-columns-repeated", lastElem.getAttribute("table:number-columns-repeated") * 1 - tableConfig.nbCol);
                beforeLastElem.setAttribute("table:number-columns-spanned", beforeLastElem.getAttribute("table:number-columns-spanned") * 1 - tableConfig.nbCol);
                newLastElem.setAttribute("table:number-columns-repeated", newLastElem.getAttribute("table:number-columns-repeated") * 1 - sheetConfig.lsTable[i].nbCol);
                newBeforeLastElem.setAttribute("table:number-columns-spanned", newBeforeLastElem.getAttribute("table:number-columns-spanned") * 1 - sheetConfig.lsTable[i].nbCol);

                lastElem.insertAdjacentElement("afterend", newBeforeLastElem);
                newBeforeLastElem.insertAdjacentElement("afterend", newLastElem);
            }
            else {
                // add new table bottom
                for (let j = 1; j <= tableauConfig.nbHeadRow + tableConfig.nbRow; j++) {
                    let copyIndex;
                    if (j <= tableauConfig.nbHeadRow) {
                        // add head row
                        copyIndex = j;
                    }
                    else if (j <= tableauConfig.nbHeadRow + tableConfig.nbBodyRow) {
                        // add body row
                        copyIndex = tableauConfig.nbHeadRow + 1 + j % 2;
                    }
                    else {
                        // add foot row
                        copyIndex = sheetConfig.lsTable[i - 1].yIndexEndTable + j - (tableauConfig.nbHeadRow + tableConfig.nbRow);
                    }
                    addRow(newSheet, copyIndex, sheetConfig.lsTable[i - 1].yIndexEndTable + j);
                }
            }

            return;
        }

        if (tableConfig.nbRow - tableConfig.nbBodyRow == 1) {
            // remove last line
            rmRow(newSheet, tableauConfig.nbHeadRow + tableauConfig.nbBodyRow + tableauConfig.nbFootRow);
        }
        // add or remove line
        let nbRowDif = tableConfig.nbBodyRow - tableauConfig.nbBodyRow;
        for (let i = 0; i < Math.abs(nbRowDif); i++) {
            if (nbRowDif > 0) {
                let index = tableauConfig.nbHeadRow + tableauConfig.nbBodyRow + i;
                addRow(newSheet, index - 1, index + 1);
            }
            else {
                rmRow(newSheet, tableauConfig.nbHeadRow + tableauConfig.nbBodyRow - i);
            }
        }

        if (tableConfig.nbCol - tableConfig.nbBodyCol == 1) {
            // remove last column
            rmCol(newSheet, tableauConfig.nbHeadCol + tableauConfig.nbBodyCol + tableauConfig.nbFootCol);
        }
        // add or remove column
        let nbColDif = tableConfig.nbBodyCol - tableauConfig.nbBodyCol;
        for (let i = 0; i < Math.abs(nbColDif); i++) {
            if (nbColDif > 0) {
                let index = tableauConfig.nbHeadCol + tableauConfig.nbBodyCol + i + 1;
                addCol(newSheet, index - 1, index);
            }
            else {
                rmCol(newSheet, tableauConfig.nbHeadCol + tableauConfig.nbBodyCol - i);
            }
        }
    });

    // fill sheet
    let htmlTable = getHtmlTable(sheetConfig.dataFile.fileName);
    htmlTable.forEach((rowData, y) => {
        rowData.forEach((colData, x) => {
            setValue(newSheet, x + 1, y + 1, colData);
        });
    });

    // add chart
    let lsGraph = getLsGraph(i, sheetName);
    // add table:shapes tag in sheet
    let graphRow = sheetGraph.getElementsByTagName("table:table-row")[0].cloneNode(true);
    let graphCell = graphRow.firstElementChild;
    // empty graphCell
    graphCell.replaceChildren();
    newSheet.insertAdjacentElement('beforeend', graphRow);
    // fill table shape whith graph list
    let svgY = 0;
    lsGraph.forEach(graph => {
        nbObjectGen++;
        let newObjectFileName = "Object gen " + nbObjectGen;
        let graphElem = sheetGraph.getElementsByTagName("draw:frame")[config.spreadSheetTemplate.graph[graph.type] - 1].cloneNode(true);
        let drawObject = graphElem.getElementsByTagName("draw:object")[0];

        let currentObjectFileName = drawObject.getAttribute("xlink:href").split("./")[1];
        drawObject.setAttribute("draw:notify-on-update-of-ranges", graph.range);
        drawObject.setAttribute("xlink:href", "./" + newObjectFileName);

        // set graph relative height
        graphElem.setAttribute("svg:y", svgY + "cm");
        svgY += graphElem.getAttribute("svg:height").split("cm")[0] * 1;

        graphCell.appendChild(graphElem);

        // create object file for chart
        // duplicate chart folder
        copydirSync(config.path.unziped.ods + currentObjectFileName, config.path.unziped.ods + newObjectFileName);

        // content.xml
        // open
        let contentFilePath = `${config.path.unziped.ods}${newObjectFileName}/content.xml`;
        let chartContentDocument = getDocument(contentFilePath);

        // edit
        // titre
        let titleElem = chartContentDocument.getElementsByTagName("chart:title")[0].firstElementChild;
        titleElem.textContent = graph.title;
        // range
        let lsRangeElem = chartContentDocument.querySelectorAll("[table:cell-range-address], [chart:values-cell-range-address]");
        let graphRange = lsRangeElem[0];
        let keyRange = lsRangeElem[1];
        let valueRange = lsRangeElem[2];
        // add object in manifest
        let lsElemToDuplicate = ods.metaInfManifestDocument.querySelectorAll(`[manifest:full-path^='${currentObjectFileName}']`);
        lsElemToDuplicate.forEach(e => {
            let newElem = e.cloneNode(true);

            let elemAttr = newElem.getAttribute("manifest:full-path");
            newElem.setAttribute("manifest:full-path", elemAttr.replace(currentObjectFileName, newObjectFileName));

            e.insertAdjacentElement("beforebegin", newElem);
        });
        
        graphRange.setAttribute("table:cell-range-address", graph.range);
        keyRange.setAttribute("table:cell-range-address", graph.range.split(" ")[0]);
        valueRange.setAttribute("chart:values-cell-range-address", graph.range.split(" ")[1]);

        saveDocument(chartContentDocument, contentFilePath);
    });

    // insert sheet
    sheetTableau.insertAdjacentElement('beforebegin', newSheet);
});

// remove template sheets
sheetTableau.remove();
sheetGraph.remove();

// normalize ods
saveDocument(ods.contentDocument, config.path.unziped.ods + "content.xml");
saveDocument(ods.settingsDocument, config.path.unziped.ods + "settings.xml");
saveDocument(ods.metaInfManifestDocument, config.path.unziped.ods + "META-INF/manifest.xml");

// save ods and xlsx final file
zip("zip", config.path.unziped.ods, `${config.path.default}TableStatsRes.ods`);
fs.rmSync(config.path.unziped.ods, { recursive: true, force: true });
let lsCommand = {
    Windows_NT: {
        toXlsx: ``,
        toOds: ``,
    },
    Linux: {
        toXlsx: `libreoffice --convert-to xlsx TableStatsRes.ods`,
        toOds: `soffice --headless --convert-to ods *.xlsx`,
    },
};
execSync(lsCommand[os.type()].toXlsx);
fs.renameSync("TableStatsRes.xlsx", "TableStatsNormalize.xlsx");
execSync(lsCommand[os.type()].toOds);
fs.renameSync("TableStatsNormalize.xlsx", "TableStatsRes.xlsx");
zip("unzip", "TableStatsNormalize.ods", config.path.unziped.ods);

// generate odt
// parse files
ods = {
    contentDocument: getDocument(config.path.unziped.ods + "content.xml"),
    metaInfManifestDocument: getDocument(config.path.unziped.ods + "META-INF/manifest.xml"),
}
let odt = {
    contentDocument: getDocument(config.path.unziped.odt + "content.xml"),
    metaInfManifestDocument: getDocument(config.path.unziped.odt + "META-INF/manifest.xml"),
}

// select elements
let docTitle = odt.contentDocument.getElementsByTagName("text:h")[0];
let docObject = odt.contentDocument.getElementsByTagName("draw:frame")[0].parentElement;

let lsSheet = ods.contentDocument.getElementsByTagName("table:table");
for (let i = 0; i < lsSheet.length; i++) {
    let sheet = lsSheet[i];

    // add sheet in odt
    // add title
    let newTitle = docTitle.cloneNode();
    let htmlTable = getHtmlTable(config.lsSheet[i].dataFile.fileName);
    newTitle.textContent = `${i + 1}. ${htmlTable[0][0]}`;

    docTitle.insertAdjacentElement("beforebegin", newTitle);

    // add table
    let TableDrawFrame = docObject.cloneNode(true);
    let objectName = "Object gen " + (i + 1);
    TableDrawFrame
        .getElementsByTagName("draw:object")[0]
        .setAttribute("xlink:href", objectName);
    
    docTitle.insertAdjacentElement("beforebegin", TableDrawFrame);
    // add table object directory
    let objectPath = config.path.unziped.odt + objectName;
    copydirSync(config.path.unziped.odt + "Object 1", objectPath);
    // update manifest
    let lsElemToDuplicate = odt.metaInfManifestDocument.querySelectorAll(`[manifest:full-path^='Object 1/']`);
    lsElemToDuplicate.forEach(e => {
        let newElem = e.cloneNode(true);

        let elemAttr = newElem.getAttribute("manifest:full-path");
        newElem.setAttribute("manifest:full-path", elemAttr.replace('Object 1', `${objectName}`));

        e.insertAdjacentElement("beforebegin", newElem);
    });
    // open table object content file
    let objectContentPath = objectPath + "/content.xml";
    let objectContentDocument = getDocument(objectContentPath);
    // edit table object content file
    // define new elements
    // sheet
    let baseObject = sheet.cloneNode(true);
    // remove first line
    // baseObject.getElementsByTagName("table:table-row")[0].replaceChildren();
    // remove all drawable
    let lsDrawable = baseObject.getElementsByTagName("draw:frame");
    for (let j = 0; j < lsDrawable.length; j++) {
        lsDrawable[j].parentElement.parentElement.remove();
    }
    // other
    let sheetFont = ods.contentDocument.getElementsByTagName("office:font-face-decls")[0].cloneNode(true);
    let sheetStyle = ods.contentDocument.getElementsByTagName("office:automatic-styles")[0].cloneNode(true);

    let replaced = objectContentDocument.getElementsByTagName("office:font-face-decls")[0];
    replaced.parentElement.replaceChild(sheetFont, replaced);
    replaced = objectContentDocument.getElementsByTagName("office:automatic-styles")[0];
    replaced.parentElement.replaceChild(sheetStyle, replaced);
    replaced = objectContentDocument.getElementsByTagName("table:table")[0];
    replaced.parentElement.replaceChild(baseObject, replaced);
    // save table object content file
    saveDocument(objectContentDocument, objectContentPath);

    // add chart
    let lsChart = sheet.getElementsByTagName("draw:frame");
    for (let j = 0; j < lsChart.length; j++) {
        const chart = lsChart[j];
        
        let drawObject = chart.getElementsByTagName("draw:object")[0];
        let chartData = {
            width: chart.getAttribute("svg:width"),
            height: chart.getAttribute("svg:height"),
            href: drawObject.getAttribute("xlink:href"),
        }

        // edit main content
        let newDocObject = docObject.cloneNode(true);
        // size
        let chartDrawFrame = newDocObject.getElementsByTagName("draw:frame")[0];
        chartDrawFrame.setAttribute("svg:width", chartData.width);
        chartDrawFrame.setAttribute("svg:height", chartData.height);
        // xlink:href
        let chartDirName = chartData.href.replace(" ", " gen chart ");
        newDocObject.getElementsByTagName("draw:object")[0].setAttribute("xlink:href", chartDirName);

        docTitle.insertAdjacentElement("beforebegin", newDocObject);

        // copy chart directory
        copydirSync(path.join(config.path.unziped.ods, chartData.href), path.join(config.path.unziped.odt, chartDirName));

        // update manifest
        lsElemToDuplicate = ods.metaInfManifestDocument.querySelectorAll(`[manifest:full-path^='Object 1/']`);
        lsElemToDuplicate.forEach(e => {
            let newElem = e.cloneNode(true);

            let elemAttr = newElem.getAttribute("manifest:full-path");
            newElem.setAttribute("manifest:full-path", elemAttr.replace('Object 1', chartDirName.split("/")[1]));

            odt.metaInfManifestDocument.getElementsByTagName("manifest:manifest")[0].insertAdjacentElement("beforeend", newElem);
        });

        // add empty paragraph
        let separator = docObject.cloneNode();
        docTitle.insertAdjacentElement("beforebegin", separator);
    }
}

// remove template doc
docTitle.remove();
docObject.remove();
// remove object replacement and thumbnails
let lsDirName = ["ObjectReplacements", "Thumbnails"];
lsDirName.forEach(dirName => {
    let dirPath = config.path.unziped.odt + dirName;
    fs.rmSync(dirPath, { recursive: true, force: true });
    fs.mkdirSync(dirPath);
});
// save document as xml
saveDocument(odt.contentDocument, config.path.unziped.odt + "content.xml");
saveDocument(odt.metaInfManifestDocument, config.path.unziped.odt + "META-INF/manifest.xml");

// save odt final file
zip("zip", config.path.unziped.odt, `${config.path.default}TableStatsRes.odt`);

cleanSpace();