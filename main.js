/******************************************************************************
 *
 * Project : XML
 * File    : main.js
 * Version : 1.0.0
 *
 ******************************************************************************/

"use strict";

//=============================================================================
// DOM
//=============================================================================

const fileInput = document.getElementById("fileInput");
const convertBtn = document.getElementById("convertBtn");
const statusOutput = document.getElementById("statusOutput");
const downloadBtn = document.getElementById("downloadBtn");

//=============================================================================
// APP
//=============================================================================

const App = {

    workbook: null,
    outputFileName: "",

    init() {

        UI.hideDownload();

        convertBtn.addEventListener("click", App.convert);

        UI.setStatus("Ready");

    },

    async convert() {

        try {

            UI.hideDownload();

            if (fileInput.files.length === 0) {
                UI.error("Please select XML file.");
                return;
            }

            const file = fileInput.files[0];

            App.outputFileName =
                file.name.replace(/\.xml$/i, "") + "_output.xlsx";

            UI.info("Reading XML...");

            const xmlText = await file.text();

            UI.info("Parsing SpreadsheetML...");

            const workbookData = Parser.parse(xmlText);

            UI.info("Creating Workbook...");

            App.workbook = ExcelBuilder.create(workbookData);

            UI.success("Completed.");

            UI.showDownload();

        }
        catch (ex) {

            console.error(ex);

            UI.error(ex.message);

        }

    }

};

//=============================================================================
// UI
//=============================================================================

const UI = {

    setStatus(text, color = "#222") {

        statusOutput.textContent = text;
        statusOutput.style.color = color;

    },

    info(text) {

        UI.setStatus(text, "#2563eb");

    },

    success(text) {

        UI.setStatus(text, "#16a34a");

    },

    error(text) {

        UI.setStatus(text, "#dc2626");

    },

    hideDownload() {

        downloadBtn.style.display = "none";

        downloadBtn.removeAttribute("href");

    },

    showDownload() {

        const wbout = XLSX.write(
            App.workbook,
            {
                bookType: "xlsx",
                type: "array"
            }
        );

        const blob = new Blob(
            [wbout],
            {
                type:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
        );

        const url = URL.createObjectURL(blob);

        downloadBtn.href = url;

        downloadBtn.download = App.outputFileName;

        downloadBtn.style.display = "block";

    }

};

//=============================================================================
// PARSER
//=============================================================================

const Parser = {

    parse(xmlText) {

        const dom = new DOMParser()
            .parseFromString(xmlText, "text/xml");

        if (dom.querySelector("parsererror"))
            throw new Error("Invalid XML.");

        const workbook = {

            worksheets: []

        };

        const worksheets = Utils.childrenByName(
            dom.documentElement,
            "Worksheet"
        );

        if (worksheets.length === 0)
            throw new Error("Worksheet not found.");

        for (const ws of worksheets) {

            workbook.worksheets.push(

                Parser.parseWorksheet(ws)

            );

        }

        return workbook;

    },

    parseWorksheet(node) {

        const result = {

            name:
                Utils.attribute(node, "Name") || "Sheet",

            rows: []

        };

        const table =
            Utils.child(node, "Table");

        if (!table)
            return result;

        const rows =
            Utils.childrenByName(table, "Row");

        for (const row of rows) {

            result.rows.push(

                Parser.parseRow(row)

            );

        }

        return result;

    },

    parseRow(rowNode) {

        const row = [];

        const cells =
            Utils.childrenByName(rowNode, "Cell");

        for (const cell of cells) {

            row.push(

                Parser.parseCell(cell)

            );

        }

        return row;

    },

    parseCell(cellNode) {

        const data =
            Utils.child(cellNode, "Data");

        if (!data)
            return "";

        return Utils.convertValue(
            data.textContent,
            Utils.attribute(data, "Type")
        );

    }

};

//=============================================================================
// EXCEL
//=============================================================================

const ExcelBuilder = {

    create(data) {

        const wb = XLSX.utils.book_new();

        for (const ws of data.worksheets) {

            const sheet =
                XLSX.utils.aoa_to_sheet(ws.rows);

            XLSX.utils.book_append_sheet(

                wb,

                sheet,

                ws.name

            );

        }

        return wb;

    }

};

//=============================================================================
// UTILITIES
//=============================================================================

const Utils = {

    child(node, name) {

        return Utils.childrenByName(node, name)[0] || null;

    },

    childrenByName(node, name) {

        return [...node.children]
            .filter(x => x.localName === name);

    },

    attribute(node, name) {

        for (const a of node.attributes) {

            if (a.localName === name)
                return a.value;

        }

        return null;

    },

    convertValue(value, type) {

        if (value == null)
            return "";

        value = value.trim();

        if (value.startsWith("<"))
            return "";

        switch ((type || "").toLowerCase()) {

            case "number":

                return Number(value);

            case "boolean":

                return value === "1";

            case "datetime": {

                const d = new Date(value);

                if (!isNaN(d))
                    return d;

                return value;

            }

            default:

                return value;

        }

    }

};

//=============================================================================
// INITIALIZE
//=============================================================================

App.init();