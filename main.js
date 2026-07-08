/******************************************************************************
 *
 * Project : XML
 * File    : main.js
 * Version : 1.0.0
 * Status  : READY TO REPLACE
 *
 ******************************************************************************/

"use strict";

//=============================================================================
// DOM
//=============================================================================

const DOM = {
    fileInput: document.getElementById("fileInput"),
    convertBtn: document.getElementById("convertBtn"),
    statusOutput: document.getElementById("statusOutput"),
    downloadBtn: document.getElementById("downloadBtn")
};

//=============================================================================
// UI
//=============================================================================

const UI = {

    setStatus(text, color = "#222") {
        DOM.statusOutput.textContent = text;
        DOM.statusOutput.style.color = color;
    },

    hideDownload() {
        DOM.downloadBtn.style.display = "none";
        DOM.downloadBtn.removeAttribute("href");
    },

    showDownload(blob, filename) {

        const url = URL.createObjectURL(blob);

        DOM.downloadBtn.href = url;
        DOM.downloadBtn.download = filename;
        DOM.downloadBtn.style.display = "block";

    }

};

//=============================================================================
// MAIN
//=============================================================================

async function convertXML() {

    try {

        UI.hideDownload();

        if (DOM.fileInput.files.length === 0) {

            UI.setStatus("Please select XML file.", "red");
            return;

        }

        UI.setStatus("Reading XML...", "#2563eb");

        const file = DOM.fileInput.files[0];

        const xmlText = await file.text();

        UI.setStatus("Parsing XML...", "#2563eb");

        const xml = new DOMParser().parseFromString(
            xmlText,
            "text/xml"
        );

        if (xml.querySelector("parsererror")) {
            throw new Error("Invalid XML.");
        }

        const worksheet = findNode(xml.documentElement, "Worksheet");

        if (!worksheet) {
            throw new Error("Worksheet not found.");
        }

        const table = findNode(worksheet, "Table");

        if (!table) {
            throw new Error("Table not found.");
        }

        const rows = findChildren(table, "Row");

        const aoa = [];

        for (const row of rows) {

            const rowData = [];

            const cells = findChildren(row, "Cell");

            for (const cell of cells) {

                const data = findNode(cell, "Data");

                if (!data) {

                    rowData.push("");
                    continue;

                }

                rowData.push(data.textContent ?? "");

            }

            aoa.push(rowData);

        }

        UI.setStatus("Creating workbook...", "#2563eb");

        const wb = XLSX.utils.book_new();

        const ws = XLSX.utils.aoa_to_sheet(aoa);

        const sheetName =
            getAttribute(worksheet, "Name") || "Sheet1";

        XLSX.utils.book_append_sheet(
            wb,
            ws,
            sheetName
        );

        const wbout = XLSX.write(
            wb,
            {
                bookType: "xlsx",
                type: "array",
                compression: true
            }
        );

        const blob = new Blob(
            [wbout],
            {
                type:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
        );

        const filename =
            file.name.replace(/\.xml$/i, "") + "_output.xlsx";

        UI.showDownload(blob, filename);

        UI.setStatus(
            `Done (${aoa.length} rows)`,
            "#16a34a"
        );

    }
    catch (err) {

        console.error(err);

        UI.setStatus(err.message, "red");

    }

}

//=============================================================================
// XML HELPERS
//=============================================================================

function findNode(parent, localName) {

    for (const node of parent.children) {

        if (node.localName === localName)
            return node;

    }

    return null;

}

function findChildren(parent, localName) {

    const result = [];

    for (const node of parent.children) {

        if (node.localName === localName)
            result.push(node);

    }

    return result;

}

function getAttribute(node, name) {

    for (const attr of node.attributes) {

        if (attr.localName === name)
            return attr.value;

    }

    return null;

}

//=============================================================================
// START
//=============================================================================

UI.hideDownload();

UI.setStatus("Ready");

DOM.convertBtn.addEventListener(
    "click",
    convertXML
);