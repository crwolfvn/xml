/******************************************************************************
 *
 * Project : XML
 * File    : main.js
 * Version : 1.0.0
 * Status  : READY TO REPLACE
 *
 ******************************************************************************/

"use strict";

// Nạp Thư viện Keyword để dò cột cần chuyển sang Number
const response = await fetch("keywords.txt");
const keywords = (await response.text())
    .split(/\r?\n/)
    .map(x => x.trim().toLowerCase())
    .filter(x => x && !x.startsWith("#"));
//=============================================================================
// DOM
//=============================================================================

const DOM = {
    fileInput: document.getElementById("fileInput"),
    convertBtn: document.getElementById("convertBtn"),
    statusOutput: document.getElementById("statusOutput"),
    downloadBtn: document.getElementById("downloadBtn"),
    xmlSize: document.getElementById("xmlSize"),
    xlsxSize: document.getElementById("xlsxSize"),
    compressionRatio: document.getElementById("compressionRatio"), 
    progressBar: document.getElementById("progressBar"),
    progressText: document.getElementById("progressText"),
};


//=============================================================================
// UI
//=============================================================================

const UI = {
        downloadUrl: null,
    setStatus(text, color = "#222") {
        DOM.statusOutput.textContent = text;
        DOM.statusOutput.style.color = color;  },
    hideDownload() {
        DOM.downloadBtn.style.display = "none";
        DOM.downloadBtn.removeAttribute("href");  },
    showDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        DOM.downloadBtn.href = url;
        DOM.downloadBtn.download = filename;
        DOM.downloadBtn.style.display = "block";  },
    setProgress(percent) {
        DOM.progressBar.value = percent;
        DOM.progressText.textContent = percent + "%";}
};

//=============================================================================
// MAIN
//=============================================================================
function formatSize(bytes){
    if(bytes < 1024)
        return bytes + " B";
    if(bytes < 1024 * 1024)
        return (bytes / 1024).toFixed(1) + " KB";
    if(bytes < 1024 * 1024 * 1024)
        return (bytes / 1024 / 1024).toFixed(1) + " MB";
    return (bytes / 1024 / 1024 / 1024).toFixed(2) + " GB";   }

function shouldConvertToNumber(columnName) {
    if (!columnName) return false;
    return keywords.some(keyword =>  columnName.toLowerCase().includes(keyword)  );  }

function convertNumberCell(value) {
    if (typeof value !== "string") return value;
    const text = value.trim();
    if (text === "") return "";
    const number = Number(text.replace(/,/g, ""));
    return Number.isNaN(number) ? value : number; }


async function convertXML() {
    DOM.xlsxSize.textContent = "-";
    DOM.compressionRatio.textContent = "-";
    UI.setProgress(0);
    try {
        UI.hideDownload();
        if (DOM.fileInput.files.length === 0) {
            UI.setStatus("Chưa chọn file XML cần chuyển!", "red");
            return;   }
        UI.setStatus("Đang nạp file XML...", "#2563eb");
        UI.setProgress(6);
        const file = DOM.fileInput.files[0];
        DOM.xmlSize.textContent = formatSize(file.size);
        const xmlText = await file.text();
        UI.setStatus("Đang xử lý file XML...", "#2563eb");
        UI.setProgress(20);
        const xml = new DOMParser().parseFromString(
            xmlText,
            "text/xml"  );
        if (xml.querySelector("parsererror")) {  throw new Error("Invalid XML.");  } 
        const worksheet = findNode(xml.documentElement, "Worksheet");
        if (!worksheet) {  throw new Error("Worksheet not found.");  }
        const table = findNode(worksheet, "Table");
        if (!table) {  throw new Error("Table not found.");  }
        const rows = findChildren(table, "Row");
        UI.setProgress(30);
        const aoa = [];
        const totalRows = rows.length;
            for (let i = 0; i < totalRows; i++) {
            const row = rows[i];
            const rowData = [];
            const cells = findChildren(row, "Cell");
            for (const cell of cells) {
                const data = findNode(cell, "Data");
                if (!data) {
                    rowData.push("");
                    continue;  }
                rowData.push(data.textContent ?? "");       }
                aoa.push(rowData);  
                if (i % 200 === 0) {UI.setProgress(30 + Math.floor(i / totalRows * 50));
                await new Promise(requestAnimationFrame); }}

    const headers = aoa[0] || [];
    const numberColumns = headers.map(shouldConvertToNumber);
                    
        //=============================================================================
        // Smart ISO Date + XML Cleanup
        //=============================================================================

    function parseISODateTime(text) {
        if (typeof text !== "string")  return null;

        // XML chuẩn: YYYY-MM-DDTHH:MM:SS
        if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(text))
            return null;
        const [datePart, timePart] = text.split("T");
        const [y, m, d] = datePart.split("-").map(Number);
        const [hh, mm, ss] = timePart.split(":").map(Number);
        return new Date(y, m - 1, d, hh, mm, ss);  }

    for (let r = 1; r < aoa.length; r++) {
        for (let c = 0; c < aoa[r].length; c++) {
            const v = aoa[r][c];
            if (typeof v !== "string")
                continue;
            // XML Cleanup
            if (v.startsWith("<")) {
                aoa[r][c] = "";
                continue;  }
            // ISO DateTime
            const dt = parseISODateTime(v);
            if (dt) {aoa[r][c] = dt; continue;}
            if (numberColumns[c]) {aoa[r][c] = convertNumberCell(v);}  }
            if (r % 200 === 0) {UI.setProgress(80 + Math.floor(r / aoa.length * 12));
                await new Promise(requestAnimationFrame); }
    }

    UI.setStatus("Đang dựng file Xlsx...", "#2563eb");
    UI.setProgress(94);
    await new Promise(requestAnimationFrame);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (const addr in ws) {  if (addr.startsWith("!"))  continue;
    const cell = ws[addr];
    if (cell.v instanceof Date) {
        cell.t = "d";
        cell.z =
            cell.v.getHours() ||
            cell.v.getMinutes() ||
            cell.v.getSeconds()
                ? "yyyy-mm-dd hh:mm:ss"
                : "yyyy-mm-dd hh:mm:ss";  }  }
        const sheetName =
            getAttribute(worksheet, "Name") || "Sheet1";
        XLSX.utils.book_append_sheet(
            wb,
            ws,
            sheetName  );
        const wbout = XLSX.write(
        wb, {
        bookType: "xlsx",
        type: "array",
        compression: true,
        cellDates: true  }  );
        UI.setProgress(97);
        await new Promise(requestAnimationFrame);
        const blob = new Blob(
            [wbout],  {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } ); 
        DOM.xlsxSize.textContent = formatSize(blob.size);
        DOM.compressionRatio.textContent = (blob.size / file.size * 100).toFixed(2) + "%";
        const filename =  file.name.replace(/\.xml$/i, "") + "_output.xlsx";
        UI.showDownload(blob, filename);
        UI.setStatus(  `Đã xử lý xong ! (${aoa.length} rows)`,  "#16a34a" );
        UI.setProgress(100);  }

    // Kết thúc hàm TRY     
    catch (err) {
        console.error(err);
        UI.setStatus(err.message, "red"); }  }

//=============================================================================
// XML HELPERS
//=============================================================================

function findNode(parent, localName) {
    for (const node of parent.children) {
        if (node.localName === localName)  return node;  }
    return null; }

function findChildren(parent, localName) {
    const result = [];
    for (const node of parent.children) {
        if (node.localName === localName)
            result.push(node);  }
    return result;  }

function getAttribute(node, name) {
    for (const attr of node.attributes) {
        if (attr.localName === name)
            return attr.value;      }
    return null;    }

//=============================================================================
// START
//=============================================================================

UI.hideDownload();
UI.setStatus("Ready");
DOM.convertBtn.addEventListener(  "click", convertXML );
DOM.fileInput.addEventListener("change", () => {
    if (DOM.fileInput.files.length === 0) {
        DOM.xmlSize.textContent = "-";
        return;     }
    const file = DOM.fileInput.files[0];
    DOM.xmlSize.textContent = formatSize(file.size);    });