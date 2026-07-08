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

    status(text, color = "#222") {

        DOM.statusOutput.textContent = text;
        DOM.statusOutput.style.color = color;

    }

};

//=============================================================================
// APP
//=============================================================================

const App = {

    init() {

        UI.status("Ready");

        DOM.downloadBtn.style.display = "none";

        DOM.convertBtn.addEventListener("click", App.convert);

    },

    convert() {

        console.clear();

        console.log("Convert clicked");

        if (!DOM.fileInput.files.length) {

            UI.status("Please select XML file.", "red");

            return;

        }

        UI.status("Reading XML...", "#2563eb");

    }

};

//=============================================================================
// START
//=============================================================================

App.init();