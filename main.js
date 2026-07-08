const fileInput = document.getElementById("fileInput");
const convertBtn = document.getElementById("convertBtn");
const status = document.getElementById("status");
const downloadBtn = document.getElementById("downloadBtn");

convertBtn.addEventListener("click", convertXML);

function setStatus(text, color = "black") {
    status.textContent = text;
    status.style.color = color;
}

async function convertXML() {

    if (fileInput.files.length === 0) {
        setStatus("⚠ Please select an XML file.", "red");
        return;
    }

    try {

        setStatus("⏳ Reading XML...", "blue");

        const file = fileInput.files[0];
        const xmlText = await file.text();

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, "text/xml");

        if (xmlDoc.querySelector("parsererror")) {
            throw new Error("Invalid XML");
        }

        const worksheet = xmlDoc.getElementsByTagName("Worksheet")[0];

        if (!worksheet)
            throw new Error("Worksheet not found.");

        const table = worksheet.getElementsByTagName("Table")[0];

        if (!table)
            throw new Error("Table not found.");

        const rows = [];

        const xmlRows = table.getElementsByTagName("Row");

        for (const row of xmlRows) {

            const rowData = [];

            const cells = row.getElementsByTagName("Cell");

            for (const cell of cells) {

                const data = cell.getElementsByTagName("Data")[0];

                let value = data ? data.textContent : "";

                //--------------------------------------------------
                // Remove XML text
                //--------------------------------------------------

                if (typeof value === "string" && value.startsWith("<"))
                    value = "";

                //--------------------------------------------------
                // ISO Date
                //--------------------------------------------------

                if (typeof value === "string" &&
                    /^\d{4}-\d{2}-\d{2}T/.test(value)) {

                    const d = new Date(value);

                    if (!isNaN(d))
                        value = d;
                }

                rowData.push(value);
            }

            rows.push(rowData);

        }

        //------------------------------------------------------
        // Create Workbook
        //------------------------------------------------------

        const wb = XLSX.utils.book_new();

        const ws = XLSX.utils.aoa_to_sheet(rows);

        XLSX.utils.book_append_sheet(
            wb,
            ws,
            worksheet.getAttribute("ss:Name") || "Sheet1"
        );

        //------------------------------------------------------
        // Download
        //------------------------------------------------------

        const outputName =
            file.name.replace(/\.xml$/i, "") + "_output.xlsx";

        XLSX.writeFile(wb, outputName);

        setStatus("✅ Completed!", "green");

    }
    catch (err) {

        console.error(err);

        setStatus("❌ " + err.message, "red");

    }

}