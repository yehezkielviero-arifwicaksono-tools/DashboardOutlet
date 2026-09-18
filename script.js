const EXCEL_FILE = "01. Informasi_Outlet_2026-R17.xlsx";

let excelData = [];
let headers = [];


// ==========================================
// LOAD EXCEL
// ==========================================

async function loadExcel() {

    try {

        const response = await fetch(EXCEL_FILE);

        if (!response.ok) {
            throw new Error(
                `File Excel tidak dapat ditemukan. Status: ${response.status}`
            );
        }

        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });


        // Ambil sheet pertama
        const firstSheetName = workbook.SheetNames[0];

        const worksheet =
            workbook.Sheets[firstSheetName];


        // Convert Excel → JSON
        excelData = XLSX.utils.sheet_to_json(
            worksheet,
            {
                defval: ""
            }
        );


        if (excelData.length === 0) {

            throw new Error(
                "Sheet Excel tidak memiliki data."
            );

        }


        // Ambil nama kolom
        headers = Object.keys(excelData[0]);


        // Update summary
        document.getElementById("totalData")
            .textContent = excelData.length.toLocaleString("id-ID");


        document.getElementById("totalColumns")
            .textContent = headers.length;


        // Tampilkan tabel
        renderTable(excelData);


        // Hide loading
        document.getElementById("loading")
            .style.display = "none";


    } catch (error) {

        console.error(error);

        document.getElementById("loading")
            .style.display = "none";

        document.getElementById("error")
            .textContent =
            "Gagal memuat Excel: " + error.message;

    }

}



// ==========================================
// RENDER TABLE
// ==========================================

function renderTable(data) {

    const tableHead =
        document.getElementById("tableHead");

    const tableBody =
        document.getElementById("tableBody");


    // Clear existing
    tableHead.innerHTML = "";
    tableBody.innerHTML = "";


    // HEADER
    const headerRow =
        document.createElement("tr");


    headers.forEach(header => {

        const th =
            document.createElement("th");

        th.textContent = header;

        headerRow.appendChild(th);

    });


    tableHead.appendChild(headerRow);


    // BODY
    data.forEach(row => {

        const tr =
            document.createElement("tr");


        headers.forEach(header => {

            const td =
                document.createElement("td");

            td.textContent =
                row[header] ?? "";

            tr.appendChild(td);

        });


        tableBody.appendChild(tr);

    });

}



// ==========================================
// SEARCH
// ==========================================

document
    .getElementById("searchInput")
    .addEventListener("input", function () {

        const keyword =
            this.value.toLowerCase();


        const filteredData =
            excelData.filter(row => {

                return headers.some(header => {

                    const value =
                        String(row[header] ?? "")
                            .toLowerCase();

                    return value.includes(keyword);

                });

            });


        renderTable(filteredData);

    });



// ==========================================
// START
// ==========================================

loadExcel();
