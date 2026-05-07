let fileName = "data.xlsx";

// Load existing Excel file on startup
window.onload = function () {
    fetch(fileName)
        .then(res => res.arrayBuffer())
        .then(buffer => {
            let workbook = XLSX.read(buffer, { type: "array" });
            let sheet = workbook.Sheets["Classes"];
            let data = XLSX.utils.sheet_to_json(sheet);
            renderTable(data);
        })
        .catch(() => console.log("No existing file found, starting fresh."));
};

function saveData() {
    let entry = {
        ClassName: document.getElementById("className").value,
        ClassDay: document.getElementById("classDay").value,
        Duration: document.getElementById("duration").value,
        Children: document.getElementById("children").value,
        Notes: document.getElementById("notes").value
    };

    fetch(fileName)
        .then(res => res.arrayBuffer())
        .then(buffer => {
            let workbook = XLSX.read(buffer, { type: "array" });
            let sheet = workbook.Sheets["Classes"];
            let data = XLSX.utils.sheet_to_json(sheet);
            data.push(entry);

            let newSheet = XLSX.utils.json_to_sheet(data);
            workbook.Sheets["Classes"] = newSheet;

            XLSX.writeFile(workbook, fileName);
            renderTable(data);
        })
        .catch(() => {
            let workbook = XLSX.utils.book_new();
            let sheet = XLSX.utils.json_to_sheet([entry]);
            XLSX.utils.book_append_sheet(workbook, sheet, "Classes");
            XLSX.writeFile(workbook, fileName);
            renderTable([entry]);
        });
}

function renderTable(data) {
    let tbody = document.querySelector("#dashboard tbody");
    tbody.innerHTML = "";

    data.forEach(row => {
        let tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${row.ClassName}</td>
            <td>${row.ClassDay}</td>
            <td>${row.Duration}</td>
            <td>${row.Children}</td>
            <td>${row.Notes}</td>
        `;
        tbody.appendChild(tr);
    });
}
