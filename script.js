let allData = JSON.parse(localStorage.getItem("allData")) || [];

// ================= UPLOAD EXCEL =================
document.getElementById("uploadBtn").addEventListener("click", uploadExcel);

function uploadExcel() {
    const file = document.getElementById("excelFile").files[0];

    if (!file) {
        alert("Please select a file");
        return;
    }

    let reader = new FileReader();

    reader.onload = function (e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: "array" });

        let sheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(sheet);

        // ================= UPSERT LOGIC =================
        jsonData.forEach(newItem => {

            let key = (newItem.Employee || newItem.EmployeeID) + "_" + newItem.Task;

            let index = allData.findIndex(item =>
                ((item.Employee || item.EmployeeID) + "_" + item.Task) === key
            );

            if (index !== -1) {
                allData[index] = {
                    ...allData[index],
                    ...newItem
                };
            } else {
                allData.push(newItem);
            }
        });

        localStorage.setItem("allData", JSON.stringify(allData));

        renderTable(allData);
        updateDashboard(allData);
        populateDepartments(allData);
    };

    reader.readAsArrayBuffer(file);
}


// ================= FORMAT EXCEL DATE =================
function formatExcelDate(value) {
    if (!value) return "-";

    // already normal date string
    if (isNaN(value)) return value;

    const date = new Date((value - 25569) * 86400 * 1000);

    return date.toISOString().split("T")[0];
}


// ================= TABLE =================
function renderTable(data) {
    if (!data.length) return;

    let html = `
    <table>
        <tr>
            <th>Employee</th>
            <th>Department</th>
            <th>Task</th>
            <th>Start Date</th>
            <th>End Date</th>
            <th>Deadline</th>
            <th>Progress</th>
            <th>Delay</th>
        </tr>
    `;

    data.forEach(row => {

        const start = formatExcelDate(row["Start Date"]);
        const end = formatExcelDate(row["End Date"]);
        const deadline = formatExcelDate(row.Deadline);

        const delay = getDelay(deadline, end);

        html += `
        <tr>
            <td>${row.Employee || row.EmployeeID || "-"}</td>
            <td>${row.Department || "-"}</td>
            <td>${row.Task || "-"}</td>

            <td>${start}</td>
            <td>${end}</td>
            <td>${deadline}</td>

            <td>
                <div class="progress-container">
                    <div class="progress-fill" style="width:${row.Progress || 0}%">
                        ${row.Progress || 0}%
                    </div>
                </div>
            </td>

            <td>
                ${
                    delay > 0
                    ? `<span class="delay">${delay} days</span>`
                    : `<span class="ontime">On Time</span>`
                }
            </td>
        </tr>
        `;
    });

    html += "</table>";

    document.getElementById("output").innerHTML = html;
}


// ================= DASHBOARD =================
function updateDashboard(data) {

    let employees = new Set();
    let departments = new Set();
    let delayed = 0;

    data.forEach(item => {

        employees.add(item.Employee || item.EmployeeID);
        departments.add(item.Department);

        const end = formatExcelDate(item["End Date"]);
        const deadline = formatExcelDate(item.Deadline);

        if (getDelay(deadline, end) > 0) {
            delayed++;
        }
    });

    document.getElementById("employeeCount").innerText = employees.size;
    document.getElementById("taskCount").innerText = data.length;
    document.getElementById("departmentCount").innerText = departments.size;
    document.getElementById("delayCount").innerText = delayed;
}


// ================= FILTER =================
document.getElementById("departmentFilter").addEventListener("change", function () {
    let value = this.value;

    if (value === "all") {
        renderTable(allData);
        return;
    }

    let filtered = allData.filter(item => item.Department === value);
    renderTable(filtered);
});


// ================= POPULATE FILTER =================
function populateDepartments(data) {

    let select = document.getElementById("departmentFilter");

    let departments = [...new Set(data.map(i => i.Department))];

    select.innerHTML = `<option value="all">All Departments</option>`;

    departments.forEach(dep => {
        select.innerHTML += `<option value="${dep}">${dep}</option>`;
    });
}


// ================= DELAY CALC =================
function getDelay(deadline, endDate) {

    if (!deadline || !endDate) return 0;

    const d1 = new Date(deadline);
    const d2 = new Date(endDate);

    let diff = Math.ceil((d2 - d1) / (1000 * 60 * 60 * 24));

    return diff > 0 ? diff : 0;
}


// ================= INIT =================
renderTable(allData);
updateDashboard(allData);
populateDepartments(allData);
document.getElementById("clearBtn").addEventListener("click", clearAllData);

function clearAllData() {
    let confirmDelete = confirm("Are you sure you want to delete ALL data?");

    if (!confirmDelete) return;

    // Clear storage
    localStorage.removeItem("allData");

    // Reset variable
    allData = [];

    // Clear UI
    document.getElementById("output").innerHTML = "";

    // Reset dashboard
    document.getElementById("employeeCount").innerText = 0;
    document.getElementById("taskCount").innerText = 0;
    document.getElementById("departmentCount").innerText = 0;
    document.getElementById("delayCount").innerText = 0;

    // Reset dropdown
    document.getElementById("departmentFilter").innerHTML =
        `<option value="all">All Departments</option>`;
}
