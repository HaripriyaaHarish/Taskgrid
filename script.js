let allData = JSON.parse(localStorage.getItem("allData")) || [];


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
        console.log(allData);
        localStorage.setItem("allData", JSON.stringify(allData));


allData = JSON.parse(localStorage.getItem("allData")) || [];

refreshUI(allData);
    };

    reader.readAsArrayBuffer(file);
}



function formatExcelDate(value) {
    if (!value) return "-";


    if (isNaN(value)) return value;

    const date = new Date((value - 25569) * 86400 * 1000);

    return date.toISOString().split("T")[0];
}



function renderTable(data) {
    if (!data.length) return;

    let allDates = [];

   
    data.forEach(item => {
        let start = new Date(formatExcelDate(item["Start Date"]));
        let end = new Date(formatExcelDate(item["End Date"]));

        if (!isNaN(start)) allDates.push(start);
        if (!isNaN(end)) allDates.push(end);
    });

    let minDate = new Date(Math.min(...allDates));
    let maxDate = new Date(Math.max(...allDates));

   
    let timeline = generateTimeline(minDate, maxDate);

    let html = `<table class="gantt">
        <tr>
            <th>Employee</th>
            <th>Task</th>`;

    timeline.forEach(date => {
        html += `<th>
        ${date.toLocaleDateString("en-GB", {
            day: "2-digit",
            month: "short"
        })}
    </th>`;
    });

    html += `</tr>`;

  
    data.forEach(item => {

        let start = new Date(formatExcelDate(item["Start Date"]));
        let end = new Date(formatExcelDate(item["End Date"]));
        let progress = item.Progress || 0;

        html += `<tr>
            <td>${item.Employee}</td>
            <td>${item.Task}</td>`;

        timeline.forEach(date => {

            let cellDate = new Date(date);

            let inRange = cellDate >= start && cellDate <= end;

            let color = "";

            if (inRange) {
                let progressLimit = Math.floor((progress / 100) * (end - start) / (1000 * 60 * 60 * 24));

                let currentDiff = (cellDate - start) / (1000 * 60 * 60 * 24);

                color = currentDiff <= progressLimit ? "#4ade80" : "#60a5fa";
            }

            html += `<td style="background:${inRange ? color : "transparent"}"></td>`;
        });

        html += `</tr>`;
    });

    html += "</table>";

    document.getElementById("output").innerHTML = html;
}

function updateDashboard(data) {

    let employees = new Set();
    let departments = new Set();
    let delayed = 0;

    data.forEach(item => {

        let employee =
            item.Employee ||
            item.EmployeeID ||
            item["Employee Name"] ||
            item.employee ||
            item.Name;

        let department =
            item.Department ||
            item.department;

        if (employee) {
            employees.add(employee);
        }

        if (department) {
            departments.add(department);
        }

        let deadline = formatExcelDate(item.Deadline);
        let endDate = formatExcelDate(item["End Date"]);

        if (getDelay(deadline, endDate) > 0) {
            delayed++;
        }
    });

    document.getElementById("employeeCount").textContent = employees.size;

    document.getElementById("taskCount").textContent = data.length;

    document.getElementById("departmentCount").textContent = departments.size;

    document.getElementById("delayCount").textContent = delayed;
}
document.getElementById("departmentFilter")
    .addEventListener("change", function () {

        let value = this.value;

        let filtered =
            value === "all"
                ? allData
                : allData.filter(item =>
                    item.Department === value
                );

        renderTable(filtered);

        renderGantt(filtered);

        updateDashboard(filtered);
    });



function populateDepartments(data) {

    let select = document.getElementById("departmentFilter");

    let departments = [...new Set(data.map(i => i.Department))];

    select.innerHTML = `<option value="all">All Departments</option>`;

    departments.forEach(dep => {
        select.innerHTML += `<option value="${dep}">${dep}</option>`;
    });
}



function getDelay(deadline, endDate) {

    if (!deadline || !endDate) return 0;

    const d1 = new Date(deadline);
    const d2 = new Date(endDate);

    let diff = Math.ceil((d2 - d1) / (1000 * 60 * 60 * 24));

    return diff > 0 ? diff : 0;
}



renderTable(allData);
renderGantt(allData);
updateDashboard(allData);
populateDepartments(allData);
document.getElementById("clearBtn").addEventListener("click", clearAllData);

function clearAllData() {

    let confirmDelete = confirm(
        "Are you sure you want to delete all data?"
    );

    if (!confirmDelete) return;

    localStorage.removeItem("allData");


    allData = [];


    document.getElementById("output").innerHTML = "";


    if (document.getElementById("ganttChart")) {
        document.getElementById("ganttChart").innerHTML = "";
    }

    if (document.getElementById("ganttHeader")) {
        document.getElementById("ganttHeader").innerHTML = "";
    }


    document.getElementById("employeeCount").innerText = 0;
    document.getElementById("taskCount").innerText = 0;
    document.getElementById("departmentCount").innerText = 0;
    document.getElementById("delayCount").innerText = 0;

  
    document.getElementById("departmentFilter").innerHTML =
        `<option value="all">All Departments</option>`;


    document.getElementById("excelFile").value = "";
    document.getElementById("ganttChart").innerHTML = "";

    document.getElementById("ganttHeader").innerHTML = "";
    alert("All data cleared successfully.");
}
function generateTimeline(start, end) {
    let dates = [];
    let current = new Date(start);
    let last = new Date(end);

    while (current <= last) {
        dates.push(new Date(current));
        current.setDate(current.getDate() + 1);
    }

    return dates;
}
function renderGanttHeader(minDate, maxDate) {
    const header = document.getElementById("ganttHeader");
    header.innerHTML = "";

    let current = new Date(minDate);

    while (current <= maxDate) {

        let cell = document.createElement("div");
        cell.className = "gantt-date";

        cell.innerText = current.toLocaleDateString("en-GB", {
            day: "2-digit",
            month: "short"
        });

        header.appendChild(cell);

        current.setDate(current.getDate() + 1);
    }
}
function renderGantt(data) {
    const container = document.getElementById("ganttChart");
    container.innerHTML = "";

    if (!data.length) return;

    let minDate = new Date(Math.min(...data.map(d =>
        new Date(formatExcelDate(d["Start Date"]))
    )));

    let maxDate = new Date(Math.max(...data.map(d =>
        new Date(formatExcelDate(d["End Date"]))
    )));

    let totalDays = (maxDate - minDate) / (1000 * 60 * 60 * 24);

    renderGanttHeader(minDate, maxDate);

    data.forEach(item => {

        let start = new Date(formatExcelDate(item["Start Date"]));
        let end = new Date(formatExcelDate(item["End Date"]));

        let startOffset = (start - minDate) / (1000 * 60 * 60 * 24);
        let duration = (end - start) / (1000 * 60 * 60 * 24);

        let progress = item.Progress || 0;

        let bar = document.createElement("div");
        bar.className = "gantt-row";

        bar.innerHTML = `
            <div class="gantt-label">
                <b>${item.Employee}</b> - ${item.Task}
            </div>

            <div class="gantt-track">
                <div class="gantt-bar"
                    style="
                        left:${(startOffset / totalDays) * 100}%;
                        width:${(duration / totalDays) * 100}%;
                        background:#3b82f6;
                    ">
                    
                    <div class="gantt-progress"
                        style="width:${progress}%; background:#22c55e;">
                    </div>

                </div>
            </div>
        `;

        container.appendChild(bar);
    });
}
function refreshUI(data) {

    renderTable(data);

    renderGantt(data);

    updateDashboard(data);

    populateDepartments(data);
}
