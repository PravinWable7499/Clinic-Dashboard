let data = [];
let filtered = [];
let currentList = [];
let currentIndex = 0;

const excelFile = "./Hospitals data.xlsx";

/* ================= LOAD EXCEL DATA ================= */
fetch(excelFile)
    .then(res => res.arrayBuffer())
    .then(buffer => {
        let workbook = XLSX.read(buffer, { type: "array" });

        let skincareSheet = workbook.Sheets["Skincare Cinics"];
        let generalSheet = workbook.Sheets["General Hospitals"];

        let skincareData = XLSX.utils.sheet_to_json(skincareSheet);
        let generalData = XLSX.utils.sheet_to_json(generalSheet);

        let skincare = skincareData.map(row => ({
            name: row["clinic_name"] || "-",
            type: "Skincare Clinic",
            mobile_no: row["contact_number"] || "Not Available",
            address: row["address"] || "-",
            area: row["Area"] || "Other"
        }));

        let general = generalData.map(row => ({
            name: row["Hospital Name"] || "-",
            type: "General Hospital",
            mobile_no: row["Phone Number"] || "Not Available",
            address: row["address"] || "-",
            area: row["Area"] || "Other"
        }));

        data = [...general, ...skincare];

        let pageType = document.body.getAttribute("data-page");

        if (pageType === "general") {
            filtered = data.filter(d => d.type === "General Hospital");
        } else if (pageType === "skincare") {
            filtered = data.filter(d => d.type === "Skincare Clinic");
        } else {
            filtered = data;
        }

        loadFilters();
        updateKPIs();
        showEmptyDetails();
    });

/* ================= HELPERS ================= */
function isPhoneAvailable(phone) {
    if (!phone) return false;

    let val = String(phone).toLowerCase().trim();

    return val !== "" &&
           val !== "not available" &&
           val !== "n/a" &&
           val !== "na" &&
           val !== "-";
}

/* ================= LOAD FILTERS ================= */
function loadFilters() {
    let nameFilter = document.getElementById("nameFilter");
    let phoneFilter = document.getElementById("phoneFilter");

    if (nameFilter) {
        nameFilter.innerHTML = `<option value="">Hospitals Name</option>`;

        filtered.forEach(d => {
            let opt = document.createElement("option");
            opt.value = d.name;
            opt.textContent = d.name;
            nameFilter.appendChild(opt);
        });
    }

}

/* ================= APPLY FILTER ================= */
function applyFilters() {
    let pageType = document.body.getAttribute("data-page");

    let baseData = data;
    
    if (pageType === "general") {
        baseData = data.filter(d =>
            d.type === "General Hospital"
        );
    }
    
    else if (pageType === "skincare") {
        baseData = data.filter(d =>
            d.type === "Skincare Clinic"
        );
    }

    else if (pageType === "overview") {
        baseData = data;
}

    let nameValue = document.getElementById("nameFilter")?.value || "";
    let phoneValue = document.getElementById("phoneFilter")?.value || "";

    filtered = baseData.filter(d => {
    return (
        (!nameValue || d.name === nameValue) &&
        (!phoneValue || d.mobile_no === phoneValue)
    );
});

    updateKPIs();
    renderDetails(filtered);

    if (nameValue || phoneValue) {

    document.getElementById("details")
    .scrollIntoView({
        behavior: "smooth",
        block: "start"
    });
    }
}

/* ================= KPI ================= */
function updateKPIs() {
    let total = filtered.length;

    let phoneAvailable = filtered.filter(d => isPhoneAvailable(d.mobile_no)).length;
    let phoneNotAvailable = total - phoneAvailable;

    let areaCount = new Set(
        filtered.map(d => d.area).filter(Boolean)
    ).size;

    setText("totalCount", total);
    setText("phoneAvailable", phoneAvailable);
    setText("phoneNotAvailable", phoneNotAvailable);
    setText("totalArea", areaCount);
    buildAreaCards();
}

function setText(id, value) {
    let el = document.getElementById(id);
    if (el) el.textContent = value;
}

/* ================= SEARCH ================= */
function handleSearch(event){

    if(event.key === "Enter"){
        runSearch();
    }
}

function runSearch() {

    let input = document.getElementById("site-search")
        ?.value
        .toLowerCase()
        .trim() || "";

    if (!input) {

        showEmptyDetails();
        return;
    }

    let result = filtered.filter(d =>

        d.name.toLowerCase().includes(input) ||

        d.area.toLowerCase().includes(input) ||

        d.mobile_no.toLowerCase().includes(input) ||

        d.address.toLowerCase().includes(input)
    );

    renderDetails(result);

    document.getElementById("details")
    .scrollIntoView({
        behavior: "smooth",
        block: "start"
    });
}

function showEmptyDetails(){

    let details =
        document.getElementById("details");

    if(!details) return;

    details.innerHTML = `
        <p>Select any hospitals to view details</p>
    `;
}

/* ================= DETAILS LIST ================= */
function renderDetails(list) {
    currentList = list;

    let details = document.getElementById("details");
    if (!details) return;

    if (list.length === 0) {
        details.innerHTML = `<p>No data found</p>`;
        return;
    }

    let html = "";

    list.forEach((d, index) => {
        let cleanNumber = getCleanNumber(d.mobile_no);

        html += `
            <div class="company-row" onclick="viewDetails(${data.indexOf(d)})">

                <div class="company-name">
                    ${d.name || "-"}
                </div>

                <div class="company-right">

                    <div class="company-phone">
                        ${
                            isPhoneAvailable(d.mobile_no)
                            ? `<a href="tel:${cleanNumber}" onclick="event.stopPropagation()">${d.mobile_no}</a>`
                            : "Not Available"
                        }
                    </div>

                    ${
                        cleanNumber.length === 10
                        ? `
                        <a href="tel:${cleanNumber}" 
                        class="call-btn"
                        onclick="event.stopPropagation()">
                        <i class="fa-solid fa-phone"></i>
                        </a>
                        
                        
                        <a href="https://wa.me/91${cleanNumber}?text=Hello"
                        target="_blank"
                        class="whatsapp-btn"
                        onclick="event.stopPropagation()">
                        <i class="fa-brands fa-whatsapp"></i>
                        </a>
                        `
                        : ""
                    }

                </div>

            </div>
        `;
    });

    details.innerHTML = html;
}

/* ================= FULL DETAILS ================= */
function viewDetails(index) {
    currentIndex = index;

    let d = data[index];
    let details = document.getElementById("details");

    details.innerHTML = `
        <div class="full-details">

            <div class="nav-buttons">
                <button class="back-btn" onclick="goBack()">⬅ Back</button>
                <button class="nav-btn" onclick="prevItem()">⬅ Previous</button>
                <button class="nav-btn" onclick="nextItem()">Next ➡</button>
            </div>

            <div class="detail-row">
                <div class="detail-label">Name</div>
                <div class="detail-value">${d.name || "-"}</div>
            </div>

            <div class="detail-row">
                <div class="detail-label">Type</div>
                <div class="detail-value">${d.type || "-"}</div>
            </div>

            <div class="detail-row">
                <div class="detail-label">Area</div>
                <div class="detail-value">${d.area || "-"}</div>
            </div>

            <div class="detail-row">
                <div class="detail-label">Phone Number</div>
                <div class="detail-value">
                    ${
                        isPhoneAvailable(d.mobile_no)
                        ? `<a href="tel:${getCleanNumber(d.mobile_no)}">${d.mobile_no}</a>`
                        : "-"
                    }
                </div>
            </div>

            <div class="detail-row">
                <div class="detail-label">Address</div>
                <div class="detail-value">
                    <a href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(d.address)}" target="_blank">
                        ${d.address || "-"}
                    </a>
                </div>
            </div>

        </div>
    `;
}

function goBack() {
    renderDetails(currentList);
}

function nextItem() {
    let currentObj = data[currentIndex];
    let index = currentList.indexOf(currentObj);

    if (index < currentList.length - 1) {
        viewDetails(data.indexOf(currentList[index + 1]));
    }
}

function prevItem() {
    let currentObj = data[currentIndex];
    let index = currentList.indexOf(currentObj);

    if (index > 0) {
        viewDetails(data.indexOf(currentList[index - 1]));
    }
}

/* ================= CLEAN PHONE ================= */
function getCleanNumber(raw) {
    if (!raw) return "";

    let clean = String(raw).replace(/\D/g, "");

    if (clean.length > 10 && clean.startsWith("91")) {
        clean = clean.slice(-10);
    }

    return clean.slice(-10);
}

/* ================= EXPORT CSV ================= */
function exportCSV() {
    let rows = [["Name", "Type", "Phone Number", "Area", "Address"]];

    currentList.forEach(d => {
        rows.push([
            d.name || "",
            d.type || "",
            d.mobile_no || "",
            d.area || "",
            d.address || ""
        ]);
    });

    let csvContent = rows.map(row =>
        row.map(value => `"${String(value).replace(/"/g, '""')}"`).join(",")
    ).join("\n");

    let blob = new Blob([csvContent], {
        type: "text/csv;charset=utf-8;"
    });

    let link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Hospital_Details.csv";
    link.click();
}
function buildAreaDropdown() {
    let areaBox = document.getElementById("areaDropdown");
    if (!areaBox) return;

    let areaCounts = {};

    filtered.forEach(d => {
        let area = (d.area || "Other").trim();
        areaCounts[area] = (areaCounts[area] || 0) + 1;
    });

    let html = "";

    Object.entries(areaCounts)
        .sort((a, b) => b[1] - a[1])
        .forEach(([area, count]) => {
            html += `
                <div class="area-item" onclick="showAreaData('${area}')">
                    ✅ ${area} - ${count}
                </div>
            `;
        });

    areaBox.innerHTML = html;
}

function toggleAreaDropdown() {
    document.getElementById("areaDropdown").classList.toggle("show");
}

function showAreaData(area) {
    let list = filtered.filter(d => (d.area || "Other").trim() === area);
    renderDetails(list);

    document.getElementById("areaDropdown").classList.remove("show");
}

function buildAreaCards(){

    let container = document.getElementById("areaCards");
    if(!container) return;

    let areaCounts = {};

    filtered.forEach(d => {

        let area = (d.area || "Others").trim();

        if(!area) area = "Others";

        areaCounts[area] = (areaCounts[area] || 0) + 1;
    });

    container.innerHTML = "";

    Object.entries(areaCounts)

        .sort((a, b) => {

            if (a[0] === "Other" || a[0] === "Others") return 1;
            if (b[0] === "Other" || b[0] === "Others") return -1;

            return b[1] - a[1];
        })

        .forEach(([area,count]) => {

            container.innerHTML += `

                <div class="area-card-box"
                     onclick="showAreaData('${area}')">

                    <div>
                        <div class="area-name">${area}</div>

                        <div class="area-count">
                            ${count} Hospitals
                        </div>
                    </div>

                    <div class="area-icon">🏥</div>

                </div>
            `;
        });
}

function showAreaData(area){
    let list = filtered.filter(d => (d.area || "Others").trim() === area);
    renderDetails(list);
    document.getElementById("details").scrollIntoView({ behavior:"smooth" });
}

function showKPIData(type) {

    let list = [];

    switch(type){

        case "totalCount":
            list = filtered;
            break;

        case "phoneAvailable":
            list = filtered.filter(d =>
                isPhoneAvailable(d.mobile_no)
            );
            break;

        case "phoneNotAvailable":
            list = filtered.filter(d =>
                !isPhoneAvailable(d.mobile_no)
            );
            break;
    }

    renderDetails(list);

    document.getElementById("details").scrollIntoView({
        behavior: "smooth",
        block: "start"
    });
}

function toggleExportMenu(){
    let menu = document.getElementById("exportMenu");
    menu.style.display = menu.style.display === "block" ? "none" : "block";
}

function exportCSV(){

    let rows = [["Name", "Phone Number"]];

    currentList.forEach(d => {
        rows.push([
            d.name || "",
            d.mobile_no || ""
        ]);
    });

    let csvContent = rows.map(row =>
        row.map(value => `"${String(value).replace(/"/g, '""')}"`).join(",")
    ).join("\n");

    let blob = new Blob([csvContent], { type:"text/csv;charset=utf-8;" });
    let link = document.createElement("a");

    link.href = URL.createObjectURL(blob);
    link.download = "Institute_Details.csv";
    link.click();

    document.getElementById("exportMenu").style.display = "none";
}

function exportWebsiteStylePDF(){

    const detailsBox = document.querySelector(".details-box");

    html2canvas(detailsBox, {
        scale: 2,
        useCORS: true
    }).then(canvas => {

        const imgData = canvas.toDataURL("image/png");
        const { jsPDF } = window.jspdf;

        const pdf = new jsPDF("p", "mm", "a4");

        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();

        const imgWidth = pageWidth - 20;
        const imgHeight = canvas.height * imgWidth / canvas.width;

        let heightLeft = imgHeight;
        let position = 10;

        pdf.addImage(imgData, "PNG", 10, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;

        while(heightLeft > 0){
            position = heightLeft - imgHeight + 10;
            pdf.addPage();
            pdf.addImage(imgData, "PNG", 10, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;
        }

        pdf.save("Hospital_Details.pdf");

        document.getElementById("exportMenu").style.display = "none";
    });
}