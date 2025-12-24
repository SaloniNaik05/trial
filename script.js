let data = [];
let workbook;
let worksheet;
let fileName = "";

document.getElementById("createExcel").onclick = () => {
    document.getElementById("formSection").classList.remove("hidden");

    const now = new Date();
    const ist = new Date(
        now.toLocaleString("en-US", { timeZone: "Asia/Kolkata" })
    );

    const date = ist.toLocaleDateString("en-GB").replaceAll("/", "-");
    const time = ist
        .toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit", hour12: true })
        .replaceAll(":", "-")
        .replaceAll(" ", "-");

    fileName = `${date}_${time}.xlsx`;

    data = [];
    data.push([
        "Date",
        "Time (IST)",
        "Product",
        "Quantity",
        "Rate (₹)",
        "Payment Mode"
    ]);
};

document.getElementById("enter").onclick = () => {
    const product = document.getElementById("product").value;
    const quantityValue = document.getElementById("quantityValue").value;
    const quantityUnit = document.getElementById("quantityUnit").value;
    const rate = document.getElementById("rate").value;
    const payment = document.getElementById("payment").value;

    if (!product || !quantityValue || !quantityUnit || !rate || !payment) {
        alert("Please fill all fields");
        return;
    }

    const now = new Date();
    const ist = new Date(
        now.toLocaleString("en-US", { timeZone: "Asia/Kolkata" })
    );

    const quantity = `${quantityValue} ${quantityUnit}`;

    data.push([
        ist.toLocaleDateString("en-GB"),
        ist.toLocaleTimeString("en-US"),
        product,
        quantity,
        rate,
        payment
    ]);

    document.getElementById("product").value = "";
    document.getElementById("quantityValue").value = "";
    document.getElementById("quantityUnit").value = "";
    document.getElementById("rate").value = "";
    document.getElementById("payment").value = "";
};

document.getElementById("submitExcel").onclick = () => {
    workbook = XLSX.utils.book_new();
    worksheet = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Daily Sales");

    document.getElementById("downloadSection").classList.remove("hidden");
};

document.getElementById("downloadExcel").onclick = () => {
    XLSX.writeFile(workbook, fileName);
};
