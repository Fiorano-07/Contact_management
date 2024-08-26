var selectedRow = null;

/* ---- Driver ---- */
function user_data() {
    if (validate()) {
        console.log("if valid is true");
        var formData = readFormData();
        if (selectedRow == null)
            insertNewRecord(formData);
        else
            updateRecord(formData);
        resetForm();
        exportToExcel(); // Export data to Excel whenever a change is made
    }
}

/* ---- Deletion ---- */
function onDelete(td) {
    if (confirm('Are you sure to delete this record?')) {
        row = td.parentElement.parentElement;
        document.getElementById("contactlist").deleteRow(row.rowIndex);
        resetForm();
        exportToExcel(); // Export data to Excel whenever a deletion is made
    }
}

/* ----- Updation & Edit---- */
function onEdit(td) {
    selectedRow = td.parentElement.parentElement;
    document.getElementById("username").value = selectedRow.cells[0].innerHTML;
    document.getElementById("dob").value = selectedRow.cells[1].innerHTML;
    document.getElementById("gender").value = selectedRow.cells[2].innerHTML;
    document.getElementById("number").value = selectedRow.cells[3].innerHTML;
    document.getElementById("address").value = selectedRow.cells[4].innerHTML;
    document.getElementById("email").value = selectedRow.cells[5].innerHTML;
}

/* ---- Updates an existing record ---- */
function updateRecord(formData) {
    selectedRow.cells[0].innerHTML = formData.username;
    selectedRow.cells[1].innerHTML = formData.dob;
    selectedRow.cells[2].innerHTML = formData.gender;
    selectedRow.cells[3].innerHTML = formData.number;
    selectedRow.cells[4].innerHTML = formData.address;
    selectedRow.cells[5].innerHTML = formData.email;
}

/* ---- Inserts a new record ---- */
function insertNewRecord(new_entry) {
    console.log("inserting a record");
    var table = document.getElementById("contactlist").getElementsByTagName('tbody')[0];
    var newRow = table.insertRow(table.length);
    cell1 = newRow.insertCell(0);
    cell1.innerHTML = new_entry.username;
    cell2 = newRow.insertCell(1);
    cell2.innerHTML = new_entry.dob;
    cell3 = newRow.insertCell(2);
    cell3.innerHTML = new_entry.gender;
    cell4 = newRow.insertCell(3);
    cell4.innerHTML = new_entry.number;
    cell5 = newRow.insertCell(4);
    cell5.innerHTML = new_entry.address;
    cell6 = newRow.insertCell(5);
    cell6.innerHTML = new_entry.email;
    cell7 = newRow.insertCell(6);
    cell7.innerHTML = `<a class="edits" onClick="onEdit(this)"><img class="pen" src="./edit.png"></a>
                       <a class="edits" onClick="onDelete(this)"><img class="pen" src="./delete.png"></a>`;
}

/* ---- Gets the input from user ---- */
function readFormData() {
    console.log("reading input data");
    var formData = {};
    formData["username"] = document.getElementById("username").value;
    formData["dob"] = document.getElementById("dob").value;
    formData["gender"] = document.getElementById("gender").value;
    formData["number"] = document.getElementById("number").value;
    formData["address"] = document.getElementById("address").value;
    formData["email"] = document.getElementById("email").value;
    return formData;
}

/* ---- Validation ---- */
function validate() {
    isValid = true;
    if(document.getElementById("username").value == "") {
        isValid = false;
        console.log("invalid");
        alert("Enter the name of Contact");
    } else {
        isValid = true;
        console.log("valid");
    }
    return isValid;
}

/* ---- Resets the form fields ---- */
function resetForm() {
    document.getElementById("username").value = "";
    document.getElementById("dob").value = "";
    document.getElementById("gender").value = "";
    document.getElementById("number").value = "";
    document.getElementById("address").value = "";
    document.getElementById("email").value = "";
    selectedRow = null;
}

/* ---- Search functionality ---- */
document.getElementById("search").addEventListener("keyup", function() {
    let searchQuery = this.value.toLowerCase();
    let rows = document.querySelectorAll("#contactlist tbody tr");

    rows.forEach(row => {
        let cells = row.getElementsByTagName("td");
        let isMatch = false;

        for (let i = 0; i < cells.length - 1; i++) { // Exclude the last cell with options
            if (cells[i].innerText.toLowerCase().includes(searchQuery)) {
                isMatch = true;
                break;
            }
        }

        row.style.display = isMatch ? "" : "none";
    });
});

/* ---- Export to Excel ---- */
function exportToExcel() {
    var wb = XLSX.utils.book_new();
    var ws_data = [];
    var table = document.getElementById("contactlist");
    
    // Add headers
    var headers = [];
    for (var i = 0; i < table.rows[0].cells.length; i++) {
        headers.push(table.rows[0].cells[i].innerText);
    }
    ws_data.push(headers);

    // Add rows
    for (var i = 1; i < table.rows.length; i++) {
        var rowData = [];
        for (var j = 0; j < table.rows[i].cells.length - 1; j++) { // Exclude the last cell with options
            rowData.push(table.rows[i].cells[j].innerText);
        }
        ws_data.push(rowData);
    }

    var ws = XLSX.utils.aoa_to_sheet(ws_data);
    XLSX.utils.book_append_sheet(wb, ws, "Contacts");

    XLSX.writeFile(wb, "contacts.xlsx");
}
