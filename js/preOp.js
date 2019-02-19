function ExcelExport(event) {
    var input = event.target;
    var reader = new FileReader();
    reader.onload = function () {
        var fileData = reader.result;
        var wb = XLSX.read(fileData, { type: 'binary' });

        wb.SheetNames.forEach(function (sheetName) {
            var rowObj = XLSX.utils.sheet_to_row_object_array(wb.Sheets[sheetName]);

            switch (sheetName) {
                case "Sheet1":
                    parseInfo(rowObj);
                    break;
                case "Sheet2":
                    parseSupply(rowObj);
                    break;
            }
        })
    };
    reader.readAsBinaryString(input.files[0]);
};
function parseInfo(jsonObj) {
    // console.log(jsonObj);
    document.getElementById("p-name").value = jsonObj[0]["Patient Name"];
    document.getElementById("p-gen").value = jsonObj[0]["Patient Gender"];
    document.getElementById("p-dob").value = jsonObj[0]["Patient DOB"];
    document.getElementById("s-name").value = jsonObj[0]["Surgery Name"];
    document.getElementById("d-code").value = jsonObj[0]["Surgeon Code"];
    document.getElementById("d-name").value = jsonObj[0]["Surgeon Name"];
    document.getElementById("s-date").value = jsonObj[0]["Surgery Assigned Date"] + " - " + jsonObj[0]["Surgery Assigned Time"];
    document.getElementById("s-w").value = jsonObj[0]["Priority Number"];
}
// done
function parseSupply(jsonStr) {
    // console.log(jsonObj);
    var table = document.getElementById('supply').getElementsByTagName('tbody')[0];
    for (var i = 0; i < jsonStr.length; i++) {
        var newRow = table.insertRow(table.rows.length);
        var newColumn = newRow.insertCell(0);
        newColumn.appendChild(document.createTextNode(i + 1));
        newColumn = newRow.insertCell(1);
        newColumn.appendChild(document.createTextNode(jsonStr[i].Code));
        newColumn = newRow.insertCell(2);
        newColumn.appendChild(document.createTextNode(jsonStr[i].Name));
        newColumn = newRow.insertCell(3);
        newColumn.appendChild(document.createTextNode(jsonStr[i].Quantity));
    }
}
function saveSurgeryProfile() {
    //imprort file check
    // alert("Surgery Profile Saved");
    window.location.replace("import.html");
    alert("Surgery Profile Saved");

    //TODO: send JSON object to sever and save profile and do appropriate actions.
};