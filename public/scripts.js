document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.querySelector('input[type="file"]');
    const excelDataElement = document.getElementById('excelData');
    const downloadButton = document.getElementById('downloadExcel');

    fileInput.addEventListener('change', function(event) {
        const file = event.target.files[0];
        const reader = new FileReader();

        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Assuming there's only one sheet in the Excel file
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Convert worksheet to JSON format
            const excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Clear previous data
            excelDataElement.innerHTML = '';

            // Populate table with Excel data
            excelData.forEach((row, rowIndex) => {
                const tr = document.createElement('tr');
                row.forEach((cellData, columnIndex) => {
                    const td = document.createElement('td');
                    td.textContent = cellData;
                    tr.appendChild(td);

                    // Add edit button to each cell
                    const editButton = document.createElement('button');
                    editButton.textContent = 'Edit';
                    editButton.addEventListener('click', function() {
                        const newValue = prompt('Enter new value:');
                        if (newValue !== null) {
                            excelData[rowIndex][columnIndex] = newValue;
                            td.textContent = newValue;
                        }
                    });
                    td.appendChild(editButton);
                });
                excelDataElement.appendChild(tr);
            });

            // Enable download button
            downloadButton.disabled = false;
        };

        reader.readAsArrayBuffer(file);
    });

    downloadButton.addEventListener('click', function() {
        // Write edited Excel data to a new file and download it
        const editedWorkbook = XLSX.utils.book_new();
        const editedWorksheet = XLSX.utils.aoa_to_sheet(excelData);
        XLSX.utils.book_append_sheet(editedWorkbook, editedWorksheet, 'Sheet1');
        XLSX.writeFile(editedWorkbook, 'edited_excel.xlsx');
    });
});




document.addEventListener('DOMContentLoaded', function() {
    const uploadForm = document.getElementById('uploadForm');
    const excelDataElement = document.getElementById('excelData');
    const downloadButton = document.getElementById('downloadExcel');

    uploadForm.addEventListener('submit', function(event) {
        event.preventDefault();

        const formData = new FormData(this);

        fetch('/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(excelData => {
            // Clear previous data
            excelDataElement.innerHTML = '';

            // Populate table with Excel data
            excelData.forEach(row => {
                const tr = document.createElement('tr');
                row.forEach(cellData => {
                    const td = document.createElement('td');
                    td.textContent = cellData;
                    tr.appendChild(td);
                });
                excelDataElement.appendChild(tr);
            });

            // Enable download button
            downloadButton.disabled = false;
        })
        .catch(error => console.error('Error:', error));
    });

    downloadButton.addEventListener('click', function() {
        // Write edited Excel data to a new file and download it
        const editedWorkbook = XLSX.utils.book_new();
        const editedWorksheet = XLSX.utils.aoa_to_sheet(excelData);
        XLSX.utils.book_append_sheet(editedWorkbook, editedWorksheet, 'Sheet1');
        XLSX.writeFile(editedWorkbook, 'edited_excel.xlsx');
    });
});
