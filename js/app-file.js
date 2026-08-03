// Handle file selection and initiate workbook parsing
function handleFileSelect(event) {
    const file = event.target.files[0];

    if (!file) {
        return;
    }

    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        showStatus('Please select a valid XLSX or XLS file.', 'error');
        return;
    }

    fileName.textContent = `Selected: ${file.name}`;
    showStatus('Reading file...', 'info');

    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbookData = XLSX.read(data, { type: 'array' });

            parseWorkbook(workbookData);
            showStatus('File loaded successfully!', 'success');
            downloadBtn.disabled = false;
        } catch (error) {
            showStatus('Error reading file: ' + error.message, 'error');
            downloadBtn.disabled = true;
        }
    };

    reader.onerror = function() {
        showStatus('Error reading file.', 'error');
        downloadBtn.disabled = true;
    };

    reader.readAsArrayBuffer(file);
}

// Parse workbook and extract data for preview and report generation
function parseWorkbook(workbook) {
    try {
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        parsedData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        displayPreview(parsedData);

        generateWordPreview(parsedData);
        generateTextPreview(parsedData);
        reportPreviewSection.style.display = 'block';
    } catch (error) {
        showStatus('Error parsing workbook: ' + error.message, 'error');
    }
}
