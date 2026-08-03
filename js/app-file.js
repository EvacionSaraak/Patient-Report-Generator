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

// Parse workbook into a canonical result and set up all previews.
// parsedData is set to the canonical object returned by parseNormalSheet:
//   { format, headerRowIndex, columns, records, rawRows }
function parseWorkbook(workbook) {
    try {
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Normal format is tested first; it returns null when not detected.
        const canonical = parseNormalSheet(rawRows);
        if (!canonical) {
            throw new Error(
                'Could not detect a supported report format in this workbook. ' +
                'Ensure the file contains headers such as PT ID., Patient Name, and Visit Date.'
            );
        }

        parsedData = canonical;

        // Pass the sheet slice starting at the header row so that displayPreview
        // receives headers in row 0 (the only place data[0] / data.slice(1) is used).
        displayPreview(rawRows.slice(canonical.headerRowIndex));

        generateWordPreview(parsedData);
        generateTextPreview(parsedData);
        reportPreviewSection.style.display = 'block';
    } catch (error) {
        showStatus('Error parsing workbook: ' + error.message, 'error');
    }
}
