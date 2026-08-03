// Normalize whitespace and trim a cell value to a plain string
function cleanText(value) {
    return String(value ?? '')
        .replace(/\u00a0/g, ' ')
        .replace(/[\t\r\n]+/g, ' ')
        .replace(/\s{2,}/g, ' ')
        .trim();
}

// Escape HTML characters to prevent XSS in OPG output
function opgEscapeHtml(value) {
    return String(value ?? '')
        .replace(/[&<>'"]/g, char => ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            "'": '&#39;',
            '"': '&quot;'
        }[char]));
}

// Normalise a header string to lowercase trimmed text for comparison
function normalizeHeader(value) {
    return cleanText(value).toLowerCase();
}

// Return the cell value or empty string if null/undefined
function getCellValue(value) {
    return (value === undefined || value === null) ? '' : value;
}

// Return true when a cell has a meaningful (non-empty) value
function hasCellValue(value) {
    return (
        value !== undefined &&
        value !== null &&
        value !== false &&
        cleanText(value) !== ''
    );
}

// Display a status message in the OPG status element
function showExampleStatus(message, type) {
    const status = document.getElementById('exampleStatus');
    status.textContent = message;
    status.className = `status-message mt-2 ${type}`;
}
