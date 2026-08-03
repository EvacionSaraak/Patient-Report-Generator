// Show status message
function showStatus(message, type) {
    statusDiv.textContent = message;
    statusDiv.className = 'status-message ' + type;
}

// Escape HTML to prevent XSS
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Filter out empty rows from data
function filterEmptyRows(rows) {
    return rows.filter(row => {
        return row && row.some(cell => cell != null && String(cell).trim() !== '');
    });
}

// Helper function to convert Excel serial date to readable format
function excelDateToJSDate(serial) {
    if (typeof serial === 'string' && isNaN(serial)) {
        return serial;
    }

    if (typeof serial === 'number' || !isNaN(serial)) {
        const utc_days = Math.floor(serial - 25569);
        const utc_value = utc_days * 86400;
        const date_info = new Date(utc_value * 1000);

        const fractional_day = serial - Math.floor(serial) + 0.0000001;
        let total_seconds = Math.floor(86400 * fractional_day);
        const seconds = total_seconds % 60;
        total_seconds -= seconds;
        const hours = Math.floor(total_seconds / (60 * 60));
        const minutes = Math.floor(total_seconds / 60) % 60;

        const date = new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);

        const day = date.getDate();
        const monthNames = ["January", "February", "March", "April", "May", "June",
                           "July", "August", "September", "October", "November", "December"];
        const month = monthNames[date.getMonth()];
        const year = date.getFullYear();

        return `${day} ${month} ${year}`;
    }

    return String(serial);
}

// Helper function to format date for header (e.g., "Jan 21")
function formatHeaderDate(dateValue) {
    if (!dateValue) return '';

    let date;

    if (typeof dateValue === 'number' || !isNaN(dateValue)) {
        const utc_days = Math.floor(dateValue - 25569);
        const utc_value = utc_days * 86400;
        date = new Date(utc_value * 1000);
    } else {
        date = new Date(dateValue);
    }

    if (isNaN(date.getTime())) {
        return String(dateValue);
    }

    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const month = monthNames[date.getMonth()];
    const day = date.getDate();

    return `${month} ${day}`;
}

// Helper function to get min and max dates from data
function getDateRange(data) {
    if (!data || data.length <= 1) return { min: '', max: '' };

    const headers = data[0] || [];
    const rows = data.slice(1);
    const visitDateIndex = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));

    if (visitDateIndex === -1) return { min: '', max: '' };

    const dates = rows
        .map(row => row[visitDateIndex])
        .filter(date => date !== undefined && date !== null && date !== '');

    if (dates.length === 0) return { min: '', max: '' };

    const comparableDates = dates.map(d => {
        if (typeof d === 'number') return d;
        const parsed = new Date(d);
        return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
    });

    const minValue = Math.min(...comparableDates);
    const maxValue = Math.max(...comparableDates);

    const minIndex = comparableDates.indexOf(minValue);
    const maxIndex = comparableDates.indexOf(maxValue);

    return {
        min: formatHeaderDate(dates[minIndex]),
        max: formatHeaderDate(dates[maxIndex])
    };
}

// Helper function to format any date value
function formatDate(dateValue) {
    if (!dateValue) return '';

    const str = String(dateValue);
    if (str.match(/\d{1,2}\s+\w+\s+\d{4}/)) {
        return str;
    }

    return excelDateToJSDate(dateValue);
}
