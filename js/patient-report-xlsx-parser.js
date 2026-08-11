// Strict XLSX parser for the Patient Report Generator.
// Requires SheetJS (window.XLSX) to be loaded before this script.
//
// Usage:
//   const result = await parsePatientReportXlsx(arrayBuffer);
//   // result: { format: 'xlsx-patient-report', records: [...], warnings: [...] }
//
// Throws a descriptive Error on any structural failure.

(function () {
    'use strict';

    // ── Required headers ──────────────────────────────────────────────────────
    // Columns A:H must match these exact strings in this exact order.
    const PR_XLSX_REQUIRED_HEADERS = [
        'PT ID.',
        'VISIT ID.',
        'Patient Name',
        'Visit Date',
        'Doctor',
        'Personal Reminders',
        'Query',
        'Status'
    ];

    // ── Month names for output formatting ─────────────────────────────────────
    const MONTH_NAMES = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];

    // ── Date helpers ──────────────────────────────────────────────────────────

    // Format year/month(0-based)/day into "DD Month YYYY"
    function formatDateParts(year, month, day) {
        const d = String(day).padStart(2, '0');
        return `${d} ${MONTH_NAMES[month]} ${year}`;
    }

    // Convert an Excel serial number to "DD Month YYYY".
    // Uses UTC arithmetic to avoid timezone shifts.
    function excelSerialToString(serial) {
        const utcDays = Math.floor(serial - 25569);
        const utcMs = utcDays * 86400 * 1000;
        const dt = new Date(utcMs);
        return formatDateParts(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate());
    }

    // Convert any cell value that represents a date into "DD Month YYYY".
    // Returns '' for null/empty values.
    // Throws nothing — invalid dates return '' (caller decides to warn/skip).
    function convertVisitDate(value) {
        if (value == null) return '';

        // JS Date object (SheetJS cellDates:true path, unlikely here but defensive)
        if (value instanceof Date) {
            if (isNaN(value.getTime())) return '';
            return formatDateParts(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate());
        }

        // Excel serial number
        if (typeof value === 'number') {
            if (!isFinite(value) || value < 1) return '';
            return excelSerialToString(value);
        }

        // String value
        if (typeof value === 'string') {
            const trimmed = value.trim();
            if (!trimmed) return '';
            // Already "DD Month YYYY" or similar human-readable form
            if (/^\d{1,2}\s+\w+\s+\d{4}$/.test(trimmed)) return trimmed;
            // Attempt ISO / other formats
            const parsed = new Date(trimmed);
            if (!isNaN(parsed.getTime())) {
                return formatDateParts(
                    parsed.getUTCFullYear(),
                    parsed.getUTCMonth(),
                    parsed.getUTCDate()
                );
            }
            return ''; // unparseable
        }

        return '';
    }

    // ── Main parser ───────────────────────────────────────────────────────────

    // Parses an XLSX ArrayBuffer with strict template validation.
    //
    // Returns:
    //   {
    //     format: 'xlsx-patient-report',
    //     records: [
    //       {
    //         fileNumber, visitId, patientName, visitDate,
    //         doctor, personalReminders, query, status
    //       }, …
    //     ],
    //     warnings: [ 'Row N was skipped: missing …', … ]
    //   }
    //
    // Throws a descriptive Error when:
    //   • the buffer is not a valid XLSX workbook;
    //   • no worksheet has the required headers in the required order;
    //   • no valid data rows exist after skipping invalid rows.
    function parsePatientReportXlsx(arrayBuffer) {
        if (typeof XLSX === 'undefined') {
            throw new Error('SheetJS (XLSX) library is not loaded. Refresh the page and try again.');
        }

        // ── 1. Parse workbook ─────────────────────────────────────────────────
        let wb;
        try {
            wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array', cellDates: false });
        } catch (e) {
            throw new Error(
                'The file is not a valid XLSX workbook. ' +
                (e && e.message ? e.message : String(e))
            );
        }

        if (!wb.SheetNames || wb.SheetNames.length === 0) {
            throw new Error('Invalid Patient Report template. The workbook contains no worksheets.');
        }

        // ── 2. Find first worksheet whose row 1 has the exact required headers ─
        let rawRows = null;

        for (const sheetName of wb.SheetNames) {
            const ws = wb.Sheets[sheetName];
            if (!ws) continue;

            const rows = XLSX.utils.sheet_to_json(ws, {
                header: 1,
                raw: true,
                defval: null,
                blankrows: true
            });

            if (!rows || rows.length === 0) continue;
            const headerRow = rows[0] || [];

            // Must have at least 8 columns
            if (headerRow.length < PR_XLSX_REQUIRED_HEADERS.length) continue;

            // Check exact match for columns A:H (indices 0-7)
            let allMatch = true;
            for (let col = 0; col < PR_XLSX_REQUIRED_HEADERS.length; col++) {
                const cellVal = headerRow[col] == null ? '' : String(headerRow[col]).trim();
                if (cellVal !== PR_XLSX_REQUIRED_HEADERS[col]) {
                    allMatch = false;
                    break;
                }
            }

            if (allMatch) {
                rawRows = rows;
                break;
            }
        }

        if (!rawRows) {
            throw new Error(
                'Invalid Patient Report template. Expected columns A:H to be: ' +
                PR_XLSX_REQUIRED_HEADERS.join(', ') + '.'
            );
        }

        // ── 3. Extract and validate records ───────────────────────────────────
        const warnings = [];
        const records = [];

        for (let i = 1; i < rawRows.length; i++) {
            const row = rawRows[i];
            const excelRowNum = i + 1; // 1-based Excel row number

            // Silently skip completely empty rows
            const hasAnyContent = row && row.some(
                cell => cell != null && String(cell).trim() !== ''
            );
            if (!hasAnyContent) continue;

            const getString = (idx) => {
                const v = row && row[idx] != null ? row[idx] : null;
                if (v == null) return '';
                return String(v).trim();
            };

            const fileNumber        = getString(0);
            const visitId           = getString(1);
            const patientName       = getString(2);
            const visitDateRaw      = row ? row[3] : null;
            const doctor            = getString(4);
            const personalReminders = getString(5);
            const query             = getString(6);
            const status            = getString(7);

            // Convert visit date
            const visitDate = convertVisitDate(visitDateRaw);
            const visitDateOk = visitDate !== '';

            // Validate required fields
            const missing = [];
            if (!fileNumber)  missing.push('PT ID.');
            if (!visitId)     missing.push('VISIT ID.');
            if (!patientName) missing.push('Patient Name');
            if (!visitDateOk) missing.push('Visit Date');
            if (!doctor)      missing.push('Doctor');

            if (missing.length > 0) {
                warnings.push(
                    `Row ${excelRowNum} was skipped: missing ${missing.join(' and ')}.`
                );
                continue;
            }

            records.push({
                fileNumber,
                visitId,
                patientName,
                visitDate,
                visitDateRaw,
                doctor,
                personalReminders,
                query,
                status
            });
        }

        if (records.length === 0) {
            throw new Error(
                'Invalid Patient Report template. No valid data rows exist after validation.'
            );
        }

        return {
            format: 'xlsx-patient-report',
            records,
            warnings
        };
    }

    // Expose globally
    window.parsePatientReportXlsx = parsePatientReportXlsx;

}());
