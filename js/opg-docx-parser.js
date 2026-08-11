// ─── Label normalisation ──────────────────────────────────────────────────────
// Lowercase, trim, then strip every character that is not a-z or 0-9.
// This handles punctuation (colons, dots), spaces, underscores, hyphens, etc.
// Examples:
//   "Date:"           → "date"
//   "File Number:"    → "filenumber"
//   "Patient name:"   → "patientname"
//   "Doctor Name:"    → "doctorname"
function normalizeLabelKey(str) {
    return String(str ?? '')
        .toLowerCase()
        .trim()
        .replace(/[^a-z0-9]/g, '');
}

// ─── Alias tables ─────────────────────────────────────────────────────────────
const DOCX_DATE_ALIASES = new Set(['date', 'visitdate']);

const DOCX_FILE_NUMBER_ALIASES = new Set([
    'filenumber', 'fileno', 'fileid', 'ptid', 'ptno', 'patientid'
]);

const DOCX_PATIENT_NAME_ALIASES = new Set(['patientname', 'ptname', 'name']);

const DOCX_DOCTOR_NAME_ALIASES = new Set([
    'doctorname', 'doctor', 'clinicianname', 'clinician'
]);

// ─── Helpers ──────────────────────────────────────────────────────────────────

// Collect all w:t text nodes inside an element (namespace-safe)
function getWmlCellText(el) {
    const nodes = el.getElementsByTagNameNS('*', 't');
    let text = '';
    for (let i = 0; i < nodes.length; i++) {
        text += nodes[i].textContent || '';
    }
    return text;
}

// Normalize whitespace; collapse NBSP, tabs, multiple spaces
function cleanDocxText(str) {
    return String(str ?? '')
        .replace(/\u00a0/g, ' ')
        .replace(/[\t\r\n]+/g, ' ')
        .replace(/\s{2,}/g, ' ')
        .trim();
}

// ─── Main parser ──────────────────────────────────────────────────────────────

// Parse a patient-report DOCX ArrayBuffer into canonical patient records.
// Returns:
//   { sourceFormat: "docx-normal-report", date: string, records: [...] }
// Throws a descriptive Error on any failure.
async function parseOPGInputDocx(arrayBuffer) {
    // ── 1. Unzip ───────────────────────────────────────────────────────────────
    if (typeof JSZip === 'undefined') {
        throw new Error('JSZip library is not loaded. Refresh the page and try again.');
    }

    let zip;
    try {
        zip = await JSZip.loadAsync(arrayBuffer);
    } catch (err) {
        throw new Error(
            'The file is not a valid DOCX (ZIP archive). ' +
            (err && err.message ? err.message : String(err))
        );
    }

    // ── 2. Extract word/document.xml ──────────────────────────────────────────
    const xmlEntry = zip.file('word/document.xml');
    if (!xmlEntry) {
        throw new Error(
            'word/document.xml is missing from the DOCX archive. ' +
            'The file may not be a standard .docx document.'
        );
    }

    let xmlString;
    try {
        xmlString = await xmlEntry.async('string');
    } catch (err) {
        throw new Error(
            'Failed to read word/document.xml: ' +
            (err && err.message ? err.message : String(err))
        );
    }

    // ── 3. Parse XML ──────────────────────────────────────────────────────────
    let xmlDoc;
    try {
        const domParser = new DOMParser();
        xmlDoc = domParser.parseFromString(xmlString, 'application/xml');

        // DOMParser sets a <parsererror> element on failure
        const parseErr = xmlDoc.getElementsByTagNameNS(
            'http://www.mozilla.org/newlayout/xml/parsererror.xml', 'parseerror'
        )[0] || xmlDoc.querySelector('parsererror');
        if (parseErr) {
            throw new Error(parseErr.textContent || 'Unknown XML error');
        }
    } catch (err) {
        throw new Error(
            'XML parsing failed: ' +
            (err && err.message ? err.message : String(err))
        );
    }

    // ── 4. Read tables in document order ─────────────────────────────────────
    const tables = xmlDoc.getElementsByTagNameNS('*', 'tbl');
    if (tables.length === 0) {
        throw new Error(
            'This DOCX does not contain the patient tables required by the OPG generator. ' +
            'Expected one two-column patient table per record with labels such as ' +
            '"File Number", "Patient Name", and "Date".'
        );
    }

    const records = [];

    for (let tblIdx = 0; tblIdx < tables.length; tblIdx++) {
        const tbl = tables[tblIdx];
        const rows = tbl.getElementsByTagNameNS('*', 'tr');
        if (rows.length === 0) continue; // skip fully empty tables

        // Collect {left, right} for each row
        const parsedRows = [];
        for (let rIdx = 0; rIdx < rows.length; rIdx++) {
            const row = rows[rIdx];
            const cells = row.getElementsByTagNameNS('*', 'tc');
            if (cells.length === 0) continue;

            const leftText  = cleanDocxText(getWmlCellText(cells[0]));
            const rightText = cells.length > 1
                ? cleanDocxText(getWmlCellText(cells[1]))
                : '';

            parsedRows.push({ left: leftText, right: rightText });
        }

        if (parsedRows.length === 0) continue;

        // Classify rows
        const fields = {};
        let personalReminders = null; // set on first unlabelled row

        for (const { left, right } of parsedRows) {
            const key = normalizeLabelKey(left);

            if (DOCX_DATE_ALIASES.has(key)) {
                fields.date = right;
            } else if (DOCX_FILE_NUMBER_ALIASES.has(key)) {
                fields.fileNumber = right;
            } else if (DOCX_PATIENT_NAME_ALIASES.has(key)) {
                fields.patientName = right;
            } else if (DOCX_DOCTOR_NAME_ALIASES.has(key)) {
                fields.doctorName = right;
            } else if (personalReminders === null && !key) {
                // First row whose left cell has no recognised label
                // → treat right cell as Personal Reminders (may be empty)
                personalReminders = right;
            }
        }

        // Require both fileNumber and patientName; skip otherwise
        const fn = cleanDocxText(fields.fileNumber || '');
        const pn = cleanDocxText(fields.patientName || '');
        if (!fn || !pn) continue;

        records.push({
            date:               cleanDocxText(fields.date || ''),
            fileNumber:         fn,
            patientName:        pn,
            doctorName:         cleanDocxText(fields.doctorName || ''),
            personalReminders:  cleanDocxText(personalReminders || '')
        });
    }

    // ── 5. Validate results ───────────────────────────────────────────────────
    if (records.length === 0) {
        throw new Error(
            'This DOCX does not contain the patient tables required by the OPG generator. ' +
            'No records with a File Number and Patient Name were found.'
        );
    }

    // ── 6. Validate dates ─────────────────────────────────────────────────────
    const uniqueDates = [...new Set(records.map(r => r.date).filter(Boolean))];

    if (uniqueDates.length > 1) {
        throw new Error(
            'Multiple different dates were found in the input: ' +
            uniqueDates.join(', ') +
            '. All patient records must share the same date to generate an OPG report. ' +
            'Please correct the input document and try again.'
        );
    }

    return {
        sourceFormat: 'docx-normal-report',
        date:         uniqueDates[0] || '',
        records
    };
}
