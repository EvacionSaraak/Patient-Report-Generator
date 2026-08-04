// ─── Label normalisation ──────────────────────────────────────────────────────
// Lowercase, trim, strip every character that is not a-z or 0-9.
function normalizePatientReportLabel(str) {
    return String(str ?? '')
        .toLowerCase()
        .trim()
        .replace(/[^a-z0-9]/g, '');
}

// ─── Alias tables ─────────────────────────────────────────────────────────────
const PR_DATE_ALIASES = new Set(['date', 'visitdate', 'encounterdate']);

const PR_FILE_NUMBER_ALIASES = new Set([
    'filenumber', 'fileno', 'fileid', 'ptid', 'ptno', 'patientid'
]);

const PR_PATIENT_NAME_ALIASES = new Set(['patientname', 'ptname', 'name']);

const PR_DOCTOR_NAME_ALIASES = new Set([
    'doctorname', 'doctor', 'clinicianname', 'clinician'
]);

const PR_REMINDERS_ALIASES = new Set([
    'personalreminders', 'personalreminder', 'reminders', 'reminder'
]);

// ─── Helpers ──────────────────────────────────────────────────────────────────

// Collect all w:t text nodes inside an element (namespace-safe)
function getPRCellText(el) {
    const nodes = el.getElementsByTagNameNS('*', 't');
    let text = '';
    for (let i = 0; i < nodes.length; i++) {
        text += nodes[i].textContent || '';
    }
    return text;
}

// Normalize whitespace; collapse NBSP, tabs, multiple spaces
function cleanPRText(str) {
    return String(str ?? '')
        .replace(/\u00a0/g, ' ')
        .replace(/[\t\r\n]+/g, ' ')
        .replace(/\s{2,}/g, ' ')
        .trim();
}

// ─── Main parser ──────────────────────────────────────────────────────────────

// Parse a Patient Report input DOCX ArrayBuffer into canonical patient records.
//
// Expects the DOCX to contain one two-column table per patient with labelled rows:
//   Date:            | <value>
//   File Number:     | <value>
//   Patient Name:    | <value>
//   Doctor Name:     | <value>
//   Personal Reminders: | <value>   (optional; may also be an unlabelled row)
//
// Unlike parseOPGInputDocx, records are allowed to have different visit dates.
//
// Returns:
//   { format: 'docx-patient-report', records: [{visitDate, fileNumber, patientName, doctor, personalReminders}] }
//
// Throws a descriptive Error on any failure.
async function parsePatientReportInputDocx(arrayBuffer) {
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
        const parseErr =
            xmlDoc.getElementsByTagNameNS(
                'http://www.mozilla.org/newlayout/xml/parseerror.xml', 'parseerror'
            )[0] || xmlDoc.querySelector('parseerror');
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
            'This DOCX does not match the expected Patient Report input format. ' +
            'No tables were found. Expected one two-column patient table per record ' +
            'with labels such as "File Number", "Patient Name", and "Date".'
        );
    }

    const records = [];

    for (let tblIdx = 0; tblIdx < tables.length; tblIdx++) {
        const tbl = tables[tblIdx];
        const rows = tbl.getElementsByTagNameNS('*', 'tr');
        if (rows.length === 0) continue;

        // Collect {left, right} for each row
        const parsedRows = [];
        for (let rIdx = 0; rIdx < rows.length; rIdx++) {
            const row = rows[rIdx];
            const cells = row.getElementsByTagNameNS('*', 'tc');
            if (cells.length === 0) continue;

            const leftText  = cleanPRText(getPRCellText(cells[0]));
            const rightText = cells.length > 1
                ? cleanPRText(getPRCellText(cells[1]))
                : '';

            parsedRows.push({ left: leftText, right: rightText });
        }

        if (parsedRows.length === 0) continue;

        // Classify rows by label
        const fields = {};
        let personalReminders = null;

        for (const { left, right } of parsedRows) {
            const key = normalizePatientReportLabel(left);

            if (PR_DATE_ALIASES.has(key)) {
                fields.visitDate = right;
            } else if (PR_FILE_NUMBER_ALIASES.has(key)) {
                fields.fileNumber = right;
            } else if (PR_PATIENT_NAME_ALIASES.has(key)) {
                fields.patientName = right;
            } else if (PR_DOCTOR_NAME_ALIASES.has(key)) {
                fields.doctor = right;
            } else if (PR_REMINDERS_ALIASES.has(key)) {
                personalReminders = right;
            } else if (personalReminders === null && !key) {
                // First unlabelled row → Personal Reminders (may be empty)
                personalReminders = right;
            }
        }

        // Require at least fileNumber and patientName; skip tables that lack both
        const fn = cleanPRText(fields.fileNumber || '');
        const pn = cleanPRText(fields.patientName || '');
        if (!fn || !pn) continue;

        records.push({
            visitDate:         cleanPRText(fields.visitDate || ''),
            fileNumber:        fn,
            patientName:       pn,
            doctor:            cleanPRText(fields.doctor || ''),
            personalReminders: cleanPRText(personalReminders || '')
        });
    }

    // ── 5. Validate results ───────────────────────────────────────────────────
    if (records.length === 0) {
        throw new Error(
            'This DOCX does not match the expected Patient Report input format. ' +
            'No patient records with a File Number and Patient Name were found. ' +
            'Expected one two-column labelled table per patient.'
        );
    }

    return {
        format: 'docx-patient-report',
        records
    };
}
