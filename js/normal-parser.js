// ─── Header normalisation ─────────────────────────────────────────────────────
// Lower-case, trim, then strip every character that is not a-z or 0-9.
// This handles punctuation (dots, commas), spaces, underscores, hyphens, etc.
// Examples:
//   "PT ID."          → "ptid"
//   "VISIT ID."       → "visitid"
//   "Patient Name"    → "patientname"
//   "Visit Date"      → "visitdate"
//   "Personal Reminders" → "personalreminders"
//   "File No."        → "fileno"
//   "Patient_ID"      → "patientid"
function normalizeHeaderKey(str) {
    return String(str ?? '')
        .toLowerCase()
        .trim()
        .replace(/[^a-z0-9]/g, '');
}

// ─── Alias table ─────────────────────────────────────────────────────────────
// Each key is a canonical field name; each value is the ordered list of
// normalised aliases that map to it.  Earlier aliases take priority.
const NORMAL_FORMAT_ALIASES = {
    fileNumber: [
        'ptid',
        'ptno',
        'patientid',
        'fileno',
        'filenumber',
        'fileid',
    ],
    visitId: [
        'visitid',
        'visitno',
        'visitnumber',
    ],
    patientName: [
        'patientname',
        'name',
    ],
    visitDate: [
        'visitdate',
        'encounterdate',
        'date',
    ],
    doctor: [
        'doctor',
        'doctorname',
        'clinician',
        'clinicianname',
    ],
    personalReminders: [
        'personalreminders',
        'personalreminder',
        'reminders',
        'reminder',
    ],
};

// ─── Column resolution ────────────────────────────────────────────────────────
// Given a raw header row array, returns an object mapping each canonical field
// name to its column index (-1 when no alias matched).
function resolveColumns(headerRow) {
    const normalized = (headerRow || []).map(normalizeHeaderKey);
    const result = {};
    for (const [field, aliases] of Object.entries(NORMAL_FORMAT_ALIASES)) {
        let idx = -1;
        for (const alias of aliases) {
            const found = normalized.indexOf(alias);
            if (found !== -1) { idx = found; break; }
        }
        result[field] = idx;
    }
    return result;
}

// ─── Format detection ─────────────────────────────────────────────────────────
// A worksheet qualifies as normal format when the resolved column map contains
// all three required fields.  Doctor and Personal Reminders are optional.
function isNormalFormat(columns) {
    return (
        columns.fileNumber  !== -1 &&
        columns.patientName !== -1 &&
        columns.visitDate   !== -1
    );
}

// ─── Single-row extraction ────────────────────────────────────────────────────
function extractNormalRecord(row, columns) {
    const get = (idx) => (idx !== -1 && row[idx] != null) ? row[idx] : null;

    return {
        fileNumber:        get(columns.fileNumber)        != null ? String(get(columns.fileNumber)).trim()        : '',
        visitId:           get(columns.visitId)           != null ? String(get(columns.visitId)).trim()           : '',
        patientName:       get(columns.patientName)       != null ? String(get(columns.patientName)).trim()       : '',
        visitDate:         get(columns.visitDate),          // keep raw; may be an Excel serial number
        doctor:            get(columns.doctor)            != null ? String(get(columns.doctor)).trim()            : '',
        personalReminders: get(columns.personalReminders) != null ? String(get(columns.personalReminders)).trim() : '',
    };
}

// ─── Main sheet parser ────────────────────────────────────────────────────────
// Accepts the full raw 2-D array from XLSX.utils.sheet_to_json({ header: 1 }).
//
// Returns a canonical result object:
//   {
//     format: 'normal',
//     headerRowIndex: <number>,
//     columns: { fileNumber, visitId, patientName, visitDate, doctor, personalReminders },
//     records: [ { fileNumber, visitId, patientName, visitDate, doctor, personalReminders }, … ],
//     rawRows: <original 2-D array>
//   }
//
// Returns null when no normal-format header row is found (caller may try another
// format detector).
//
// Throws a descriptive Error when a header row is found but a required column is
// missing (indicating a malformed normal-format file rather than a different format).
function parseNormalSheet(rawRows) {
    if (!rawRows || rawRows.length === 0) return null;

    // Search the first 20 non-empty rows for a header row.
    let headerRowIndex = -1;
    let columns = null;
    let nonEmptyCount = 0;

    for (let i = 0; i < rawRows.length; i++) {
        const row = rawRows[i];
        const hasContent = row && row.some(
            cell => cell != null && String(cell).trim() !== ''
        );
        if (!hasContent) continue;

        nonEmptyCount++;
        if (nonEmptyCount > 20) break;

        const candidate = resolveColumns(row);
        if (isNormalFormat(candidate)) {
            headerRowIndex = i;
            columns = candidate;
            break;
        }
    }

    if (headerRowIndex === -1 || columns === null) {
        // Not a normal-format sheet; caller should try a different detector.
        return null;
    }

    // Belt-and-suspenders: validate required fields (already guaranteed by
    // isNormalFormat, but produce specific error messages if something is off).
    if (columns.fileNumber === -1) {
        throw new Error(
            'Normal report format was detected, but the required PT ID/File Number column could not be found.'
        );
    }
    if (columns.patientName === -1) {
        throw new Error(
            'Normal report format was detected, but the required Patient Name column could not be found.'
        );
    }
    if (columns.visitDate === -1) {
        throw new Error(
            'Normal report format was detected, but the required Visit Date column could not be found.'
        );
    }

    // Extract records; skip fully empty rows.
    const records = [];
    for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
        const row = rawRows[i];
        if (!row || !row.some(cell => cell != null && String(cell).trim() !== '')) {
            continue;
        }
        records.push(extractNormalRecord(row, columns));
    }

    return {
        format: 'normal',
        headerRowIndex,
        columns,
        records,
        rawRows,
    };
}

// ─── Inline self-test ─────────────────────────────────────────────────────────
// Run from the browser console: runNormalParserTest()
// Returns true when all assertions pass.
function runNormalParserTest() {
    const matrix = [
        ['PT ID.', 'VISIT ID.', 'Patient Name', 'Visit Date', 'Doctor', 'Personal Reminders'],
        ['TVIP00879410', 'IVMCV260822812', 'TEST PATIENT', 46235, 'Dr. Test', 'E OPG'],
    ];

    let allPassed = true;

    function assert(label, got, expected) {
        const pass = got === expected;
        console[pass ? 'log' : 'error'](
            `[NormalParser] ${pass ? 'PASS' : 'FAIL'} ${label}: got "${got}", expected "${expected}"`
        );
        if (!pass) allPassed = false;
    }

    try {
        const result = parseNormalSheet(matrix);

        if (!result || result.format !== 'normal') {
            console.error('[NormalParser] FAIL: parseNormalSheet returned null or wrong format');
            return false;
        }

        const record = result.records[0];

        assert('fileNumber',  record.fileNumber,  'TVIP00879410');
        assert('visitId',     record.visitId,     'IVMCV260822812');
        assert('patientName', record.patientName, 'TEST PATIENT');
        assert('visitDate (raw serial)', record.visitDate, 46235);

        // Variant spellings
        const variants = [
            ['PT ID',    'VISIT ID',   'Patient Name', 'Visit Date'],
            ['Pt Id.',   'Visit Id.',  'Patient Name', 'Visit Date'],
            ['File No.', 'Visit No.',  'Patient Name', 'Visit Date'],
            ['Patient_ID', 'Visit_ID', 'Patient Name', 'Visit Date'],
        ];
        variants.forEach((hdrs, vi) => {
            const m2 = [hdrs, ['F123', 'V456', 'Test Patient', 46235]];
            const r2 = parseNormalSheet(m2);
            const ok = r2 && r2.records[0].fileNumber === 'F123';
            console[ok ? 'log' : 'error'](
                `[NormalParser] ${ok ? 'PASS' : 'FAIL'} variant[${vi}] (${hdrs[0]})`
            );
            if (!ok) allPassed = false;
        });

        // Missing required header → returns null (not normal format)
        const noFileNo = [
            ['VISIT ID.', 'Patient Name', 'Visit Date'],
            ['V1', 'Test', 46235],
        ];
        const noFileNoResult = parseNormalSheet(noFileNo);
        const nullOk = noFileNoResult === null;
        console[nullOk ? 'log' : 'error'](
            `[NormalParser] ${nullOk ? 'PASS' : 'FAIL'} missing fileNumber → null`
        );
        if (!nullOk) allPassed = false;

    } catch (e) {
        console.error('[NormalParser] FAIL: unexpected exception:', e.message);
        allPassed = false;
    }

    console.log(`[NormalParser] Self-test ${allPassed ? 'PASSED ✓' : 'FAILED ✗'}`);
    return allPassed;
}
