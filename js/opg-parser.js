// Parse the full OPG workbook and return accepted/invalid row sets
function parseOPGWorkbook(workbook) {
    const found = findOPGHeaderRow(workbook);

    if (!found) {
        throw new Error(
            `No worksheet contains all required headers: ${OPG_REQUIRED_HEADERS.join(', ')}.`
        );
    }

    const { rows, headerRow, headerMap } = found;
    const accepted = [];
    const invalid = [];

    rows.slice(headerRow + 1).forEach((row, index) => {
        const rowNumber = headerRow + index + 2;

        const values = Object.fromEntries(
            OPG_REQUIRED_HEADERS.map(header => [
                header,
                getCellValue(row[headerMap[header]])
            ])
        );

        if (OPG_REQUIRED_HEADERS.every(header => !hasCellValue(values[header]))) {
            return;
        }

        if (isOPGGroupRow(values)) {
            return;
        }

        const missing = OPG_REQUIRED_HEADERS.filter(header => !hasCellValue(values[header]));

        let parsedPatient = { fileNumber: '', patientName: '' };
        let doctor = '';
        let date = '';

        if (!missing.includes('Patient')) {
            parsedPatient = parsePatientCell(values.Patient);

            if (!parsedPatient.fileNumber) {
                missing.push('Patient (File Number)');
            }

            if (!parsedPatient.patientName) {
                missing.push('Patient (Name)');
            }
        }

        if (!missing.includes('Performing Clinician')) {
            doctor = parseClinicianCell(values['Performing Clinician']);

            if (!doctor) {
                missing.push('Performing Clinician (Name)');
            }
        }

        if (!missing.includes('Date')) {
            date = formatOPGDate(values.Date);

            if (!date) {
                missing.push('Date (Invalid)');
            }
        }

        const claimId = cleanText(values['Claim ID']);

        if (missing.length) {
            invalid.push({
                claimId,
                rowNumber,
                missing: [...new Set(missing)]
            });
            return;
        }

        accepted.push({
            claimId,
            fileNumber: parsedPatient.fileNumber,
            patientName: parsedPatient.patientName,
            doctor,
            date,
            lastModifiedBy: cleanText(values['Last Modified By']),
            sourceRow: rowNumber
        });
    });

    return { accepted, invalid };
}

// Scan all sheets for a row containing all required OPG headers
function findOPGHeaderRow(workbook) {
    let best = null;

    for (const sheetName of workbook.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
            header: 1,
            raw: true,
            defval: ''
        });

        for (let i = 0; i < Math.min(rows.length, 100); i++) {
            const normalized = rows[i].map(normalizeHeader);
            const headerMap = {};
            const matched = [];

            for (const header of OPG_REQUIRED_HEADERS) {
                const index = normalized.indexOf(normalizeHeader(header));

                if (index !== -1) {
                    headerMap[header] = index;
                    matched.push(header);
                }
            }

            if (!best || matched.length > best.matched.length) {
                best = { rows, headerRow: i, headerMap, matched };
            }

            if (matched.length === OPG_REQUIRED_HEADERS.length) {
                return { rows, headerRow: i, headerMap, sheetName };
            }
        }
    }

    if (best && best.matched.length) {
        const missing = OPG_REQUIRED_HEADERS.filter(
            header => !best.matched.includes(header)
        );

        throw new Error(
            `Missing required header${missing.length === 1 ? '' : 's'}: ${missing.join(', ')}.`
        );
    }

    return null;
}

// Detect rows that are date-group headings rather than claim records
function isOPGGroupRow(values) {
    const filled = OPG_REQUIRED_HEADERS.filter(header => hasCellValue(values[header]));

    if (filled.length !== 1 || filled[0] !== 'Claim ID') {
        return false;
    }

    return /^\s*\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\s*\(\d+\)\s*$/.test(
        cleanText(values['Claim ID'])
    );
}

// Strip suffix tags and the "Dr." prefix from a clinician cell value
function parseClinicianCell(value) {
    return cleanText(value)
        .replace(/\s*\[[^\]]+\]\s*$/, '')
        .replace(/^dr\.?\s*/i, '')
        .trim();
}

// Extract the file number and patient name from a combined patient cell value
function parsePatientCell(value) {
    const raw = cleanText(value);
    const match = raw.match(/\[([^\]]+)\]/);

    const fileNumber = match ? cleanText(match[1]) : '';

    let patientName = match
        ? raw.slice((match.index || 0) + match[0].length).trim()
        : raw;

    patientName = patientName
        .replace(/\s*\(\s*\d+\s*[YMD]\s*\/\s*[MF]\s*\)\s*$/i, '')
        .replace(/^(?:(?:Mr|Mrs|Miss|Ms|Mstr|Master|Baby|Dr)\.?\s+)+/i, '')
        .trim();

    return { fileNumber, patientName };
}

// Convert a date cell value (Date object, serial number, or string) to "MONTH DD, YYYY"
function formatOPGDate(value) {
    let date = null;

    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        date = value;
    } else if (typeof value === 'number' && typeof XLSX !== 'undefined' && XLSX.SSF) {
        const parsed = XLSX.SSF.parse_date_code(value);

        if (parsed) {
            date = new Date(parsed.y, parsed.m - 1, parsed.d);
        }
    } else {
        const text = cleanText(value);

        if (!text) {
            return '';
        }

        const parsed = new Date(text);

        if (!Number.isNaN(parsed.getTime())) {
            date = parsed;
        } else {
            return text.toUpperCase();
        }
    }

    return date
        ? `${date.toLocaleString('en-US', { month: 'long' }).toUpperCase()} ${String(date.getDate()).padStart(2, '0')}, ${date.getFullYear()}`
        : '';
}
