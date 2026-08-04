// ─── OPG output format constants (derived from example output DOCX) ───────────
// Table width and column widths in dxa (twips)
const OPG_TABLE_W    = 9913;               // total table width
const OPG_COL_W      = [2117, 5386, 2410]; // File #, Pt. Name, Dr.
const OPG_ROW_H      = 92;                 // info-row height (auto minimum)
const OPG_IMG_H      = 4891;               // image-placeholder height (exact)
const OPG_BORDER_SZ  = 12;                 // border thickness (1/8 pt units)
const OPG_PAGE_W     = 12240;              // US Letter width  (twips)
const OPG_PAGE_H     = 15840;              // US Letter height (twips)
const OPG_MARGIN     = 1440;               // all page margins (twips)
const OPG_HEADER_FTR = 708;               // header/footer distance (twips)

// ─── Lazy-load docx ESM module ────────────────────────────────────────────────
function loadOPGDocxModule() {
    if (!opgDocxModulePromise) {
        opgDocxModulePromise = import('https://cdn.jsdelivr.net/npm/docx@8.2.2/+esm');
    }
    return opgDocxModulePromise;
}

// ─── Strip leading "Dr." so we never emit "Dr. Dr. …" ─────────────────────────
function opgNormaliseDoctorName(name) {
    return String(name || '').trim().replace(/^dr\.?\s*/i, '');
}

// ─── Generate and download the OPG Word report ───────────────────────────────
async function downloadOPGReport() {
    const button = document.getElementById('exampleDownloadBtn');

    try {
        if (!opgRecords || !opgRecords.length) {
            throw new Error('Upload a patient-report DOCX with at least one complete record first.');
        }

        button.disabled = true;
        showExampleStatus('Generating OPG Report…', 'info');

        const {
            Document, Packer, Paragraph, TextRun,
            Table, TableRow, TableCell,
            WidthType, TableLayoutType, BorderStyle,
            HeightRule, VerticalAlign
        } = await loadOPGDocxModule();

        // ── Border definition (single, sz=12, black) ──────────────────────────
        const makeBorder = () => ({
            style: BorderStyle.SINGLE,
            size:  OPG_BORDER_SZ,
            color: '000000'
        });
        const allBorders = {
            top:             makeBorder(),
            bottom:          makeBorder(),
            left:            makeBorder(),
            right:           makeBorder(),
            insideHorizontal: makeBorder(),
            insideVertical:   makeBorder()
        };

        // ── Text run helper (bold Arial 10 pt = size 20 half-points) ──────────
        const run = text => new TextRun({
            text:  String(text || ''),
            font:  'Arial',
            size:  20,
            bold:  true,
            color: '000000'
        });

        // ── Build the three-column patient info table ─────────────────────────
        const makeCell = (text, colIdx) => new TableCell({
            width:          { size: OPG_COL_W[colIdx], type: WidthType.DXA },
            verticalAlign:  VerticalAlign.CENTER,
            margins:        { top: 35, bottom: 35, left: 75, right: 75 },
            children:       [new Paragraph({
                spacing:    { before: 0, after: 0, line: 240 },
                children:   [run(text)]
            })]
        });

        const makeInfoTable = record => new Table({
            width:    { size: OPG_TABLE_W, type: WidthType.DXA },
            indent:   { size: -5,          type: WidthType.DXA },
            layout:   TableLayoutType.FIXED,
            borders:  allBorders,
            rows:     [new TableRow({
                cantSplit: true,
                height:    { value: OPG_ROW_H, rule: HeightRule.AUTO },
                children:  [
                    makeCell(`File #  ${record.fileNumber}`, 0),
                    makeCell(`Pt. Name - ${record.patientName}`, 1),
                    makeCell(`Dr. ${opgNormaliseDoctorName(record.doctorName)}`, 2)
                ]
            })]
        });

        // ── Build the borderless personal-reminders paragraph ─────────────────
        const makeReminderPara = record => {
            const reminder = (record.personalReminders || '').trim();
            return new Paragraph({
                spacing: { before: 0, after: 0, line: 240 },
                children: reminder ? [run(reminder)] : []
            });
        };

        // ── Build the empty X-ray image placeholder table ─────────────────────
        const makeImageTable = () => new Table({
            width:   { size: OPG_TABLE_W, type: WidthType.DXA },
            indent:  { size: -5,          type: WidthType.DXA },
            layout:  TableLayoutType.FIXED,
            borders: allBorders,
            rows:    [new TableRow({
                cantSplit: true,
                height:    { value: OPG_IMG_H, rule: HeightRule.EXACT },
                children:  [new TableCell({
                    width:    { size: OPG_TABLE_W, type: WidthType.DXA },
                    children: [new Paragraph({ children: [] })]
                })]
            })]
        });

        // ── Build blank spacing paragraph ─────────────────────────────────────
        const blankPara = () => new Paragraph({
            spacing: { before: 0, after: 0, line: 240 },
            children: []
        });

        // ── Assemble document body ────────────────────────────────────────────
        // Pattern per record: infoTable → blankPara → imageTable → blankPara → blankPara
        // (matches the structure of REPORT FOR OPG JULY 08, 2026.docx)
        const children = [];

        opgRecords.forEach(record => {
            children.push(makeInfoTable(record), makeReminderPara(record), makeImageTable(), blankPara(), blankPara());
        });

        const document = new Document({
            sections: [{
                properties: {
                    page: {
                        size:   { width: OPG_PAGE_W, height: OPG_PAGE_H },
                        margin: {
                            top:    OPG_MARGIN,
                            right:  OPG_MARGIN,
                            bottom: OPG_MARGIN,
                            left:   OPG_MARGIN,
                            header: OPG_HEADER_FTR,
                            footer: OPG_HEADER_FTR,
                            gutter: 0
                        }
                    }
                },
                children
            }]
        });

        // ── Derive filename from report date ──────────────────────────────────
        // Use the date from the first record (all records share the same date
        // after validation in parsePatientReportDocx).
        const reportDate = (opgRecords[0] && opgRecords[0].date) ? opgRecords[0].date : 'OPG';
        const filename   = `REPORT FOR OPG - ${reportDate}`
            .replace(/[\\/:*?"<>|]/g, '-')
            .concat('.docx');

        saveAs(await Packer.toBlob(document), filename);

        showExampleStatus('OPG Report downloaded.', 'success');
    } catch (err) {
        console.error('OPG report generation error:', err);
        showExampleStatus(`Error generating OPG Report: ${err.message}`, 'error');
    } finally {
        button.disabled = !opgRecords || !opgRecords.length;
    }
}
