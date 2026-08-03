// Lazily load the docx ESM module (cached after first load)
function loadOPGDocxModule() {
    if (!opgDocxModulePromise) {
        opgDocxModulePromise = import('https://cdn.jsdelivr.net/npm/docx@8.2.2/+esm');
    }

    return opgDocxModulePromise;
}

// Generate and trigger download of the OPG Word report
async function downloadOPGReport() {
    const button = document.getElementById('exampleDownloadBtn');

    try {
        if (!opgRows.length) {
            throw new Error('Upload a claim XLSX/XLS file with at least one complete row first.');
        }

        button.disabled = true;
        showExampleStatus('Generating OPG Report...', 'info');

        const {
            Document, Packer, Paragraph, TextRun,
            Table, TableRow, TableCell, PageBreak,
            WidthType, TableLayoutType, BorderStyle,
            HeightRule, VerticalAlign
        } = await loadOPGDocxModule();

        const border = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
        const borders = {
            top: border, bottom: border, left: border, right: border,
            insideHorizontal: border, insideVertical: border
        };

        const run = text => new TextRun({
            text: String(text || ''),
            font: 'Arial',
            size: 20,
            bold: true,
            color: '000000'
        });

        const infoCell = (label, value, width) => new TableCell({
            width: { size: width, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 35, bottom: 35, left: 75, right: 75 },
            children: [new Paragraph({
                spacing: { before: 0, after: 0, line: 240 },
                children: [run(label), run(value)]
            })]
        });

        const patientTable = row => new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            layout: TableLayoutType.FIXED,
            borders,
            rows: [new TableRow({
                cantSplit: true,
                children: [
                    infoCell('File #  ', row.fileNumber, 22),
                    infoCell('Pt. Name -  ', row.patientName, 55),
                    infoCell('Dr. ', row.doctor, 23)
                ]
            })]
        });

        const blankReminder = () => new Paragraph({
            keepNext: true,
            spacing: { before: 0, after: 0, line: 240 },
            children: [run(' ')]
        });

        const emptyOPGTable = () => new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            layout: TableLayoutType.FIXED,
            borders,
            rows: [new TableRow({
                cantSplit: true,
                height: { value: 6000, rule: HeightRule.EXACT },
                children: [new TableCell({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    children: [new Paragraph({ children: [] })]
                })]
            })]
        });

        const children = [];

        opgRows.forEach((row, index) => {
            if (index > 0 && index % 2 === 0) {
                children.push(new Paragraph({ children: [new PageBreak()] }));
            }

            children.push(patientTable(row), blankReminder(), emptyOPGTable());

            if (index % 2 === 0 && index < opgRows.length - 1) {
                children.push(new Paragraph({
                    spacing: { before: 0, after: 480 },
                    children: []
                }));
            }
        });

        const document = new Document({
            sections: [{
                properties: {
                    page: {
                        size: { width: 11906, height: 16838 },
                        margin: { top: 650, right: 270, bottom: 650, left: 270, header: 0, footer: 0, gutter: 0 }
                    }
                },
                children
            }]
        });

        const date = opgRows[0].date || 'OPG';
        const filename = `REPORT FOR OPG - ${date}`.replace(/[\\/:*?"<>|]/g, '-') + '.docx';

        saveAs(await Packer.toBlob(document), filename);

        showExampleStatus('OPG Report downloaded.', 'success');
    } catch (error) {
        console.error('OPG report error:', error);
        showExampleStatus(`Error generating OPG Report: ${error.message}`, 'error');
    } finally {
        button.disabled = !opgRows.length;
    }
}
