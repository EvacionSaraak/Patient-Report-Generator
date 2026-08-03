// Dispatch to the correct download format handler
async function generateReport() {
    if (selectedDownloadFormat === 'txt') {
        generateTextDocument();
        return;
    }

    await generateWordDocument();
}

// Generate Word document using docxtemplater and the patient_report_template.docx
async function generateWordDocument() {
    if (!parsedData || parsedData.length === 0) {
        showStatus('No data to export.', 'error');
        return;
    }

    try {
        showStatus('Generating Word document...', 'info');
        downloadBtn.disabled = true;

        if (typeof PizZip === 'undefined' || typeof docxtemplater === 'undefined') {
            throw new Error('Templating libraries not loaded. Please refresh the page.');
        }
        if (typeof PATIENT_REPORT_TEMPLATE_B64 === 'undefined') {
            throw new Error('Report template not found. Please refresh the page.');
        }

        const headers = parsedData[0] || [];
        const rows = filterEmptyRows(parsedData.slice(1));

        const ptNoIdx      = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
        const nameIdx      = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
        const dateIdx      = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
        const drIdx        = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
        const remindersIdx = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));

        const patients = rows.map(row => {
            const reminder = remindersIdx >= 0 && row[remindersIdx] !== undefined
                ? String(row[remindersIdx]).trim() : '';
            const remUpper = reminder.toUpperCase();
            return {
                file_no:    ptNoIdx  >= 0 && row[ptNoIdx]  !== undefined ? String(row[ptNoIdx])  : '',
                pt_name:    nameIdx  >= 0 && row[nameIdx]  !== undefined ? String(row[nameIdx])  : '',
                visit_date: dateIdx  >= 0 && row[dateIdx]  !== undefined ? formatDate(row[dateIdx]) : '',
                dr:         drIdx    >= 0 && row[drIdx]    !== undefined ? String(row[drIdx]).trim() : '',
                reminder,
                is_yellow: remUpper.startsWith('LAST VISIT'),
                is_green:  remUpper.startsWith('NEW PATIENT'),
            };
        });

        const zip = new PizZip(PATIENT_REPORT_TEMPLATE_B64, { base64: true });
        const doc = new docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
        doc.render({ patients });

        const dateRange = getDateRange(parsedData);
        let filename = 'PATIENT REPORT DATED ';
        filename += dateRange.min ? dateRange.min : 'Unknown';
        filename += '.docx';

        const blob = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
        saveAs(blob, filename);

        showStatus('Word document generated successfully!', 'success');
        downloadBtn.disabled = false;
    } catch (error) {
        showStatus('Error generating document: ' + error.message, 'error');
        console.error('Error details:', error);
        downloadBtn.disabled = false;
    }
}

// Generate plain-text document from the current text preview content
function generateTextDocument() {
    if (!parsedData || parsedData.length === 0) {
        showStatus('No data to export.', 'error');
        return;
    }

    try {
        showStatus('Generating text document...', 'info');
        downloadBtn.disabled = true;

        const dateRange = getDateRange(parsedData);
        let filename = 'PATIENT REPORT _ DATED ';
        if (dateRange.min && dateRange.max) {
            filename += `${dateRange.min} - ${dateRange.max}.txt`;
        } else {
            filename += 'Unknown.txt';
        }

        const content = textPreview.textContent || '';
        const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
        saveAs(blob, filename);

        showStatus('Text document generated successfully!', 'success');
        downloadBtn.disabled = false;
    } catch (error) {
        showStatus('Error generating text document: ' + error.message, 'error');
        console.error('Error details:', error);
        downloadBtn.disabled = false;
    }
}
