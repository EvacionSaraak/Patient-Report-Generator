// Dispatch to the correct download format handler
async function generateReport() {
    if (selectedDownloadFormat === 'txt') {
        generateTextDocument();
        return;
    }

    await generateWordDocument();
}

// Generate Word document using docxtemplater and the patient_report_template.docx.
// Consumes canonical records from parsedData (set by parseNormalSheet).
async function generateWordDocument() {
    if (!parsedData || !parsedData.records || parsedData.records.length === 0) {
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

        const patients = parsedData.records.map(record => {
            const reminder = record.personalReminders;
            const remUpper = reminder.toUpperCase();
            return {
                file_no:    record.fileNumber,
                pt_name:    record.patientName,
                visit_date: formatDate(record.visitDate),
                dr:         record.doctor,
                reminder,
                is_yellow: remUpper.startsWith('LAST VISIT'),
                is_green:  remUpper.startsWith('NEW PATIENT'),
            };
        });

        const zip = new PizZip(PATIENT_REPORT_TEMPLATE_B64, { base64: true });
        const doc = new docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
        doc.render({ patients });

        const dateRange = getDateRange(parsedData.records);
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

// Generate plain-text document from the current text preview content.
function generateTextDocument() {
    if (!parsedData || !parsedData.records || parsedData.records.length === 0) {
        showStatus('No data to export.', 'error');
        return;
    }

    try {
        showStatus('Generating text document...', 'info');
        downloadBtn.disabled = true;

        const dateRange = getDateRange(parsedData.records);
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
