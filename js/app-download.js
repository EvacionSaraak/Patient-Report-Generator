// Dispatch to the correct download format handler
async function generateReport() {
    if (selectedDownloadFormat === 'txt') {
        generateTextDocument();
        return;
    }

    await generateWordDocument();
}

async function generateWordDocument() {
    if (!parsedData || !parsedData.records || parsedData.records.length === 0) {
        showStatus('No data to export.', 'error');
        return;
    }

    try {
        showStatus('Generating Word document...', 'info');
        downloadBtn.disabled = true;

        let lib = window.docx;
        if (!lib && typeof docx !== 'undefined') {
            lib = docx;
        }

        if (!lib) {
            throw new Error('docx library is not loaded. Please refresh the page and try again.');
        }

        const contentToUse = createNormalDocumentContent(parsedData, lib);
        const doc = new lib.Document({
            sections: [{
                properties: {},
                children: contentToUse
            }]
        });

        const dateRange = getDateRange(parsedData.records);
        const formattedDateRange = formatDateRange(dateRange);
        let filename = 'PATIENT REPORT _ DATED ';

        if (formattedDateRange) {
            filename += `${formattedDateRange}.docx`;
        } else {
            filename += 'Unknown.docx';
        }

        const blob = await lib.Packer.toBlob(doc);
        saveAs(blob, filename);

        showStatus('Word document generated successfully!', 'success');
        downloadBtn.disabled = false;
    } catch (error) {
        showStatus('Error generating document: ' + error.message, 'error');
        console.error('Error details:', error);
        downloadBtn.disabled = false;
    }
}

function generateTextDocument() {
    if (!parsedData || !parsedData.records || parsedData.records.length === 0) {
        showStatus('No data to export.', 'error');
        return;
    }

    try {
        showStatus('Generating text document...', 'info');
        downloadBtn.disabled = true;

        const dateRange = getDateRange(parsedData.records);
        const formattedDateRange = formatDateRange(dateRange);
        let filename = 'PATIENT REPORT _ DATED ';

        if (formattedDateRange) {
            filename += `${formattedDateRange}.txt`;
        } else {
            filename += 'Unknown.txt';
        }

        const content = generateNormalTextReport(parsedData);
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
