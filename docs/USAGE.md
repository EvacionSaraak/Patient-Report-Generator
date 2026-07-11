# Patient Report Generator - Usage Guide

## Overview
The Patient Report Generator is a web-based tool that converts Excel spreadsheets (XLSX format) into formatted Word documents. It runs entirely in your browser with no server required.

## Features

### 📤 File Upload
- Drag and drop or click to select XLSX/XLS files
- Automatic file validation
- Visual feedback when file is selected

### 👁️ Data Preview
- See your data in a table before exporting
- Shows first 10 rows for quick verification
- Displays total row count

### 📄 Word Document Generation
- Creates professionally formatted Word documents
- Includes:
  - Document title
  - Generation timestamp
  - Data presented in a formatted table
  - Proper styling and formatting

### 📝 Text Report Generation
- Creates lightweight plain-text report files
- Uses the same patient fields shown in the Word report
- Useful for quick sharing or archiving

### 💾 Download
- One-click download
- Choose between `.docx` and `.txt`
- Compatible with Microsoft Word and Google Docs

## How to Use

1. **Open the Application**
   - Visit the GitHub Pages URL
   - Or open `index.html` locally in your browser

2. **Upload Your XLSX File**
   - Click the "Choose XLSX File" button
   - Select your Excel file
   - Wait for the preview to appear

3. **Review the Data**
   - Check the preview table
   - Ensure your data looks correct

4. **Select Report Format**
   - Use the **Word (.docx)** and **Text (.txt)** tabs in the report preview panel

5. **Download Report**
   - Click the format-specific download button
   - The selected file type downloads automatically

## Expected XLSX Format

The application works best with spreadsheets that have:
- **First row**: Column headers
- **Subsequent rows**: Data entries

Example structure:
```
| Patient ID | Name        | Age | Condition    | Date       | Treatment           |
|-----------|-------------|-----|--------------|------------|---------------------|
| P001      | John Doe    | 45  | Hypertension | 2024-01-15 | Medication prescribed|
| P002      | Jane Smith  | 32  | Diabetes     | 2024-01-16 | Insulin therapy     |
```

## Browser Compatibility

Works on all modern browsers:
- ✅ Chrome/Edge (v90+)
- ✅ Firefox (v88+)
- ✅ Safari (v14+)
- ✅ Opera (v76+)

## Privacy & Security

- 🔒 **All processing happens locally** - your data never leaves your browser
- 🔒 **No server uploads** - files are processed entirely client-side
- 🔒 **No data storage** - nothing is saved or transmitted

## Technical Notes

- File size limit: Depends on your browser's memory (typically up to 100MB)
- Supported formats: .xlsx, .xls
- Output format: .docx (Microsoft Word 2007+)

## Troubleshooting

**File won't upload?**
- Ensure it's a valid XLSX or XLS file
- Check the file isn't corrupted
- Try a different file

**Download button disabled?**
- Wait for the file to finish loading
- Check for error messages in the status area
- Refresh the page and try again

**Preview looks wrong?**
- Verify your Excel file has data in the first sheet
- Ensure headers are in the first row
- Check for merged cells or complex formatting

## Support

For issues or questions, please open an issue on the GitHub repository.
