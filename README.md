# Patient Report Generator

A tool to automate generation of patient reports because I'm lazy.

## Features

- 📊 Upload XLSX/XLS files
- 👁️ Preview data before export
- 📄 Generate Word documents with one click
- 📝 Generate plain text reports with one click
- 🔁 Switch between Word and Text report tabs before download
- 🌐 Works entirely in the browser (no server needed)
- ✨ Clean, modern interface

## How to Use

1. Open the [Patient Report Generator](https://evacionsaraak.github.io/Patient-Report-Generator/) on GitHub Pages
2. Click "Choose XLSX File" to upload your Excel file
3. Preview the data in the preview section
4. Select either the **Word (.docx)** or **Text (.txt)** tab
5. Click the matching download button to save the report

## Local Development

Simply open `index.html` in a web browser. No build process or server required.

## Repository Structure

- `css/` - Stylesheets
- `js/` - JavaScript application logic
- `assets/icons/` - Icon assets
- `docs/` - Additional documentation

## Technologies Used

- HTML5, CSS3, JavaScript (ES6+)
- [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) - For reading Excel files
- [docx](https://github.com/dolanmiu/docx) - For generating Word documents
- [FileSaver.js](https://github.com/eligrey/FileSaver.js) - For downloading files

## GitHub Pages Deployment

This site is automatically deployed to GitHub Pages. To deploy:

1. Go to repository Settings
2. Navigate to Pages section
3. Select branch (e.g., `main` or `copilot/add-xlsx-to-word-export-tool`)
4. Click Save

The site will be available at: `https://evacionsaraak.github.io/Patient-Report-Generator/`
