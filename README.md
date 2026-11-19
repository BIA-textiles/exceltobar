# BarcodeStudio — Excel to Barcode Generator

A modern, professional static web application built with **HTML5**, **CSS3**, and **Vanilla JavaScript** to convert words and values from any Excel or CSV file into scannable barcodes.

---

## 🌟 Key Features

1. **Automatic Column 1 Extraction**:
   - Drag & drop or upload any `.xlsx`, `.xls`, or `.csv` spreadsheet.
   - Automatically detects the **Column 1 Header Name** (e.g., `SKU`, `Product Code`, `Item ID`) and includes it at the top of every barcode card/label.
   - Reads every word/text entry in Column 1 and renders high-resolution barcodes.

2. **Professional & Scannable Output**:
   - Built with **CODE 128** (universal alphanumeric support for words, codes, and serial numbers).
   - Also supports **CODE 39**, **EAN-13**, **UPC-A**, **EAN-8**, **ITF-14**, **MSI**, and **Pharmacode**.
   - Vector SVG barcodes render with crystal-clear precision at any scale (300+ DPI equivalent for optical laser and camera scanners).

3. **Label Customization**:
   - Adjust **Barcode Height** (30px – 120px) and **Bar Width Scale** in real-time.
   - Toggle **Column 1 Name Header** on/off.
   - Toggle **Value Text** below the barcode.
   - Toggle **Item Counter Index** (`#1`, `#2`, etc.).
   - Trim whitespace automatically.

4. **Print & Sticker Sheet Ready**:
   - Dedicated print stylesheet (`@media print`) designed for standard A4, Letter, and adhesive sticker sheets.
   - Clean print mode automatically hides navigation, controls, and buttons, formatting barcodes into 1, 2, 3, or 4-column sticker grids without awkward page breaks.

5. **Batch Export**:
   - **Download All (ZIP)**: Exports all generated barcodes into a `.zip` archive containing individual high-res PNG images.
   - **Download Single PNG / SVG**: Export individual barcodes with one click.
   - **One-Click Copy**: Copy barcode values to your clipboard.

6. **Instant Testing**:
   - Click **"Try Demo Data"** to preview 12 realistic barcodes instantly.
   - Click **"Excel Template"** to download a ready-to-fill `.xlsx` template.

7. **100% Client-Side & Private**:
   - Your data is processed entirely in your browser; no files or data are ever uploaded to an external server.
   - Includes a built-in offline SVG barcode generator engine and CSV parser fallback.

---

## 🚀 How to Run

Since this is a static web page, no installation or backend server is required:

1. Simply **double-click** `index.html` to open it in your preferred web browser (Google Chrome, Microsoft Edge, Firefox, Brave, Safari).
2. Upload your Excel spreadsheet or click **Try Demo Data**.
3. Customize your layout and click **Print Labels** or **Download All (ZIP)**.

---

## 📄 File Structure

- `index.html`: Clean, semantic HTML structure with accessible UI controls.
- `styles.css`: Modern responsive stylesheet with print-specific layout rules.
- `app.js`: Complete application logic for parsing Excel files, barcode rendering, filtering, and export.
- `README.md`: Documentation and usage guide.
