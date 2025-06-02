// File: lib/css/generators/table-styles.js
// Table style generation

const { convertTwipToPt } = require("../../utils/unit-converter");
const { getBorderStyle } = require("./base-styles");

/**
 * Generate table styles from style information
 */
function generateTableStyles(styleInfo) {
  let css = "\n/* Table Styles */\n";
  
  // Add default table styling for docx-table-default class
  css += generateDefaultTableStyles();
  
  // Add table-responsive wrapper styles
  css += generateTableResponsiveStyles();
  
  // Generate styles for specific table styles from DOCX
  Object.entries(styleInfo.styles?.table || {}).forEach(([id, style]) => {
    const className = `docx-t-${id
      .replace(/[^a-zA-Z0-9-_]/g, "-")
      .toLowerCase()}`;
    css += `
.${className} {
  border-collapse: ${
    style.cellSpacing === "0" ? "collapse" : "separate"
  }; /* Handle cell spacing */
  border-spacing: ${
    style.cellSpacing && style.cellSpacing !== "0"
      ? convertTwipToPt(style.cellSpacing) + "pt"
      : "0"
  };
  width: ${style.width || "auto"}; /* auto or specific width */
  margin-bottom: 1em;
}
.${className} td, .${className} th {
  padding: ${
    style.cellPadding ? convertTwipToPt(style.cellPadding) + "pt" : "5pt"
  };
  text-align: ${style.alignment || "left"}; /* Default cell alignment */
  ${getBorderStyle(style, "top")} ${getBorderStyle(style, "bottom")}
  ${getBorderStyle(style, "left")} ${getBorderStyle(style, "right")}
  ${getBorderStyle(style, "insideH")} ${getBorderStyle(style, "insideV")}
}
`;
  });
// Add comprehensive table data attribute styles
  css += generateTableDataAttributeStyles();
  return css;
}

/**
 * Generate default table styles for tables without specific DOCX styling
 */
function generateDefaultTableStyles() {
  return `
/* Default Table Styles */
.docx-table-default {
  border-collapse: collapse;
  width: 100%;
  margin: 1em 0;
  font-size: inherit;
  background-color: transparent;
}

.docx-table-default th,
.docx-table-default td {
  padding: 8pt 12pt;
  text-align: left;
  vertical-align: top;
  border: 1pt solid #d0d0d0;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

.docx-table-default th {
  background-color: #f8f9fa;
  font-weight: bold;
  border-bottom: 2pt solid #d0d0d0;
}

.docx-table-default tbody tr:nth-child(even) {
  background-color: #f8f9fa;
}

.docx-table-default tbody tr:hover {
  background-color: #e9ecef;
}

/* Table caption styling */
.docx-table-default caption {
  caption-side: top;
  padding: 8pt 0;
  font-weight: bold;
  text-align: left;
  color: #495057;
}
`;
}

/**
 * Generate responsive table wrapper styles
 */
function generateTableResponsiveStyles() {
  return `
/* Responsive Table Wrapper */
.table-responsive {
  display: block;
  width: 100%;
  overflow-x: auto;
  -webkit-overflow-scrolling: touch;
  margin: 1em 0;
}

.table-responsive > table {
  margin-bottom: 0;
}

/* Ensure tables don't break layout on small screens */
@media (max-width: 768px) {
  .table-responsive {
    font-size: 0.875em;
  }
  
  .docx-table-default th,
  .docx-table-default td {
    padding: 6pt 8pt;
  }
}
`;
}
/**
 * Generate comprehensive table data attribute styles
 */
function generateTableDataAttributeStyles() {
  return `
/* Enhanced Table Formatting - Data Attribute Styles */

/* Table properties */
table[data-tbl-w] { /* Table width applied via data attributes */ }
table[data-tbl-ind] { /* Table indentation applied via data attributes */ }
table[data-tbl-layout] { /* Table layout applied via data attributes */ }
table[data-tbl-overlap] { /* Table overlap applied via data attributes */ }

.docx-table { 
  border-collapse: collapse;
  width: 100%;
  margin: 1em 0;
}

.docx-table-ind { 
  /* Table with indentation */
}

.docx-table-layout { 
  /* Table with specific layout */
}

.docx-table-overlap { 
  /* Table with overlap settings */
}

/* Table row properties */
tr[data-tr-height] { /* Row height applied via data attributes */ }
tr[data-tr-height-rule] { /* Row height rule applied via data attributes */ }
tr[data-tr-cant-split] { page-break-inside: avoid; }
tr[data-tr-header] { /* Header row applied via data attributes */ }

.docx-tr { 
  /* Enhanced table row */
}

.docx-tr-cant-split { 
  page-break-inside: avoid;
}

.docx-tr-header { 
  font-weight: bold;
  background-color: #f8f9fa;
}

/* Table cell properties */
td[data-tc-w] { /* Cell width applied via data attributes */ }
td[data-tc-borders] { /* Cell borders applied via data attributes */ }
td[data-tc-shd] { /* Cell shading applied via data attributes */ }
td[data-tc-mar] { /* Cell margins applied via data attributes */ }
td[data-tc-valign] { /* Cell vertical alignment applied via data attributes */ }
td[data-tc-fit-text] { /* Cell fit text applied via data attributes */ }
td[data-tc-no-wrap] { white-space: nowrap; }
td[data-tc-grid-span] { /* Cell grid span applied via data attributes */ }
td[data-tc-vmerge] { /* Cell vertical merge applied via data attributes */ }

.docx-tc { 
  padding: 8pt 12pt;
  text-align: left;
  vertical-align: top;
  border: 1pt solid #d0d0d0;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

.docx-tc-borders { 
  /* Cell with enhanced borders */
}

.docx-tc-shd { 
  /* Cell with shading */
}

.docx-tc-margins { 
  /* Cell with custom margins */
}

.docx-tc-valign { 
  /* Cell with vertical alignment */
}

.docx-tc-fit-text { 
  font-size: smaller;
  white-space: nowrap;
}

.docx-tc-no-wrap { 
  white-space: nowrap;
}

.docx-tc-grid-span { 
  /* Cell with grid span */
}

.docx-tc-vmerge { 
  /* Cell with vertical merge */
}

/* Table style utility classes */
.docx-tbl-grid { 
  border-collapse: collapse;
  border: 1pt solid #000000;
}

.docx-tbl-grid td, .docx-tbl-grid th { 
  border: 1pt solid #000000;
  padding: 4pt 8pt;
}

.docx-tbl-list { 
  border-collapse: collapse;
  border: none;
}

.docx-tbl-list td, .docx-tbl-list th { 
  border: none;
  padding: 2pt 4pt;
}

.docx-tbl-colorful { 
  border-collapse: collapse;
  border: 2pt solid #4472C4;
}

.docx-tbl-colorful th { 
  background-color: #4472C4;
  color: white;
  font-weight: bold;
  padding: 8pt 12pt;
}

.docx-tbl-colorful td { 
  border: 1pt solid #4472C4;
  padding: 6pt 10pt;
}

.docx-tbl-colorful tbody tr:nth-child(even) { 
  background-color: #F2F2F2;
}

/* Enhanced table border styles */
.docx-border-single { border-style: solid; }
.docx-border-double { border-style: double; }
.docx-border-dotted { border-style: dotted; }
.docx-border-dashed { border-style: dashed; }
.docx-border-thick { border-width: 2pt; }
.docx-border-thin { border-width: 0.5pt; }

/* Enhanced table alignment */
.docx-table-center { margin-left: auto; margin-right: auto; }
.docx-table-left { margin-left: 0; margin-right: auto; }
.docx-table-right { margin-left: auto; margin-right: 0; }

/* Enhanced cell alignment */
.docx-cell-top { vertical-align: top; }
.docx-cell-middle { vertical-align: middle; }
.docx-cell-bottom { vertical-align: bottom; }
.docx-cell-left { text-align: left; }
.docx-cell-center { text-align: center; }
.docx-cell-right { text-align: right; }
.docx-cell-justify { text-align: justify; }
`;
}

module.exports = {
  generateTableStyles,
  generateDefaultTableStyles,
  generateTableResponsiveStyles,
  generateTableDataAttributeStyles,
};
