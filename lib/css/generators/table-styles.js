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

module.exports = {
  generateTableStyles,
  generateDefaultTableStyles,
  generateTableResponsiveStyles,
};
