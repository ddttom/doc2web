// File: lib/css/generators/table-styles.js
// Table style generation

const { convertTwipToPt } = require("../../utils/unit-converter");
const { getBorderStyle } = require("./base-styles");

/**
 * Generate table styles from style information
 */
function generateTableStyles(styleInfo) {
  let css = "\n/* Table Styles */\n";
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

module.exports = {
  generateTableStyles,
};