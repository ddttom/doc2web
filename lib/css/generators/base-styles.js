// File: lib/css/generators/base-styles.js
// Base document styles and utilities

const {
  convertTwipToPt,
  convertBorderSizeToPt,
  getBorderTypeValue,
} = require("../../utils/unit-converter");

/**
 * Generate base document styles including body and container
 */
function generateBaseStyles(styleInfo) {
  // Calculate page margins from DOCX introspection
  const pageMargins = styleInfo.settings?.pageMargins;
  let bodyMargins = "20px"; // Default fallback
  
  if (pageMargins) {
    const topMargin = convertTwipToPt(pageMargins.top);
    const rightMargin = convertTwipToPt(pageMargins.right);
    const bottomMargin = convertTwipToPt(pageMargins.bottom);
    const leftMargin = convertTwipToPt(pageMargins.left);
    
    // Use the extracted margins to match Word layout
    bodyMargins = `${topMargin}pt ${rightMargin}pt ${bottomMargin}pt ${leftMargin}pt`;
  }

  return `
/* Document defaults - using DOCX page margins */
body {
  font-family: "${styleInfo.theme?.fonts?.minor || "Calibri"}", sans-serif;
  font-size: ${styleInfo.documentDefaults?.character?.fontSize || "11pt"};
  line-height: ${
    styleInfo.documentDefaults?.paragraph?.lineHeight || "1.2"
  }; /* Use specific or default */
  margin: ${bodyMargins}; /* Use actual DOCX page margins */
  padding: 0;
  color: ${styleInfo.theme?.colors?.tx1 || "#000000"}; /* Default text color */
  background-color: ${
    styleInfo.theme?.colors?.bg1 || "#FFFFFF"
  }; /* Default bg color */
  max-width: ${pageMargins?.pageSize ? convertTwipToPt(pageMargins.pageSize.width) + 'pt' : 'none'}; /* Match page width */
  box-sizing: border-box;
}

/* Container for content to match Word's content area */
.docx-content {
  width: 100%;
  box-sizing: border-box;
}
`;
}

/**
 * Get font family from style with fallbacks
 */
function getFontFamily(style, styleInfo) {
  if (!style.font) {
    return styleInfo.theme?.fonts?.minor || "Calibri";
  }
  return (
    style.font.ascii ||
    style.font.hAnsi ||
    styleInfo.theme?.fonts?.minor ||
    "Calibri"
  );
}

/**
 * Generate border style CSS property
 */
function getBorderStyle(style, side) {
  if (!style.borders || !style.borders[side]) {
    return "";
  }
  const border = style.borders[side];
  if (
    !border ||
    !border.value ||
    border.value === "nil" ||
    border.value === "none"
  ) {
    return `border-${side}: none;`;
  }
  const width = border.size ? convertBorderSizeToPt(border.size) : 1;
  const color =
    border.color && border.color !== "auto"
      ? `#${border.color}`
      : "currentColor";
  const borderStyle = getBorderTypeValue(border.value);
  return `border-${side}: ${width}pt ${borderStyle} ${color};`;
}

/**
 * Generate fallback CSS for error cases
 */
function generateFallbackCSS(styleInfo) {
  return `
body { font-family: Calibri, sans-serif; font-size: 11pt; line-height: 1.15; margin: 20px; }
h1 { font-size: 16pt; } h2 { font-size: 13pt; } p { margin: 10pt 0; }
`;
}

module.exports = {
  generateBaseStyles,
  getFontFamily,
  getBorderStyle,
  generateFallbackCSS,
};