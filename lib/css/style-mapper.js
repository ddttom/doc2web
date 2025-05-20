// lib/css/style-mapper.js - Style mapping functions for DOCX to HTML/CSS

/**
 * Create a custom style map based on extracted styles
 * Maps DOCX style names to HTML elements with appropriate CSS classes
 * 
 * @param {Object} styleInfo - Extracted style information
 * @returns {Array<string>} - Style map entries for mammoth.js
 */
function createStyleMap(styleInfo) {
  const styleMap = [];
  
  // Map paragraph styles
  Object.entries(styleInfo.styles.paragraph || {}).forEach(([id, style]) => {
    styleMap.push(
      `p[style-name='${style.name}'] => p.docx-p-${id.toLowerCase()}`
    );
  });
  
  // Map character styles
  Object.entries(styleInfo.styles.character || {}).forEach(([id, style]) => {
    styleMap.push(
      `r[style-name='${style.name}'] => span.docx-c-${id.toLowerCase()}`
    );
  });
  
  // Map table styles
  Object.entries(styleInfo.styles.table || {}).forEach(([id, style]) => {
    styleMap.push(
      `table[style-name='${style.name}'] => table.docx-t-${id.toLowerCase()}`
    );
  });
  
  // Heading mappings
  styleMap.push("p[style-name='heading 1'] => h1.docx-heading1");
  styleMap.push("p[style-name='heading 2'] => h2.docx-heading2");
  styleMap.push("p[style-name='heading 3'] => h3.docx-heading3");
  styleMap.push("p[style-name='heading 4'] => h4.docx-heading4");
  styleMap.push("p[style-name='heading 5'] => h5.docx-heading5");
  styleMap.push("p[style-name='heading 6'] => h6.docx-heading6");
  
  // Handle headings with styles like "Heading1"
  styleMap.push("p[style-name='Heading1'] => h1.docx-heading1");
  styleMap.push("p[style-name='Heading2'] => h2.docx-heading2");
  styleMap.push("p[style-name='Heading3'] => h3.docx-heading3");
  
  // Handle TOC specific styles
  styleMap.push("p[style-name='TOC Heading'] => h2.docx-toc-heading");
  styleMap.push("p[style-name='TOC1'] => p.docx-toc-entry.docx-toc-level-1");
  styleMap.push("p[style-name='TOC2'] => p.docx-toc-entry.docx-toc-level-2");
  styleMap.push("p[style-name='TOC3'] => p.docx-toc-entry.docx-toc-level-3");
  styleMap.push("p[style-name='toc 1'] => p.docx-toc-entry.docx-toc-level-1");
  styleMap.push("p[style-name='toc 2'] => p.docx-toc-entry.docx-toc-level-2");
  styleMap.push("p[style-name='toc 3'] => p.docx-toc-entry.docx-toc-level-3");
  
  // Additional custom mappings for specific elements
  styleMap.push("p:fresh => p");
  styleMap.push("r[bold] => strong");
  styleMap.push("r[italic] => em");
  styleMap.push("r[underline] => span.docx-underline");
  styleMap.push("r[strikethrough] => span.docx-strike");
  styleMap.push("r[subscript] => sub");
  styleMap.push("r[superscript] => sup");
  
  // General paragraph styles
  styleMap.push("p[style-name='Normal Web'] => p.docx-normalweb"); 
  styleMap.push("p[style-name='Body Text'] => p.docx-bodytext");
  
  return styleMap;
}

/**
 * Create document transformer function for style preservation
 * This function can be used to enhance the document with additional styling
 * 
 * @param {Object} styleInfo - Style information
 * @returns {Function} - Document transformer function for mammoth.js
 */
function createDocumentTransformer(styleInfo) {
  return function (document) {
    // In a complete implementation, this would walk the document tree
    // and enhance elements with style attributes based on the styleInfo
    
    // Process paragraphs, runs, tables and other elements
    // Add class names and style attributes as needed
    
    return document;
  };
}

module.exports = {
  createStyleMap,
  createDocumentTransformer
};