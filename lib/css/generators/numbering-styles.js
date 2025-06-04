// File: lib/css/generators/numbering-styles.js
// Generate CSS for DOCX numbering and bullet lists

const { calculateDefaultParagraphMargin } = require("./utility-styles");

/**
 * Generate CSS for DOCX numbering styles
 */
function generateDOCXNumberingStyles(numberingDefs, styleInfo) {
  if (!numberingDefs || !numberingDefs.abstractNums) {
    return "";
  }

  let css = "\n/* DOCX Numbering Styles */\n";

  // Generate styles for each abstract numbering definition
  Object.entries(numberingDefs.abstractNums).forEach(([abstractNumId, abstractNum]) => {
    if (!abstractNum.levels) return;

    Object.entries(abstractNum.levels).forEach(([level, levelDef]) => {
      const selector = `[data-abstract-num="${abstractNumId}"][data-num-level="${level}"]`;
      
      css += `
${selector} {
  margin-left: ${(parseInt(level) * 18) || 0}pt;
  text-indent: ${levelDef.textIndent || '-18pt'};
  padding-left: ${levelDef.paddingLeft || '18pt'};
  position: relative;
}
`;

      // Add numbering format styles
      if (levelDef.format === 'bullet') {
        css += `
${selector}::before {
  content: "${levelDef.bulletChar || '•'}";
  position: absolute;
  left: ${levelDef.leftIndent || '0pt'};
  width: 18pt;
  text-align: left;
}
`;
      }
      // Note: Decimal and other numbering formats are handled by the original DOCX content
      // No CSS counter generation needed as numbers are already in the HTML
    });
  });

  // Add special handling for bullet lists inside table cells
  css += `

/* Special handling for bullet lists inside table cells */
table ul.docx-bullet-list,
td ul.docx-bullet-list,
th ul.docx-bullet-list {
  margin: 0 0 0 20pt !important; /* Reset to relative positioning within table cells */
}

table li.docx-list-item,
td li.docx-list-item,
th li.docx-list-item,
table li[data-format="bullet"],
td li[data-format="bullet"],
th li[data-format="bullet"] {
  position: relative !important;
  margin-left: 0pt !important; /* Reset level-specific margins in table cells */
  padding-left: 0 !important;
  text-indent: 0 !important;
}

table li.docx-list-item::before,
td li.docx-list-item::before,
th li.docx-list-item::before,
table li[data-format="bullet"]::before,
td li[data-format="bullet"]::before,
th li[data-format="bullet"]::before {
  content: "•" !important;
  position: absolute !important;
  left: -15pt !important; /* Closer to text in table cells */
  top: 0 !important;
  width: 10pt !important;
  text-align: left !important;
  font-weight: normal !important;
  display: inline-block !important;
  color: inherit !important;
}

table li.docx-bullet-level-1,
td li.docx-bullet-level-1,
th li.docx-bullet-level-1 {
  margin-left: 18pt !important; /* Relative to table cell, not document */
}
`;

  return css;
}

/**
 * Generate enhanced list styles
 */
function generateEnhancedListStyles(styleInfo) {
  let css = "\n/* Enhanced List Styles */\n";
  
  css += `
/* Reset default list styles */
ul, ol {
  margin: 0;
  padding: 0;
  list-style: none;
}

/* Level-based indentation for numbered elements */
[data-num-level="0"] { margin-left: 0pt; }
[data-num-level="1"] { margin-left: 18pt; }
[data-num-level="2"] { margin-left: 36pt; }
[data-num-level="3"] { margin-left: 54pt; }
[data-num-level="4"] { margin-left: 72pt; }
[data-num-level="5"] { margin-left: 90pt; }
[data-num-level="6"] { margin-left: 108pt; }
[data-num-level="7"] { margin-left: 126pt; }
[data-num-level="8"] { margin-left: 144pt; }

/* Basic positioning for numbered elements */
[data-num-id] {
  position: relative;
}

/* List item support */
li[data-abstract-num] {
  display: block; 
}
`;
  // Add special handling for bullet lists inside table cells
  css += `

/* Special handling for bullet lists inside table cells */
table ul.docx-bullet-list,
td ul.docx-bullet-list,
th ul.docx-bullet-list {
  margin: 0 0 0 20pt !important; /* Reset to relative positioning within table cells */
}

table li.docx-list-item,
td li.docx-list-item,
th li.docx-list-item,
table li[data-format="bullet"],
td li[data-format="bullet"],
th li[data-format="bullet"] {
  position: relative !important;
  margin-left: 0pt !important; /* Reset level-specific margins in table cells */
  padding-left: 0 !important;
  text-indent: 0 !important;
}

table li.docx-list-item::before,
td li.docx-list-item::before,
th li.docx-list-item::before,
table li[data-format="bullet"]::before,
td li[data-format="bullet"]::before,
th li[data-format="bullet"]::before {
  content: "•" !important;
  position: absolute !important;
  left: -15pt !important; /* Closer to text in table cells */
  top: 0 !important;
  width: 10pt !important;
  text-align: left !important;
  font-weight: normal !important;
  display: inline-block !important;
  color: inherit !important;
}

table li.docx-bullet-level-1,
td li.docx-bullet-level-1,
th li.docx-bullet-level-1 {
  margin-left: 18pt !important; /* Relative to table cell, not document */
}
`;

  return css;
}

/**
 * Generate CSS for bullet lists
 */
function generateBulletListStyles(styleInfo) {
  let css = "\n/* Bullet List Styles */\n";
  
  // Calculate the document's base margin to ensure bullet lists align with regular content
  const defaultParagraphMargin = calculateDefaultParagraphMargin(styleInfo);
  const baseMarginValue = parseFloat(defaultParagraphMargin) || 0;
  const bulletListMargin = baseMarginValue + 20; // Add 20pt for bullet indentation
  
  // Base bullet list styling with enhanced specificity
  css += `
/* Bullet list containers - add proper indentation */
ul.docx-bullet-list {
  list-style-type: none !important;
  padding-left: 0 !important;
  margin: 0 0 0 ${bulletListMargin}pt !important;
  position: relative !important;
}

/* High specificity selectors for bullet list items */
ul.docx-bullet-list li.docx-list-item,
ul.docx-bullet-list li[data-format="bullet"],
ul.docx-bullet-list li[class*="docx-bullet-"],
ul.docx-bullet-list li[class*="docx-list-item"],
li.docx-list-item[data-format="bullet"],
li[data-format="bullet"].docx-list-item {
  position: relative !important;
  margin-left: 0 !important;
  margin-bottom: 0.5em !important;
  line-height: 1.4 !important;
  list-style: none !important;
  padding-left: 20pt !important;
  text-indent: -20pt !important;
  word-break: normal !important;
  overflow-wrap: normal !important;
  word-wrap: normal !important;
  hyphens: none !important;
}

/* High specificity selectors for bullet characters */
ul.docx-bullet-list li.docx-list-item::before,
ul.docx-bullet-list li[data-format="bullet"]::before,
ul.docx-bullet-list li[class*="docx-bullet-"]::before,
ul.docx-bullet-list li[class*="docx-list-item"]::before,
li.docx-list-item[data-format="bullet"]::before,
li[data-format="bullet"].docx-list-item::before {
  content: "•" !important;
  position: relative !important;
  left: 0 !important;
  top: 0 !important;
  width: auto !important;
  text-align: left !important;
  font-weight: normal !important;
  display: inline !important;
  color: inherit !important;
  margin-right: 0.5em !important;
}

/* Additional fallback selectors for maximum compatibility */
li.docx-list-item,
li[class*="docx-bullet-"],
li[class*="docx-list-item"],
li[data-format="bullet"] {
  position: relative !important;
  padding-left: 20pt !important;
  margin-bottom: 0.5em !important;
  line-height: 1.4 !important;
  list-style: none !important;
  text-indent: -20pt !important;
  word-break: normal !important;
  overflow-wrap: normal !important;
  word-wrap: normal !important;
  hyphens: none !important;
}

li.docx-list-item::before,
li[class*="docx-bullet-"]::before,
li[class*="docx-list-item"]::before,
li[data-format="bullet"]::before {
  content: "•" !important;
  position: relative !important;
  left: 0 !important;
  top: 0 !important;
  width: auto !important;
  text-align: left !important;
  font-weight: normal !important;
  display: inline !important;
  margin-right: 0.5em !important;
}
`;

  // Generate level-specific bullet styles
  for (let level = 0; level <= 8; level++) {
    const indent = baseMarginValue + (level * 18); // Base margin + level-specific indentation
    const bulletChar = getBulletCharForLevel(level);
    
    css += `
li.docx-bullet-level-${level} {
  margin-left: ${indent}pt;
}

li.docx-bullet-level-${level}::before {
  content: "${bulletChar}" !important;
}

/* Override DOCX abstract numbering styles with higher specificity */
li[data-abstract-num][data-num-level="${level}"]::before {
  content: "${bulletChar}" !important;
  position: relative !important;
  display: inline !important;
  margin-right: 0.5em !important;
  left: 0 !important;
  top: 0 !important;
  width: auto !important;
}
`;
  }
  
  // Add special handling for bullet lists inside table cells
  css += `

/* Special handling for bullet lists inside table cells */
table ul.docx-bullet-list,
td ul.docx-bullet-list,
th ul.docx-bullet-list {
  margin: 0 0 0 20pt !important; /* Reset to relative positioning within table cells */
}

table li.docx-list-item,
td li.docx-list-item,
th li.docx-list-item,
table li[data-format="bullet"],
td li[data-format="bullet"],
th li[data-format="bullet"] {
  position: relative !important;
  margin-left: 0pt !important; /* Reset level-specific margins in table cells */
  padding-left: 0 !important;
  text-indent: 0 !important;
}

table li.docx-list-item::before,
td li.docx-list-item::before,
th li.docx-list-item::before,
table li[data-format="bullet"]::before,
td li[data-format="bullet"]::before,
th li[data-format="bullet"]::before {
  content: "•" !important;
  position: absolute !important;
  left: -15pt !important; /* Closer to text in table cells */
  top: 0 !important;
  width: 10pt !important;
  text-align: left !important;
  font-weight: normal !important;
  display: inline-block !important;
  color: inherit !important;
}

table li.docx-bullet-level-1,
td li.docx-bullet-level-1,
th li.docx-bullet-level-1 {
  margin-left: 18pt !important; /* Relative to table cell, not document */
}
`;

  return css;
}

/**
 * Get bullet character for specific level
 */
function getBulletCharForLevel(level) {
  const bullets = ["•", "○", "•", "○", "•", "○", "•", "○"];
  return bullets[level % bullets.length];
}

/**
 * Generate CSS for custom bullet characters
 */
function generateCustomBulletStyles(numberingDefs) {
  let css = "";
  
  if (!numberingDefs || !numberingDefs.abstractNums) {
    return css;
  }
  
  Object.entries(numberingDefs.abstractNums).forEach(([id, abstractNum]) => {
    if (!abstractNum.levels) return;
    
    Object.entries(abstractNum.levels).forEach(([level, levelDef]) => {
      if (levelDef.format === 'bullet' && levelDef.bulletChar) {
        css += `
li[data-abstract-num="${id}"][data-num-level="${level}"]::before {
  content: "${levelDef.bulletChar}";
  ${levelDef.runProps?.color ? `color: ${levelDef.runProps.color};` : ''}
  ${levelDef.runProps?.fontSize ? `font-size: ${levelDef.runProps.fontSize};` : ''}
  ${levelDef.runProps?.fontFamily ? `font-family: "${levelDef.runProps.fontFamily}";` : ''}
}
`;
      }
    });
  });
  
  // Add special handling for bullet lists inside table cells
  css += `

/* Special handling for bullet lists inside table cells */
table ul.docx-bullet-list,
td ul.docx-bullet-list,
th ul.docx-bullet-list {
  margin: 0 0 0 20pt !important; /* Reset to relative positioning within table cells */
}

table li.docx-list-item,
td li.docx-list-item,
th li.docx-list-item,
table li[data-format="bullet"],
td li[data-format="bullet"],
th li[data-format="bullet"] {
  position: relative !important;
  margin-left: 0pt !important; /* Reset level-specific margins in table cells */
  padding-left: 0 !important;
  text-indent: 0 !important;
}

table li.docx-list-item::before,
td li.docx-list-item::before,
th li.docx-list-item::before,
table li[data-format="bullet"]::before,
td li[data-format="bullet"]::before,
th li[data-format="bullet"]::before {
  content: "•" !important;
  position: absolute !important;
  left: -15pt !important; /* Closer to text in table cells */
  top: 0 !important;
  width: 10pt !important;
  text-align: left !important;
  font-weight: normal !important;
  display: inline-block !important;
  color: inherit !important;
}

table li.docx-bullet-level-1,
td li.docx-bullet-level-1,
th li.docx-bullet-level-1 {
  margin-left: 18pt !important; /* Relative to table cell, not document */
}
`;

  return css;
}

module.exports = {
  generateDOCXNumberingStyles,
  generateEnhancedListStyles,
  generateBulletListStyles,
  generateCustomBulletStyles,
  getBulletCharForLevel,
};
