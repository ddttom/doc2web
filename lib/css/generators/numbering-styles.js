// File: lib/css/generators/numbering-styles.js
// Generate CSS for DOCX numbering and bullet lists

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
  return css;
}

/**
 * Generate CSS for bullet lists
 */
function generateBulletListStyles(styleInfo) {
  let css = "\n/* Bullet List Styles */\n";
  
  // Base bullet list styling with enhanced specificity
  css += `
/* Bullet list containers - add proper indentation */
ul.docx-bullet-list {
  list-style-type: none !important;
  padding-left: 0 !important;
  margin: 0 0 0 20pt !important;
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
  padding-left: 20pt !important;
  margin-left: 0 !important;
  margin-bottom: 0.5em !important;
  line-height: 1.4 !important;
  list-style: none !important;
  text-indent: 0 !important;
}

/* High specificity selectors for bullet characters */
ul.docx-bullet-list li.docx-list-item::before,
ul.docx-bullet-list li[data-format="bullet"]::before,
ul.docx-bullet-list li[class*="docx-bullet-"]::before,
ul.docx-bullet-list li[class*="docx-list-item"]::before,
li.docx-list-item[data-format="bullet"]::before,
li[data-format="bullet"].docx-list-item::before {
  content: "•" !important;
  position: absolute !important;
  left: 0 !important;
  top: 0 !important;
  width: 15pt !important;
  text-align: left !important;
  font-weight: normal !important;
  display: inline-block !important;
  color: inherit !important;
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
}

li.docx-list-item::before,
li[class*="docx-bullet-"]::before,
li[class*="docx-list-item"]::before,
li[data-format="bullet"]::before {
  content: "•" !important;
  position: absolute !important;
  left: 0 !important;
  top: 0 !important;
  width: 15pt !important;
  text-align: left !important;
  font-weight: normal !important;
  display: inline-block !important;
}
`;

  // Generate level-specific bullet styles
  for (let level = 0; level <= 8; level++) {
    const indent = level * 18; // 0.25 inch per level
    const bulletChar = getBulletCharForLevel(level);
    
    css += `
li.docx-bullet-level-${level} {
  margin-left: ${indent}pt;
}

li.docx-bullet-level-${level}::before {
  content: "${bulletChar}";
}
`;
  }
  
  return css;
}

/**
 * Get bullet character for specific level
 */
function getBulletCharForLevel(level) {
  const bullets = ["•", "○", "■", "□", "▪", "▫", "►", "⇒"];
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
  
  return css;
}

module.exports = {
  generateDOCXNumberingStyles,
  generateEnhancedListStyles,
  generateBulletListStyles,
  generateCustomBulletStyles,
  getBulletCharForLevel,
};
