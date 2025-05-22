// lib/css/css-generator.js - Enhanced CSS generation with proper DOCX indentation support
const { convertTwipToPt, convertBorderSizeToPt, getBorderTypeValue } = require('../utils/unit-converter');
const { getCSSCounterContent } = require('../parsers/numbering-parser');

/**
 * Generate CSS from extracted style information with enhanced DOCX numbering support
 * Creates a comprehensive CSS stylesheet based on the extracted DOCX styles and numbering context
 * 
 * @param {Object} styleInfo - Style information with numbering context
 * @returns {string} - CSS stylesheet with DOCX-derived numbering
 */
function generateCssFromStyleInfo(styleInfo) {
  let css = '';
  
  try {
    // Add document defaults
    css += `
/* Document defaults */
body {
  font-family: "${styleInfo.theme.fonts.minor || 'Calibri'}", sans-serif;
  font-size: ${styleInfo.documentDefaults.character.fontSize || '11pt'};
  line-height: 1.15;
  margin: 20px;
  padding: 0;
}
`;

    // Add enhanced paragraph styles with proper indentation
    css += generateParagraphStyles(styleInfo);

    // Add character styles
    Object.entries(styleInfo.styles.character || {}).forEach(([id, style]) => {
      const className = `docx-c-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`;
      
      css += `
.${className} {
  ${style.font ? `font-family: "${getFontFamily(style, styleInfo)}", sans-serif;` : ''}
  ${style.fontSize ? `font-size: ${style.fontSize};` : ''}
  ${style.bold ? 'font-weight: bold;' : ''}
  ${style.italic ? 'font-style: italic;' : ''}
  ${style.color ? `color: ${style.color};` : ''}
  ${style.underline ? `text-decoration: ${style.underline.type === 'none' ? 'none' : 'underline'};` : ''}
}
`;
    });

    // Add table styles
    Object.entries(styleInfo.styles.table || {}).forEach(([id, style]) => {
      const className = `docx-t-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`;
      
      css += `
.${className} {
  border-collapse: collapse;
  width: 100%;
}
.${className} td, .${className} th {
  padding: 5pt;
  ${getBorderStyle(style, 'top')}
  ${getBorderStyle(style, 'bottom')}
  ${getBorderStyle(style, 'left')}
  ${getBorderStyle(style, 'right')}
}
`;
    });

    // Enhanced TOC styles based on extracted information
    if (styleInfo.tocStyles) {
      css += generateTOCStyles(styleInfo.tocStyles);
    }

    // ENHANCED: Generate CSS for DOCX-derived numbering definitions with proper indentation
    if (styleInfo.numberingDefs && Object.keys(styleInfo.numberingDefs.abstractNums).length > 0) {
      css += generateDOCXNumberingStyles(styleInfo.numberingDefs, styleInfo.numberingContext);
    }

    // Enhanced list styling based on extracted numbering definitions
    css += generateEnhancedListStyles(styleInfo);

    // Add utility styles
    css += generateUtilityStyles(styleInfo);

    // Add accessibility styles
    css += generateAccessibilityStyles(styleInfo);
    
    // Add track changes styles
    css += generateTrackChangesStyles(styleInfo);

  } catch (error) {
    console.error('Error generating CSS:', error);
    
    // Provide fallback CSS
    css = generateFallbackCSS(styleInfo);
  }

  return css;
}

/**
 * Generate enhanced paragraph styles with proper indentation handling
 */
function generateParagraphStyles(styleInfo) {
  let css = `
/* Paragraph styles with DOCX indentation */
`;
  
  Object.entries(styleInfo.styles.paragraph || {}).forEach(([id, style]) => {
    const className = `docx-p-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`;
    
    // Calculate indentation values
    let marginLeft = 0;
    let paddingLeft = 0;
    let textIndent = 0;
    
    if (style.indentation) {
      // Left indentation
      if (style.indentation.left) {
        marginLeft = convertTwipToPt(style.indentation.left);
      }
      
      // Hanging indent (negative first line indent)
      if (style.indentation.hanging) {
        paddingLeft = convertTwipToPt(style.indentation.hanging);
        textIndent = -convertTwipToPt(style.indentation.hanging);
      }
      
      // First line indent (positive)
      if (style.indentation.firstLine) {
        textIndent = convertTwipToPt(style.indentation.firstLine);
      }
      
      // Start indentation (for numbered items)
      if (style.indentation.start) {
        marginLeft = Math.max(marginLeft, convertTwipToPt(style.indentation.start));
      }
    }
    
    css += `
.${className} {
  ${style.font ? `font-family: "${getFontFamily(style, styleInfo)}", sans-serif;` : ''}
  ${style.fontSize ? `font-size: ${style.fontSize};` : ''}
  ${style.bold ? 'font-weight: bold;' : ''}
  ${style.italic ? 'font-style: italic;' : ''}
  ${style.color ? `color: ${style.color};` : ''}
  ${style.alignment ? `text-align: ${style.alignment};` : ''}
  ${style.spacing?.before ? `margin-top: ${convertTwipToPt(style.spacing.before)}pt;` : ''}
  ${style.spacing?.after ? `margin-bottom: ${convertTwipToPt(style.spacing.after)}pt;` : ''}
  ${marginLeft > 0 ? `margin-left: ${marginLeft}pt;` : ''}
  ${paddingLeft > 0 ? `padding-left: ${paddingLeft}pt;` : ''}
  ${style.indentation?.right ? `margin-right: ${convertTwipToPt(style.indentation.right)}pt;` : ''}
  ${textIndent !== 0 ? `text-indent: ${textIndent}pt;` : ''}
  ${style.indentation?.hanging ? 'position: relative;' : ''}
}
`;
    
    // Add numbering-specific styles if this style has numbering
    if (style.numbering && style.numbering.hasNumbering) {
      css += generateNumberingStylesForParagraph(className, style, styleInfo);
    }
  });
  
  return css;
}

/**
 * Generate numbering styles for paragraph with numbering
 * @param {string} className - CSS class name
 * @param {Object} style - Paragraph style
 * @param {Object} styleInfo - Complete style information
 * @returns {string} - CSS for numbered paragraph
 */
function generateNumberingStylesForParagraph(className, style, styleInfo) {
  const numId = style.numbering.id;
  const numLevel = parseInt(style.numbering.level, 10);
  
  // Find the numbering definition
  const numDef = styleInfo.numberingDefs?.nums?.[numId];
  if (!numDef) return '';
  
  const abstractNum = styleInfo.numberingDefs?.abstractNums?.[numDef.abstractNumId];
  if (!abstractNum) return '';
  
  const levelDef = abstractNum.levels?.[numLevel];
  if (!levelDef) return '';
  
  let css = '';
  
  // Add specific styles for this numbered paragraph
  css += `
/* Numbering styles for ${className} (numId: ${numId}, level: ${numLevel}) */
.${className}[data-num-id="${numId}"][data-num-level="${numLevel}"] {
  ${levelDef.indentation?.left ? `margin-left: ${levelDef.indentation.left}pt;` : ''}
  ${levelDef.indentation?.hanging ? `padding-left: ${levelDef.indentation.hanging}pt;` : ''}
  ${levelDef.indentation?.hanging ? `text-indent: -${levelDef.indentation.hanging}pt;` : ''}
  ${levelDef.indentation?.firstLine ? `text-indent: ${levelDef.indentation.firstLine}pt;` : ''}
  counter-increment: docx-counter-${numId}-${numLevel};
  position: relative;
}

.${className}[data-num-id="${numId}"][data-num-level="${numLevel}"]::before {
  content: ${getCSSCounterContent(levelDef)};
  position: absolute;
  left: 0;
  ${levelDef.runProps?.bold ? 'font-weight: bold;' : ''}
  ${levelDef.runProps?.italic ? 'font-style: italic;' : ''}
  ${levelDef.runProps?.fontSize ? `font-size: ${levelDef.runProps.fontSize};` : ''}
  ${levelDef.runProps?.color ? `color: ${levelDef.runProps.color};` : ''}
  ${levelDef.indentation?.hanging ? `width: ${levelDef.indentation.hanging}pt;` : ''}
  text-align: ${levelDef.alignment || 'left'};
}
`;

  return css;
}

/**
 * Generate CSS for DOCX-derived numbering definitions with proper indentation
 * @param {Object} numberingDefs - Numbering definitions from DOCX
 * @param {Array} numberingContext - Resolved numbering context
 * @returns {string} - CSS for DOCX numbering
 */
function generateDOCXNumberingStyles(numberingDefs, numberingContext = []) {
  let css = `
/* DOCX-derived numbering styles with proper indentation */
`;

  // Generate counter reset rules for each numbering definition
  Object.entries(numberingDefs.nums).forEach(([numId, numDef]) => {
    const abstractNum = numberingDefs.abstractNums[numDef.abstractNumId];
    if (!abstractNum) return;
    
    // Generate counter resets for all levels
    const counterNames = Object.keys(abstractNum.levels).map(level => 
      `docx-counter-${numId}-${level}`
    ).join(' ');
    
    css += `
/* Counter resets for numbering ${numId} */
[data-num-id="${numId}"] {
  counter-reset: ${counterNames};
}
`;

    // Generate styles for each level with proper indentation
    Object.entries(abstractNum.levels).forEach(([level, levelDef]) => {
      const levelNum = parseInt(level, 10);
      
      // Calculate total indentation for this level
      let marginLeft = 0;
      let paddingLeft = 0;
      let textIndent = 0;
      
      if (levelDef.indentation) {
        // Left margin is the base indentation
        marginLeft = levelDef.indentation.left || 0;
        
        // If there's a hanging indent, add padding and negative text-indent
        if (levelDef.indentation.hanging) {
          paddingLeft = levelDef.indentation.hanging;
          textIndent = -levelDef.indentation.hanging;
        }
        
        // If there's a first line indent, use it for text-indent
        if (levelDef.indentation.firstLine) {
          textIndent = levelDef.indentation.firstLine;
        }
      }
      
      css += `
/* Level ${level} styles for numbering ${numId} */
[data-num-id="${numId}"][data-num-level="${level}"] {
  ${marginLeft > 0 ? `margin-left: ${marginLeft}pt !important;` : ''}
  ${paddingLeft > 0 ? `padding-left: ${paddingLeft}pt;` : ''}
  ${textIndent !== 0 ? `text-indent: ${textIndent}pt;` : ''}
  counter-increment: docx-counter-${numId}-${level};
  position: relative;
}

[data-num-id="${numId}"][data-num-level="${level}"]::before {
  content: ${getCSSCounterContent(levelDef)};
  position: absolute;
  left: 0;
  top: 0;
  ${paddingLeft > 0 ? `width: ${paddingLeft}pt;` : ''}
  text-align: ${levelDef.alignment || 'left'};
  ${levelDef.runProps?.bold ? 'font-weight: bold;' : ''}
  ${levelDef.runProps?.italic ? 'font-style: italic;' : ''}
  ${levelDef.runProps?.fontSize ? `font-size: ${levelDef.runProps.fontSize};` : ''}
  ${levelDef.runProps?.color ? `color: ${levelDef.runProps.color};` : ''}
  ${levelDef.runProps?.font?.ascii ? `font-family: "${levelDef.runProps.font.ascii}", sans-serif;` : ''}
}
`;

      // Reset deeper level counters when this level increments
      const deeperLevels = Object.keys(abstractNum.levels)
        .map(l => parseInt(l, 10))
        .filter(l => l > levelNum)
        .map(l => `docx-counter-${numId}-${l}`)
        .join(' ');
      
      if (deeperLevels) {
        css += `
[data-num-id="${numId}"][data-num-level="${level}"] {
  counter-reset: ${deeperLevels};
}
`;
      }
    });
  });

  // Add styles for maintaining indentation context
  css += `
/* Maintain indentation for content following numbered items */
.docx-indent-inherit {
  margin-left: inherit;
}

/* Specific indentation levels for non-numbered content */
.docx-indent-level-1 {
  margin-left: 72pt;
}

.docx-indent-level-2 {
  margin-left: 108pt;
}

.docx-indent-level-3 {
  margin-left: 144pt;
}

.docx-indent-level-4 {
  margin-left: 180pt;
}

/* Special handling for content after numbered items */
.docx-continues-numbering {
  margin-left: inherit;
  margin-top: 0.5em;
}

/* Styles for paragraphs that follow numbered items */
p[data-follows-numbered="true"] {
  margin-left: inherit;
}
`;

  return css;
}

/**
 * Generate enhanced list styles based on DOCX numbering
 * @param {Object} styleInfo - Style information
 * @returns {string} - Enhanced list CSS
 */
function generateEnhancedListStyles(styleInfo) {
  let css = `
/* Enhanced List Styles based on DOCX numbering */
ol.docx-numbered-list,
ol.docx-alpha-list,
ol.docx-roman-list,
ul.docx-bulleted-list {
  list-style-type: none;
  padding-left: 0;
  margin: 0.5em 0;
}

ol.docx-numbered-list > li,
ol.docx-alpha-list > li,
ol.docx-roman-list > li,
ul.docx-bulleted-list > li {
  position: relative;
  margin-bottom: 0.4em;
  min-height: 1.2em;
}

/* Nested list indentation */
ol.docx-numbered-list ol,
ol.docx-alpha-list ol,
ol.docx-roman-list ol,
ul.docx-bulleted-list ul {
  margin-left: 36pt;
  margin-top: 0.2em;
}

/* Maintain proper indentation inheritance */
li[data-numbering-id] {
  position: relative;
}

li[data-numbering-id][data-level="0"] {
  margin-left: 0;
}

li[data-numbering-id][data-level="1"] {
  margin-left: 36pt;
}

li[data-numbering-id][data-level="2"] {
  margin-left: 72pt;
}

li[data-numbering-id][data-level="3"] {
  margin-left: 108pt;
}

/* Special paragraph styles within lists */
.docx-list-special-paragraph {
  margin-left: inherit;
  margin-top: 0.5em;
  margin-bottom: 0.5em;
}
`;

  return css;
}

/**
 * Generate TOC styles
 * @param {Object} tocStyles - TOC style information
 * @returns {string} - TOC CSS
 */
function generateTOCStyles(tocStyles) {
  const leaderStyle = tocStyles.leaderStyle || {};
  
  return `
/* Table of Contents Styles */
.docx-toc {
  margin: 1em 0 2em 0;
  width: 100%;
  padding: 0;
}

.docx-toc-heading {
  ${tocStyles.tocHeadingStyle?.fontFamily ? `font-family: "${tocStyles.tocHeadingStyle.fontFamily}", sans-serif;` : ''}
  ${tocStyles.tocHeadingStyle?.fontSize ? `font-size: ${tocStyles.tocHeadingStyle.fontSize};` : 'font-size: 14pt;'}
  font-weight: bold;
  margin-bottom: 12pt;
  text-align: ${tocStyles.tocHeadingStyle?.alignment || 'center'};
}

/* Enhanced TOC entry styles */
.docx-toc-entry {
  display: flex;
  flex-wrap: nowrap;
  align-items: baseline;
  position: relative;
  width: 100%;
  margin-bottom: 4pt;
  line-height: 1.2;
  overflow: hidden;
  white-space: nowrap;
}

.docx-toc-text {
  flex-grow: 0;
  flex-shrink: 1;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  padding-right: 3px;
}

.docx-toc-dots {
  flex-grow: 1;
  position: relative;
  margin: 0 2px;
  height: 1em;
  min-width: 20px;
  border-bottom: none;
}

/* Create leader dots using CSS background */
.docx-toc-dots::after {
  content: "";
  position: absolute;
  bottom: 0.3em;
  left: 0;
  right: 0;
  height: 1px;
  background-image: ${leaderStyle.character === '.' ? 
    'radial-gradient(circle, #000 1px, transparent 0)' : 
    leaderStyle.character === '-' ? 
    'linear-gradient(to right, #000 2px, transparent 2px)' : 
    'linear-gradient(to right, #000 3px, transparent 0)'};
  background-position: bottom;
  background-size: ${leaderStyle.character === '.' ? '4px 1px' : '6px 1px'};
  background-repeat: repeat-x;
}

.docx-toc-pagenum {
  flex-shrink: 0;
  text-align: right;
  padding-left: 4px;
  font-weight: normal;
}

/* Specific styles for different TOC levels */
${tocStyles.tocEntryStyles?.map((style, index) => {
  const level = style.level || (index + 1);
  const leftIndent = style.indentation?.left ? 
                   convertTwipToPt(style.indentation.left) : 
                   (level - 1) * 20;
  
  return `
.docx-toc-level-${level} {
  ${style.fontFamily ? `font-family: "${style.fontFamily}", sans-serif;` : ''}
  ${style.fontSize ? `font-size: ${style.fontSize};` : ''}
  ${level === 1 ? 'font-weight: bold;' : ''}
  ${level === 3 ? 'font-style: italic;' : ''}
  margin-left: ${leftIndent}pt;
}`;
}).join('') || ''}
`;
}

/**
 * Generate utility styles
 * @param {Object} styleInfo - Style information
 * @returns {string} - Utility CSS
 */
function generateUtilityStyles(styleInfo) {
  return `
/* Utility styles */
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }
.docx-tab { display: inline-block; width: ${convertTwipToPt(styleInfo.settings?.defaultTabStop || '720')}pt; }
.docx-rtl { direction: rtl; unicode-bidi: embed; }
.docx-image { max-width: 100%; height: auto; display: block; margin: 10px 0; }
.docx-table-default { width: 100%; border-collapse: collapse; margin: 10px 0; }
.docx-table-default td, .docx-table-default th { border: 1px solid #ddd; padding: 5pt; }
.docx-placeholder { 
  font-weight: bold; 
  text-align: center; 
  padding: 15px; 
  margin: 20px 0; 
  background-color: #f0f0f0; 
  border: 1px dashed #999; 
  border-radius: 5px; 
}

/* Enhanced heading styles with numbering support */
.docx-heading1, .docx-heading2, .docx-heading3, .docx-heading4, .docx-heading5, .docx-heading6 {
  font-family: "${styleInfo.theme.fonts?.major || 'Calibri Light'}", sans-serif;
  color: #2F5496;
  margin-top: 1.2em;
  margin-bottom: 0.6em;
}

.docx-heading1 { font-size: 16pt; }
.docx-heading2 { font-size: 14pt; }
.docx-heading3 { font-size: 13pt; font-style: italic; }
.docx-heading4 { font-size: 12pt; }
.docx-heading5 { font-size: 11pt; font-style: italic; }
.docx-heading6 { font-size: 11pt; }

/* Enhanced heading numbering from DOCX */
.heading-number {
  display: inline-block;
  margin-right: 0.3em;
  font-weight: bold;
  color: inherit;
}

/* Structural pattern styles */
.docx-structural-pattern {
  margin-bottom: 0.7em;
}

.docx-word_colon {
  font-weight: bold;
  margin-bottom: 0.5em;
}

.docx-word_comma {
  font-weight: normal;
  margin-bottom: 0.5em;
}

.docx-word_parenthesis {
  font-weight: bold;
  margin-bottom: 0.7em;
}

/* Document footer */
.docx-footer {
  text-align: center;
  margin-top: 2em;
  font-size: 0.9em;
  color: #666;
}

/* Improved table styles */
.table-responsive {
  overflow-x: auto;
  margin: 1em 0;
}

/* Print styles */
@media print {
  body {
    margin: 1cm;
  }
  
  .docx-toc-dots::after {
    background-image: ${styleInfo.tocStyles?.leaderStyle?.character === '.' ? 
      'radial-gradient(circle, #000 0.7px, transparent 0)' : 
      styleInfo.tocStyles?.leaderStyle?.character === '-' ? 
      'linear-gradient(to right, #000 2px, transparent 1px)' : 
      'linear-gradient(to right, #000 3px, transparent 0)'};
    background-size: ${styleInfo.tocStyles?.leaderStyle?.character === '.' ? '3px 1px' : '5px 1px'};
  }
  
  .table-responsive {
    overflow-x: visible;
  }
}
`;
}

/**
 * Generate accessibility styles
 * Creates CSS for accessibility features
 * 
 * @param {Object} styleInfo - Style information
 * @returns {string} - CSS styles for accessibility
 */
function generateAccessibilityStyles(styleInfo) {
  return `
/* Accessibility Styles */
.skip-link {
  position: absolute;
  top: -40px;
  left: 0;
  padding: 8px;
  background-color: #f8f9fa;
  color: #005a9c;
  z-index: 100;
  transition: top 0.3s;
}

.skip-link:focus {
  top: 0;
  outline: 2px solid #4d90fe;
  outline-offset: 2px;
}

.keyboard-focusable:focus {
  outline: 2px solid #4d90fe;
  outline-offset: 2px;
}

@media (prefers-reduced-motion: reduce) {
  * {
    transition: none !important;
    animation: none !important;
  }
}

/* High contrast mode */
@media (prefers-contrast: more) {
  body {
    color: #000000;
    background-color: #ffffff;
  }
  
  a {
    color: #0000EE;
    text-decoration: underline;
  }
  
  a:visited {
    color: #551A8B;
  }
  
  th, td {
    border: 1px solid #000000;
  }
  
  .docx-toc-dots::after {
    background-image: linear-gradient(to right, #000 1px, transparent 0);
    background-position: bottom;
    background-size: 4px 1px;
    background-repeat: repeat-x;
  }
  
  .docx-insertion {
    background-color: #BBDEFB !important;
    outline: 1px solid #2196F3 !important;
  }
  
  .docx-deletion {
    background-color: #FFCDD2 !important;
    outline: 1px solid #F44336 !important;
    text-decoration: line-through;
  }
}

/* Screen reader only text */
.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border-width: 0;
}

/* Focus visible utility */
*:focus-visible {
  outline: 2px solid #4d90fe;
  outline-offset: 2px;
}

/* Table accessibility enhancements */
table {
  border-collapse: collapse;
  width: 100%;
  margin: 1em 0;
}

caption {
  text-align: left;
  font-weight: bold;
  margin-bottom: 0.5em;
}

th {
  text-align: left;
  background-color: #f8f9fa;
}

table.striped tr:nth-child(even) {
  background-color: #f8f9fa;
}

/* Figure and image accessibility */
figure {
  margin: 1em 0;
}

figcaption {
  font-style: italic;
  margin-top: 0.5em;
}
`;
}

/**
 * Generate track changes styles
 * Creates CSS for track changes visualization
 * 
 * @param {Object} styleInfo - Style information
 * @returns {string} - CSS styles for track changes
 */
function generateTrackChangesStyles(styleInfo) {
  return `
/* Track Changes Styles */
.docx-track-changes-legend {
  margin: 1em 0;
  padding: 0.5em;
  border: 1px solid #BDBDBD;
  border-radius: 4px;
  background-color: #F5F5F5;
}

.docx-track-changes-toggle {
  padding: 0.25em 0.5em;
  margin-right: 1em;
  border: 1px solid #BDBDBD;
  border-radius: 3px;
  cursor: pointer;
}

.docx-track-changes-toggle:hover {
  background-color: #EEEEEE;
}

.docx-track-changes-toggle:focus {
  outline: 2px solid #4d90fe;
  outline-offset: 1px;
}

/* Track Changes Mode: Show */
.docx-track-changes-show .docx-insertion {
  background-color: #E6F4FF;
  position: relative;
}

.docx-track-changes-show .docx-insertion::after {
  content: attr(data-author) " (" attr(data-date) ")";
  display: none;
  position: absolute;
  top: 100%;
  left: 0;
  background-color: #FFFFFF;
  border: 1px solid #BDBDBD;
  border-radius: 3px;
  padding: 0.25em 0.5em;
  font-size: 0.8em;
  z-index: 10;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  white-space: nowrap;
}

.docx-track-changes-show .docx-insertion:hover::after {
  display: block;
}

.docx-track-changes-show .docx-deletion {
  background-color: #FFEBEE;
  text-decoration: line-through;
  position: relative;
}

.docx-track-changes-show .docx-deletion::after {
  content: attr(data-author) " (" attr(data-date) ")";
  display: none;
  position: absolute;
  top: 100%;
  left: 0;
  background-color: #FFFFFF;
  border: 1px solid #BDBDBD;
  border-radius: 3px;
  padding: 0.25em 0.5em;
  font-size: 0.8em;
  z-index: 10;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  white-space: nowrap;
}

.docx-track-changes-show .docx-deletion:hover::after {
  display: block;
}

.docx-track-changes-show .docx-move-from {
  border-bottom: 2px dashed #9575CD;
  position: relative;
}

.docx-track-changes-show .docx-move-to {
  border-bottom: 2px solid #9575CD;
  position: relative;
}

.docx-track-changes-show .docx-move-from::after,
.docx-track-changes-show .docx-move-to::after {
  content: "Moved by " attr(data-author) " (" attr(data-date) ")";
  display: none;
  position: absolute;
  top: 100%;
  left: 0;
  background-color: #FFFFFF;
  border: 1px solid #BDBDBD;
  border-radius: 3px;
  padding: 0.25em 0.5em;
  font-size: 0.8em;
  z-index: 10;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  white-space: nowrap;
}

.docx-track-changes-show .docx-move-from:hover::after,
.docx-track-changes-show .docx-move-to:hover::after {
  display: block;
}

.docx-track-changes-show .docx-formatting-change {
  border-bottom: 1px dotted #FFA000;
  position: relative;
}

.docx-track-changes-show .docx-formatting-change::after {
  content: attr(data-formatting) " by " attr(data-author) " (" attr(data-date) ")";
  display: none;
  position: absolute;
  top: 100%;
  left: 0;
  background-color: #FFFFFF;
  border: 1px solid #BDBDBD;
  border-radius: 3px;
  padding: 0.25em 0.5em;
  font-size: 0.8em;
  z-index: 10;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  white-space: nowrap;
}

.docx-track-changes-show .docx-formatting-change:hover::after {
  display: block;
}

/* Track Changes Mode: Hide */
.docx-track-changes-hide .docx-insertion,
.docx-track-changes-hide .docx-deletion,
.docx-track-changes-hide .docx-move-from,
.docx-track-changes-hide .docx-move-to,
.docx-track-changes-hide .docx-formatting-change {
  background-color: transparent;
  text-decoration: none;
  border-bottom: none;
}

.docx-track-changes-hide .docx-deleted-content {
  display: none;
}

/* Deleted content area */
.docx-deleted-content {
  margin: 1em 0;
  padding: 0.5em;
  border: 1px solid #FFCDD2;
  border-radius: 4px;
  background-color: #FFEBEE;
}

.docx-deleted-content::before {
  content: "Deleted Content";
  display: block;
  font-weight: bold;
  margin-bottom: 0.5em;
}
`;
}

/**
 * Generate fallback CSS in case of errors
 * @param {Object} styleInfo - Style information
 * @returns {string} - Fallback CSS
 */
function generateFallbackCSS(styleInfo) {
  return `
body { font-family: Calibri, sans-serif; font-size: 11pt; line-height: 1.15; margin: 20px; }
h1, h2, h3, h4, h5, h6 { font-family: "Calibri Light", sans-serif; color: #2F5496; }
h1 { font-size: 16pt; }
h2 { font-size: 13pt; }
h3 { font-size: 12pt; }
p { margin: 10pt 0; }
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }
.docx-rtl { direction: rtl; }
.docx-image { max-width: 100%; height: auto; display: block; margin: 10px 0; }
table { width: 100%; border-collapse: collapse; margin: 10px 0; }
td, th { border: 1px solid #ddd; padding: 5pt; }

/* Basic TOC styles */
.docx-toc-entry {
  display: flex;
  flex-wrap: nowrap;
  align-items: baseline;
  width: 100%;
  margin-bottom: 4pt;
  white-space: nowrap;
  overflow: hidden;
}
.docx-toc-text {
  white-space: nowrap;
  overflow: hidden;
  flex-grow: 0;
}
.docx-toc-dots {
  flex-grow: 1;
  margin: 0 4pt;
  height: 1em;
  position: relative;
}
.docx-toc-dots::after {
  content: "";
  position: absolute;
  bottom: 0.3em;
  left: 0;
  right: 0;
  height: 1px;
  background-image: radial-gradient(circle, #000 1px, transparent 0);
  background-position: bottom;
  background-size: 4px 1px;
  background-repeat: repeat-x;
}
.docx-toc-pagenum {
  flex-shrink: 0;
  text-align: right;
}

/* Basic list styles */
ol.docx-numbered-list { list-style-type: none; padding-left: 0; }
ol.docx-numbered-list > li { position: relative; padding-left: 2.5em; margin-bottom: 0.5em; }
ol.docx-numbered-list > li::before { position: absolute; left: 0; content: attr(data-prefix) "."; font-weight: bold; }
ol.docx-alpha-list { list-style-type: none; padding-left: 0; }
ol.docx-alpha-list > li { position: relative; padding-left: 2.5em; margin-bottom: 0.5em; }
ol.docx-alpha-list > li::before { position: absolute; left: 0; content: attr(data-prefix) "."; font-weight: normal; }

/* Basic indentation levels */
.docx-indent-level-1 { margin-left: 72pt; }
.docx-indent-level-2 { margin-left: 108pt; }
.docx-indent-level-3 { margin-left: 144pt; }

/* Accessibility styles */
.skip-link {
  position: absolute;
  top: -40px;
  left: 0;
  padding: 8px;
  background-color: #f8f9fa;
  color: #005a9c;
  z-index: 100;
  transition: top 0.3s;
}

.skip-link:focus {
  top: 0;
  outline: 2px solid #4d90fe;
}

/* Track changes styles */
.docx-insertion { background-color: #E6F4FF; }
.docx-deletion { background-color: #FFEBEE; text-decoration: line-through; }
`;
}

/**
 * Get font family based on style and theme
 * Determines the appropriate font family to use based on style and theme settings
 * 
 * @param {Object} style - Style information
 * @param {Object} styleInfo - Style information including theme
 * @returns {string} - Font family
 */
function getFontFamily(style, styleInfo) {
  if (!style.font) {
    return styleInfo.theme.fonts.minor;
  }
  
  // Use ASCII font if available, fall back to theme
  return style.font.ascii || style.font.hAnsi || styleInfo.theme.fonts.minor;
}

/**
 * Get border style
 * Creates CSS border property based on style information
 * 
 * @param {Object} style - Style information
 * @param {string} side - Border side (top, bottom, left, right)
 * @returns {string} - CSS border property
 */
function getBorderStyle(style, side) {
  if (!style.borders || !style.borders[side]) {
    return '';
  }
  
  const border = style.borders[side];
  if (!border || !border.value || border.value === 'nil' || border.value === 'none') {
    return `border-${side}: none;`;
  }
  
  const width = border.size ? convertBorderSizeToPt(border.size) : 1;
  const color = border.color ? `#${border.color}` : 'black';
  const borderStyle = getBorderTypeValue(border.value);
  
  return `border-${side}: ${width}pt ${borderStyle} ${color};`;
}

module.exports = {
  generateCssFromStyleInfo,
  getFontFamily,
  getBorderStyle,
  generateAccessibilityStyles,
  generateTrackChangesStyles,
  generateDOCXNumberingStyles,
  generateEnhancedListStyles
};