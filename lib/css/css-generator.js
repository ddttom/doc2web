// lib/css/css-generator.js - CSS generation functions
const { convertTwipToPt, convertBorderSizeToPt, getBorderTypeValue } = require('../utils/unit-converter');
const { getCSSCounterContent } = require('../parsers/numbering-parser');

/**
 * Generate CSS from extracted style information with better TOC and list handling
 * Creates a comprehensive CSS stylesheet based on the extracted DOCX styles
 * 
 * @param {Object} styleInfo - Style information
 * @returns {string} - CSS stylesheet
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

    // Add paragraph styles
    Object.entries(styleInfo.styles.paragraph || {}).forEach(([id, style]) => {
      const className = `docx-p-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`;
      
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
  ${style.indentation?.left ? `margin-left: ${convertTwipToPt(style.indentation.left)}pt;` : ''}
  ${style.indentation?.right ? `margin-right: ${convertTwipToPt(style.indentation.right)}pt;` : ''}
  ${style.indentation?.firstLine ? `text-indent: ${convertTwipToPt(style.indentation.firstLine)}pt;` : ''}
  ${style.indentation?.hanging ? `text-indent: -${convertTwipToPt(style.indentation.hanging)}pt;` : ''}
}
`;
    });

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
      const tocHeaderStyle = styleInfo.tocStyles.tocHeadingStyle;
      const tocEntryStyles = styleInfo.tocStyles.tocEntryStyles;
      const leaderStyle = styleInfo.tocStyles.leaderStyle;
      
      // Main TOC container
      css += `
/* Table of Contents Styles */
.docx-toc {
  margin: 1em 0 2em 0;
  width: 100%;
  padding: 0;
}

.docx-toc-heading {
  ${tocHeaderStyle.fontFamily ? `font-family: "${tocHeaderStyle.fontFamily}", sans-serif;` : ''}
  ${tocHeaderStyle.fontSize ? `font-size: ${tocHeaderStyle.fontSize};` : 'font-size: 14pt;'}
  font-weight: bold;
  margin-bottom: 12pt;
  text-align: ${tocHeaderStyle.alignment || 'center'};
}
`;

      // TOC entries with leader lines - Enhanced to match Word's formatting
      css += `
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
`;

      // Specific styles for different TOC levels
      tocEntryStyles.forEach(style => {
        const level = style.level || 1;
        const leftIndent = style.indentation?.left ? 
                         convertTwipToPt(style.indentation.left) : 
                         (level - 1) * 20;
        
        css += `
/* TOC Level ${level} */
.docx-toc-level-${level} {
  ${style.fontFamily ? `font-family: "${style.fontFamily}", sans-serif;` : ''}
  ${style.fontSize ? `font-size: ${style.fontSize};` : ''}
  ${level === 1 ? 'font-weight: bold;' : ''}
  ${level === 3 ? 'font-style: italic;' : ''}
  margin-left: ${leftIndent}pt;
}
`;
      });
    }

    // Enhanced list styling based on extracted numbering definitions
    if (styleInfo.numberingDefs && Object.keys(styleInfo.numberingDefs.abstractNums).length > 0) {
      // Basic list containers
      css += `
/* Enhanced List and Numbering Styles */
ol.docx-numbered-list,
ol.docx-alpha-list,
ul.docx-bulleted-list {
  list-style-type: none;
  padding-left: 0;
  margin-top: 0.5em;
  margin-bottom: 0.5em;
  counter-reset: list-level-1;
}

ol.docx-numbered-list > li,
ol.docx-alpha-list > li,
ul.docx-bulleted-list > li {
  position: relative;
  margin-bottom: 0.4em;
  min-height: 1.2em;
}

/* Counter resets for nested lists */
ol.docx-numbered-list ol,
ol.docx-alpha-list ol {
  counter-reset: list-level-2;
}

ol.docx-numbered-list ol ol,
ol.docx-alpha-list ol ol {
  counter-reset: list-level-3;
}
`;

      // Generate specific styles for each numbering definition
      const abstractNums = styleInfo.numberingDefs.abstractNums;
      
      Object.entries(abstractNums).forEach(([id, abstractNum]) => {
        Object.entries(abstractNum.levels).forEach(([level, levelDef]) => {
          const levelNum = parseInt(level, 10);
          const className = `docx-num-${id}-${level}`;
          
          // Calculate indentation
          const leftIndent = levelDef.indentation?.left ? 
                           convertTwipToPt(levelDef.indentation.left) : 
                           levelNum * 36; // Default 36pt per level
                           
          const hangingIndent = levelDef.indentation?.hanging ? 
                              convertTwipToPt(levelDef.indentation.hanging) : 
                              24; // Default 24pt
          
          // Create CSS counter reset and list item styles
          css += `
/* Numbering style for abstract num ${id}, level ${level} */
li.${className} {
  margin-left: ${leftIndent}pt;
  padding-left: 0;
  counter-increment: list-level-${levelNum + 1};
}

li.${className}::before {
  position: absolute;
  left: 0;
  top: 0;
  width: ${hangingIndent}pt;
  text-align: ${levelDef.alignment || 'left'};
  margin-left: ${leftIndent - hangingIndent}pt;
  ${levelDef.runProps?.bold ? 'font-weight: bold;' : ''}
  ${levelDef.runProps?.italic ? 'font-style: italic;' : ''}
  ${levelDef.runProps?.fontSize ? `font-size: ${levelDef.runProps.fontSize};` : ''}
  ${levelDef.runProps?.color ? `color: ${levelDef.runProps.color};` : ''}
  content: "${getCSSCounterContent(levelDef)}";
}
`;
        });
      });
      
      // Add specific list type styles for common formats
      css += `
/* Special list formatting for different formats */
.docx-decimal-list > li::before {
  content: counter(list-level-1) ".";
}

.docx-decimal-list > li > ol > li::before {
  content: counter(list-level-1) "." counter(list-level-2) ".";
}

.docx-alpha-list > li::before {
  content: counter(list-level-1, lower-alpha) ".";
}

.docx-alpha-list > li > ol > li::before {
  content: counter(list-level-1, lower-alpha) "." counter(list-level-2) ".";
}

.docx-roman-list > li::before {
  content: counter(list-level-1, lower-roman) ".";
}

.docx-bulleted-list > li::before {
  content: "•";
}

/* Special paragraph types within lists */
.docx-list-special-paragraph {
  margin-left: 2.5em;
  margin-top: 0.5em;
  margin-bottom: 0.5em;
}

/* Generic special paragraph styles - based on structural patterns */
.docx-special-pattern {
  font-style: italic;
  color: #444;
  margin-left: 2.5em;
  margin-top: 0.5em;
  margin-bottom: 0.7em;
  border-left: 2px solid #eee;
  padding-left: 0.7em;
}

/* Style for paragraphs matching the "word_for_word" pattern */
.docx-word_for_word {
  font-style: italic;
  color: #444;
  margin-left: 2.5em;
  margin-top: 0.4em;
  margin-bottom: 0.6em;
}

/* Style for paragraphs matching the "word_comma" pattern */
.docx-word_comma {
  font-weight: normal;
  margin-bottom: 0.5em;
}

/* Style for paragraphs matching the "word_parenthesis" pattern */
.docx-word_parenthesis {
  font-weight: bold;
  margin-bottom: 0.7em;
}
`;
    }

    // Add utility styles
    css += `
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

/* Special paragraph styles */
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

/* Heading numbering */
.heading-number {
  display: inline-block;
  margin-right: 5px;
  font-weight: bold;
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

    // Add accessibility styles
    css += generateAccessibilityStyles(styleInfo);
    
    // Add track changes styles
    css += generateTrackChangesStyles(styleInfo);

  } catch (error) {
    console.error('Error generating CSS:', error);
    
    // Provide fallback CSS
    css = `
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

  return css;
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
  generateTrackChangesStyles
};
