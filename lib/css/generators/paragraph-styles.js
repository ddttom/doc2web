// File: lib/css/generators/paragraph-styles.js
// Paragraph style generation

const { convertTwipToPt } = require("../../utils/unit-converter");
const { getFontFamily } = require("./base-styles");

/**
 * Generate paragraph styles from style information
 */
function generateParagraphStyles(styleInfo) {
  let css = "\n/* Paragraph Styles - Enhanced for Word-like indentation */\n";
  Object.entries(styleInfo.styles?.paragraph || {}).forEach(([id, style]) => {
    const className = `docx-p-${id
      .replace(/[^a-zA-Z0-9-_]/g, "-")
      .toLowerCase()}`;
    let marginLeft = 0,
      marginRight = 0,
      paddingLeft = 0,
      textIndent = 0;
    let marginTop = style.spacing?.before
      ? convertTwipToPt(style.spacing.before)
      : null;
    let marginBottom = style.spacing?.after
      ? convertTwipToPt(style.spacing.after)
      : null;
    let lineHeight = style.spacing?.line
      ? parseInt(style.spacing.line, 10) / 240
      : null; // line value is in 240ths of a line

    // Enhanced indentation handling to match Word behavior exactly
    if (style.indentation) {
      // Left margin - this is the paragraph's left edge position
      marginLeft = style.indentation.left
        ? convertTwipToPt(style.indentation.left)
        : 0;

      // Right margin - this is the paragraph's right edge position
      marginRight = style.indentation.right
        ? convertTwipToPt(style.indentation.right)
        : 0;

      // Handle hanging indent (negative first line indent)
      if (style.indentation.hanging) {
        const hangingIndent = convertTwipToPt(style.indentation.hanging);
        // In Word, hanging indent creates space for numbering/bullets
        paddingLeft = hangingIndent;
        textIndent = -hangingIndent; // First line pulls back
      }
      // Handle positive first line indent
      else if (style.indentation.firstLine) {
        textIndent = convertTwipToPt(style.indentation.firstLine);
        // No padding needed for positive first line indent
        paddingLeft = 0;
      }

      // Handle special case where both left and start are specified (Word 2010+ format)
      if (style.indentation.start !== undefined) {
        marginLeft = convertTwipToPt(style.indentation.start);
      }
      if (style.indentation.end !== undefined) {
        marginRight = convertTwipToPt(style.indentation.end);
      }
    }
    css += generateParagraphStyleRule(className, style, styleInfo, {
      marginLeft,
      marginRight,
      paddingLeft,
      textIndent,
      marginTop,
      marginBottom,
      lineHeight
    });

    // If this is a heading style, also generate heading-specific CSS
    if (style.isHeading) {
      const headingLevels = ["1", "2", "3", "4", "5", "6"];
      headingLevels.forEach((level) => {
        // Check if this style matches heading level
        if (
          (style.name &&
            style.name.toLowerCase().includes(`heading ${level}`)) ||
          (style.id && style.id.toLowerCase().includes(`heading${level}`))
        ) {
          css += `
/* Apply paragraph indentation to heading level ${level} */
h${level}.${className} {
  ${marginLeft !== 0 ? `margin-left: ${marginLeft}pt;` : ""}
  ${marginRight !== 0 ? `margin-right: ${marginRight}pt;` : ""}
  ${paddingLeft !== 0 ? `padding-left: ${paddingLeft}pt;` : ""}
  ${textIndent !== 0 ? `text-indent: ${textIndent}pt;` : ""}
  ${paddingLeft > 0 || textIndent !== 0 ? "position: relative;" : ""}
  /* Word-like heading behavior with proper text wrapping */
  word-wrap: break-word;
  overflow-wrap: break-word;
  hyphens: auto;
}
`;
        }
      });
    }
  });
  return css;
}

module.exports = {
  generateParagraphStyles,
  generateGlobalParagraphDataAttributeStyles,
};

/**
 * Generate comprehensive paragraph style rule
 */
function generateParagraphStyleRule(className, style, styleInfo, computedValues) {
  const { marginLeft, marginRight, paddingLeft, textIndent, marginTop, marginBottom, lineHeight } = computedValues;
  
  let css = `\n.${className} {\n`;
  
  // Font properties
  if (style.font) {
    css += `  font-family: "${getFontFamily(style, styleInfo)}", sans-serif;\n`;
  }
  if (style.fontSize) css += `  font-size: ${style.fontSize};\n`;
  if (lineHeight) css += `  line-height: ${lineHeight};\n`;
  
  // Basic formatting
  if (style.bold) css += `  font-weight: bold;\n`;
  if (style.italic) css += `  font-style: italic;\n`;
  if (style.color) css += `  color: ${style.color};\n`;
  
  // Alignment
  if (style.alignment) css += `  text-align: ${style.alignment};\n`;
  
  // Spacing
  if (marginTop !== null) css += `  margin-top: ${marginTop}pt;\n`;
  if (marginBottom !== null) css += `  margin-bottom: ${marginBottom}pt;\n`;
  if (marginLeft !== 0) css += `  margin-left: ${marginLeft}pt;\n`;
  if (marginRight !== 0) css += `  margin-right: ${marginRight}pt;\n`;
  if (paddingLeft !== 0) css += `  padding-left: ${paddingLeft}pt;\n`;
  if (textIndent !== 0) css += `  text-indent: ${textIndent}pt;\n`;
  
  // Enhanced spacing properties
  if (style.spacing) {
    if (style.spacing.lineRule) {
      css += `  /* Line spacing rule: ${style.spacing.lineRule} */\n`;
    }
    if (style.spacing.beforeAuto) css += `  /* Auto spacing before */\n`;
    if (style.spacing.afterAuto) css += `  /* Auto spacing after */\n`;
  }
  
  // Shading properties
  if (style.shading) {
    if (style.shading.fill) css += `  background-color: #${style.shading.fill};\n`;
    if (style.shading.pattern && style.shading.pattern !== 'clear') {
      css += `  /* Shading pattern: ${style.shading.pattern} */\n`;
    }
  }
  
  // Page properties
  if (style.pageProperties) {
    if (style.pageProperties.pageBreakBefore) css += `  page-break-before: always;\n`;
    if (style.pageProperties.keepNext) css += `  page-break-after: avoid;\n`;
    if (style.pageProperties.keepLines) css += `  page-break-inside: avoid;\n`;
  }
  
  // Widow control
  if (style.widowControl !== undefined) {
    css += `  orphans: ${style.widowControl ? 2 : 1};\n`;
    css += `  widows: ${style.widowControl ? 2 : 1};\n`;
  }
  
  // Frame properties
  if (style.frameProperties) {
    const frame = style.frameProperties;
    css += `  /* Frame properties */\n`;
    if (frame.w !== undefined) css += `  width: ${convertTwipToPt(frame.w)}pt;\n`;
    if (frame.h !== undefined) css += `  height: ${convertTwipToPt(frame.h)}pt;\n`;
    if (frame.x !== undefined || frame.y !== undefined) {
      css += `  position: absolute;\n`;
      if (frame.x !== undefined) css += `  left: ${convertTwipToPt(frame.x)}pt;\n`;
      if (frame.y !== undefined) css += `  top: ${convertTwipToPt(frame.y)}pt;\n`;
    }
    if (frame.anchor) css += `  /* Anchor: ${frame.anchor} */\n`;
  }
  
  // Box model and behavior
  css += `  box-sizing: border-box;\n`;
  
  if (paddingLeft > 0 || textIndent !== 0 || (style.numbering && style.numbering.hasNumbering)) {
    css += `  position: relative;\n`;
  }
  
  // Word-like paragraph behavior
  css += `  /* Word-like paragraph behavior */\n`;
  css += `  word-wrap: break-word;\n`;
  css += `  overflow-wrap: break-word;\n`;
  css += `  hyphens: auto;\n`;
  
  css += `}\n`;
  
  // Add data attribute styles for this paragraph
  css += generateParagraphDataAttributeStyles(className, style);
  
  return css;
}

/**
 * Generate data attribute styles for paragraph formatting
 */
function generateParagraphDataAttributeStyles(className, style) {
  let css = '';
  
  // Generate data attribute selectors for enhanced formatting
  css += `
/* Enhanced Paragraph Formatting - Data Attribute Styles for ${className} */

/* Alignment styles */
.${className}[data-jc="left"] { text-align: left; }
.${className}[data-jc="center"] { text-align: center; }
.${className}[data-jc="right"] { text-align: right; }
.${className}[data-jc="both"] { text-align: justify; }
.${className}[data-jc="distribute"] { text-align: justify; text-align-last: justify; }

/* Indentation styles */
.${className}[data-ind-left] { /* Left indentation applied via data attributes */ }
.${className}[data-ind-right] { /* Right indentation applied via data attributes */ }
.${className}[data-ind-first-line] { /* First line indentation applied via data attributes */ }
.${className}[data-ind-hanging] { /* Hanging indentation applied via data attributes */ }
.${className}[data-ind-start] { /* Start indentation applied via data attributes */ }
.${className}[data-ind-end] { /* End indentation applied via data attributes */ }

/* Spacing styles */
.${className}[data-spacing-before] { /* Before spacing applied via data attributes */ }
.${className}[data-spacing-after] { /* After spacing applied via data attributes */ }
.${className}[data-spacing-line] { /* Line spacing applied via data attributes */ }
.${className}[data-spacing-line-rule] { /* Line spacing rule applied via data attributes */ }
.${className}[data-spacing-before-auto] { /* Auto before spacing applied via data attributes */ }
.${className}[data-spacing-after-auto] { /* Auto after spacing applied via data attributes */ }

/* Shading styles */
.${className}[data-shd-fill] { /* Shading fill applied via data attributes */ }
.${className}[data-shd-color] { /* Shading color applied via data attributes */ }
.${className}[data-shd-pattern] { /* Shading pattern applied via data attributes */ }

/* Page properties styles */
.${className}[data-page-break-before] { page-break-before: always; }
.${className}[data-keep-next] { page-break-after: avoid; }
.${className}[data-keep-lines] { page-break-inside: avoid; }
.${className}[data-widow-control] { orphans: 2; widows: 2; }

/* Frame properties styles */
.${className}[data-frame-w] { /* Frame width applied via data attributes */ }
.${className}[data-frame-h] { /* Frame height applied via data attributes */ }
.${className}[data-frame-x] { /* Frame X position applied via data attributes */ }
.${className}[data-frame-y] { /* Frame Y position applied via data attributes */ }
.${className}[data-frame-anchor] { /* Frame anchor applied via data attributes */ }
`;
  
  return css;
}

/**
 * Generate global data attribute styles for paragraph formatting
 */
function generateGlobalParagraphDataAttributeStyles() {
  return `
/* Global Enhanced Paragraph Formatting - Data Attribute Styles */

/* Enhanced Hanging Indentation for Numbered Paragraphs */
p[data-num-id] {
  text-indent: -18pt !important;  /* Negative indent to pull number left */
  padding-left: 18pt !important;  /* Positive padding to indent content */
  margin-left: 0pt !important;    /* Reset margin to prevent compounding */
  position: relative;
}

/* Special handling for numbered headings with span structure */
h1[data-num-id],
h2[data-num-id],
h3[data-num-id],
h4[data-num-id],
h5[data-num-id],
h6[data-num-id] {
  position: relative;
  display: flex;
  align-items: flex-start;
  flex-wrap: wrap;
}

/* Level-based indentation for numbered headings */
h1[data-num-level="0"], h2[data-num-level="0"], h3[data-num-level="0"], 
h4[data-num-level="0"], h5[data-num-level="0"], h6[data-num-level="0"] { 
  margin-left: 0pt !important; 
}
h1[data-num-level="1"], h2[data-num-level="1"], h3[data-num-level="1"], 
h4[data-num-level="1"], h5[data-num-level="1"], h6[data-num-level="1"] { 
  margin-left: 18pt !important; 
}
h1[data-num-level="2"], h2[data-num-level="2"], h3[data-num-level="2"], 
h4[data-num-level="2"], h5[data-num-level="2"], h6[data-num-level="2"] { 
  margin-left: 36pt !important; 
}
h1[data-num-level="3"], h2[data-num-level="3"], h3[data-num-level="3"], 
h4[data-num-level="3"], h5[data-num-level="3"], h6[data-num-level="3"] { 
  margin-left: 54pt !important; 
}
h1[data-num-level="4"], h2[data-num-level="4"], h3[data-num-level="4"], 
h4[data-num-level="4"], h5[data-num-level="4"], h6[data-num-level="4"] { 
  margin-left: 72pt !important; 
}
h1[data-num-level="5"], h2[data-num-level="5"], h3[data-num-level="5"], 
h4[data-num-level="5"], h5[data-num-level="5"], h6[data-num-level="5"] { 
  margin-left: 90pt !important; 
}
h1[data-num-level="6"], h2[data-num-level="6"], h3[data-num-level="6"], 
h4[data-num-level="6"], h5[data-num-level="6"], h6[data-num-level="6"] { 
  margin-left: 108pt !important; 
}
h1[data-num-level="7"], h2[data-num-level="7"], h3[data-num-level="7"], 
h4[data-num-level="7"], h5[data-num-level="7"], h6[data-num-level="7"] { 
  margin-left: 126pt !important; 
}
h1[data-num-level="8"], h2[data-num-level="8"], h3[data-num-level="8"], 
h4[data-num-level="8"], h5[data-num-level="8"], h6[data-num-level="8"] { 
  margin-left: 144pt !important; 
}

/* Position the heading number span */
h1[data-num-id] .heading-number,
h2[data-num-id] .heading-number,
h3[data-num-id] .heading-number,
h4[data-num-id] .heading-number,
h5[data-num-id] .heading-number,
h6[data-num-id] .heading-number {
  position: absolute;
  left: 0;
  top: 0;
  width: 18pt;
  text-align: left;
  white-space: nowrap;
}

/* Position the heading content span */
h1[data-num-id] .heading-content,
h2[data-num-id] .heading-content,
h3[data-num-id] .heading-content,
h4[data-num-id] .heading-content,
h5[data-num-id] .heading-content,
h6[data-num-id] .heading-content {
  margin-left: 18pt;
  flex: 1;
  min-width: 0; /* Allow text to wrap */
}

/* Specific hanging indent adjustments for different numbering formats */
p[data-format="lowerLetter"] {
  text-indent: -24pt !important;  /* Slightly more space for letters */
  padding-left: 24pt !important;
  margin-left: 0pt !important;    /* Reset margin to prevent compounding */
}

p[data-format="upperLetter"] {
  text-indent: -24pt !important;
  padding-left: 24pt !important;
  margin-left: 0pt !important;
}

p[data-format="lowerRoman"] {
  text-indent: -30pt !important;  /* More space for Roman numerals */
  padding-left: 30pt !important;
  margin-left: 0pt !important;
}

p[data-format="upperRoman"] {
  text-indent: -30pt !important;
  padding-left: 30pt !important;
  margin-left: 0pt !important;
}

/* Alignment styles */
p[data-jc="left"] { text-align: left; }
p[data-jc="center"] { text-align: center; }
p[data-jc="right"] { text-align: right; }
p[data-jc="both"] { text-align: justify; }
p[data-jc="distribute"] { text-align: justify; text-align-last: justify; }

/* Indentation utility classes */
.docx-ind-left { /* Left indentation utility */ }
.docx-ind-right { /* Right indentation utility */ }
.docx-ind-first-line { /* First line indentation utility */ }
.docx-ind-hanging { /* Hanging indentation utility */ }
.docx-ind-start { /* Start indentation utility */ }
.docx-ind-end { /* End indentation utility */ }

/* Spacing utility classes */
.docx-spacing-before { /* Before spacing utility */ }
.docx-spacing-after { /* After spacing utility */ }
.docx-spacing-line { /* Line spacing utility */ }
.docx-spacing-line-rule { /* Line spacing rule utility */ }
.docx-spacing-before-auto { /* Auto before spacing utility */ }
.docx-spacing-after-auto { /* Auto after spacing utility */ }

/* Shading utility classes */
.docx-shd-fill { /* Shading fill utility */ }
.docx-shd-color { /* Shading color utility */ }
.docx-shd-pattern { /* Shading pattern utility */ }

/* Page properties utility classes */
.docx-page-break-before { page-break-before: always; }
.docx-keep-next { page-break-after: avoid; }
.docx-keep-lines { page-break-inside: avoid; }
.docx-widow-control { orphans: 2; widows: 2; }

/* Frame utility classes */
.docx-frame { position: relative; }
.docx-frame[data-frame-w] { /* Frame width utility */ }
.docx-frame[data-frame-h] { /* Frame height utility */ }
.docx-frame[data-frame-x] { /* Frame X position utility */ }
.docx-frame[data-frame-y] { /* Frame Y position utility */ }
.docx-frame[data-frame-anchor] { /* Frame anchor utility */ }

/* Alignment utility classes */
.docx-align-left { text-align: left; }
.docx-align-center { text-align: center; }
.docx-align-right { text-align: right; }
.docx-align-justify { text-align: justify; }
.docx-align-distribute { text-align: justify; text-align-last: justify; }
`;
}
