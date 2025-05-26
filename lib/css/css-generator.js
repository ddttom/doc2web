// File: ddttom/doc2web/doc2web-c1a38debdb5cb80b84caa0caa610b6ba6a166799/lib/css/css-generator.js

const {
  convertTwipToPt,
  convertBorderSizeToPt,
  getBorderTypeValue,
} = require("../utils/unit-converter");
const {
  getCSSCounterContent,
  getCSSCounterFormat,
} = require("../parsers/numbering-parser");

/**
 * Generate CSS from extracted style information.
 */
function generateCssFromStyleInfo(styleInfo) {
  let css = "";
  try {
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

    css += `
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
    css += generateParagraphStyles(styleInfo);
    css += generateCharacterStyles(styleInfo);
    css += generateTableStyles(styleInfo);
    if (styleInfo.tocStyles) {
      css += generateTOCStyles(styleInfo);
    } // Pass full styleInfo
    if (
      styleInfo.numberingDefs &&
      Object.keys(styleInfo.numberingDefs.abstractNums || {}).length > 0
    ) {
      css += generateDOCXNumberingStyles(styleInfo.numberingDefs, styleInfo); // Pass styleInfo for context
    }
    css += generateEnhancedListStyles(styleInfo);
    css += generateUtilityStyles(styleInfo);
    css += generateHeaderStyles(styleInfo);
    css += generateAccessibilityStyles(styleInfo);
    css += generateTrackChangesStyles(styleInfo);
  } catch (error) {
    console.error("Error generating CSS:", error.message, error.stack);
    css = generateFallbackCSS(styleInfo);
  }
  return css;
}
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
     css += `
.${className} {
  ${
    style.font
      ? `font-family: "${getFontFamily(style, styleInfo)}", sans-serif;`
      : ""
  }
  ${style.fontSize ? `font-size: ${style.fontSize};` : ""}
  ${lineHeight ? `line-height: ${lineHeight};` : ""}
  ${style.bold ? "font-weight: bold;" : ""}
  ${style.italic ? "font-style: italic;" : ""}
  ${style.color ? `color: ${style.color};` : ""}
  ${style.alignment ? `text-align: ${style.alignment};` : ""}
  ${marginTop !== null ? `margin-top: ${marginTop}pt;` : ""}
  ${marginBottom !== null ? `margin-bottom: ${marginBottom}pt;` : ""}
  ${marginLeft !== 0 ? `margin-left: ${marginLeft}pt;` : ""}
  ${marginRight !== 0 ? `margin-right: ${marginRight}pt;` : ""}
  ${paddingLeft !== 0 ? `padding-left: ${paddingLeft}pt;` : ""}
  ${textIndent !== 0 ? `text-indent: ${textIndent}pt;` : ""}
  box-sizing: border-box;
  ${
    paddingLeft > 0 ||
    textIndent !== 0 ||
    (style.numbering && style.numbering.hasNumbering)
      ? "position: relative;"
      : ""
  }
  /* Word-like paragraph behavior */
  word-wrap: break-word;
  overflow-wrap: break-word;
  hyphens: auto;
}
`;

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
}
`;
         }
       });
     }
   });
   return css;
 }
function generateCharacterStyles(styleInfo) {
  let css = "\n/* Character Styles */\n";
  Object.entries(styleInfo.styles?.character || {}).forEach(([id, style]) => {
    const className = `docx-c-${id
      .replace(/[^a-zA-Z0-9-_]/g, "-")
      .toLowerCase()}`;
    css += `
.${className} {
  ${
    style.font
      ? `font-family: "${getFontFamily(style, styleInfo)}", sans-serif;`
      : ""
  }
  ${style.fontSize ? `font-size: ${style.fontSize};` : ""}
  ${style.bold ? "font-weight: bold;" : ""}
  ${style.italic ? "font-style: italic;" : ""}
  ${style.color ? `color: ${style.color};` : ""}
  ${
    style.underline && style.underline.type !== "none"
      ? `text-decoration: underline;`
      : ""
  }
  ${
    style.highlight ? `background-color: ${style.highlight};` : ""
  } /* Added highlight */
}
`;
  });
  return css;
}

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

// Enhanced generateDOCXNumberingStyles function for lib/css/css-generator.js
function generateDOCXNumberingStyles(numberingDefs, styleInfo) {
  let css = `\n/* DOCX Numbering Styles */\n`;
  
  // Add global box model fixes for all numbered elements
  css += `
/* Global box model fixes for numbered elements */
[data-abstract-num][data-num-level] {
  box-sizing: border-box;
  overflow-wrap: break-word;
  word-break: normal;
  hyphens: auto;
}

/* Fix for section IDs with Roman numerals */
[id^="section-"][data-format="upperRoman"],
[id^="section-"][data-format="lowerRoman"] {
  padding-left: 2em; /* Extra space for Roman numerals */
}
`;
  
  // Process each abstract numbering definition
  Object.values(numberingDefs.abstractNums || {}).forEach((abstractNum) => {
    // Set up counter resets for all levels in this abstract numbering
    const allLevelCountersForAbstractNum = Object.keys(abstractNum.levels || {})
      .map((levelIndex) => `docx-counter-${abstractNum.id}-${levelIndex} 0`)
      .join(" ");
    
    if (allLevelCountersForAbstractNum) {
      css += `
/* Counter resets for abstract numbering ID ${abstractNum.id} */
body { 
  counter-reset: ${allLevelCountersForAbstractNum};
}

/* More targeted counter reset for specific containers with this abstract numbering */
ol[data-abstract-num="${abstractNum.id}"], 
ul[data-abstract-num="${abstractNum.id}"], 
div[data-abstract-num="${abstractNum.id}"],
nav[data-abstract-num="${abstractNum.id}"] {
  counter-reset: ${allLevelCountersForAbstractNum};
}
`;
    }

    // Process each level in the abstract numbering
    Object.entries(abstractNum.levels || {}).forEach(([levelIndexStr, levelDef]) => {
      const levelIndex = parseInt(levelIndexStr, 10);
      const itemSelector = `[data-abstract-num="${abstractNum.id}"][data-num-level="${levelIndex}"]`;
      
      // Calculate indentation and spacing based on DOCX definition with enhanced precision
      let pIndentLeft = 0;
      let pTextIndent = 0;
      let numRegionWidth = 24; // Base width for numbering region
      
      // Extract precise indentation from DOCX
      if (levelDef.indentation) {
        // Left indentation - where the paragraph starts
        if (levelDef.indentation.left !== undefined) {
          pIndentLeft = convertTwipToPt(levelDef.indentation.left);
        } else {
          // Fallback to level-based indentation
          pIndentLeft = levelIndex * 36; // 0.5 inch per level
        }
        
        // Handle hanging indent (most common for numbered lists)
        if (levelDef.indentation.hanging !== undefined) {
          const hangingIndent = convertTwipToPt(levelDef.indentation.hanging);
          numRegionWidth = hangingIndent;
          pTextIndent = 0; // Text starts at the left margin
          // The number will be positioned in the hanging indent space
        } 
        // Handle first line indent
        else if (levelDef.indentation.firstLine !== undefined) {
          pTextIndent = convertTwipToPt(levelDef.indentation.firstLine);
          numRegionWidth = Math.max(24, Math.abs(pTextIndent)); // Ensure enough space
        }
      } else {
        // Default indentation based on level
        pIndentLeft = levelIndex * 36; // 0.5 inch per level
        numRegionWidth = 24; // Default numbering width
      }
      
      // Adjust numbering width based on format
      if (levelDef.format === "upperRoman" || levelDef.format === "lowerRoman") {
        numRegionWidth = Math.max(numRegionWidth, 48); // Wider for Roman numerals
      } else if (levelDef.format === "bullet") {
        numRegionWidth = Math.max(numRegionWidth, 18); // Bullets need less space
      } else {
        numRegionWidth = Math.max(numRegionWidth, 36); // Numbers need moderate space
      }
      
      // Create CSS for the numbered element with precise Word-like positioning
      css += `
/* Styling for level ${levelIndex} of abstract numbering ${abstractNum.id} */
${itemSelector} {
  display: block;
  position: relative;
  margin-left: ${pIndentLeft}pt;
  padding-left: ${numRegionWidth}pt;
  text-indent: ${pTextIndent}pt;
  counter-increment: docx-counter-${abstractNum.id}-${levelIndex};
  box-sizing: border-box;
  overflow-wrap: break-word;
  word-wrap: break-word;
  hyphens: auto;
  /* Match Word's paragraph spacing */
  margin-top: 0;
  margin-bottom: 0;
}

/* Section ID based styling for easy targeting */
${itemSelector}[id^="section-"] {
  scroll-margin-top: 20px; /* For smooth scrolling to sections */
}

/* Number or bullet display using ::before pseudo-element with precise positioning */
${itemSelector}::before {
  content: ${getCSSCounterContent(levelDef, abstractNum.id, numberingDefs)};
  position: absolute;
  left: ${levelDef.indentation?.hanging ? -numRegionWidth : 0}pt;
  top: 0;
  width: ${numRegionWidth}pt;
  text-align: ${levelDef.alignment || "left"};
  box-sizing: border-box;
  white-space: nowrap;
  overflow: hidden;
  /* Spacing after number based on DOCX suffix */
  padding-right: ${
    levelDef.suffix === "tab" ? "0.5em" : 
    levelDef.suffix === "space" ? "0.25em" : 
    "0.125em"
  };
  /* Apply DOCX formatting to numbering */
  ${levelDef.runProps?.bold ? "font-weight: bold;" : ""}
  ${levelDef.runProps?.italic ? "font-style: italic;" : ""}
  ${levelDef.runProps?.fontSize ? `font-size: ${levelDef.runProps.fontSize};` : ""}
  ${levelDef.runProps?.color ? `color: ${levelDef.runProps.color};` : ""}
  font-family: ${levelDef.runProps?.font?.ascii ? `"${levelDef.runProps.font.ascii}", ` : ""}${levelDef.runProps?.font?.hAnsi ? `"${levelDef.runProps.font.hAnsi}", ` : ""}sans-serif;
}
`;
      
      // Reset deeper level counters when this level increments
      const deeperLevelsToReset = Object.keys(abstractNum.levels)
        .map((l) => parseInt(l, 10))
        .filter((l) => l > levelIndex)
        .map((l) => `docx-counter-${abstractNum.id}-${l} 0`)
        .join(" ");
      
      if (deeperLevelsToReset) {
        css += `
/* Reset deeper level counters when level ${levelIndex} increments */
${itemSelector} {
  counter-reset: ${deeperLevelsToReset};
}
`;
      }
    });
  });
  
  return css;
}

function generateEnhancedListStyles(styleInfo) {
  let css = `\n/* List Container Styles */\n`;
  css += `
ol[data-abstract-num], ul[data-abstract-num] {
  list-style-type: none; 
  padding-left: 0; 
  margin-top: 0.5em; 
  margin-bottom: 0.5em;
}
li[data-abstract-num] { /* Targeting <li> if they are the numbered items */
  display: block; 
  /* Indentation and numbering are handled by the [data-abstract-num][data-num-level] styles */
}
`;
  return css;
}

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

// Fixed and enhanced TOC styles function
function generateTOCStyles(styleInfo) {
  const tocStyles = styleInfo.tocStyles || {};
  const leaderStyle = tocStyles.leaderStyle || {
    character: ".",
    position: 468, // default 6.5 inches in points
    spacesBetween: 3,
  };

  let css = `
/* Table of Contents Styles */
.docx-toc {
  margin: 1em 0 2em 0;
  width: 100%;
  padding: 0;
  /* Force single-column layout */
  column-count: 1 !important;
  -webkit-column-count: 1 !important;
  -moz-column-count: 1 !important;
  columns: 1 !important; /* Additional property */
  column-width: auto !important; /* Prevent auto width calculation */
}

/* Ensure parent containers don't interfere */
.docx-toc, .docx-toc * {
  box-sizing: border-box;
}

.docx-toc-heading {
  font-family: "${
    tocStyles.tocHeadingStyle?.fontFamily ||
    styleInfo.theme?.fonts?.major ||
    "Calibri Light"
  }", sans-serif;
  font-size: ${tocStyles.tocHeadingStyle?.fontSize || "16pt"}; 
  font-weight: bold;
  margin-bottom: 1em;
  text-align: ${tocStyles.tocHeadingStyle?.alignment || "left"};
  column-span: all;
  -webkit-column-span: all;
  color: ${
    tocStyles.tocHeadingStyle?.color ||
    styleInfo.theme?.colors?.dk1 ||
    "#2F5496"
  };
}

/* Structure TOC entries as a flex container with proper alignment */
.docx-toc-entry {
  display: flex !important;
  flex-wrap: nowrap !important;
  align-items: baseline !important;
  width: 100% !important;
  max-width: 100%; /* Prevent overflow */
  margin-bottom: 0.5em !important;
  line-height: 1.4 !important;
  position: relative !important;
  overflow: hidden !important;
  page-break-inside: avoid; /* Keep entries together on page breaks */
}

/* TOC text part - flex-grow: 0 to prevent it from expanding */
.docx-toc-text {
  flex: 0 0 auto !important;
  white-space: nowrap !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  padding-right: 0.5em !important;
  max-width: 80% !important;
}

/* For clickable TOC entries */
.docx-toc-text a {
  text-decoration: none !important;
  color: inherit !important;
}

/* TOC dots - flex-grow: 1 to fill available space with dots */
.docx-toc-dots {
  flex: 1 1 auto !important;
  position: relative !important;
  height: 1.2em !important;
  margin: 0 0.5em !important;
  overflow: hidden !important;
}

/* Create dots using background image with proper spacing based on leaderStyle */
.docx-toc-dots::after {
  content: "";
  position: absolute;
  bottom: 0.5em; /* Position closer to text baseline */
  left: 0;
  right: 0;
  height: 1px; /* Use 1px for consistent rendering */
  background-image: radial-gradient(circle, currentColor 1px, transparent 1.5px);
  background-position: bottom;
  background-size: 6px 1px; /* Fixed size for consistency */
  background-repeat: repeat-x;
}

/* TOC page number - flex-grow: 0 to maintain size and right-align */
.docx-toc-pagenum {
  flex: 0 0 auto !important;
  text-align: right !important;
  white-space: nowrap !important;
  font-variant-numeric: tabular-nums !important;
}

/* Print-specific styles for TOC */
@media print {
  .docx-toc-dots::after {
    background-image: radial-gradient(circle, black 1px, transparent 1.5px);
  }
}

/* Force TOC entry paragraphs to use flex */
p.docx-toc-entry {
  display: flex !important;
  margin: 0.5em 0 !important; /* Override paragraph margins */
}

/* Browser Compatibility */
@supports not (background-image: radial-gradient(circle, currentColor 1px, transparent 1.5px)) {
  .docx-toc-dots::after {
    background-image: none;
    border-bottom: 1px dotted currentColor;
  }
}
`;

  // Add specific styles for each TOC level
  (tocStyles.tocEntryStyles || []).forEach((style, index) => {
    const level = style.level || index + 1;
    const baseIndent = (level - 1) * 1.5; // Base indent in em (e.g., 1.5em per level)
    const leftIndentPt = style.indentation?.left || 0; // Use points from DOCX if available

    css += `
.docx-toc-level-${level} {
  font-family: "${
    style.fontFamily || styleInfo.theme?.fonts?.minor || "Calibri"
  }", sans-serif;
  font-size: ${style.fontSize || "10pt"};
  font-weight: ${
    level === 1
      ? style.bold === false
        ? "normal"
        : "bold"
      : style.bold
      ? "bold"
      : "normal"
  };
  font-style: ${style.italic ? "italic" : "normal"};
  color: ${style.color || styleInfo.theme?.colors?.tx1 || "inherit"};
  padding-left: ${leftIndentPt > 0 ? leftIndentPt + "pt" : baseIndent + "em"};
}

/* Handle TOC entries that are numbered by the main numbering system */
.docx-toc-level-${level}[data-abstract-num] {
  padding-left: 0;
  margin-left: ${leftIndentPt > 0 ? leftIndentPt + "pt" : baseIndent + "em"};
}
`;
  });

  // Special handling for potential "Rationale" lines or similar entries
  // that might not have regular TOC styling but need consistent layout
  css += `
/* Special handling for TOC entries without explicit TOC styles */
.docx-toc-entry:not([class*="docx-toc-level-"]) {
  font-family: "${styleInfo.theme?.fonts?.minor || "Calibri"}", sans-serif;
  font-size: 10pt;
  padding-left: 0;
  margin-top: 0.5em;
  margin-bottom: 0.5em;
}

/* Force all TOC entries to use flex layout for consistent alignment */
p.docx-toc-entry {
  display: flex !important;
}

/* Ensure all TOC entries have dots and page numbers aligned consistently */
.docx-toc p, .docx-toc-entry {
  display: flex !important;
  flex-wrap: nowrap !important;
  align-items: baseline !important;
  width: 100% !important;
  margin-bottom: 0.5em !important;
  line-height: 1.4 !important;
  position: relative !important;
  overflow: hidden !important;
}

/* Special handling for TOC entries with numbering - adjust ::before positioning for flex layout */
.docx-toc-entry[data-abstract-num][data-num-level]::before {
  position: relative !important;
  left: 0 !important;
  margin-right: 0.5em !important;
  width: auto !important;
  display: inline !important;
}
`;

  return css;
}

function generateUtilityStyles(styleInfo) {
  // Calculate default paragraph margin based on DOCX introspection
  let defaultParagraphMargin = "0pt";
  
  // Method 1: Check document defaults for paragraph indentation
  if (styleInfo.documentDefaults?.paragraph?.indentation?.left) {
    defaultParagraphMargin = convertTwipToPt(styleInfo.documentDefaults.paragraph.indentation.left) + "pt";
  }
  // Method 2: Check the most common paragraph style indentation
  else if (styleInfo.styles?.paragraph) {
    const paragraphStyles = Object.values(styleInfo.styles.paragraph);
    const indentations = paragraphStyles
      .map(style => style.indentation?.left)
      .filter(indent => indent && parseInt(indent) > 0)
      .map(indent => convertTwipToPt(indent));
    
    if (indentations.length > 0) {
      // Use the most common indentation, or the first one if all are different
      const indentCounts = {};
      indentations.forEach(indent => {
        indentCounts[indent] = (indentCounts[indent] || 0) + 1;
      });
      
      const mostCommonIndent = Object.keys(indentCounts)
        .reduce((a, b) => indentCounts[a] > indentCounts[b] ? a : b);
      
      defaultParagraphMargin = mostCommonIndent + "pt";
    }
  }
  // Method 3: Check if numbered paragraphs have consistent indentation
  if (defaultParagraphMargin === "0pt" && styleInfo.numberingDefs?.abstractNums) {
    const abstractNums = Object.values(styleInfo.numberingDefs.abstractNums);
    const levelIndentations = [];
    
    abstractNums.forEach(abstractNum => {
      if (abstractNum.levels) {
        Object.values(abstractNum.levels).forEach(level => {
          if (level.indentation?.left) {
            levelIndentations.push(convertTwipToPt(level.indentation.left));
          }
        });
      }
    });
    
    if (levelIndentations.length > 0) {
      // Use the minimum indentation from numbered lists as the base paragraph margin
      const minIndent = Math.min(...levelIndentations);
      if (minIndent > 0) {
        defaultParagraphMargin = minIndent + "pt";
      }
    }
  }
  
  // Fallback: Use a reasonable default based on typical Word documents
  if (defaultParagraphMargin === "0pt") {
    defaultParagraphMargin = "0pt"; // Keep 0 if we can't determine from DOCX
  }

  return `
/* Utility Styles */
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }

/* Section navigation styling */
[id^="section-"]:target {
  background-color: #fffbcc;
  transition: background-color 0.3s ease;
}

/* Improved spacing for section IDs */
[id^="section-"] {
  margin-top: 0.5em;
  clear: both;
}

/* Fix for Roman numeral sections */
[id^="section-v"], 
[id^="section-vi"],
[id^="section-vii"],
[id^="section-viii"],
[id^="section-ix"],
[id^="section-x"] {
  padding-left: 0.5em;
  margin-left: 0.5em;
}

/* Enhanced styling for Roman numeral sections */
.roman-numeral-section,
.roman-numeral-heading,
[data-format="upperRoman"],
[data-format="lowerRoman"] {
  padding-left: 0.5em !important;
  margin-left: 0.5em !important;
  position: relative;
}

/* Ensure proper spacing for Roman numeral headings */
.roman-numeral-heading::before,
[data-format="upperRoman"]::before,
[data-format="lowerRoman"]::before {
  margin-right: 0.5em !important;
  white-space: nowrap !important;
}

/* Add spacing between heading number and content */
.heading-number {
  white-space: nowrap;
  display: inline;
  margin-right: 0;
}

/* Ensure heading content can wrap but stays connected to numbering */
.heading-content {
  display: inline;
  white-space: normal;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

/* Ensure headings allow full text display with proper wrapping */
h1:has(.heading-number), h2:has(.heading-number), h3:has(.heading-number),
h4:has(.heading-number), h5:has(.heading-number), h6:has(.heading-number) {
  white-space: normal;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

/* Prevent line break specifically between number and content start */
.heading-number + .heading-content {
  margin-left: 0;
}

/* Ensure proper spacing between numbering and content */
[data-num-id][data-abstract-num] {
  overflow-wrap: break-word;
  word-wrap: break-word;
  word-break: normal;
  hyphens: auto;
}

/* Default paragraph styling with DOCX-derived margins */
p {
  overflow-wrap: break-word;
  word-wrap: break-word;
  max-width: 100%;
  box-sizing: border-box;
  /* Apply consistent left margin to all paragraphs based on DOCX introspection */
  margin-left: ${defaultParagraphMargin};
}

/* Override for paragraphs that already have specific styling */
p[class*="docx-p-"] {
  /* Styled paragraphs use their own margin-left from generateParagraphStyles */
  margin-left: unset;
}

/* Override for TOC entries and other special paragraphs */
p.docx-toc-entry,
p[class*="docx-toc-"],
nav.docx-toc p {
  /* TOC entries have their own indentation logic */
  margin-left: 0;
}

/* Apply default page margins to TOC link text */
.docx-toc-text a {
  /* Inherit the default paragraph margin for consistent spacing */
  margin-left: ${defaultParagraphMargin};
  display: inline-block;
}

h1,h2,h3,h4,h5,h6 { 
  font-family:"${styleInfo.theme?.fonts?.major || "Calibri Light"}", Arial, sans-serif; 
  color: ${styleInfo.theme?.colors?.dk1 || "#2E74B5"}; 
  overflow-wrap: break-word;
  word-wrap: break-word;
  /* Apply default paragraph margin to all headings */
  margin-left: ${defaultParagraphMargin};
  /* Restore default browser margins for headings */
  margin-top: 0.83em;
  margin-bottom: 0.83em;
}

/* Specific margins for each heading level to match browser defaults */
h1 { 
  margin-top: 0.67em; 
  margin-bottom: 0.67em; 
  font-size: 2em; 
}
h2 { 
  margin-top: 0.83em; 
  margin-bottom: 0.83em; 
  font-size: 1.5em; 
}
h3 { 
  margin-top: 1em; 
  margin-bottom: 1em; 
  font-size: 1.17em; 
}
h4 { 
  margin-top: 1.33em; 
  margin-bottom: 1.33em; 
  font-size: 1em; 
}
h5 { 
  margin-top: 1.67em; 
  margin-bottom: 1.67em; 
  font-size: 0.83em; 
}
h6 { 
  margin-top: 2.33em; 
  margin-bottom: 2.33em; 
  font-size: 0.67em; 
}
/* ... more utility styles ... */
`;
}

function generateAccessibilityStyles(styleInfo) {
  return `
/* Accessibility Styles */
.skip-link { position: absolute; top: -100px; left: 0; padding: 10px; background-color: #f0f0f0; color: #333; z-index: 9999; transition: top 0.3s ease-in-out; text-decoration: none; border: 1px solid #ccc; border-top: none; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }
.skip-link:focus { top: 0; outline: 2px solid #005a9c; outline-offset: -2px; }
.keyboard-focusable:focus, *:focus-visible { outline: 2px dashed #005a9c !important; outline-offset: 2px !important; box-shadow: 0 0 0 2px rgba(0,90,156,0.3) !important; }
@media (prefers-reduced-motion: reduce) { * { transition-duration: 0.001ms !important; animation-duration: 0.001ms !important; animation-iteration-count: 1 !important; scroll-behavior: auto !important; } }
@media (prefers-contrast: more) {
  body { color: #000000 !important; background-color: #ffffff !important; }
  a { color: #0000EE !important; text-decoration: underline !important; } a:visited { color: #551A8B !important; }
  th, td { border: 1px solid #000000 !important; }
  .docx-toc-dots::after { background-image: linear-gradient(to right, #000 1px, transparent 0) !important; }
  .docx-insertion { background-color: #cce5ff !important; border: 1px solid #007bff !important; color: #004085 !important; }
  .docx-deletion { background-color: #f8d7da !important; border: 1px solid #dc3545 !important; color: #721c24 !important; text-decoration: line-through !important; }
}
.sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border-width: 0; }
table caption.sr-only { position: static; width: auto; height: auto; margin: 0; overflow: visible; clip: auto; white-space: normal; }
figure { margin: 1em 0; }
figcaption { font-style: italic; margin-top: 0.5em; }
`;
}

function generateTrackChangesStyles(styleInfo) {
  return `
/* Track Changes Styles */
.docx-track-changes-legend { margin: 1em 0; padding: 0.5em; border: 1px solid #BDBDBD; border-radius: 4px; background-color: #F5F5F5; font-size: 0.9em; }
.docx-track-changes-toggle { padding: 0.3em 0.6em; margin-right: 1em; border: 1px solid #BDBDBD; border-radius: 3px; cursor: pointer; background-color: #e9e9e9; }
.docx-track-changes-toggle:hover { background-color: #dcdcdc; }
.docx-track-changes-show .docx-insertion { background-color: #e6ffed; border-bottom: 1px solid #a2d5ab; position: relative; }
.docx-track-changes-show .docx-insertion:hover::after, .docx-track-changes-show .docx-deletion:hover::after { content: attr(data-author) " (" attr(data-date) ")"; display: block; position: absolute; bottom: 100%; left: 0; margin-bottom: 2px; background-color: #333; color: #fff; border-radius: 3px; padding: 0.25em 0.5em; font-size: 0.8em; z-index: 100; white-space: nowrap; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
.docx-track-changes-show .docx-deletion { background-color: #ffe6e6; border-bottom: 1px solid #f5b7b1; text-decoration: line-through; color: #d9534f; position: relative; }
.docx-track-changes-hide .docx-insertion, .docx-track-changes-hide .docx-deletion { background-color: transparent; text-decoration: none; border-bottom: none; color: inherit; }
.docx-track-changes-hide .docx-deleted-content { display: none; }
.docx-deleted-content { margin: 1em 0; padding: 0.5em; border: 1px dashed #FFCDD2; border-radius: 4px; background-color: #FFEBEE; color: #b71c1c; }
.docx-deleted-content::before { content: "Deleted Content Section"; display: block; font-weight: bold; margin-bottom: 0.5em; }
`;
}

function generateHeaderStyles(styleInfo) {
  return `
/* Document Header Styles */
.docx-document-header {
  margin-bottom: 2em;
  padding-bottom: 1em;
  border-bottom: 1px solid #e0e0e0;
  text-align: center;
  page-break-after: avoid;
}

.docx-header-paragraph {
  margin: 0.5em 0;
  line-height: 1.3;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

.docx-header-text {
  display: block;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

/* Default header styling for common header styles */
.docx-header-paragraph[data-style-id*="header"],
.docx-header-paragraph[data-style-id*="Header"] {
  font-size: 16pt;
  font-weight: bold;
  margin: 1em 0;
}

.docx-header-paragraph[data-style-id*="title"],
.docx-header-paragraph[data-style-id*="Title"] {
  font-size: 18pt;
  font-weight: bold;
  margin: 1em 0;
}

.docx-header-paragraph[data-style-id*="subtitle"],
.docx-header-paragraph[data-style-id*="Subtitle"] {
  font-size: 14pt;
  font-style: italic;
  margin: 0.5em 0;
}

/* Header type specific styles */
.docx-header-paragraph[data-header-type="first"] {
  /* Styles for first page header */
  font-weight: bold;
}

.docx-header-paragraph[data-header-type="even"] {
  /* Styles for even page header */
  text-align: left;
}

.docx-header-paragraph[data-header-type="default"] {
  /* Styles for default header */
  text-align: center;
}

/* Ensure header images are properly sized */
.docx-document-header img {
  max-width: 100%;
  height: auto;
  display: block;
  margin: 0.5em auto;
}

/* Print-specific header styles */
@media print {
  .docx-document-header {
    page-break-after: avoid;
    break-after: avoid;
  }
  
  .docx-header-paragraph {
    page-break-inside: avoid;
    break-inside: avoid;
  }
}

/* Responsive header styles */
@media (max-width: 768px) {
  .docx-document-header {
    margin-bottom: 1.5em;
    padding-bottom: 0.5em;
  }
  
  .docx-header-paragraph[data-style-id*="title"],
  .docx-header-paragraph[data-style-id*="Title"] {
    font-size: 16pt;
  }
  
  .docx-header-paragraph[data-style-id*="subtitle"],
  .docx-header-paragraph[data-style-id*="Subtitle"] {
    font-size: 12pt;
  }
  
  .docx-header-paragraph[data-style-id*="header"],
  .docx-header-paragraph[data-style-id*="Header"] {
    font-size: 14pt;
  }
}
`;
}

function generateFallbackCSS(styleInfo) {
  return `
body { font-family: Calibri, sans-serif; font-size: 11pt; line-height: 1.15; margin: 20px; }
h1 { font-size: 16pt; } h2 { font-size: 13pt; } p { margin: 10pt 0; }
`;
}

module.exports = {
  generateCssFromStyleInfo,
  getFontFamily,
  getBorderStyle,
  generateAccessibilityStyles,
  generateTrackChangesStyles,
  generateDOCXNumberingStyles,
  generateEnhancedListStyles,
  generateHeaderStyles,
};
