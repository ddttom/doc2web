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
    css += `
/* Document defaults */
body {
  font-family: "${styleInfo.theme?.fonts?.minor || "Calibri"}", sans-serif;
  font-size: ${styleInfo.documentDefaults?.character?.fontSize || "11pt"};
  line-height: ${
    styleInfo.documentDefaults?.paragraph?.lineHeight || "1.2"
  }; /* Use specific or default */
  margin: ${
    styleInfo.settings?.pageMargins?.margin || "20px"
  }; /* Use actual margins if available */
  padding: 0;
  color: ${styleInfo.theme?.colors?.tx1 || "#000000"}; /* Default text color */
  background-color: ${
    styleInfo.theme?.colors?.bg1 || "#FFFFFF"
  }; /* Default bg color */
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
    css += generateAccessibilityStyles(styleInfo);
    css += generateTrackChangesStyles(styleInfo);
  } catch (error) {
    console.error("Error generating CSS:", error.message, error.stack);
    css = generateFallbackCSS(styleInfo);
  }
  return css;
}

function generateParagraphStyles(styleInfo) {
  let css = "\n/* Paragraph Styles */\n";
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

    if (style.indentation) {
      marginLeft = style.indentation.left
        ? convertTwipToPt(style.indentation.left)
        : 0;
      marginRight = style.indentation.right
        ? convertTwipToPt(style.indentation.right)
        : 0;
      if (style.indentation.hanging) {
        paddingLeft = convertTwipToPt(style.indentation.hanging);
        textIndent = -paddingLeft;
      } else if (style.indentation.firstLine) {
        textIndent = convertTwipToPt(style.indentation.firstLine);
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
  ${marginLeft > 0 ? `margin-left: ${marginLeft}pt;` : ""}
  ${marginRight > 0 ? `margin-right: ${marginRight}pt;` : ""}
  ${paddingLeft > 0 ? `padding-left: ${paddingLeft}pt;` : ""}
  ${textIndent !== 0 ? `text-indent: ${textIndent}pt;` : ""}
  ${
    paddingLeft > 0 || (style.numbering && style.numbering.hasNumbering)
      ? "position: relative;"
      : ""
  }
}
`;
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
      
      // Calculate indentation and spacing based on DOCX definition
      let pIndentLeft = levelDef.indentation?.left || (levelIndex * 36); // Default indentation if not specified
      let pTextIndent = 0;
      let numRegionWidth = levelDef.format === "bullet" ? 18 : 36; // Default width for number/bullet region
      
      // Handle special indentation cases from DOCX
      if (levelDef.indentation) {
        if (levelDef.indentation.hanging) {
          numRegionWidth = convertTwipToPt(levelDef.indentation.hanging);
          pTextIndent = -numRegionWidth; // Negative text indent for hanging indentation
        } else if (levelDef.indentation.firstLine) {
          pTextIndent = convertTwipToPt(levelDef.indentation.firstLine);
        }
        
        if (levelDef.indentation.left) {
          pIndentLeft = convertTwipToPt(levelDef.indentation.left);
        }
      }
      
      // Create CSS for the numbered element
      css += `
/* Styling for level ${levelIndex} of abstract numbering ${abstractNum.id} */
${itemSelector} {
  display: block;
  position: relative;
  margin-left: ${pIndentLeft}pt;
  padding-left: ${numRegionWidth}pt;
  text-indent: ${pTextIndent}pt;
  counter-increment: docx-counter-${abstractNum.id}-${levelIndex};
}

/* Number or bullet display using ::before pseudo-element */
${itemSelector}::before {
  content: ${getCSSCounterContent(levelDef, abstractNum.id, numberingDefs)};
  position: absolute;
  left: 0;
  top: 0;
  width: ${numRegionWidth}pt;
  text-align: ${levelDef.alignment || "left"};
  box-sizing: border-box;
  padding-right: ${levelDef.suffix === "tab" ? "0.25em" : levelDef.suffix === "space" ? "0.25em" : "0"};
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
  /* ... same as before ... */
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
  /* ... same as before ... */
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
}// Enhanced generateTOCStyles function for lib/css/css-generator.js
function generateTOCStyles(styleInfo) {
  const tocStyles = styleInfo.tocStyles || {};
  const leaderStyle = tocStyles.leaderStyle || {
    character: ".",
    position: 468, /* default 6.5 inches in points */
    spacesBetween: 3
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
}

.docx-toc-heading {
  font-family: "${tocStyles.tocHeadingStyle?.fontFamily || styleInfo.theme?.fonts?.major || "Calibri Light"}", sans-serif;
  font-size: ${tocStyles.tocHeadingStyle?.fontSize || "16pt"}; 
  font-weight: bold;
  margin-bottom: 1em;
  text-align: ${tocStyles.tocHeadingStyle?.alignment || "left"};
  column-span: all;
  -webkit-column-span: all;
  color: ${tocStyles.tocHeadingStyle?.color || styleInfo.theme?.colors?.dk1 || "#2F5496"};
}

/* Structure TOC entries as a flex container with proper alignment */
.docx-toc-entry {
  display: flex;
  flex-wrap: nowrap;
  align-items: baseline;
  width: 100%;
  margin-bottom: 0.5em;
  line-height: 1.4;
  position: relative;
  overflow: hidden;
}

/* TOC text part - flex-grow: 0 to prevent it from expanding */
.docx-toc-text {
  flex: 0 0 auto;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  padding-right: 0.5em;
  max-width: 80%;
}

/* For clickable TOC entries */
.docx-toc-text a {
  text-decoration: none;
  color: inherit;
}

/* TOC dots - flex-grow: 1 to fill available space with dots */
.docx-toc-dots {
  flex: 1 1 auto;
  position: relative;
  height: 1.2em;
  margin: 0 0.5em;
  overflow: hidden;
}

/* Create dots using background image with proper spacing based on leaderStyle */
.docx-toc-dots::after {
  content: "";
  position: absolute;
  bottom: 0.5em;
  left: 0;
  right: 0;
  height: 1px;
  background-image: ${leaderStyle.character === "." 
    ? `radial-gradient(circle, currentColor ${leaderStyle.spacesBetween <= 2 ? "1px" : "0.8px"}, transparent ${leaderStyle.spacesBetween <= 2 ? "1.5px" : "1.2px"})` 
    : leaderStyle.character === "-" 
      ? "linear-gradient(to right, currentColor 33%, transparent 0%)"
      : leaderStyle.character === "_"
        ? "linear-gradient(to right, currentColor 100%, transparent 0%)"
        : "radial-gradient(circle, currentColor 0.8px, transparent 1.2px)"};
  background-position: bottom;
  background-size: ${leaderStyle.character === "." 
    ? `${leaderStyle.spacesBetween * 2 + 2}px 1px` 
    : leaderStyle.character === "-" 
      ? "9px 1px" 
      : "100% 1px"};
  background-repeat: repeat-x;
}

/* TOC page number - flex-grow: 0 to maintain size and right-align */
.docx-toc-pagenum {
  flex: 0 0 auto;
  text-align: right;
  white-space: nowrap;
  font-variant-numeric: tabular-nums;
}

/* Print-specific styles for TOC */
@media print {
  .docx-toc-dots::after {
    background-image: ${leaderStyle.character === "." 
      ? "radial-gradient(circle, black 0.8px, transparent 1.2px)" 
      : leaderStyle.character === "-" 
        ? "linear-gradient(to right, black 33%, transparent 0%)"
        : leaderStyle.character === "_"
          ? "linear-gradient(to right, black 100%, transparent 0%)"
          : "radial-gradient(circle, black 0.8px, transparent 1.2px)"};
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
  font-family: "${style.fontFamily || styleInfo.theme?.fonts?.minor || "Calibri"}", sans-serif;
  font-size: ${style.fontSize || "10pt"};
  font-weight: ${level === 1 ? (style.bold === false ? "normal" : "bold") : (style.bold ? "bold" : "normal")};
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

  return css;
}
function generateUtilityStyles(styleInfo) {
  /* ... (same as previous version, check theme font/color usage) ... */ return `
/* Utility Styles */
.docx-underline { text-decoration: underline; } .docx-strike { text-decoration: line-through; }
h1,h2,h3,h4,h5,h6 { font-family:"${
    styleInfo.theme?.fonts?.major || "Calibri Light"
  }", Arial, sans-serif; color: ${styleInfo.theme?.colors?.dk1 || "#2E74B5"}; }
/* ... more utility styles ... */
`;
}
function generateAccessibilityStyles(styleInfo) {
  /* ... (same as previous, complete) ... */ return `
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
  /* ... (same as previous, complete) ... */ return `
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
function generateFallbackCSS(styleInfo) {
  /* ... (same as previous, complete) ... */ return `
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
};
