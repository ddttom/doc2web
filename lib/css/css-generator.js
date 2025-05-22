// File: ddttom/doc2web/doc2web-804421e293a28695ca8c4527a7c79f342ff0c562/lib/css/css-generator.js

const {
  convertTwipToPt,
  convertBorderSizeToPt,
  getBorderTypeValue,
} = require("../utils/unit-converter");
// getCSSCounterContent is now imported from numbering-parser
const {
  getCSSCounterContent,
  getCSSCounterFormat,
} = require("../parsers/numbering-parser");

/**
 * Generate CSS from extracted style information with enhanced DOCX numbering support
 * Creates a comprehensive CSS stylesheet based on the extracted DOCX styles and numbering context
 * * @param {Object} styleInfo - Style information with numbering context
 * @returns {string} - CSS stylesheet with DOCX-derived numbering
 */
function generateCssFromStyleInfo(styleInfo) {
  let css = "";

  try {
    css += `
/* Document defaults */
body {
  font-family: "${styleInfo.theme?.fonts?.minor || "Calibri"}", sans-serif;
  font-size: ${styleInfo.documentDefaults?.character?.fontSize || "11pt"};
  line-height: 1.15; /* A sensible default */
  margin: 20px; /* Default page margin */
  padding: 0;
}
`;

    css += generateParagraphStyles(styleInfo);

    Object.entries(styleInfo.styles?.character || {}).forEach(([id, style]) => {
      const className = `docx-c-${id
        .replace(/[^a-zA-Z0-9-_]/g, "-")
        .toLowerCase()}`; // Allow underscore and hyphen in class
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
}
`;
    });

    Object.entries(styleInfo.styles?.table || {}).forEach(([id, style]) => {
      const className = `docx-t-${id
        .replace(/[^a-zA-Z0-9-_]/g, "-")
        .toLowerCase()}`;
      css += `
.${className} {
  border-collapse: collapse;
  width: 100%; /* Default to full width */
  margin-bottom: 1em; /* Space after tables */
}
.${className} td, .${className} th {
  padding: ${
    style.cellPadding || "5pt"
  }; /* Use extracted cell padding if available */
  ${getBorderStyle(style, "top")}
  ${getBorderStyle(style, "bottom")}
  ${getBorderStyle(style, "left")}
  ${getBorderStyle(style, "right")}
}
`;
    });

    if (styleInfo.tocStyles) {
      css += generateTOCStyles(styleInfo.tocStyles);
    }

    if (
      styleInfo.numberingDefs &&
      Object.keys(styleInfo.numberingDefs.abstractNums || {}).length > 0
    ) {
      css += generateDOCXNumberingStyles(styleInfo.numberingDefs, styleInfo); // Pass full styleInfo
    }

    css += generateEnhancedListStyles(styleInfo);
    css += generateUtilityStyles(styleInfo);
    css += generateAccessibilityStyles(styleInfo);
    css += generateTrackChangesStyles(styleInfo);
  } catch (error) {
    console.error("Error generating CSS:", error);
    css = generateFallbackCSS(styleInfo);
  }

  return css;
}

function generateParagraphStyles(styleInfo) {
  let css = `
/* Paragraph styles with DOCX indentation */
`;
  Object.entries(styleInfo.styles?.paragraph || {}).forEach(([id, style]) => {
    const className = `docx-p-${id
      .replace(/[^a-zA-Z0-9-_]/g, "-")
      .toLowerCase()}`;
    let marginLeft = 0,
      paddingLeft = 0,
      textIndent = 0;
    let marginRight = 0;

    if (style.indentation) {
      marginLeft = style.indentation.left
        ? convertTwipToPt(style.indentation.left)
        : 0;
      marginRight = style.indentation.right
        ? convertTwipToPt(style.indentation.right)
        : 0;

      if (style.indentation.hanging) {
        // For a hanging indent, the text starts at the 'left' margin,
        // but the number hangs into space defined by 'hanging'.
        // CSS: p { margin-left: left; padding-left: hanging; text-indent: -hanging; }
        // The ::before for number is positioned at `left: -hanging` relative to text start (i.e. within padding-left).
        // However, our ::before is simpler: left:0 relative to padding box.
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
  ${
    style.lineHeight ? `line-height: ${style.lineHeight};` : ""
  } /* Added line height */
  ${style.bold ? "font-weight: bold;" : ""}
  ${style.italic ? "font-style: italic;" : ""}
  ${style.color ? `color: ${style.color};` : ""}
  ${style.alignment ? `text-align: ${style.alignment};` : ""}
  ${
    style.spacing?.before
      ? `margin-top: ${convertTwipToPt(style.spacing.before)}pt;`
      : ""
  }
  ${
    style.spacing?.after
      ? `margin-bottom: ${convertTwipToPt(style.spacing.after)}pt;`
      : ""
  }
  ${marginLeft > 0 ? `margin-left: ${marginLeft}pt;` : ""}
  ${marginRight > 0 ? `margin-right: ${marginRight}pt;` : ""}
  ${paddingLeft > 0 ? `padding-left: ${paddingLeft}pt;` : ""}
  ${textIndent !== 0 ? `text-indent: ${textIndent}pt;` : ""}
  ${paddingLeft > 0 ? "position: relative;" : ""}
}
`;
  });
  return css;
}

function generateDOCXNumberingStyles(numberingDefs, styleInfo) {
  // Pass full styleInfo if needed by getCSSCounterContent indirectly
  let css = `
/* DOCX-derived numbering styles */
`;

  Object.values(numberingDefs.abstractNums || {}).forEach((abstractNum) => {
    // Reset counters for this abstract numbering definition (list group)
    // This should ideally be on a list container element (ol, ul)
    // For now, let's ensure a base reset, assuming lists are structured.
    const listContainerSelector = `[data-num-id^="${abstractNum.id}-container"]`; // Hypothetical container
    // Actually, reset on first item is tricky.
    // Let's rely on counter-reset on ol/ul tags later.

    Object.entries(abstractNum.levels || {}).forEach(
      ([levelIndexStr, levelDef]) => {
        const levelIndex = parseInt(levelIndexStr, 10);

        let pIndentLeft = 0;
        let pTextIndent = 0; // Paragraph's text indent (first line indent or hanging indent)
        let numRegionWidth = 0; // Width allocated for the number/bullet (from hanging usually)

        if (levelDef.indentation) {
          pIndentLeft = levelDef.indentation.left || 0; // Overall indent of the paragraph block
          if (levelDef.indentation.hanging) {
            numRegionWidth = levelDef.indentation.hanging;
            pTextIndent = -numRegionWidth; // Text first line aligns with number
          } else if (levelDef.indentation.firstLine) {
            pTextIndent = levelDef.indentation.firstLine; // Number and text first line share this indent
            // If firstLine is used, number is typically not hanging. numRegionWidth might be small or default.
            numRegionWidth = levelDef.format === "bullet" ? 18 : 24; // Default space for non-hanging numbers
          } else {
            numRegionWidth = levelDef.format === "bullet" ? 18 : 24; // Default if no hanging/firstLine
          }
        } else {
          numRegionWidth = levelDef.format === "bullet" ? 18 : 24; // Default if no indentation info
        }

        const itemSelector = `[data-num-id][data-num-level="${levelIndex}"][data-abstract-num="${abstractNum.id}"]`;
        // More specific selector if numId is available and maps to this abstractNum
        // For now, use abstractNumId directly in counter names to group them.

        css += `
/* Styles for AbstractNum ${abstractNum.id}, Level ${levelIndex} */
${itemSelector} {
  display: block; 
  position: relative;
  margin-left: ${pIndentLeft}pt;
  padding-left: ${numRegionWidth}pt; /* Space for the number/bullet */
  text-indent: ${pTextIndent}pt;   /* Aligns first line of text correctly */
  counter-increment: docx-counter-${abstractNum.id}-${levelIndex};
}

${itemSelector}::before {
  content: ${getCSSCounterContent(levelDef, abstractNum.id, numberingDefs)};
  position: absolute;
  left: 0; /* Position at the start of the padding-left region */
  top: 0; /* Adjust for vertical alignment if needed */
  width: ${numRegionWidth}pt; /* Width of the number/bullet region */
  text-align: ${
    levelDef.alignment || "left"
  }; /* As defined in DOCX for the number */
  box-sizing: border-box;
  ${levelDef.runProps?.bold ? "font-weight: bold;" : ""}
  ${levelDef.runProps?.italic ? "font-style: italic;" : ""}
  ${
    levelDef.runProps?.fontSize
      ? `font-size: ${levelDef.runProps.fontSize};`
      : ""
  }
  ${levelDef.runProps?.color ? `color: ${levelDef.runProps.color};` : ""}
  ${
    levelDef.runProps?.font?.ascii
      ? `font-family: "${levelDef.runProps.font.ascii}", sans-serif;`
      : levelDef.runProps?.font?.hAnsi
      ? `font-family: "${levelDef.runProps.font.hAnsi}", sans-serif;`
      : ""
  }
}
`;
        // Counter reset for deeper levels
        const deeperLevelsToReset = Object.keys(abstractNum.levels)
          .map((l) => parseInt(l, 10))
          .filter((l) => l > levelIndex)
          .map((l) => `docx-counter-${abstractNum.id}-${l} 0`)
          .join(" ");

        if (deeperLevelsToReset) {
          css += `
${itemSelector} {
  counter-reset: ${deeperLevelsToReset};
}
`;
        }
      }
    );
  });

  // Add initial counter resets for each abstract numbering set.
  // These should apply to the containers of lists (e.g., <ol>, <ul>) if they are generated.
  // Or, can be applied to body, but that's less specific.
  Object.values(numberingDefs.abstractNums || {}).forEach((abstractNum) => {
    const allLevelCountersForAbstractNum = Object.keys(abstractNum.levels || {})
      .map((levelIndex) => `docx-counter-${abstractNum.id}-${levelIndex} 0`)
      .join(" ");
    if (allLevelCountersForAbstractNum) {
      // This selector needs to target the actual list container element generated by content-processors.
      // Assuming list containers (ol/ul) will also get a data-abstract-num attribute.
      css += `
ol[data-abstract-num="${abstractNum.id}"], ul[data-abstract-num="${abstractNum.id}"] {
    counter-reset: ${allLevelCountersForAbstractNum};
}
`;
    }
  });

  return css;
}

function generateEnhancedListStyles(styleInfo) {
  let css = `
/* Enhanced List Styles (structural) */
ol[data-abstract-num], ul[data-abstract-num] { /* Target generated list containers */
  list-style-type: none; 
  padding-left: 0; 
  margin-top: 0.5em; /* Default list spacing */
  margin-bottom: 0.5em;
}

li[data-num-id] { /* Target generated list items */
  display: block; /* Behaves like a paragraph */
  /* position: relative; is now part of the [data-num-id][data-num-level] style */
  /* margin/padding/indent for li is handled by the [data-num-id][data-num-level] style */
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
      : "currentColor"; // Use currentColor for 'auto'
  const borderStyle = getBorderTypeValue(border.value);
  return `border-${side}: ${width}pt ${borderStyle} ${color};`;
}

function generateTOCStyles(tocStyles) {
  const leaderStyle = tocStyles.leaderStyle || {};
  let css = `
/* Table of Contents Styles */
.docx-toc {
  margin: 1em 0 2em 0;
  width: 100%;
  padding: 0;
  column-count: 1; /* Explicitly prevent multi-column layout */
  -webkit-column-count: 1;
  -moz-column-count: 1;
}

.docx-toc-heading {
  font-family: "${
    tocStyles.tocHeadingStyle?.fontFamily || "Calibri Light"
  }", sans-serif;
  font-size: ${tocStyles.tocHeadingStyle?.fontSize || "14pt"};
  font-weight: bold;
  margin-bottom: 12pt;
  text-align: ${tocStyles.tocHeadingStyle?.alignment || "center"};
  column-span: all; /* Ensure heading spans all columns if parent somehow becomes multi-column */
  -webkit-column-span: all;
}

.docx-toc-entry {
  display: flex;
  flex-wrap: nowrap;
  align-items: baseline;
  position: relative;
  width: 100%;
  margin-bottom: 4pt;
  line-height: 1.3; /* Adjusted for readability */
  overflow: hidden; /* Prevents long lines from breaking layout */
  /* page-break-inside: avoid; /* Try to keep entries on one page when printing */
}

.docx-toc-text {
  order: 1;
  white-space: nowrap; /* Keep on one line */
  overflow: hidden; /* Hide overflow */
  text-overflow: ellipsis; /* Add ... for overflow */
  margin-right: 0.5em; /* Space before dots */
}

.docx-toc-dots {
  order: 2;
  flex-grow: 1; /* Allow dots to fill space */
  position: relative;
  height: 1em; /* Align with text */
  min-width: 1em; /* Minimum space for dots */
  border-bottom: none;
}

.docx-toc-dots::after {
  content: "";
  position: absolute;
  bottom: 0.35em; /* Adjust vertical position of dots */
  left: 0;
  right: 0;
  height: 1px; /* Thickness of the dot line */
  background-image: ${
    leaderStyle.character === "."
      ? "radial-gradient(circle, currentColor 1px, transparent 1.5px)" /* Slightly spaced dots */
      : leaderStyle.character === "-"
      ? "linear-gradient(to right, currentColor 100%, transparent 0%)" /* Solid line for hyphen */
      : "linear-gradient(to right, currentColor 3px, transparent 0)"
  }; /* Default for underscore */
  background-position: bottom;
  background-size: ${
    leaderStyle.character === "."
      ? "3px 1px"
      : leaderStyle.character === "-"
      ? "100% 1px"
      : "6px 1px"
  };
  background-repeat: ${
    leaderStyle.character === "-" ? "no-repeat" : "repeat-x"
  };
}

.docx-toc-pagenum {
  order: 3;
  text-align: right;
  padding-left: 0.5em; /* Space after dots */
  font-variant-numeric: tabular-nums; /* Align page numbers nicely */
}
`;
  (tocStyles.tocEntryStyles || []).forEach((style, index) => {
    const level = style.level || index + 1;
    const leftIndent = style.indentation?.left
      ? typeof style.indentation.left === "number"
        ? style.indentation.left
        : parseFloat(style.indentation.left) || 0
      : (level - 1) * 18; // Base indent per level (e.g. 18pt)

    css += `
.docx-toc-level-${level} {
  font-family: "${style.fontFamily || "Calibri"}", sans-serif;
  font-size: ${style.fontSize || "11pt"};
  ${
    level === 1 && !style.bold
      ? "font-weight: normal;"
      : style.bold
      ? "font-weight: bold;"
      : ""
  } /* Adjust bolding based on level/style */
  ${style.italic ? "font-style: italic;" : ""}
  padding-left: ${leftIndent}pt; /* Use padding for indent to affect flex items correctly */
}`;
  });
  return css;
}

function generateUtilityStyles(styleInfo) {
  /* ... (largely same as previous, ensure fonts/colors are from theme if possible) ... */
  return `
/* Utility styles */
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }
.docx-tab { display: inline-block; width: ${convertTwipToPt(
    styleInfo.settings?.defaultTabStop || "720"
  )}pt; }
.docx-rtl { direction: rtl; unicode-bidi: bidi-override; } /* unicode-bidi for stronger RTL */
.docx-image { max-width: 100%; height: auto; display: block; margin: 1em auto; } 
figure.docx-image-figure { margin: 1em auto; text-align: center; }
figcaption.docx-image-caption { font-size: 0.9em; font-style: italic; margin-top: 0.5em; color: #555; }
.docx-table-default { width: 100%; border-collapse: collapse; margin: 1em 0; }
.docx-table-default td, .docx-table-default th { border: 1px solid #ccc; padding: 6pt; text-align: left; vertical-align: top; }
.docx-table-default th { background-color: #f2f2f2; font-weight: bold; }
.table-responsive { overflow-x: auto; margin: 1em 0; -webkit-overflow-scrolling: touch; }

/* Base Heading styles (can be overridden by specific .docx-headingX or numbered heading styles) */
h1, h2, h3, h4, h5, h6 {
  font-family: "${
    styleInfo.theme?.fonts?.major || "Calibri Light"
  }", sans-serif;
  color: ${
    styleInfo.theme?.colors?.tx2 || "#1F497D"
  }; /* Using a theme text color */
  margin-top: 1.5em; 
  margin-bottom: 0.5em;
  line-height: 1.2;
  font-weight: bold; /* Common default for headings */
}
h1 { font-size: 20pt; } h2 { font-size: 16pt; } h3 { font-size: 14pt; }
h4 { font-size: 12pt; } h5 { font-size: 11pt; } h6 { font-size: 10pt; }

/* Specific docx-heading classes if generated by Mammoth styleMap */
.docx-heading1 { font-size: 20pt; color: ${
    styleInfo.theme?.colors?.dk1 || "#2F5496"
  }; }
.docx-heading2 { font-size: 16pt; color: ${
    styleInfo.theme?.colors?.dk1 || "#2F5496"
  }; }
.docx-heading3 { font-size: 14pt; color: ${
    styleInfo.theme?.colors?.dk1 || "#4F81BD"
  }; font-style: italic;}
/* ... etc for other docx-headingX classes ... */

@media print {
  body { margin: 0.75in; font-size: 10pt; line-height: 1.2; color: #000; background-color: #fff; }
  .docx-toc-dots::after { background-image: radial-gradient(circle, #000 0.5px, transparent 0) !important; background-size: 2.5px 1px !important;}
  .table-responsive { overflow-x: visible; }
  a { text-decoration: none !important; color: inherit !important; }
  h1,h2,h3,h4,h5,h6 { page-break-after: avoid; page-break-inside: avoid; }
  table, figure { page-break-inside: avoid; }
}
`;
}

function generateAccessibilityStyles(styleInfo) {
  /* ... (same as previous version, ensure it's complete) ... */
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
  /* ... (same as previous version, ensure it's complete) ... */
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

function generateFallbackCSS(styleInfo) {
  /* ... (same as previous version, ensure it's complete) ... */
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
};
