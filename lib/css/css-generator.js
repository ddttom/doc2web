// File: ddttom/doc2web/lib/css/css-generator.js
// Refine generateDOCXNumberingStyles and generateNumberingStylesForParagraph

const {
  convertTwipToPt,
  convertBorderSizeToPt,
  getBorderTypeValue,
} = require("../utils/unit-converter");
const {
  getCSSCounterContent,
  getCSSCounterFormat,
} = require("../parsers/numbering-parser"); // Ensure getCSSCounterFormat is imported

/**
 * Generate CSS from extracted style information with enhanced DOCX numbering support
 * Creates a comprehensive CSS stylesheet based on the extracted DOCX styles and numbering context
 * @param {Object} styleInfo - Style information with numbering context
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
  line-height: 1.15;
  margin: 20px;
  padding: 0;
}
`;

    css += generateParagraphStyles(styleInfo);

    Object.entries(styleInfo.styles?.character || {}).forEach(([id, style]) => {
      const className = `docx-c-${id
        .replace(/[^a-zA-Z0-9]/g, "-")
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
    style.underline
      ? `text-decoration: ${
          style.underline.type === "none" ? "none" : "underline"
        };`
      : ""
  }
}
`;
    });

    Object.entries(styleInfo.styles?.table || {}).forEach(([id, style]) => {
      const className = `docx-t-${id
        .replace(/[^a-zA-Z0-9]/g, "-")
        .toLowerCase()}`;
      css += `
.${className} {
  border-collapse: collapse;
  width: 100%;
}
.${className} td, .${className} th {
  padding: 5pt;
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
      css += generateDOCXNumberingStyles(
        styleInfo.numberingDefs,
        styleInfo.numberingContext
      );
    }

    css += generateEnhancedListStyles(styleInfo); // For HTML lists if they are generated differently
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
      .replace(/[^a-zA-Z0-9]/g, "-")
      .toLowerCase()}`;
    let marginLeft = 0,
      paddingLeft = 0,
      textIndent = 0;

    if (style.indentation) {
      marginLeft = style.indentation.left
        ? convertTwipToPt(style.indentation.left)
        : 0;
      if (style.indentation.hanging) {
        paddingLeft = convertTwipToPt(style.indentation.hanging);
        textIndent = -paddingLeft; // Hanging indent means first line is less indented (or outdented)
      } else if (style.indentation.firstLine) {
        textIndent = convertTwipToPt(style.indentation.firstLine);
      }
      // 'start' indent usually aligns with 'left' for lists, let data-num-id specific styles handle list indents.
    }

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
  ${paddingLeft > 0 ? `padding-left: ${paddingLeft}pt;` : ""}
  ${
    style.indentation?.right
      ? `margin-right: ${convertTwipToPt(style.indentation.right)}pt;`
      : ""
  }
  ${textIndent !== 0 ? `text-indent: ${textIndent}pt;` : ""}
  ${
    paddingLeft > 0
      ? "position: relative;"
      : "" /* Needed for ::before positioning if it hangs */
  }
}
`;
    // Numbering styles for paragraphs with specific numbering definitions are handled by generateDOCXNumberingStyles
  });
  return css;
}

function generateDOCXNumberingStyles(numberingDefs, numberingContext = []) {
  let css = `
/* DOCX-derived numbering styles with proper indentation */
`;

  Object.values(numberingDefs.abstractNums || {}).forEach((abstractNum) => {
    Object.entries(abstractNum.levels || {}).forEach(
      ([levelIndex, levelDef]) => {
        // Find all numIds that use this abstractNum
        const affectedNumIds = Object.entries(numberingDefs.nums || {})
          .filter(
            ([_, numConcrete]) => numConcrete.abstractNumId === abstractNum.id
          )
          .map(([numId, _]) => numId);

        if (affectedNumIds.length === 0 && !levelDef.isReferencedByStyle)
          return; // Skip if not used

        let marginLeft = 0;
        let paddingLeft = 0; // Space for the number/bullet
        let textIndent = 0; // For first-line indent relative to the number/bullet

        if (levelDef.indentation) {
          marginLeft = levelDef.indentation.left || 0; // Overall indent from page margin
          if (levelDef.indentation.hanging) {
            paddingLeft = levelDef.indentation.hanging; // Width of the hanging region
            textIndent = -paddingLeft; // Text starts at margin, number hangs to its left
          } else if (levelDef.indentation.firstLine) {
            // If firstLine is present, it dictates where the number is, and text follows.
            // This usually means the number is indented by firstLine, and text by firstLine too.
            // For CSS, this means padding-left for the text, and ::before positioned by firstLine.
            // Simpler: text-indent for the whole line including number.
            textIndent = levelDef.indentation.firstLine;
            // If firstLine is used, the number is often *not* hanging.
            // We need to ensure text flows correctly after the number.
            // A common pattern is that the number is at `firstLine` and text continues there.
            // Let's adjust padding to accommodate the number.
            // This part is tricky as Word's model isn't a direct CSS map.
            // If numFmt isn't 'bullet', assume number width of ~18-24pt (0.25-0.33 inch)
            if (levelDef.format !== "bullet") {
              paddingLeft = Math.max(paddingLeft, 24); // Ensure some space for numbers
            } else {
              paddingLeft = Math.max(paddingLeft, 18); // Space for bullet
            }
          }
        }

        // Create selectors for all numIds using this abstractNum and level
        const selectors = affectedNumIds.map(
          (numId) => `[data-num-id="${numId}"][data-num-level="${levelIndex}"]`
        );

        if (selectors.length === 0) return; // No elements will match

        css += `
/* Styles for abstractNum ${abstractNum.id}, level ${levelIndex} */
${selectors.join(",\n")} {
  display: block; /* Ensure proper block behavior for p, h*, li */
  position: relative;
  ${marginLeft > 0 ? `margin-left: ${marginLeft}pt;` : "margin-left: 0;"}
  ${paddingLeft > 0 ? `padding-left: ${paddingLeft}pt;` : "padding-left: 0;"}
  ${textIndent !== 0 ? `text-indent: ${textIndent}pt;` : "text-indent: 0;"}
  counter-increment: docx-counter-${abstractNum.id}-${levelIndex};
}

${selectors.join(",\n")}::before {
  content: ${getCSSCounterContent(levelDef, abstractNum.id, levelIndex)};
  position: absolute;
  left: ${
    textIndent < 0 ? textIndent : 0
  }pt; /* Position number in hanging indent or at text-indent start */
  top: 0; /* Adjust as needed for vertical alignment */
  width: ${
    paddingLeft > 0 ? paddingLeft : levelDef.indentation?.hanging || 24
  }pt; /* Width for number, fallback if no padding */
  text-align: ${levelDef.alignment || "left"};
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
      : ""
  }
  box-sizing: border-box; /* Ensure padding/border are included in width */
}
`;
        // Counter reset for deeper levels
        const deeperLevels = Object.keys(abstractNum.levels)
          .map((l) => parseInt(l, 10))
          .filter((l) => l > parseInt(levelIndex, 10))
          .map((l) => `docx-counter-${abstractNum.id}-${l} 0`) // Add 0 to reset to
          .join(" ");

        if (deeperLevels) {
          css += `
${selectors.join(",\n")} {
  counter-reset: ${deeperLevels};
}
`;
        }
      }
    );
  });

  Object.keys(numberingDefs.nums || {}).forEach((numId) => {
    const abstractNumId = numberingDefs.nums[numId].abstractNumId;
    const abstractNum = numberingDefs.abstractNums[abstractNumId];
    if (abstractNum) {
      const counterResets = Object.keys(abstractNum.levels)
        .map((level) => `docx-counter-${abstractNum.id}-${level} 0`)
        .join(" ");
      if (counterResets) {
        css += `
/* Initial counter reset for numId ${numId} (abstract ${abstractNumId}) on list container */
ol[data-numbering-id="${numId}"], ul[data-numbering-id="${numId}"] {
    counter-reset: ${counterResets};
}
`;
      }
    }
  });

  css += `
/* Indentation for non-numbered content following numbered items */
.docx-continues-numbering {
  margin-top: 0.5em; 
}
p[data-follows-numbered="true"] {
  /* Indentation should be handled by .docx-indent-level-X */
}
.docx-indent-level-1 { margin-left: 36pt; } 
.docx-indent-level-2 { margin-left: 72pt; }
.docx-indent-level-3 { margin-left: 108pt; }
.docx-indent-level-4 { margin-left: 144pt; }
`;
  return css;
}

function generateEnhancedListStyles(styleInfo) {
  let css = `
/* Enhanced List Styles (structural) */
ol.docx-numbered-list,
ol.docx-alpha-list,
ol.docx-roman-list,
ul.docx-bulleted-list {
  list-style-type: none; 
  padding-left: 0; 
  margin-top: ${
    styleInfo.documentDefaults?.paragraph?.spacing?.before
      ? convertTwipToPt(styleInfo.documentDefaults.paragraph.spacing.before)
      : "0.5em"
  }pt;
  margin-bottom: ${
    styleInfo.documentDefaults?.paragraph?.spacing?.after
      ? convertTwipToPt(styleInfo.documentDefaults.paragraph.spacing.after)
      : "0.5em"
  }pt;
}

li[data-num-id] {
  display: block; 
  position: relative; 
  margin-bottom: ${
    styleInfo.documentDefaults?.paragraph?.spacing?.line
      ? parseFloat(styleInfo.documentDefaults.paragraph.spacing.line) / 240 / 2
      : "0.2"
  }em;
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
  const color = border.color ? `#${border.color}` : "black";
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
}

.docx-toc-heading {
  font-family: "${
    tocStyles.tocHeadingStyle?.fontFamily || "Calibri Light"
  }", sans-serif;
  font-size: ${tocStyles.tocHeadingStyle?.fontSize || "14pt"};
  font-weight: bold;
  margin-bottom: 12pt;
  text-align: ${tocStyles.tocHeadingStyle?.alignment || "center"};
}

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

.docx-toc-dots::after {
  content: "";
  position: absolute;
  bottom: 0.3em;
  left: 0;
  right: 0;
  height: 1px;
  background-image: ${
    leaderStyle.character === "."
      ? "radial-gradient(circle, currentColor 1px, transparent 1px)"
      : leaderStyle.character === "-"
      ? "linear-gradient(to right, currentColor 2px, transparent 2px)"
      : "linear-gradient(to right, currentColor 3px, transparent 3px)"
  };
  background-position: bottom;
  background-size: ${leaderStyle.character === "." ? "4px 1px" : "6px 1px"};
  background-repeat: repeat-x;
}

.docx-toc-pagenum {
  flex-shrink: 0;
  text-align: right;
  padding-left: 4px;
  font-weight: normal;
}
`;
  (tocStyles.tocEntryStyles || []).forEach((style, index) => {
    const level = style.level || index + 1;
    const leftIndent = style.indentation?.left
      ? typeof style.indentation.left === "string"
        ? parseFloat(style.indentation.left)
        : style.indentation.left
      : (level - 1) * 20; // Default indent if not specified

    css += `
.docx-toc-level-${level} {
  font-family: "${style.fontFamily || "Calibri"}", sans-serif;
  font-size: ${style.fontSize || "11pt"};
  ${level === 1 ? "font-weight: bold;" : ""}
  ${level > 2 ? "font-style: italic;" : ""} /* Example: deeper levels italic */
  margin-left: ${leftIndent}pt;
}`;
  });
  return css;
}

function generateUtilityStyles(styleInfo) {
  return `
/* Utility styles */
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }
.docx-tab { display: inline-block; width: ${convertTwipToPt(
    styleInfo.settings?.defaultTabStop || "720"
  )}pt; }
.docx-rtl { direction: rtl; unicode-bidi: embed; }
.docx-image { max-width: 100%; height: auto; display: block; margin: 10px auto; } /* Centered images */
figure.docx-image-figure { margin: 10px auto; text-align: center; }
figcaption.docx-image-caption { font-size: 0.9em; font-style: italic; margin-top: 0.5em; color: #555; }
.docx-table-default { width: 100%; border-collapse: collapse; margin: 10px 0; }
.docx-table-default td, .docx-table-default th { border: 1px solid #ccc; padding: 6pt; text-align: left; }
.docx-table-default th { background-color: #f2f2f2; font-weight: bold; }
.table-responsive { overflow-x: auto; margin: 1em 0; }

/* Heading styles (base) */
.docx-heading1, .docx-heading2, .docx-heading3, .docx-heading4, .docx-heading5, .docx-heading6 {
  font-family: "${
    styleInfo.theme?.fonts?.major || "Calibri Light"
  }", sans-serif;
  color: ${
    styleInfo.theme?.colors?.dk1 ? styleInfo.theme.colors.dk1 : "#2F5496"
  }; /* Using theme color */
  margin-top: 1.5em;  /* Increased top margin for better separation */
  margin-bottom: 0.5em;
  line-height: 1.2; /* Improved line height for headings */
}

.docx-heading1 { font-size: 20pt; } /* Slightly larger */
.docx-heading2 { font-size: 16pt; }
.docx-heading3 { font-size: 14pt; font-style: italic; }
.docx-heading4 { font-size: 12pt; font-weight: bold; }
.docx-heading5 { font-size: 11pt; font-style: italic; }
.docx-heading6 { font-size: 10pt; color: #555; }

/* Print styles */
@media print {
  body { margin: 0.5in; font-size: 10pt; } /* Adjust print margins */
  .docx-toc-dots::after { background-image: radial-gradient(circle, #000 0.7px, transparent 0); background-size: 3px 1px;}
  .table-responsive { overflow-x: visible; }
  a { text-decoration: none; color: inherit; }
}
`;
}

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
.skip-link:focus { top: 0; outline: 2px solid #4d90fe; outline-offset: 2px; }
.keyboard-focusable:focus, *:focus-visible { outline: 2px solid #4d90fe; outline-offset: 2px; }
@media (prefers-reduced-motion: reduce) { * { transition: none !important; animation: none !important; } }
@media (prefers-contrast: more) {
  body { color: #000000; background-color: #ffffff; }
  a { color: #0000EE; text-decoration: underline; } a:visited { color: #551A8B; }
  th, td { border: 1px solid #000000 !important; } /* Ensure high contrast borders */
  .docx-toc-dots::after { background-image: linear-gradient(to right, #000 1px, transparent 0) !important; }
  .docx-insertion { background-color: #BBDEFB !important; outline: 1px solid #2196F3 !important; }
  .docx-deletion { background-color: #FFCDD2 !important; outline: 1px solid #F44336 !important; text-decoration: line-through; }
}
.sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border-width: 0; }
table caption.sr-only { position: static; /* Ensure caption is visible for tables for context */ }
figure { margin: 1em 0; }
figcaption { font-style: italic; margin-top: 0.5em; }
`;
}

function generateTrackChangesStyles(styleInfo) {
  return `
/* Track Changes Styles */
.docx-track-changes-legend { margin: 1em 0; padding: 0.5em; border: 1px solid #BDBDBD; border-radius: 4px; background-color: #F5F5F5; }
.docx-track-changes-toggle { padding: 0.25em 0.5em; margin-right: 1em; border: 1px solid #BDBDBD; border-radius: 3px; cursor: pointer; }
.docx-track-changes-toggle:hover { background-color: #EEEEEE; }
.docx-track-changes-show .docx-insertion { background-color: #E6F4FF; position: relative; }
.docx-track-changes-show .docx-insertion:hover::after { content: attr(data-author) " (" attr(data-date) ")"; display: block; position: absolute; top: 100%; left: 0; background-color: #FFFFFF; border: 1px solid #BDBDBD; border-radius: 3px; padding: 0.25em 0.5em; font-size: 0.8em; z-index: 10; box-shadow: 0 2px 4px rgba(0,0,0,0.1); white-space: nowrap; }
.docx-track-changes-show .docx-deletion { background-color: #FFEBEE; text-decoration: line-through; position: relative; }
.docx-track-changes-show .docx-deletion:hover::after { content: attr(data-author) " (" attr(data-date) ")"; display: block; position: absolute; top: 100%; left: 0; background-color: #FFFFFF; border: 1px solid #BDBDBD; border-radius: 3px; padding: 0.25em 0.5em; font-size: 0.8em; z-index: 10; box-shadow: 0 2px 4px rgba(0,0,0,0.1); white-space: nowrap; }
/* ... other track changes styles ... */
.docx-track-changes-hide .docx-insertion, .docx-track-changes-hide .docx-deletion { background-color: transparent; text-decoration: none; border-bottom: none; }
.docx-track-changes-hide .docx-deleted-content { display: none; }
`;
}

function generateFallbackCSS(styleInfo) {
  return `
body { font-family: Calibri, sans-serif; font-size: 11pt; line-height: 1.15; margin: 20px; }
h1 { font-size: 16pt; } h2 { font-size: 13pt; } /* Basic heading fallbacks */
p { margin: 10pt 0; }
/* ... other basic fallbacks ... */
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
