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

module.exports = {
  generateParagraphStyles,
};