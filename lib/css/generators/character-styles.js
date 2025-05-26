// File: lib/css/generators/character-styles.js
// Character style generation

const { getFontFamily } = require("./base-styles");

/**
 * Generate character styles from style information
 */
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

module.exports = {
  generateCharacterStyles,
};