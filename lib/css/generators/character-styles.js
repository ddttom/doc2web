// File: lib/css/generators/character-styles.js
// Character style generation

const { getFontFamily } = require("./base-styles");

/**
 * Generate character styles from style information
 */
function generateCharacterStyles(styleInfo) {
  let css = "\n/* Character Styles */\n";
  
  // Generate comprehensive character styles
  Object.entries(styleInfo.styles?.character || {}).forEach(([id, style]) => {
    const className = `docx-c-${id
      .replace(/[^a-zA-Z0-9-_]/g, "-")
      .toLowerCase()}`;
    css += generateCharacterStyleRule(className, style, styleInfo);
  });
  
  // Generate data attribute-based styles for enhanced formatting
  css += generateDataAttributeStyles();
  
  // Add fallback italic styles for better coverage
  css += `
/* Fallback italic styles */
em, .italic, [style*="font-style: italic"] {
  font-style: italic !important;
}

/* Ensure em tags are always italic */
em {
  font-style: italic !important;
}
`;
  
  return css;
}

module.exports = {
  generateCharacterStyles,
};

/**
 * Generate comprehensive character style rule
 */
function generateCharacterStyleRule(className, style, styleInfo) {
  let css = `\n.${className} {\n`;
  
  // Font family properties
  if (style.fonts) {
    if (style.fonts.ascii) css += `  font-family: "${style.fonts.ascii}", sans-serif;\n`;
    if (style.fonts.hAnsi && style.fonts.hAnsi !== style.fonts.ascii) {
      css += `  /* Hansi font: ${style.fonts.hAnsi} */\n`;
    }
    if (style.fonts.eastAsia) css += `  /* East Asian font: ${style.fonts.eastAsia} */\n`;
    if (style.fonts.cs) css += `  /* Complex script font: ${style.fonts.cs} */\n`;
  } else if (style.font) {
    css += `  font-family: "${getFontFamily(style, styleInfo)}", sans-serif;\n`;
  }
  
  // Font size properties
  if (style.fontSize !== undefined) {
    css += `  font-size: ${typeof style.fontSize === 'number' ? style.fontSize + 'pt' : style.fontSize};\n`;
  }
  if (style.fontSizeCs !== undefined) {
    css += `  /* Complex script font size: ${typeof style.fontSizeCs === 'number' ? style.fontSizeCs + 'pt' : style.fontSizeCs} */\n`;
  }
  
  // Basic formatting
  if (style.bold) css += `  font-weight: bold;\n`;
  if (style.italic) css += `  font-style: italic !important;\n`;
  
  // Color properties
  if (style.color) {
    if (typeof style.color === 'string') {
      css += `  color: ${style.color};\n`;
    } else if (style.color.val) {
      css += `  color: #${style.color.val};\n`;
      if (style.color.themeColor) css += `  /* Theme color: ${style.color.themeColor} */\n`;
      if (style.color.themeTint) css += `  /* Theme tint: ${style.color.themeTint} */\n`;
      if (style.color.themeShade) css += `  /* Theme shade: ${style.color.themeShade} */\n`;
    }
  }
  
  // Background/highlight properties
  if (style.highlight) {
    css += `  background-color: ${style.highlight};\n`;
  }
  if (style.shading) {
    if (style.shading.fill) css += `  background-color: #${style.shading.fill};\n`;
    if (style.shading.pattern && style.shading.pattern !== 'clear') {
      css += `  /* Shading pattern: ${style.shading.pattern} */\n`;
    }
  }
  
  // Text decoration properties
  let textDecoration = [];
  if (style.underline && style.underline.type !== "none") {
    textDecoration.push('underline');
  }
  if (style.strikethrough) {
    textDecoration.push('line-through');
  }
  if (style.doubleStrikethrough) {
    textDecoration.push('line-through');
    css += `  /* Double strikethrough */\n`;
  }
  if (textDecoration.length > 0) {
    css += `  text-decoration: ${textDecoration.join(' ')};\n`;
  }
  
  // Text effects
  if (style.caps) css += `  text-transform: uppercase;\n`;
  if (style.smallCaps) css += `  font-variant: small-caps;\n`;
  if (style.vanish) css += `  visibility: hidden;\n`;
  if (style.webHidden) css += `  display: none;\n`;
  if (style.emboss) css += `  text-shadow: 1px 1px 0px rgba(0,0,0,0.3);\n`;
  if (style.imprint) css += `  text-shadow: -1px -1px 0px rgba(0,0,0,0.3);\n`;
  if (style.outline) css += `  -webkit-text-stroke: 1px currentColor;\n`;
  if (style.shadow) css += `  text-shadow: 2px 2px 4px rgba(0,0,0,0.5);\n`;
  
  // Vertical alignment
  if (style.verticalAlign === 'superscript' || style.superscript) {
    css += `  vertical-align: super;\n  font-size: 0.8em;\n`;
  }
  if (style.verticalAlign === 'subscript' || style.subscript) {
    css += `  vertical-align: sub;\n  font-size: 0.8em;\n`;
  }
  
  // Character spacing properties
  if (style.spacing !== undefined) {
    css += `  letter-spacing: ${typeof style.spacing === 'number' ? style.spacing + 'pt' : style.spacing};\n`;
  }
  if (style.w !== undefined) {
    css += `  transform: scaleX(${style.w / 100});\n`;
  }
  if (style.kern !== undefined) {
    css += `  font-kerning: ${style.kern > 0 ? 'normal' : 'none'};\n`;
  }
  if (style.position !== undefined) {
    css += `  position: relative;\n  top: ${typeof style.position === 'number' ? style.position + 'pt' : style.position};\n`;
  }
  
  css += `}\n`;
  return css;
}

/**
 * Generate data attribute-based styles for enhanced formatting
 */
function generateDataAttributeStyles() {
  let css = `
/* Enhanced Character Formatting - Data Attribute Styles */

/* Font family styles */
[data-font-ascii] { /* Font family applied via data attributes */ }
[data-font-hansi] { /* Hansi font applied via data attributes */ }
[data-font-eastasia] { /* East Asian font applied via data attributes */ }
[data-font-cs] { /* Complex script font applied via data attributes */ }

/* Font size styles */
[data-font-size] { /* Font size applied via data attributes */ }
[data-font-size-cs] { /* Complex script font size applied via data attributes */ }

/* Color styles */
[data-color] { /* Color applied via data attributes */ }
[data-color-theme] { /* Theme color applied via data attributes */ }
[data-color-tint] { /* Color tint applied via data attributes */ }
[data-color-shade] { /* Color shade applied via data attributes */ }

/* Background/highlight styles */
[data-highlight] { /* Highlight applied via data attributes */ }
[data-shd-fill] { /* Shading fill applied via data attributes */ }
[data-shd-color] { /* Shading color applied via data attributes */ }
[data-shd-pattern] { /* Shading pattern applied via data attributes */ }

/* Text effects styles */
.docx-double-strikethrough { text-decoration: line-through; }
.docx-caps { text-transform: uppercase; }
.docx-small-caps { font-variant: small-caps; }
.docx-vanish { visibility: hidden; }
.docx-web-hidden { display: none; }
.docx-emboss { text-shadow: 1px 1px 0px rgba(0,0,0,0.3); }
.docx-imprint { text-shadow: -1px -1px 0px rgba(0,0,0,0.3); }
.docx-outline { -webkit-text-stroke: 1px currentColor; }
.docx-shadow { text-shadow: 2px 2px 4px rgba(0,0,0,0.5); }

/* Character spacing styles */
[data-spacing] { /* Letter spacing applied via data attributes */ }
[data-w] { /* Character scaling applied via data attributes */ }
[data-kern] { /* Kerning applied via data attributes */ }
[data-position] { /* Character position applied via data attributes */ }

/* Enhanced underline styles */
.docx-underline { text-decoration: underline; }
.docx-underline-double { text-decoration: underline; text-decoration-style: double; }
.docx-underline-thick { text-decoration: underline; text-decoration-thickness: 2px; }
.docx-underline-dotted { text-decoration: underline; text-decoration-style: dotted; }
.docx-underline-dashed { text-decoration: underline; text-decoration-style: dashed; }
.docx-underline-wavy { text-decoration: underline; text-decoration-style: wavy; }
`;
  
  return css;
}
