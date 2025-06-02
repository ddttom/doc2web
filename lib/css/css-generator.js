// File: lib/css/css-generator.js
// Main CSS generator - orchestrates all style generation

const { generateBaseStyles, generateFallbackCSS, getFontFamily, getBorderStyle } = require("./generators/base-styles");
const { generateParagraphStyles, generateGlobalParagraphDataAttributeStyles } = require("./generators/paragraph-styles");
const { generateCharacterStyles } = require("./generators/character-styles");
const { generateTableStyles, generateTableDataAttributeStyles } = require("./generators/table-styles");
const { generateDOCXNumberingStyles, generateEnhancedListStyles, generateBulletListStyles, generateCustomBulletStyles } = require("./generators/numbering-styles");
const { generateTOCStyles } = require("./generators/toc-styles");
const { generateUtilityStyles } = require("./generators/utility-styles");
const { generateAccessibilityStyles, generateTrackChangesStyles, generateHeaderStyles } = require("./generators/specialized-styles");

/**
 * Generate CSS from extracted style information.
 */
function generateCssFromStyleInfo(styleInfo) {
  let css = "";
  try {
    // Generate base document styles
    css += generateBaseStyles(styleInfo);
    
    // Generate component-specific styles
    css += generateParagraphStyles(styleInfo);
    css += generateCharacterStyles(styleInfo);
    css += generateTableStyles(styleInfo);
    
    // Generate TOC styles if present
    if (styleInfo.tocStyles) {
      css += generateTOCStyles(styleInfo);
    }
    
    // Generate numbering styles if present
    if (
      styleInfo.numberingDefs &&
      Object.keys(styleInfo.numberingDefs.abstractNums || {}).length > 0
    ) {
      css += generateDOCXNumberingStyles(styleInfo.numberingDefs, styleInfo);
    }
    
    // Generate additional styles
    css += generateEnhancedListStyles(styleInfo);
    css += generateBulletListStyles(styleInfo);
    css += generateCustomBulletStyles(styleInfo.numberingDefs);
    css += generateUtilityStyles(styleInfo);
    css += generateHeaderStyles(styleInfo);
    css += generateAccessibilityStyles(styleInfo);
    css += generateTrackChangesStyles(styleInfo);
    
    // Generate enhanced data attribute styles for comprehensive formatting
    css += generateGlobalParagraphDataAttributeStyles();
    css += generateTableDataAttributeStyles();
    
  } catch (error) {
    console.error("Error generating CSS:", error.message, error.stack);
    css = generateFallbackCSS(styleInfo);
  }
  return css;
}

// Export the main function and utility functions for backward compatibility
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
