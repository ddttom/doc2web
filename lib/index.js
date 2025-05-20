// lib/index.js - Main entry point that re-exports public API
// This file maintains backward compatibility with existing code

// Re-export the main functions from the HTML generator
const { extractAndApplyStyles } = require('./html/html-generator');

// Re-export the main functions from the style parser and CSS generator
const { parseDocxStyles } = require('./parsers/style-parser');
const { generateCssFromStyleInfo } = require('./css/css-generator');

// Re-export the XML utilities that were previously exported
const { createXPathSelector, selectNodes, selectSingleNode } = require('./xml/xpath-utils');
const { convertTwipToPt } = require('./utils/unit-converter');

// Export the public API
module.exports = {
  // Main API functions
  extractAndApplyStyles,
  parseDocxStyles,
  generateCssFromStyleInfo,
  
  // XML utilities that were previously exported
  createXPathSelector,
  selectNodes,
  selectSingleNode,
  
  // Unit conversion utility that was previously exported
  convertTwipToPt
};