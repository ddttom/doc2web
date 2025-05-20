// docx-style-parser.js - Main entry point for DOCX style parsing
// This file maintains backward compatibility with existing code while using the new modular structure

// Import the main functionality from the refactored modules
const { parseDocxStyles, generateCssFromStyleInfo } = require('./lib/index');
const { createXPathSelector, selectNodes, selectSingleNode } = require('./lib/xml/xpath-utils');
const { convertTwipToPt } = require('./lib/utils/unit-converter');

// Re-export the functions to maintain backward compatibility
module.exports = {
  parseDocxStyles,
  generateCssFromStyleInfo,
  createXPathSelector,
  selectNodes,
  selectSingleNode,
  convertTwipToPt
};
