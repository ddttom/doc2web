// style-extractor.js - Main entry point for DOCX style extraction and HTML generation
// This file maintains backward compatibility with existing code while using the new modular structure

// Import the main functionality from the refactored modules
// Update the import to use the new function name
const { extractAndApplyStyles, extractImagesFromDocx } = require('./lib/html/html-generator');

// Re-export the main function to maintain backward compatibility
module.exports = {
  extractAndApplyStyles,
  extractImagesFromDocx,  // Export with the new name
};
