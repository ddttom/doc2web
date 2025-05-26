// File: lib/html/content-processors.js
// Main content processors - orchestrates all content processing

const { selectNodes, selectSingleNode } = require("../xml/xpath-utils");
const {
  processHeadings,
  applyParagraphStyleToHeading,
  findParagraphContextByText,
  findParagraphContextByPartialText,
  ensureHeadingAccessibility,
} = require("./processors/heading-processor");
const {
  processTOC,
  createTOCContainer,
  processTOCEntry,
  determineTOCLevel,
  structureTOCEntry,
  validateTOCEntryStructure,
  findMatchingHeadingId,
  generateIdFromText,
} = require("./processors/toc-processor");
const {
  processNestedNumberedParagraphs,
  processSpecialParagraphs,
  identifyListPatterns,
  isSpecialParagraph,
  isStructuralMatch,
  extractStructuralPattern,
} = require("./processors/numbering-processor");

// Re-export all functions for backward compatibility
module.exports = {
  // Heading processing
  processHeadings,
  applyParagraphStyleToHeading,
  findParagraphContextByText,
  findParagraphContextByPartialText,
  ensureHeadingAccessibility,
  
  // TOC processing
  processTOC,
  createTOCContainer,
  processTOCEntry,
  determineTOCLevel,
  structureTOCEntry,
  validateTOCEntryStructure,
  findMatchingHeadingId,
  generateIdFromText,
  
  // Numbering processing
  processNestedNumberedParagraphs,
  processSpecialParagraphs,
  identifyListPatterns,
  isSpecialParagraph,
  isStructuralMatch,
  extractStructuralPattern,
};
