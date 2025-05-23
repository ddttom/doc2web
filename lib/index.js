// lib/index.js - Main entry point that re-exports public API
// This file maintains backward compatibility with existing code

// Re-export the main functions from the HTML generator
// Change processImages to extractImagesFromDocx
const { extractAndApplyStyles, convertToStyledHtml, applyStylesToHtml, extractImagesFromDocx } = require('./html/html-generator');

// Re-export the main functions from the style parser and CSS generator
const { parseDocxStyles } = require('./parsers/style-parser');
const { generateCssFromStyleInfo, getFontFamily, getBorderStyle, generateAccessibilityStyles, generateTrackChangesStyles } = require('./css/css-generator');

// Re-export the XML utilities that were previously exported
const { createXPathSelector, selectNodes, selectSingleNode } = require('./xml/xpath-utils');
const { convertTwipToPt, convertBorderSizeToPt, getBorderTypeValue } = require('./utils/unit-converter');
const { getLeaderChar } = require('./utils/common-utils');

// Re-export the new modules
const { parseDocumentMetadata, applyMetadataToHtml, addDublinCoreMetadata, addOpenGraphMetadata, addTwitterCardMetadata, addJsonLdStructuredData } = require('./parsers/metadata-parser');
const { parseTrackChanges, processTrackChanges, processInsertions, processDeletions, processMoves, processFormattingChanges, addTrackChangesLegend } = require('./parsers/track-changes-parser');
const { processForAccessibility, processTablesForAccessibility, processImagesForAccessibility, ensureHeadingHierarchy, addAriaLandmarks, addSkipNavigation, enhanceKeyboardNavigation, enhanceColorContrast } = require('./accessibility/wcag-processor');

// Re-export the HTML element processors
const { processTables, processImages: processHtmlImages, processLanguageElements } = require('./html/element-processors');
const { processHeadings, processTOC, processNestedNumberedParagraphs, processSpecialParagraphs, identifyListPatterns, isSpecialParagraph } = require('./html/content-processors');
const { ensureHtmlStructure, addDocumentMetadata } = require('./html/structure-processor');

// Re-export the style mapper
const { createStyleMap, createDocumentTransformer } = require('./css/style-mapper');

// Re-export document parser functions
const { parseDocumentDefaults, parseSettings, analyzeDocumentStructure, getDefaultStyleInfo } = require('./parsers/document-parser');
const { parseNumberingDefinitions, getCSSCounterFormat, getCSSCounterContent } = require('./parsers/numbering-parser');
const { parseTheme, getColorValue } = require('./parsers/theme-parser');
const { parseTocStyles } = require('./parsers/toc-parser');

// Re-export enhanced numbering resolution functions
const { extractParagraphNumberingContext, resolveNumberingForParagraphs, NumberingSequenceTracker } = require('./parsers/numbering-resolver');

// Re-export header parser functions
const { extractDocumentHeader, processHeaderForHtml, extractHeaderFromXml, extractHeaderFromDocument, analyzeHeaderParagraph, extractParagraphFormatting, extractImagesFromParagraph, isLikelyHeaderContent } = require('./parsers/header-parser');

// Export the public API
module.exports = {
  // Main API functions
  extractAndApplyStyles,
  convertToStyledHtml,
  applyStylesToHtml,
  extractImagesFromDocx,   // Changed from processImages
  parseDocxStyles,
  generateCssFromStyleInfo,
  
  // XML utilities
  createXPathSelector,
  selectNodes,
  selectSingleNode,
  
  // Style utilities
  getFontFamily,
  getBorderStyle,
  
  // Unit conversion utilities
  convertTwipToPt,
  convertBorderSizeToPt,
  getBorderTypeValue,
  getLeaderChar,
  
  // Style mapper functions
  createStyleMap,
  createDocumentTransformer,
  
  // Document parsers
  parseDocumentDefaults,
  parseSettings,
  analyzeDocumentStructure,
  getDefaultStyleInfo,
  parseNumberingDefinitions,
  getCSSCounterFormat,
  getCSSCounterContent,
  parseTheme,
  getColorValue,
  parseTocStyles,
  
  // HTML processors
  processTables,
  processHtmlImages,  // Alias for processImages from element-processors
  processLanguageElements,
  processHeadings,
  processTOC,
  processNestedNumberedParagraphs,
  processSpecialParagraphs,
  identifyListPatterns,
  isSpecialParagraph,
  ensureHtmlStructure,
  addDocumentMetadata,
  
  // CSS generators
  generateAccessibilityStyles,
  generateTrackChangesStyles,
  
  // Metadata functions
  parseDocumentMetadata,
  applyMetadataToHtml,
  addDublinCoreMetadata,
  addOpenGraphMetadata,
  addTwitterCardMetadata,
  addJsonLdStructuredData,
  
  // Track changes functions
  parseTrackChanges,
  processTrackChanges,
  processInsertions,
  processDeletions,
  processMoves,
  processFormattingChanges,
  addTrackChangesLegend,
  
  // Accessibility functions
  processForAccessibility,
  processTablesForAccessibility,
  processImagesForAccessibility,
  ensureHeadingHierarchy,
  addAriaLandmarks,
  addSkipNavigation,
  enhanceKeyboardNavigation,
  enhanceColorContrast,
  
  // Header parser functions
  extractDocumentHeader,
  processHeaderForHtml,
  extractHeaderFromXml,
  extractHeaderFromDocument,
  analyzeHeaderParagraph,
  extractParagraphFormatting,
  extractImagesFromParagraph,
  isLikelyHeaderContent
};
