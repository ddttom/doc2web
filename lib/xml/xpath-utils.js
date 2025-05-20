// lib/xml/xpath-utils.js - XML and XPath utilities for DOCX parsing
const xpath = require('xpath');

/**
 * Define XML namespaces used in DOCX files
 * These namespaces are required for XPath queries to work correctly with DOCX XML structure
 * Each namespace prefix maps to its corresponding URI
 */
const NAMESPACES = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006'
};

/**
 * Create an XPath selector with namespaces registered
 * This function configures the XPath selector with the DOCX namespaces
 * 
 * @param {string} expression - XPath expression
 * @returns {xpath.XPathSelect} - XPath selector with namespaces
 */
function createXPathSelector(expression) {
  const select = xpath.useNamespaces(NAMESPACES);
  return select(expression);
}

/**
 * Select nodes using XPath with namespaces
 * This function executes an XPath query and returns matching nodes
 * It includes error handling to prevent crashes on invalid expressions
 * 
 * @param {string} expression - XPath expression
 * @param {Document} doc - XML document
 * @returns {Node[]} - Selected nodes or empty array if none found
 */
function selectNodes(expression, doc) {
  try {
    const select = xpath.useNamespaces(NAMESPACES);
    return select(expression, doc) || [];
  } catch (error) {
    console.error(`Error selecting nodes with expression "${expression}":`, error.message);
    return [];
  }
}

/**
 * Select a single node using XPath with namespaces
 * This function executes an XPath query and returns the first matching node
 * It includes error handling to prevent crashes on invalid expressions
 * 
 * @param {string} expression - XPath expression
 * @param {Document|Node} doc - XML document or node
 * @returns {Node|null} - Selected node or null if not found
 */
function selectSingleNode(expression, doc) {
  try {
    const select = xpath.useNamespaces(NAMESPACES);
    return select(expression, doc, true);
  } catch (error) {
    console.error(`Error selecting single node with expression "${expression}":`, error.message);
    return null;
  }
}

module.exports = {
  NAMESPACES,
  createXPathSelector,
  selectNodes,
  selectSingleNode
};