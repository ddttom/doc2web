// lib/parsers/numbering-parser.js - Numbering definition parsing functions
const { selectNodes, selectSingleNode } = require('../xml/xpath-utils');
const { convertTwipToPt } = require('../utils/unit-converter');

/**
 * Enhanced numbering definition parsing to capture more details
 * Extracts numbering definitions, levels, and formatting information
 * 
 * @param {Document} numberingDoc - Numbering XML document
 * @returns {Object} - Numbering definitions
 */
function parseNumberingDefinitions(numberingDoc) {
  const numberingDefs = {
    abstractNums: {},
    nums: {},
    numIdMap: {} // Maps numId to abstractNumId for quicker lookup
  };
  
  if (!numberingDoc) {
    return numberingDefs;
  }
  
  try {
    // Parse abstract numbering definitions
    const abstractNumNodes = selectNodes("//w:abstractNum", numberingDoc);
    
    abstractNumNodes.forEach(node => {
      const id = node.getAttribute('w:abstractNumId');
      if (!id) return;
      
      const abstractNum = {
        id,
        levels: {}
      };
      
      // Parse level definitions with enhanced property extraction
      const levelNodes = selectNodes("w:lvl", node);
      levelNodes.forEach(lvlNode => {
        const ilvl = lvlNode.getAttribute('w:ilvl');
        if (ilvl === null) return;
        
        const level = {
          level: parseInt(ilvl, 10),
          format: 'decimal', // Default
          text: '%1.',      // Default
          alignment: 'left', // Default
          indentation: {},
          isLegal: false,
          textFormat: '',
          restart: null,
          start: 1          // Default start value
        };
        
        // Get numbering format (decimal, lowerLetter, upperLetter, etc.)
        const numFmtNode = selectSingleNode("w:numFmt", lvlNode);
        if (numFmtNode) {
          level.format = numFmtNode.getAttribute('w:val') || 'decimal';
        }
        
        // Get level text (how the number appears)
        const lvlTextNode = selectSingleNode("w:lvlText", lvlNode);
        if (lvlTextNode) {
          level.text = lvlTextNode.getAttribute('w:val') || '%1.';
          
          // Store raw format for later use
          level.textFormat = level.text;
          
          // Replace placeholders with computed values for CSS content property
          let cssText = level.text
            .replace(/%(\d+)/g, (match, levelNum) => {
              // Convert to appropriate counter format
              return `counter(list-level-${levelNum})`;
            })
            .replace(/^\s+|\s+$/g, '');
          
          // Add specific handling for common formats
          const cssFormat = getCSSCounterFormat(level.format);
          if (cssFormat !== 'decimal') {
            cssText = cssText.replace(/counter\(list-level-(\d+)\)/g, `counter(list-level-$1, ${cssFormat})`);
          }
          
          level.cssText = cssText;
        }
        
        // Get if legal style numbering
        const isLegalNode = selectSingleNode("w:isLgl", lvlNode);
        if (isLegalNode) {
          level.isLegal = isLegalNode.getAttribute('w:val') === '1' || isLegalNode.getAttribute('w:val') === 'true';
        }
        
        // Get alignment
        const alignmentNode = selectSingleNode("w:lvlJc", lvlNode);
        if (alignmentNode) {
          level.alignment = alignmentNode.getAttribute('w:val') || 'left';
        }
        
        // Get indentation
        const indentNode = selectSingleNode("w:pPr/w:ind", lvlNode);
        if (indentNode) {
          const left = indentNode.getAttribute('w:left');
          const hanging = indentNode.getAttribute('w:hanging');
          const firstLine = indentNode.getAttribute('w:firstLine');
          
          if (left) level.indentation.left = convertTwipToPt(left);
          if (hanging) level.indentation.hanging = convertTwipToPt(hanging);
          if (firstLine) level.indentation.firstLine = convertTwipToPt(firstLine);
        }
        
        // Get restart information
        const lvlRestartNode = selectSingleNode("w:lvlRestart", lvlNode);
        if (lvlRestartNode) {
          level.restart = parseInt(lvlRestartNode.getAttribute('w:val'), 10);
        }
        
        // Get start value
        const startNode = selectSingleNode("w:start", lvlNode);
        if (startNode) {
          level.start = parseInt(startNode.getAttribute('w:val'), 10) || 1;
        }
        
        // Extract paragraph properties
        const pPrNode = selectSingleNode("w:pPr", lvlNode);
        if (pPrNode) {
          // We need to import this function to avoid circular dependencies
          const { parseParagraphProperties } = require('./style-parser');
          level.paragraphProps = parseParagraphProperties(pPrNode);
        }
        
        // Extract run properties (font, size, etc.)
        const rPrNode = selectSingleNode("w:rPr", lvlNode);
        if (rPrNode) {
          // We need to import this function to avoid circular dependencies
          const { parseRunningProperties } = require('./style-parser');
          level.runProps = parseRunningProperties(rPrNode);
        }
        
        abstractNum.levels[ilvl] = level;
      });
      
      numberingDefs.abstractNums[id] = abstractNum;
    });
    
    // Parse numbering instances
    const numNodes = selectNodes("//w:num", numberingDoc);
    
    numNodes.forEach(node => {
      const id = node.getAttribute('w:numId');
      if (!id) return;
      
      const abstractNumIdNode = selectSingleNode("w:abstractNumId", node);
      if (!abstractNumIdNode) return;
      
      const abstractNumId = abstractNumIdNode.getAttribute('w:val');
      if (!abstractNumId) return;
      
      // Enhanced with level overrides
      const levelOverrideNodes = selectNodes("w:lvlOverride", node);
      const overrides = {};
      
      levelOverrideNodes.forEach(overrideNode => {
        const ilvl = overrideNode.getAttribute('w:ilvl');
        if (ilvl === null) return;
        
        const startOverrideNode = selectSingleNode("w:startOverride", overrideNode);
        if (startOverrideNode) {
          const startVal = parseInt(startOverrideNode.getAttribute('w:val'), 10);
          overrides[ilvl] = { start: startVal };
        }
        
        // Check for complete level override
        const lvlNode = selectSingleNode("w:lvl", overrideNode);
        if (lvlNode) {
          // Complete level override - parse it like a regular level
          if (!overrides[ilvl]) overrides[ilvl] = {};
          
          // Get format override
          const numFmtNode = selectSingleNode("w:numFmt", lvlNode);
          if (numFmtNode) {
            overrides[ilvl].format = numFmtNode.getAttribute('w:val');
          }
          
          // Get text override
          const lvlTextNode = selectSingleNode("w:lvlText", lvlNode);
          if (lvlTextNode) {
            overrides[ilvl].text = lvlTextNode.getAttribute('w:val');
          }
          
          // Get paragraph property overrides
          const pPrNode = selectSingleNode("w:pPr", lvlNode);
          if (pPrNode) {
            // We need to import this function to avoid circular dependencies
            const { parseParagraphProperties } = require('./style-parser');
            overrides[ilvl].paragraphProps = parseParagraphProperties(pPrNode);
          }
        }
      });
      
      numberingDefs.nums[id] = {
        id,
        abstractNumId,
        overrides: Object.keys(overrides).length > 0 ? overrides : null
      };
      
      // Add to quick lookup map
      numberingDefs.numIdMap[id] = abstractNumId;
    });
    
  } catch (error) {
    console.error('Error parsing numbering definitions:', error);
  }
  
  return numberingDefs;
}

/**
 * Get CSS counter format from Word format name
 * Maps Word's numbering format names to CSS counter-style values
 * 
 * @param {string} format - Word numbering format
 * @returns {string} - CSS counter format
 */
function getCSSCounterFormat(format) {
  const formatMap = {
    'decimal': 'decimal',
    'lowerLetter': 'lower-alpha',
    'upperLetter': 'upper-alpha',
    'lowerRoman': 'lower-roman',
    'upperRoman': 'upper-roman',
    'bullet': 'disc',
    'chicago': 'decimal',
    'cardinalText': 'decimal',
    'ordinalText': 'decimal',
    'ordinal': 'decimal',
  };
  
  return formatMap[format] || 'decimal';
}

/**
 * Get CSS counter content for numbering
 * Converts Word numbering format to CSS counter content
 * 
 * @param {Object} levelDef - Level definition
 * @returns {string} - CSS content value
 */
function getCSSCounterContent(levelDef) {
  if (!levelDef || !levelDef.text) {
    return 'counter(list-level-1) "."';
  }
  
  // Convert the Word number format to CSS counter
  const text = levelDef.text;
  
  // Replace %1, %2, etc. with counter expressions
  const replaced = text.replace(/%(\d+)/g, (match, level) => {
    const levelNum = parseInt(level, 10);
    let format = 'decimal';
    
    // Use the appropriate format for this level
    if (levelNum === parseInt(levelDef.level, 10) + 1) {
      format = getCSSCounterFormat(levelDef.format);
    }
    
    return '" counter(list-level-' + levelNum + ', ' + format + ') "';
  });
  
  // Add any additional text
  return replaced.replace(/^"/, '').replace(/"$/, '');
}

module.exports = {
  parseNumberingDefinitions,
  getCSSCounterFormat,
  getCSSCounterContent
};