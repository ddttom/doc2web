// lib/parsers/numbering-resolver.js - Numbering resolution engine for DOCX introspection
const { selectNodes, selectSingleNode } = require('../xml/xpath-utils');

/**
 * Extract complete paragraph numbering context from document XML
 * Analyzes document.xml to extract numbering context for each paragraph
 * 
 * @param {Document} documentDoc - Document XML document
 * @param {Document} numberingDoc - Numbering XML document
 * @param {Document} styleDoc - Style XML document
 * @returns {Array} - Array of paragraph contexts with numbering information
 */
function extractParagraphNumberingContext(documentDoc, numberingDoc, styleDoc) {
  const paragraphContexts = [];
  
  try {
    // Get all paragraphs from document.xml
    const paragraphs = selectNodes("//w:p", documentDoc);
    
    paragraphs.forEach((p, index) => {
      const context = {
        paragraphIndex: index,
        paragraphId: extractParagraphId(p),
        numberingId: extractNumberingId(p),
        numberingLevel: extractNumberingLevel(p),
        styleId: extractStyleId(p),
        textContent: extractTextContent(p),
        indentation: extractIndentation(p),
        isHeading: false,
        resolvedNumbering: null
      };
      
      // Check if this is a heading paragraph
      context.isHeading = isHeadingParagraph(context.styleId, styleDoc);
      
      // Store raw paragraph node for later reference
      context.paragraphNode = p;
      
      paragraphContexts.push(context);
    });
    
  } catch (error) {
    console.error('Error extracting paragraph numbering context:', error);
  }
  
  return paragraphContexts;
}

/**
 * Extract paragraph ID from paragraph node
 * @param {Node} paragraphNode - Paragraph XML node
 * @returns {string|null} - Paragraph ID
 */
function extractParagraphId(paragraphNode) {
  try {
    const pPr = selectSingleNode("w:pPr", paragraphNode);
    if (pPr) {
      const pId = selectSingleNode("w:pId", pPr);
      if (pId) {
        return pId.getAttribute('w:val');
      }
    }
  } catch (error) {
    console.error('Error extracting paragraph ID:', error);
  }
  return null;
}

/**
 * Extract numbering ID from paragraph properties
 * @param {Node} paragraphNode - Paragraph XML node
 * @returns {string|null} - Numbering ID
 */
function extractNumberingId(paragraphNode) {
  try {
    const numPr = selectSingleNode("w:pPr/w:numPr", paragraphNode);
    if (numPr) {
      const numId = selectSingleNode("w:numId", numPr);
      if (numId) {
        return numId.getAttribute('w:val');
      }
    }
  } catch (error) {
    console.error('Error extracting numbering ID:', error);
  }
  return null;
}

/**
 * Extract numbering level from paragraph properties
 * @param {Node} paragraphNode - Paragraph XML node
 * @returns {number|null} - Numbering level (0-based)
 */
function extractNumberingLevel(paragraphNode) {
  try {
    const numPr = selectSingleNode("w:pPr/w:numPr", paragraphNode);
    if (numPr) {
      const ilvl = selectSingleNode("w:ilvl", numPr);
      if (ilvl) {
        return parseInt(ilvl.getAttribute('w:val'), 10);
      }
    }
  } catch (error) {
    console.error('Error extracting numbering level:', error);
  }
  return null;
}

/**
 * Extract style ID from paragraph properties
 * @param {Node} paragraphNode - Paragraph XML node
 * @returns {string|null} - Style ID
 */
function extractStyleId(paragraphNode) {
  try {
    const pStyle = selectSingleNode("w:pPr/w:pStyle", paragraphNode);
    if (pStyle) {
      return pStyle.getAttribute('w:val');
    }
  } catch (error) {
    console.error('Error extracting style ID:', error);
  }
  return null;
}

/**
 * Extract text content from paragraph
 * @param {Node} paragraphNode - Paragraph XML node
 * @returns {string} - Paragraph text content
 */
function extractTextContent(paragraphNode) {
  try {
    const textNodes = selectNodes(".//w:t", paragraphNode);
    let text = '';
    textNodes.forEach(t => {
      text += t.textContent || '';
    });
    return text.trim();
  } catch (error) {
    console.error('Error extracting text content:', error);
    return '';
  }
}

/**
 * Extract indentation from paragraph properties
 * @param {Node} paragraphNode - Paragraph XML node
 * @returns {Object} - Indentation properties
 */
function extractIndentation(paragraphNode) {
  const indentation = {};
  
  try {
    const ind = selectSingleNode("w:pPr/w:ind", paragraphNode);
    if (ind) {
      const left = ind.getAttribute('w:left');
      const hanging = ind.getAttribute('w:hanging');
      const firstLine = ind.getAttribute('w:firstLine');
      const start = ind.getAttribute('w:start');
      const end = ind.getAttribute('w:end');
      
      if (left) indentation.left = parseInt(left, 10);
      if (hanging) indentation.hanging = parseInt(hanging, 10);
      if (firstLine) indentation.firstLine = parseInt(firstLine, 10);
      if (start) indentation.start = parseInt(start, 10);
      if (end) indentation.end = parseInt(end, 10);
    }
  } catch (error) {
    console.error('Error extracting indentation:', error);
  }
  
  return indentation;
}

/**
 * Check if paragraph is a heading based on style
 * @param {string} styleId - Style ID
 * @param {Document} styleDoc - Style XML document
 * @returns {boolean} - True if heading paragraph
 */
function isHeadingParagraph(styleId, styleDoc) {
  if (!styleId || !styleDoc) return false;
  
  try {
    // Look for heading styles
    const headingPatterns = /^heading\d+$|^toc\d+$|^h\d+$/i;
    if (headingPatterns.test(styleId)) {
      return true;
    }
    
    // Check style definition in styles.xml
    const styleNode = selectSingleNode(`//w:style[@w:styleId='${styleId}']`, styleDoc);
    if (styleNode) {
      const nameNode = selectSingleNode("w:name", styleNode);
      if (nameNode) {
        const styleName = nameNode.getAttribute('w:val') || '';
        return /heading|title|caption/i.test(styleName);
      }
    }
  } catch (error) {
    console.error('Error checking if heading paragraph:', error);
  }
  
  return false;
}

/**
 * Resolve actual numbering for paragraphs with numbering context
 * Calculates the actual number based on sequence, restarts, and overrides
 * 
 * @param {Array} paragraphContexts - Array of paragraph contexts
 * @param {Object} numberingDefs - Numbering definitions
 * @returns {Array} - Paragraph contexts with resolved numbering
 */
function resolveNumberingForParagraphs(paragraphContexts, numberingDefs) {
  // Initialize sequence trackers
  const sequenceTrackers = {};
  
  // Initialize trackers for each numbering definition
  Object.keys(numberingDefs.nums).forEach(numId => {
    sequenceTrackers[numId] = new NumberingSequenceTracker(numId, numberingDefs);
  });
  
  // Process each paragraph in order
  paragraphContexts.forEach(context => {
    if (context.numberingId && context.numberingLevel !== null) {
      const tracker = sequenceTrackers[context.numberingId];
      if (tracker) {
        context.resolvedNumbering = tracker.getNextNumber(
          context.numberingLevel,
          context,
          numberingDefs
        );
      }
    }
  });
  
  return paragraphContexts;
}

/**
 * Class to track numbering sequences for each numbering definition
 */
class NumberingSequenceTracker {
  constructor(numberingId, numberingDefs) {
    this.numberingId = numberingId;
    this.numberingDefs = numberingDefs;
    this.levelCounters = {}; // Current counter for each level
    this.lastSequence = {}; // Last seen sequence for restart detection
    this.restartPoints = {}; // Points where numbering restarts
    
    // Initialize counters for all levels
    const numDef = numberingDefs.nums[numberingId];
    if (numDef) {
      const abstractNum = numberingDefs.abstractNums[numDef.abstractNumId];
      if (abstractNum) {
        Object.keys(abstractNum.levels).forEach(level => {
          this.levelCounters[level] = 0;
        });
      }
    }
  }
  
  /**
   * Get the next number for a specific level
   * @param {number} level - Numbering level (0-based)
   * @param {Object} context - Paragraph context
   * @param {Object} numberingDefs - Numbering definitions
   * @returns {Object} - Resolved numbering information
   */
  getNextNumber(level, context, numberingDefs) {
    const numDef = numberingDefs.nums[this.numberingId];
    if (!numDef) return null;
    
    const abstractNum = numberingDefs.abstractNums[numDef.abstractNumId];
    if (!abstractNum) return null;
    
    const levelDef = abstractNum.levels[level];
    if (!levelDef) return null;
    
    // Check for level overrides
    let effectiveLevelDef = levelDef;
    if (numDef.overrides && numDef.overrides[level]) {
      const override = numDef.overrides[level];
      if (override.completeLevel) {
        effectiveLevelDef = { ...levelDef, ...override.completeLevel };
      }
      if (override.startOverride !== undefined) {
        effectiveLevelDef = { ...effectiveLevelDef, start: override.startOverride };
      }
    }
    
    // Handle restart logic
    this.handleRestarts(level, effectiveLevelDef, context);
    
    // Increment counter for this level
    this.levelCounters[level] = (this.levelCounters[level] || 0) + 1;
    
    // Reset counters for deeper levels
    Object.keys(this.levelCounters).forEach(lvl => {
      if (parseInt(lvl, 10) > level) {
        this.levelCounters[lvl] = 0;
      }
    });
    
    // Calculate actual number
    const actualNumber = (effectiveLevelDef.start || 1) + this.levelCounters[level] - 1;
    
    // Format the number according to the level definition
    return this.formatNumbering(actualNumber, level, effectiveLevelDef, abstractNum);
  }
  
  /**
   * Handle restart logic for numbering
   * @param {number} level - Current level
   * @param {Object} levelDef - Level definition
   * @param {Object} context - Paragraph context
   */
  handleRestarts(level, levelDef, context) {
    // Check if this level should restart numbering
    if (levelDef.restart !== null && levelDef.restart !== undefined) {
      // Restart at specified intervals
      if (this.levelCounters[level] >= levelDef.restart) {
        this.levelCounters[level] = 0;
      }
    }
    
    // Check for document structure-based restarts
    // (e.g., new sections, page breaks, etc.)
    if (this.shouldRestartBasedOnContext(level, context)) {
      this.levelCounters[level] = 0;
    }
  }
  
  /**
   * Check if numbering should restart based on context
   * @param {number} level - Numbering level
   * @param {Object} context - Paragraph context
   * @returns {boolean} - True if should restart
   */
  shouldRestartBasedOnContext(level, context) {
    // This would analyze document structure for restart conditions
    // For now, implement basic logic
    
    // Restart if there's a significant gap in paragraph indices
    // (indicating possible section breaks)
    if (this.lastSequence[level] !== undefined) {
      const gap = context.paragraphIndex - this.lastSequence[level];
      if (gap > 50) { // Arbitrary threshold
        return true;
      }
    }
    
    this.lastSequence[level] = context.paragraphIndex;
    return false;
  }
  
  /**
   * Format numbering according to level definition
   * @param {number} actualNumber - The calculated number
   * @param {number} level - Numbering level
   * @param {Object} levelDef - Level definition
   * @param {Object} abstractNum - Abstract numbering definition
   * @returns {Object} - Formatted numbering information
   */
  formatNumbering(actualNumber, level, levelDef, abstractNum) {
    const formattedNumber = this.formatSingleNumber(actualNumber, levelDef.format);
    
    // Build hierarchical numbering if needed
    let fullNumbering = '';
    if (levelDef.parsedFormat) {
      fullNumbering = this.buildHierarchicalNumbering(level, levelDef, abstractNum);
    } else {
      // Fallback to simple formatting
      fullNumbering = levelDef.parsedFormat?.prefix || '';
      fullNumbering += formattedNumber;
      fullNumbering += levelDef.parsedFormat?.suffix || '.';
    }
    
    return {
      rawNumber: actualNumber,
      formattedNumber: formattedNumber,
      fullNumbering: fullNumbering,
      levelDef: levelDef,
      format: levelDef.format,
      textFormat: levelDef.textFormat || levelDef.text,
      suffix: levelDef.suffix
    };
  }
  
  /**
   * Format a single number according to the specified format
   * @param {number} number - Number to format
   * @param {string} format - Numbering format
   * @returns {string} - Formatted number
   */
  formatSingleNumber(number, format) {
    switch (format) {
      case 'decimal':
        return number.toString();
      case 'lowerLetter':
        return String.fromCharCode(96 + ((number - 1) % 26) + 1); // a, b, c, ...
      case 'upperLetter':
        return String.fromCharCode(64 + ((number - 1) % 26) + 1); // A, B, C, ...
      case 'lowerRoman':
        return this.toRoman(number).toLowerCase();
      case 'upperRoman':
        return this.toRoman(number);
      case 'decimalZero':
        return number.toString().padStart(2, '0');
      case 'bullet':
        return '•';
      default:
        return number.toString();
    }
  }
  
  /**
   * Build hierarchical numbering (e.g., "1.a.i")
   * @param {number} currentLevel - Current level
   * @param {Object} levelDef - Level definition
   * @param {Object} abstractNum - Abstract numbering definition
   * @returns {string} - Hierarchical numbering string
   */
  buildHierarchicalNumbering(currentLevel, levelDef, abstractNum) {
    if (!levelDef.parsedFormat) {
      return this.formatSingleNumber(this.levelCounters[currentLevel], levelDef.format);
    }
    
    let result = levelDef.parsedFormat.prefix || '';
    
    // Process each level in the format
    levelDef.parsedFormat.levels.forEach((formatLevel, index) => {
      if (formatLevel <= currentLevel + 1) { // +1 because levels are 1-based in format
        const level = formatLevel - 1; // Convert to 0-based
        const levelCount = this.levelCounters[level] || 0;
        
        if (levelCount > 0) {
          const levelLevelDef = abstractNum.levels[level];
          if (levelLevelDef) {
            const number = (levelLevelDef.start || 1) + levelCount - 1;
            const formatted = this.formatSingleNumber(number, levelLevelDef.format);
            result += formatted;
            
            // Add separator if not the last level
            if (index < levelDef.parsedFormat.levels.length - 1 && 
                levelDef.parsedFormat.separator) {
              result += levelDef.parsedFormat.separator;
            }
          }
        }
      }
    });
    
    result += levelDef.parsedFormat.suffix || '';
    return result;
  }
  
  /**
   * Convert number to Roman numerals
   * @param {number} num - Number to convert
   * @returns {string} - Roman numeral representation
   */
  toRoman(num) {
    const values = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
    const numerals = ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I'];
    
    let result = '';
    for (let i = 0; i < values.length; i++) {
      while (num >= values[i]) {
        result += numerals[i];
        num -= values[i];
      }
    }
    return result;
  }
}

module.exports = {
  extractParagraphNumberingContext,
  resolveNumberingForParagraphs,
  NumberingSequenceTracker,
  extractParagraphId,
  extractNumberingId,
  extractNumberingLevel,
  extractStyleId,
  extractTextContent,
  extractIndentation,
  isHeadingParagraph
};