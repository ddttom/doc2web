// lib/parsers/numbering-parser.js - Enhanced numbering definition parsing functions
const { selectNodes, selectSingleNode } = require('../xml/xpath-utils');
const { convertTwipToPt } = require('../utils/unit-converter');

/**
 * Enhanced numbering definition parsing to capture complete numbering context
 * Extracts numbering definitions, levels, and formatting information with full context
 * 
 * @param {Document} numberingDoc - Numbering XML document
 * @returns {Object} - Complete numbering definitions with resolution context
 */
function parseNumberingDefinitions(numberingDoc) {
  const numberingDefs = {
    abstractNums: {},
    nums: {},
    numIdMap: {}, // Maps numId to abstractNumId for quicker lookup
    sequenceTrackers: {} // Track numbering sequences
  };
  
  if (!numberingDoc) {
    return numberingDefs;
  }
  
  try {
    // Parse abstract numbering definitions with enhanced detail
    const abstractNumNodes = selectNodes("//w:abstractNum", numberingDoc);
    
    abstractNumNodes.forEach(node => {
      const id = node.getAttribute('w:abstractNumId');
      if (!id) return;
      
      const abstractNum = {
        id,
        levels: {},
        multiLevelType: null,
        restartNumberingAfterBreak: false
      };
      
      // Check for multi-level type
      const multiLevelTypeNode = selectSingleNode("w:multiLevelType", node);
      if (multiLevelTypeNode) {
        abstractNum.multiLevelType = multiLevelTypeNode.getAttribute('w:val');
      }
      
      // Parse level definitions with comprehensive property extraction
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
          start: 1,          // Default start value
          suffix: 'tab',     // Default suffix
          tentative: false
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
          level.textFormat = level.text;
          
          // Parse the text format to understand numbering pattern
          level.parsedFormat = parseNumberingTextFormat(level.text);
        }
        
        // Get suffix type (tab, space, nothing)
        const suffixNode = selectSingleNode("w:suff", lvlNode);
        if (suffixNode) {
          level.suffix = suffixNode.getAttribute('w:val') || 'tab';
        }
        
        // Get if legal style numbering
        const isLegalNode = selectSingleNode("w:isLgl", lvlNode);
        if (isLegalNode) {
          level.isLegal = isLegalNode.getAttribute('w:val') === '1' || isLegalNode.getAttribute('w:val') === 'true';
        }
        
        // Get tentative attribute
        const tentativeAttr = lvlNode.getAttribute('w:tentative');
        if (tentativeAttr) {
          level.tentative = tentativeAttr === '1' || tentativeAttr === 'true';
        }
        
        // Get alignment
        const alignmentNode = selectSingleNode("w:lvlJc", lvlNode);
        if (alignmentNode) {
          level.alignment = alignmentNode.getAttribute('w:val') || 'left';
        }
        
        // Get indentation with comprehensive property extraction
        const indentNode = selectSingleNode("w:pPr/w:ind", lvlNode);
        if (indentNode) {
          const left = indentNode.getAttribute('w:left');
          const hanging = indentNode.getAttribute('w:hanging');
          const firstLine = indentNode.getAttribute('w:firstLine');
          const start = indentNode.getAttribute('w:start');
          const end = indentNode.getAttribute('w:end');
          
          if (left) level.indentation.left = convertTwipToPt(left);
          if (hanging) level.indentation.hanging = convertTwipToPt(hanging);
          if (firstLine) level.indentation.firstLine = convertTwipToPt(firstLine);
          if (start) level.indentation.start = convertTwipToPt(start);
          if (end) level.indentation.end = convertTwipToPt(end);
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
          level.paragraphProps = parseParagraphPropertiesForNumbering(pPrNode);
        }
        
        // Extract run properties (font, size, etc.)
        const rPrNode = selectSingleNode("w:rPr", lvlNode);
        if (rPrNode) {
          level.runProps = parseRunPropertiesForNumbering(rPrNode);
        }
        
        abstractNum.levels[ilvl] = level;
      });
      
      numberingDefs.abstractNums[id] = abstractNum;
    });
    
    // Parse numbering instances with enhanced override handling
    const numNodes = selectNodes("//w:num", numberingDoc);
    
    numNodes.forEach(node => {
      const id = node.getAttribute('w:numId');
      if (!id) return;
      
      const abstractNumIdNode = selectSingleNode("w:abstractNumId", node);
      if (!abstractNumIdNode) return;
      
      const abstractNumId = abstractNumIdNode.getAttribute('w:val');
      if (!abstractNumId) return;
      
      // Enhanced with comprehensive level overrides
      const levelOverrideNodes = selectNodes("w:lvlOverride", node);
      const overrides = {};
      
      levelOverrideNodes.forEach(overrideNode => {
        const ilvl = overrideNode.getAttribute('w:ilvl');
        if (ilvl === null) return;
        
        const override = {};
        
        // Start override
        const startOverrideNode = selectSingleNode("w:startOverride", overrideNode);
        if (startOverrideNode) {
          override.startOverride = parseInt(startOverrideNode.getAttribute('w:val'), 10);
        }
        
        // Complete level override
        const lvlNode = selectSingleNode("w:lvl", overrideNode);
        if (lvlNode) {
          // Parse complete level definition for override
          override.completeLevel = parseLevelDefinition(lvlNode);
        }
        
        overrides[ilvl] = override;
      });
      
      numberingDefs.nums[id] = {
        id,
        abstractNumId,
        overrides: Object.keys(overrides).length > 0 ? overrides : null
      };
      
      // Add to quick lookup map
      numberingDefs.numIdMap[id] = abstractNumId;
      
      // Initialize sequence tracker for this numbering
      numberingDefs.sequenceTrackers[id] = {
        levelCounters: {},
        lastSequence: {},
        restartPoints: {}
      };
    });
    
  } catch (error) {
    console.error('Error parsing numbering definitions:', error);
  }
  
  return numberingDefs;
}

/**
 * Parse numbering text format to understand the numbering pattern
 * @param {string} textFormat - The level text format (e.g., "%1.", "%1.%2.", "(%1)")
 * @returns {Object} - Parsed format information
 */
function parseNumberingTextFormat(textFormat) {
  const parsed = {
    pattern: textFormat,
    levels: [],
    prefix: '',
    suffix: '',
    separator: ''
  };
  
  // Extract level placeholders (%1, %2, etc.)
  const levelMatches = textFormat.match(/%(\d+)/g);
  if (levelMatches) {
    parsed.levels = levelMatches.map(match => parseInt(match.replace('%', ''), 10));
  }
  
  // Extract prefix (text before first level)
  const prefixMatch = textFormat.match(/^([^%]*)/);
  if (prefixMatch) {
    parsed.prefix = prefixMatch[1];
  }
  
  // Extract suffix (text after last level)
  const suffixMatch = textFormat.match(/%\d+(.*)$/);
  if (suffixMatch) {
    parsed.suffix = suffixMatch[1];
  }
  
  // Extract separator between levels
  if (parsed.levels.length > 1) {
    const separatorMatch = textFormat.match(/%\d+([^%]*)%\d+/);
    if (separatorMatch) {
      parsed.separator = separatorMatch[1];
    }
  }
  
  return parsed;
}

/**
 * Parse paragraph properties for numbering context
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Parsed paragraph properties
 */
function parseParagraphPropertiesForNumbering(pPrNode) {
  const props = {};
  
  try {
    // Spacing
    const spacingNode = selectSingleNode("w:spacing", pPrNode);
    if (spacingNode) {
      props.spacing = {
        before: spacingNode.getAttribute('w:before'),
        after: spacingNode.getAttribute('w:after'),
        line: spacingNode.getAttribute('w:line'),
        lineRule: spacingNode.getAttribute('w:lineRule')
      };
    }
    
    // Tab stops specific to numbering
    const tabsNode = selectSingleNode("w:tabs", pPrNode);
    if (tabsNode) {
      const tabNodes = selectNodes("w:tab", tabsNode);
      props.tabs = [];
      
      tabNodes.forEach(tabNode => {
        const pos = tabNode.getAttribute('w:pos');
        const val = tabNode.getAttribute('w:val');
        const leader = tabNode.getAttribute('w:leader');
        
        if (pos && val) {
          props.tabs.push({
            position: pos,
            positionPt: convertTwipToPt(pos),
            type: val,
            leader: leader || 'none'
          });
        }
      });
    }
    
  } catch (error) {
    console.error('Error parsing paragraph properties for numbering:', error);
  }
  
  return props;
}

/**
 * Parse run properties for numbering context
 * @param {Node} rPrNode - Run properties node
 * @returns {Object} - Parsed run properties
 */
function parseRunPropertiesForNumbering(rPrNode) {
  const props = {};
  
  try {
    // Font information
    const fontNode = selectSingleNode("w:rFonts", rPrNode);
    if (fontNode) {
      props.font = {
        ascii: fontNode.getAttribute('w:ascii'),
        hAnsi: fontNode.getAttribute('w:hAnsi'),
        eastAsia: fontNode.getAttribute('w:eastAsia'),
        cs: fontNode.getAttribute('w:cs')
      };
    }
    
    // Size
    const szNode = selectSingleNode("w:sz", rPrNode);
    if (szNode) {
      const sizeHalfPoints = parseInt(szNode.getAttribute('w:val'), 10) || 22;
      props.fontSize = (sizeHalfPoints / 2) + 'pt';
    }
    
    // Bold
    const bNode = selectSingleNode("w:b", rPrNode);
    props.bold = bNode !== null;
    
    // Italic
    const iNode = selectSingleNode("w:i", rPrNode);
    props.italic = iNode !== null;
    
    // Color
    const colorNode = selectSingleNode("w:color", rPrNode);
    if (colorNode) {
      const colorVal = colorNode.getAttribute('w:val');
      if (colorVal) {
        props.color = '#' + colorVal;
      }
    }
    
  } catch (error) {
    console.error('Error parsing run properties for numbering:', error);
  }
  
  return props;
}

/**
 * Parse complete level definition for overrides
 * @param {Node} lvlNode - Level node
 * @returns {Object} - Complete level definition
 */
function parseLevelDefinition(lvlNode) {
  const level = {};
  
  // Get all the same properties as in the main level parsing
  const numFmtNode = selectSingleNode("w:numFmt", lvlNode);
  if (numFmtNode) {
    level.format = numFmtNode.getAttribute('w:val');
  }
  
  const lvlTextNode = selectSingleNode("w:lvlText", lvlNode);
  if (lvlTextNode) {
    level.text = lvlTextNode.getAttribute('w:val');
    level.textFormat = level.text;
    level.parsedFormat = parseNumberingTextFormat(level.text);
  }
  
  const startNode = selectSingleNode("w:start", lvlNode);
  if (startNode) {
    level.start = parseInt(startNode.getAttribute('w:val'), 10);
  }
  
  // Add other level properties as needed
  
  return level;
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
    'decimalZero': 'decimal-leading-zero',
    'aiueo': 'hiragana',
    'iroha': 'hiragana-iroha',
    'decimalEnclosedCircle': 'decimal',
    'decimalFullWidth': 'decimal',
    'ideographDigital': 'cjk-ideographic',
    'japaneseCounting': 'japanese-informal',
    'aiueoFullWidth': 'hiragana',
    'irohaFullWidth': 'hiragana-iroha'
  };
  
  return formatMap[format] || 'decimal';
}

/**
 * Get CSS counter content for numbering based on parsed format
 * Converts Word numbering format to CSS counter content
 * 
 * @param {Object} levelDef - Level definition with parsed format
 * @returns {string} - CSS content value
 */
function getCSSCounterContent(levelDef) {
  if (!levelDef || !levelDef.parsedFormat) {
    return 'counter(list-level-1) "."';
  }
  
  const { parsedFormat } = levelDef;
  let content = '';
  
  // Add prefix
  if (parsedFormat.prefix) {
    content += `"${parsedFormat.prefix}"`;
  }
  
  // Add level counters
  parsedFormat.levels.forEach((level, index) => {
    if (content && !content.endsWith('"')) content += ' ';
    
    const format = getCSSCounterFormat(levelDef.format);
    const counterName = `list-level-${level}`;
    
    content += `counter(${counterName}, ${format})`;
    
    // Add separator between levels
    if (index < parsedFormat.levels.length - 1 && parsedFormat.separator) {
      content += ` "${parsedFormat.separator}"`;
    }
  });
  
  // Add suffix
  if (parsedFormat.suffix) {
    if (content) content += ' ';
    content += `"${parsedFormat.suffix}"`;
  }
  
  return content || 'counter(list-level-1) "."';
}

module.exports = {
  parseNumberingDefinitions,
  parseNumberingTextFormat,
  parseParagraphPropertiesForNumbering,
  parseRunPropertiesForNumbering,
  parseLevelDefinition,
  getCSSCounterFormat,
  getCSSCounterContent
};