// lib/parsers/document-parser.js - Document structure and settings parsing
const { selectSingleNode, selectNodes } = require('../xml/xpath-utils');

/**
 * Parse document defaults from style XML
 * Extracts default paragraph and character properties
 * 
 * @param {Document} styleDoc - Style XML document
 * @returns {Object} - Document defaults for paragraph and character formatting
 */
function parseDocumentDefaults(styleDoc) {
  const defaults = {
    paragraph: {},
    character: {}
  };
  
  try {
    // Get document defaults section
    const docDefaultsNode = selectSingleNode("//w:docDefaults", styleDoc);
    if (!docDefaultsNode) {
      return defaults;
    }
    
    // Get paragraph defaults
    const pPrDefaultNode = selectSingleNode(".//w:pPrDefault/w:pPr", docDefaultsNode);
    if (pPrDefaultNode) {
      // We need to import these functions to avoid circular dependencies
      const { parseParagraphProperties } = require('./style-parser');
      defaults.paragraph = parseParagraphProperties(pPrDefaultNode);
    }
    
    // Get character defaults
    const rPrDefaultNode = selectSingleNode(".//w:rPrDefault/w:rPr", docDefaultsNode);
    if (rPrDefaultNode) {
      // We need to import these functions to avoid circular dependencies
      const { parseRunningProperties } = require('./style-parser');
      defaults.character = parseRunningProperties(rPrDefaultNode);
    }
  } catch (error) {
    console.error('Error parsing document defaults:', error);
  }
  
  return defaults;
}

/**
 * Parse document settings
 * Extracts settings like tab stops, character spacing, etc.
 * 
 * @param {Document} settingsDoc - Settings XML document
 * @returns {Object} - Document settings
 */
function parseSettings(settingsDoc) {
  const settings = {
    defaultTabStop: '720', // Default value (0.5 inch in twentieths of a point)
    characterSpacing: 'normal',
    doNotHyphenateCaps: false,
    rtlGutter: false,
    pageMargins: null // Will be populated from document.xml section properties
  };
  
  if (!settingsDoc) {
    return settings;
  }
  
  try {
    // Default tab stop
    const defaultTabStopNode = selectSingleNode("//w:defaultTabStop", settingsDoc);
    if (defaultTabStopNode) {
      const val = defaultTabStopNode.getAttribute('w:val');
      if (val) {
        settings.defaultTabStop = val;
      }
    }
    
    // Character spacing control
    const characterSpacingControlNode = selectSingleNode("//w:characterSpacingControl", settingsDoc);
    if (characterSpacingControlNode) {
      const val = characterSpacingControlNode.getAttribute('w:val');
      if (val) {
        settings.characterSpacing = val;
      }
    }
    
    // Do not hyphenate capital letters
    const doNotHyphenateCapsNode = selectSingleNode("//w:doNotHyphenateCaps", settingsDoc);
    if (doNotHyphenateCapsNode) {
      const val = doNotHyphenateCapsNode.getAttribute('w:val');
      settings.doNotHyphenateCaps = val === 'true' || val === '1';
    }
    
    // Right-to-left document gutter
    const rtlGutterNode = selectSingleNode("//w:rtlGutter", settingsDoc);
    if (rtlGutterNode) {
      const val = rtlGutterNode.getAttribute('w:val');
      settings.rtlGutter = val === 'true' || val === '1';
    }
  } catch (error) {
    console.error('Error parsing settings:', error);
  }
  
  return settings;
}

/**
 * Enhanced document structure analysis to capture more details
 * Analyzes document structure including TOC, lists, and paragraph patterns
 * 
 * @param {Document} documentDoc - Document XML
 * @param {Document} numberingDoc - Numbering XML document for reference
 * @returns {Object} - Document structure information
 */
function analyzeDocumentStructure(documentDoc, numberingDoc) {
  const structure = {
    hasToc: false,
    tocLocation: null,
    paragraphTypes: {},
    specialSections: [],
    lists: [],
    specialParagraphPatterns: []
  };
  
  try {
    // Check for TOC field
    const tocFieldNodes = selectNodes("//w:fldChar[@w:fldCharType='begin']/following-sibling::w:instrText[contains(., 'TOC')]", documentDoc);
    structure.hasToc = tocFieldNodes.length > 0;
    
    // Check for common patterns in paragraphs
    const paragraphNodes = selectNodes("//w:p", documentDoc);
    let currentList = null;
    let currentListId = null;
    let currentListLevel = -1;
    
    // Create maps to track paragraph patterns
    const paragraphPatterns = new Map();
    
    paragraphNodes.forEach((p, index) => {
      // Check for style information
      const pStyle = selectSingleNode("w:pPr/w:pStyle", p);
      const styleId = pStyle ? pStyle.getAttribute('w:val') : null;
      
      // Check for numbering information
      const numPr = selectSingleNode("w:pPr/w:numPr", p);
      const numId = numPr ? selectSingleNode("w:numId", numPr)?.getAttribute('w:val') : null;
      const ilvl = numPr ? selectSingleNode("w:ilvl", numPr)?.getAttribute('w:val') : null;
      
      // Get paragraph text
      const textNodes = selectNodes(".//w:t", p);
      let text = '';
      textNodes.forEach(t => {
        text += t.textContent || '';
      });
      
      // Check for TOC entries
      if (text.toLowerCase().includes('table of contents') || 
          text.toLowerCase().includes('contents') ||
          (styleId && styleId.toLowerCase().includes('toc'))) {
        structure.hasToc = true;
        structure.tocLocation = index;
      }
      
      // Analyze paragraph structure and patterns
      if (text.trim()) {
        // Get first line for pattern analysis
        const firstLine = text.split('\n')[0].trim();
        
        // Look for common structure patterns without hardcoding specific terms
        // Pattern 1: Word followed by "for" followed by more text
        if (/^\w+\s+for\s+.+$/i.test(firstLine)) {
          const pattern = "word_for_word";
          if (!paragraphPatterns.has(pattern)) {
            paragraphPatterns.set(pattern, { count: 0, examples: [] });
          }
          
          const info = paragraphPatterns.get(pattern);
          info.count++;
          
          if (info.examples.length < 3) {
            info.examples.push(firstLine);
          }
        }
        
        // Pattern 2: Word followed by comma
        if (/^\w+,\s*.*$/i.test(firstLine)) {
          const pattern = "word_comma";
          if (!paragraphPatterns.has(pattern)) {
            paragraphPatterns.set(pattern, { count: 0, examples: [] });
          }
          
          const info = paragraphPatterns.get(pattern);
          info.count++;
          
          if (info.examples.length < 3) {
            info.examples.push(firstLine);
          }
        }
        
        // Pattern 3: Word followed by parenthesized text
        if (/^\w+\s+\([^)]+\)\s*.*$/i.test(firstLine)) {
          const pattern = "word_parenthesis";
          if (!paragraphPatterns.has(pattern)) {
            paragraphPatterns.set(pattern, { count: 0, examples: [] });
          }
          
          const info = paragraphPatterns.get(pattern);
          info.count++;
          
          if (info.examples.length < 3) {
            info.examples.push(firstLine);
          }
        }
      }
      
      // Track style usage
      if (styleId) {
        if (!structure.paragraphTypes[styleId]) {
          structure.paragraphTypes[styleId] = { count: 0, samples: [] };
        }
        structure.paragraphTypes[styleId].count++;
        
        // Store a few samples of each style for analysis
        if (structure.paragraphTypes[styleId].samples.length < 3) {
          structure.paragraphTypes[styleId].samples.push(text.substring(0, 100));
        }
      }
      
      // Track numbering info for hierarchical lists
      if (numId && ilvl !== null) {
        const level = parseInt(ilvl, 10);
        
        // Check if this is a new list or continuation
        if (currentListId !== numId || level === 0) {
          // Start a new list
          currentListId = numId;
          currentList = {
            id: numId,
            startIndex: index,
            items: [],
            levels: {}
          };
          structure.lists.push(currentList);
        }
        
        // Track items at different levels
        if (!currentList.levels[level]) {
          currentList.levels[level] = {
            count: 0,
            items: []
          };
        }
        
        currentList.levels[level].count++;
        currentListLevel = level;
        
        // Add this item to the list
        const listItem = {
          index: index,
          level: level,
          text: text.trim(),
          hasNumId: !!numId,
          numId: numId,
          styleId: styleId
        };
        
        currentList.items.push(listItem);
        currentList.levels[level].items.push(listItem);
      } else if (currentList && text.trim()) {
        // Check if this is a non-numbered paragraph in a list - could be special
        let isSpecial = false;
        for (const [pattern, info] of paragraphPatterns.entries()) {
          if (info.count >= 2) {
            // This pattern appears multiple times, check if this paragraph matches
            if ((pattern === "word_for_word" && /^\w+\s+for\s+.+$/i.test(text)) ||
                (pattern === "word_comma" && /^\w+,\s*.*$/i.test(text)) ||
                (pattern === "word_parenthesis" && /^\w+\s+\([^)]+\)\s*.*$/i.test(text))) {
              isSpecial = true;
              
              // Add to current list as a special item
              currentList.items.push({
                index: index,
                level: currentListLevel, // Keep same level as previous item
                text: text.trim(),
                isSpecial: true,
                specialType: pattern
              });
              break;
            }
          }
        }
        
        // If not special, and not empty, this might end the list
        if (!isSpecial) {
          currentList = null;
          currentListId = null;
          currentListLevel = -1;
        }
      }
    });
    
    // Add detected paragraph patterns to structure
    for (const [pattern, info] of paragraphPatterns.entries()) {
      if (info.count >= 2) {
        // This pattern appears multiple times, so it's likely a structural element
        structure.specialParagraphPatterns.push({
          type: pattern,
          count: info.count,
          examples: info.examples
        });
      }
    }
    
  } catch (error) {
    console.error('Error analyzing document structure:', error);
  }
  
  return structure;
}

/**
 * Get default style information if parsing fails
 * Provides a fallback structure with sensible defaults
 * 
 * @returns {Object} - Default style information
 */
function getDefaultStyleInfo() {
  return {
    styles: {
      paragraph: {},
      character: {},
      table: {},
      numbering: {}
    },
    theme: {
      colors: {},
      fonts: {
        major: 'Calibri Light',
        minor: 'Calibri'
      }
    },
    documentDefaults: {
      paragraph: {},
      character: {
        fontSize: '11pt'
      }
    },
    settings: {
      defaultTabStop: '720',
      characterSpacing: 'normal',
      doNotHyphenateCaps: false,
      rtlGutter: false
    },
    tocStyles: {
      hasTableOfContents: false,
      tocHeadingStyle: {},
      tocEntryStyles: [],
      leaderStyle: {
        character: '.',
        spacesBetween: 3,
        position: 6 * 72 // Default position for right tab (6 inches)
      }
    },
    numberingDefs: {
      abstractNums: {},
      nums: {},
      numIdMap: {}
    },
    documentStructure: {
      hasToc: false,
      tocLocation: null,
      paragraphTypes: {},
      specialSections: [],
      lists: [],
      specialParagraphPatterns: []
    }
  };
}

/**
 * Extract page margins and section properties from document.xml
 * Analyzes sectPr elements to get page setup information
 * 
 * @param {Document} documentDoc - Document XML
 * @returns {Object} - Page margins and section properties
 */
function extractPageMargins(documentDoc) {
  const pageMargins = {
    top: 1440,    // Default 1 inch in twips
    bottom: 1440, // Default 1 inch in twips
    left: 1440,   // Default 1 inch in twips
    right: 1440,  // Default 1 inch in twips
    header: 720,  // Default 0.5 inch in twips
    footer: 720,  // Default 0.5 inch in twips
    gutter: 0     // Default no gutter
  };

  try {
    // Look for section properties in the document
    // Check both body sectPr and paragraph sectPr elements
    const sectPrNodes = selectNodes("//w:sectPr", documentDoc);
    
    if (sectPrNodes.length > 0) {
      // Use the last sectPr found (typically the main document section)
      const sectPr = sectPrNodes[sectPrNodes.length - 1];
      
      // Extract page margins
      const pgMarNode = selectSingleNode("w:pgMar", sectPr);
      if (pgMarNode) {
        const top = pgMarNode.getAttribute('w:top');
        const bottom = pgMarNode.getAttribute('w:bottom');
        const left = pgMarNode.getAttribute('w:left');
        const right = pgMarNode.getAttribute('w:right');
        const header = pgMarNode.getAttribute('w:header');
        const footer = pgMarNode.getAttribute('w:footer');
        const gutter = pgMarNode.getAttribute('w:gutter');
        
        if (top) pageMargins.top = parseInt(top, 10);
        if (bottom) pageMargins.bottom = parseInt(bottom, 10);
        if (left) pageMargins.left = parseInt(left, 10);
        if (right) pageMargins.right = parseInt(right, 10);
        if (header) pageMargins.header = parseInt(header, 10);
        if (footer) pageMargins.footer = parseInt(footer, 10);
        if (gutter) pageMargins.gutter = parseInt(gutter, 10);
      }
      
      // Extract page size for context
      const pgSzNode = selectSingleNode("w:pgSz", sectPr);
      if (pgSzNode) {
        const width = pgSzNode.getAttribute('w:w');
        const height = pgSzNode.getAttribute('w:h');
        const orient = pgSzNode.getAttribute('w:orient');
        
        pageMargins.pageSize = {
          width: width ? parseInt(width, 10) : 12240, // Default letter width in twips
          height: height ? parseInt(height, 10) : 15840, // Default letter height in twips
          orientation: orient || 'portrait'
        };
      }
    }
  } catch (error) {
    console.error('Error extracting page margins:', error);
  }
  
  return pageMargins;
}

/**
 * Enhanced parseSettings function that includes page margin extraction
 * 
 * @param {Document} settingsDoc - Settings XML document
 * @param {Document} documentDoc - Document XML for section properties
 * @returns {Object} - Enhanced document settings with page margins
 */
function parseSettingsWithMargins(settingsDoc, documentDoc) {
  const settings = parseSettings(settingsDoc);
  
  // Extract page margins from document.xml
  const pageMargins = extractPageMargins(documentDoc);
  settings.pageMargins = pageMargins;
  
  return settings;
}

module.exports = {
  parseDocumentDefaults,
  parseSettings,
  parseSettingsWithMargins,
  extractPageMargins,
  analyzeDocumentStructure,
  getDefaultStyleInfo
};