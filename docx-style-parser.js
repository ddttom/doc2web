// docx-style-parser.js - Enhanced to better extract DOCX styles
const JSZip = require('jszip');
const fs = require('fs');
const { DOMParser } = require('jsdom').JSDOM;
const xpath = require('xpath');
const dom = require('xmldom').DOMParser;

// Define XML namespaces used in DOCX files
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
 * @param {string} expression - XPath expression
 * @returns {xpath.XPathSelect} - XPath selector with namespaces
 */
function createXPathSelector(expression) {
  const select = xpath.useNamespaces(NAMESPACES);
  return select(expression);
}

/**
 * Select nodes using XPath with namespaces
 * @param {string} expression - XPath expression
 * @param {Document} doc - XML document
 * @returns {Node[]} - Selected nodes
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
 * @param {string} expression - XPath expression
 * @param {Document|Node} doc - XML document or node
 * @returns {Node|null} - Selected node or null
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

/**
 * Parse a DOCX file to extract detailed style information
 * @param {string} docxPath - Path to the DOCX file
 * @returns {Promise<Object>} - Detailed style information
 */
async function parseDocxStyles(docxPath) {
  try {
    // Read the DOCX file (which is a ZIP archive)
    const data = fs.readFileSync(docxPath);
    const zip = await JSZip.loadAsync(data);
    
    // Extract key files
    const styleXml = await zip.file('word/styles.xml')?.async('string');
    const documentXml = await zip.file('word/document.xml')?.async('string');
    const themeXml = await zip.file('word/theme/theme1.xml')?.async('string');
    const settingsXml = await zip.file('word/settings.xml')?.async('string');
    const numberingXml = await zip.file('word/numbering.xml')?.async('string');
    
    if (!styleXml || !documentXml) {
      throw new Error('Invalid DOCX file: missing core XML files');
    }
    
    // Parse XML content
    const styleDoc = new dom().parseFromString(styleXml);
    const documentDoc = new dom().parseFromString(documentXml);
    const themeDoc = themeXml ? new dom().parseFromString(themeXml) : null;
    const settingsDoc = settingsXml ? new dom().parseFromString(settingsXml) : null;
    const numberingDoc = numberingXml ? new dom().parseFromString(numberingXml) : null;
    
    // Parse styles from XML
    const styles = parseStyles(styleDoc);
    
    // Parse theme (colors, fonts)
    const theme = parseTheme(themeDoc);
    
    // Parse document defaults
    const documentDefaults = parseDocumentDefaults(styleDoc);
    
    // Parse document settings
    const settings = parseSettings(settingsDoc);
    
    // Parse TOC styles (improved)
    const tocStyles = parseTocStyles(documentDoc, styleDoc);
    
    // Parse numbering definitions (improved)
    const numberingDefs = parseNumberingDefinitions(numberingDoc);
    
    // Extract document structure (improved)
    const documentStructure = analyzeDocumentStructure(documentDoc, numberingDoc);
    
    // Return combined style information
    return {
      styles,
      theme,
      documentDefaults,
      settings,
      tocStyles,
      numberingDefs,
      documentStructure
    };
  } catch (error) {
    console.error('Error parsing DOCX styles:', error);
    // Return default style info if parsing fails
    return getDefaultStyleInfo();
  }
}

/**
 * IMPROVED: Enhanced TOC parsing to better capture leader lines and formatting
 * @param {Document} documentDoc - Document XML
 * @param {Document} styleDoc - Style XML
 * @returns {Object} - TOC specific style information
 */
function parseTocStyles(documentDoc, styleDoc) {
  const tocStyles = {
    hasTableOfContents: false,
    tocHeadingStyle: {},
    tocEntryStyles: [],
    leaderStyle: {
      character: '.',
      spacesBetween: 3,
      position: 0
    }
  };
  
  try {
    // Check for TOC field
    const tocFieldNodes = selectNodes("//w:fldChar[@w:fldCharType='begin']/following-sibling::w:instrText[contains(., 'TOC')]", documentDoc);
    tocStyles.hasTableOfContents = tocFieldNodes.length > 0;
    
    if (tocStyles.hasTableOfContents) {
      // Extract TOC field properties
      for (const tocNode of tocFieldNodes) {
        const instrText = tocNode.textContent || '';
        
        // Extract leader character if specified
        const leaderMatch = instrText.match(/\\[a-z]\s+([\.\_\-])/);
        if (leaderMatch) {
          tocStyles.leaderStyle.character = leaderMatch[1];
        }
        
        // Extract heading levels if specified
        const levelMatch = instrText.match(/\\o\s+"([1-9])"-"([1-9])"/);
        if (levelMatch) {
          tocStyles.startLevel = parseInt(levelMatch[1], 10);
          tocStyles.endLevel = parseInt(levelMatch[2], 10);
        }
      }
    }
    
    // Look for TOC styles in the style definitions - improved to catch more TOC styles
    const tocStyleNodes = selectNodes("//w:style[contains(@w:styleId, 'TOC') or contains(w:name/@w:val, 'TOC') or contains(w:name/@w:val, 'Contents')]", styleDoc);
    
    tocStyleNodes.forEach((node, index) => {
      const styleId = node.getAttribute('w:styleId');
      const nameNode = selectSingleNode("w:name", node);
      const name = nameNode ? nameNode.getAttribute('w:val') : styleId;
      
      // Parse specific style properties
      const pPrNode = selectSingleNode("w:pPr", node);
      const rPrNode = selectSingleNode("w:rPr", node);
      
      const style = {
        id: styleId,
        name,
        indentation: {},
        fontSize: null,
        fontFamily: null,
        tabs: [] // Enhanced to store tab information
      };
      
      // Extract paragraph properties
      if (pPrNode) {
        // Get indentation
        const indNode = selectSingleNode("w:ind", pPrNode);
        if (indNode) {
          const left = indNode.getAttribute('w:left');
          const hanging = indNode.getAttribute('w:hanging');
          const firstLine = indNode.getAttribute('w:firstLine');
          
          if (left) style.indentation.left = convertTwipToPt(left);
          if (hanging) style.indentation.hanging = convertTwipToPt(hanging);
          if (firstLine) style.indentation.firstLine = convertTwipToPt(firstLine);
        }
        
        // Get tab stops - IMPROVED to handle more tab types
        const tabsNode = selectSingleNode("w:tabs", pPrNode);
        if (tabsNode) {
          const tabNodes = selectNodes("w:tab", tabsNode);
          
          tabNodes.forEach(tabNode => {
            const pos = tabNode.getAttribute('w:pos');
            const val = tabNode.getAttribute('w:val');
            const leader = tabNode.getAttribute('w:leader');
            
            if (pos && val) {
              const tabPosition = convertTwipToPt(pos);
              style.tabs.push({
                position: tabPosition,
                type: val,
                leader: leader || 'none'
              });
              
              // If we find a right-aligned tab with leader, use it for TOC dots
              if (val === 'right' && (leader === 'dot' || leader === '1')) {
                tocStyles.leaderStyle.character = '.';
                tocStyles.leaderStyle.position = tabPosition;
              }
            }
          });
        }
      }
      
      // Extract text properties
      if (rPrNode) {
        // Font size
        const szNode = selectSingleNode("w:sz", rPrNode);
        if (szNode) {
          const size = parseInt(szNode.getAttribute('w:val'), 10) || 22; // Default 11pt
          style.fontSize = (size / 2) + 'pt';
        }
        
        // Font family
        const fontNode = selectSingleNode("w:rFonts", rPrNode);
        if (fontNode) {
          style.fontFamily = fontNode.getAttribute('w:ascii') || 
                           fontNode.getAttribute('w:hAnsi') || 
                           'Calibri';
        }
      }
      
      // Determine if this is a TOC heading or entry style
      if (name.includes('Heading') || styleId === 'TOCHeading') {
        tocStyles.tocHeadingStyle = style;
      } else {
        // Extract level info from the style ID if possible
        const levelMatch = styleId.match(/TOC(\d+)/);
        if (levelMatch) {
          style.level = parseInt(levelMatch[1], 10);
        } else {
          // Default to sequential level based on order
          style.level = index + 1;
        }
        
        tocStyles.tocEntryStyles.push(style);
      }
    });
    
    // Sort TOC entry styles by level
    tocStyles.tocEntryStyles.sort((a, b) => a.level - b.level);
    
    // IMPROVED: Scan document for actual TOC entries to get better tab/dot information
    const paragraphNodes = selectNodes("//w:p", documentDoc);
    let inTocSection = false;
    
    for (let i = 0; i < paragraphNodes.length; i++) {
      const p = paragraphNodes[i];
      
      // Check if this is a TOC heading paragraph
      const pStyle = selectSingleNode("w:pPr/w:pStyle", p);
      const styleId = pStyle ? pStyle.getAttribute('w:val') : '';
      
      // Start of TOC detection
      if (styleId === 'TOCHeading' || styleId.startsWith('TOC') || 
          (selectSingleNode(".//w:t[contains(text(), 'Table of Contents')]", p) !== null)) {
        inTocSection = true;
        continue;
      }
      
      // If we're in TOC section, analyze the paragraph for tab stops and dots
      if (inTocSection) {
        // Get text content to check if it looks like TOC entry
        const textNodes = selectNodes(".//w:t", p);
        let text = '';
        textNodes.forEach(t => text += t.textContent || '');
        
        // Check if this paragraph has page number pattern (text....number)
        const pageNumMatch = text.match(/^(.*?)[\.\s]*(\d+)$/);
        
        if (pageNumMatch) {
          // This looks like a TOC entry with page number
          const tabsNode = selectSingleNode("w:pPr/w:tabs", p);
          
          if (tabsNode) {
            const rightTab = selectSingleNode("w:tab[@w:val='right']", tabsNode);
            if (rightTab) {
              const leader = rightTab.getAttribute('w:leader');
              const pos = rightTab.getAttribute('w:pos');
              
              if (leader) {
                tocStyles.leaderStyle.character = leader === '1' ? '.' : 
                                              leader === '2' ? '-' : 
                                              leader === '3' ? '_' : '.';
              }
              
              if (pos) {
                tocStyles.leaderStyle.position = convertTwipToPt(pos);
              }
            }
          }
        } else if (text.trim() === '' || styleId === 'Normal') {
          // Empty line might indicate end of TOC
          inTocSection = false;
        }
      }
    }
    
  } catch (error) {
    console.error('Error parsing TOC styles:', error);
  }
  
  return tocStyles;
}

/**
 * IMPROVED: Enhanced numbering definition parsing to capture more details
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
          isLegal: false,  // Added property
          textFormat: '',   // Added property for display
          restart: null,    // Added property
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
          
          // Analyze text format for display
          if (level.text) {
            // Store raw format for later use
            level.textFormat = level.text;
            
            // Replace placeholders with computed values for CSS content property
            let cssText = level.text
              .replace(/%(\d+)/g, 'counter(item$1)')
              .replace(/^\s+|\s+$/g, '');
            
            // Add specific handling for common formats
            if (level.format === 'decimal') {
              // No transformation needed for decimal
            } else if (level.format === 'lowerLetter') {
              // a, b, c format
              cssText = cssText.replace(/counter\(item(\d+)\)/g, 'counter(item$1, lower-alpha)');
            } else if (level.format === 'upperLetter') {
              // A, B, C format
              cssText = cssText.replace(/counter\(item(\d+)\)/g, 'counter(item$1, upper-alpha)');
            } else if (level.format === 'lowerRoman') {
              // i, ii, iii format
              cssText = cssText.replace(/counter\(item(\d+)\)/g, 'counter(item$1, lower-roman)');
            } else if (level.format === 'upperRoman') {
              // I, II, III format
              cssText = cssText.replace(/counter\(item(\d+)\)/g, 'counter(item$1, upper-roman)');
            }
            
            level.cssText = cssText;
          }
        }
        
        // Get if legal style numbering
        const isLegalNode = selectSingleNode("w:isLgl", lvlNode);
        if (isLegalNode) {
          level.isLegal = isLegalNode.getAttribute('w:val') === '1';
        }
        
        // Get alignment
        const alignmentNode = selectSingleNode("w:lvlJc", lvlNode);
        if (alignmentNode) {
          level.alignment = alignmentNode.getAttribute('w:val') || 'left';
        }
        
        // Get indentation - IMPROVED with more detail
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
 * IMPROVED: Enhanced document structure analysis to capture more details
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
    hasRationale: false,
    specialParagraphPatterns: [] // Store regex patterns for special paragraph types
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
    
    // Find patterns for special paragraphs by analyzing the document
    const textPatterns = {};
    const specialPhrases = [
      'Rationale for Resolution',
      'Whereas,',
      'Resolved',
      'RESOLVED'
    ];
    
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
      if (text.includes('Table of Contents') || 
          text.includes('Contents') ||
          (styleId && styleId.toLowerCase().includes('toc'))) {
        structure.hasToc = true;
        structure.tocLocation = index;
      }
      
      // Identify patterns for special paragraphs
      specialPhrases.forEach(phrase => {
        if (text.includes(phrase)) {
          if (!textPatterns[phrase]) {
            textPatterns[phrase] = { count: 0, indexes: [] };
          }
          textPatterns[phrase].count++;
          textPatterns[phrase].indexes.push(index);
          
          // Check for "Rationale" specifically
          if (phrase === 'Rationale for Resolution') {
            structure.hasRationale = true;
            
            // Add to special sections list
            structure.specialSections.push({
              type: 'rationale',
              index: index,
              text: text.trim()
            });
          }
        }
      });
      
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
      
      // Track numbering info for hierarchical lists - IMPROVED to track list structure
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
      } else {
        // Not a list item - check if it interrupts a list
        if (currentList && text.trim()) {
          // Check if this looks like a special paragraph within a list
          let isSpecial = false;
          for (const phrase of specialPhrases) {
            if (text.includes(phrase)) {
              isSpecial = true;
              
              // Add to current list as a special item
              currentList.items.push({
                index: index,
                level: currentListLevel, // Keep same level as previous item
                text: text.trim(),
                isSpecial: true,
                specialType: phrase.split(' ')[0].toLowerCase() // e.g., "rationale", "whereas", "resolved"
              });
              break;
            }
          }
          
          // If not special, and not empty, this might end the list
          if (!isSpecial && !text.match(/^\s*$/)) {
            currentList = null;
            currentListId = null;
            currentListLevel = -1;
          }
        }
      }
    });
    
    // Analyze patterns and register common special paragraph patterns
    for (const [phrase, data] of Object.entries(textPatterns)) {
      if (data.count >= 2) {
        // This appears to be a recurring pattern
        let pattern;
        
        if (phrase === 'Rationale for Resolution') {
          pattern = /^Rationale for Resolution/i;
        } else if (phrase === 'Whereas,') {
          pattern = /^Whereas,/i;
        } else if (phrase === 'Resolved' || phrase === 'RESOLVED') {
          pattern = /^Resolved \([\d\.]+\)/i;
        }
        
        if (pattern) {
          structure.specialParagraphPatterns.push({
            type: phrase.split(' ')[0].toLowerCase(),
            pattern: pattern.toString(),
            count: data.count
          });
        }
      }
    }
    
  } catch (error) {
    console.error('Error analyzing document structure:', error);
  }
  
  return structure;
}

/**
 * Get default style information if parsing fails
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
      hasRationale: false,
      specialParagraphPatterns: []
    }
  };
}

/**
 * Parse styles.xml to extract style definitions
 * @param {Document} styleDoc - XML document
 * @returns {Object} - Style definitions
 */
function parseStyles(styleDoc) {
  const styles = {
    paragraph: {},
    character: {},
    table: {},
    numbering: {}
  };
  
  try {
    // XPath for different style types
    const paragraphStylesXPath = "//w:style[@w:type='paragraph']";
    const characterStylesXPath = "//w:style[@w:type='character']";
    const tableStylesXPath = "//w:style[@w:type='table']";
    const numberingStylesXPath = "//w:style[@w:type='numbering']";
    
    // Select nodes using namespace-aware selectors
    const paragraphStyleNodes = selectNodes(paragraphStylesXPath, styleDoc);
    const characterStyleNodes = selectNodes(characterStylesXPath, styleDoc);
    const tableStyleNodes = selectNodes(tableStylesXPath, styleDoc);
    const numberingStyleNodes = selectNodes(numberingStylesXPath, styleDoc);
    
    // Process paragraph styles
    paragraphStyleNodes.forEach(node => {
      const style = parseStyleNode(node);
      styles.paragraph[style.id] = style;
    });
    
    // Process character styles
    characterStyleNodes.forEach(node => {
      const style = parseStyleNode(node);
      styles.character[style.id] = style;
    });
    
    // Process table styles
    tableStyleNodes.forEach(node => {
      const style = parseStyleNode(node);
      styles.table[style.id] = style;
    });
    
    // Process numbering styles
    numberingStyleNodes.forEach(node => {
      const style = parseStyleNode(node);
      styles.numbering[style.id] = style;
    });
  } catch (error) {
    console.error('Error parsing styles:', error);
  }
  
  return styles;
}

/**
 * Parse individual style node
 * @param {Node} node - XML node for a style
 * @returns {Object} - Style properties
 */
function parseStyleNode(node) {
  try {
    const styleId = node.getAttribute('w:styleId');
    const type = node.getAttribute('w:type');
    
    // Get style name
    const nameNode = selectSingleNode("w:name", node);
    const name = nameNode ? nameNode.getAttribute('w:val') : styleId;
    
    // Get based on style
    const basedOnNode = selectSingleNode("w:basedOn", node);
    const basedOn = basedOnNode ? basedOnNode.getAttribute('w:val') : null;
    
    // Check if default style
    const defaultNode = selectSingleNode("@w:default", node);
    const isDefault = defaultNode && defaultNode.value === '1';
    
    // Parse running properties (text formatting)
    const rPrNode = selectSingleNode("w:rPr", node);
    const runningProps = rPrNode ? parseRunningProperties(rPrNode) : {};
    
    // Parse paragraph properties
    const pPrNode = selectSingleNode("w:pPr", node);
    const paragraphProps = pPrNode ? parseParagraphProperties(pPrNode) : {};
    
    // Combine all properties
    return {
      id: styleId,
      type,
      name,
      basedOn,
      isDefault,
      ...runningProps,
      ...paragraphProps
    };
  } catch (error) {
    console.error('Error parsing style node:', error);
    return { id: 'default', name: 'Default', type: 'paragraph' };
  }
}

/**
 * Parse running properties (text formatting)
 * @param {Node} rPrNode - Running properties node
 * @returns {Object} - Running properties
 */
function parseRunningProperties(rPrNode) {
  const props = {};
  
  try {
    // Font
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
      // Size is in half-points, convert to points
      const sizeHalfPoints = parseInt(szNode.getAttribute('w:val'), 10) || 22; // Default 11pt
      props.fontSize = (sizeHalfPoints / 2) + 'pt';
    }
    
    // Bold
    const bNode = selectSingleNode("w:b", rPrNode);
    props.bold = bNode !== null;
    
    // Italic
    const iNode = selectSingleNode("w:i", rPrNode);
    props.italic = iNode !== null;
    
    // Underline
    const uNode = selectSingleNode("w:u", rPrNode);
    if (uNode) {
      props.underline = {
        type: uNode.getAttribute('w:val') || 'single'
      };
    }
    
    // Color
    const colorNode = selectSingleNode("w:color", rPrNode);
    if (colorNode) {
      const colorVal = colorNode.getAttribute('w:val');
      if (colorVal) {
        props.color = '#' + colorVal;
      }
    }
    
    // Highlight
    const highlightNode = selectSingleNode("w:highlight", rPrNode);
    if (highlightNode) {
      props.highlight = highlightNode.getAttribute('w:val');
    }
  } catch (error) {
    console.error('Error parsing running properties:', error);
  }
  
  return props;
}

/**
 * Parse paragraph properties
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Paragraph properties
 */
function parseParagraphProperties(pPrNode) {
  const props = {};
  
  try {
    // Alignment
    const jcNode = selectSingleNode("w:jc", pPrNode);
    if (jcNode) {
      props.alignment = jcNode.getAttribute('w:val');
    }
    
    // Indentation
    const indNode = selectSingleNode("w:ind", pPrNode);
    if (indNode) {
      props.indentation = {
        left: indNode.getAttribute('w:left'),
        right: indNode.getAttribute('w:right'),
        firstLine: indNode.getAttribute('w:firstLine'),
        hanging: indNode.getAttribute('w:hanging')
      };
    }
    
    // Tab stops - IMPROVED to capture more details
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
            leader: leader || 'none',
            leaderChar: getLeaderChar(leader)
          });
        }
      });
    }
    
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
    
    // Borders
    const pBdrNode = selectSingleNode("w:pBdr", pPrNode);
    if (pBdrNode) {
      props.borders = parseBorders(pBdrNode);
    }
    
    // Shading
    const shadingNode = selectSingleNode("w:shd", pPrNode);
    if (shadingNode) {
      props.shading = {
        value: shadingNode.getAttribute('w:val'),
        color: shadingNode.getAttribute('w:color'),
        fill: shadingNode.getAttribute('w:fill')
      };
    }
    
    // Numbering properties - IMPROVED
    const numPrNode = selectSingleNode("w:numPr", pPrNode);
    if (numPrNode) {
      const numId = selectSingleNode("w:numId", numPrNode);
      const ilvl = selectSingleNode("w:ilvl", numPrNode);
      
      if (numId) {
        props.numbering = {
          id: numId.getAttribute('w:val'),
          level: ilvl ? ilvl.getAttribute('w:val') : '0'
        };
      }
    }
  } catch (error) {
    console.error('Error parsing paragraph properties:', error);
  }
  
  return props;
}

/**
 * Get leader character based on leader type
 * @param {string} leader - Leader type
 * @returns {string} - Leader character
 */
function getLeaderChar(leader) {
  if (!leader) return '';
  
  switch (leader) {
    case 'dot':
    case '1':
      return '.';
    case 'hyphen':
    case '2':
      return '-';
    case 'underscore':
    case '3':
      return '_';
    case 'heavy':
    case '4':
      return '=';
    case 'middleDot':
    case '5':
      return '·';
    default:
      return '';
  }
}

/**
 * Parse border information
 * @param {Node} borderNode - Border properties node
 * @returns {Object} - Border properties
 */
function parseBorders(borderNode) {
  const borders = {};
  
  try {
    // Top border
    const topNode = selectSingleNode("w:top", borderNode);
    if (topNode) {
      borders.top = {
        value: topNode.getAttribute('w:val'),
        size: topNode.getAttribute('w:sz'),
        color: topNode.getAttribute('w:color')
      };
    }
    
    // Bottom border
    const bottomNode = selectSingleNode("w:bottom", borderNode);
    if (bottomNode) {
      borders.bottom = {
        value: bottomNode.getAttribute('w:val'),
        size: bottomNode.getAttribute('w:sz'),
        color: bottomNode.getAttribute('w:color')
      };
    }
    
    // Left border
    const leftNode = selectSingleNode("w:left", borderNode);
    if (leftNode) {
      borders.left = {
        value: leftNode.getAttribute('w:val'),
        size: leftNode.getAttribute('w:sz'),
        color: leftNode.getAttribute('w:color')
      };
    }
    
    // Right border
    const rightNode = selectSingleNode("w:right", borderNode);
    if (rightNode) {
      borders.right = {
        value: rightNode.getAttribute('w:val'),
        size: rightNode.getAttribute('w:sz'),
        color: rightNode.getAttribute('w:color')
      };
    }
  } catch (error) {
    console.error('Error parsing borders:', error);
  }
  
  return borders;
}

/**
 * Parse theme information
 * @param {Document} themeDoc - Theme XML document
 * @returns {Object} - Theme properties
 */
function parseTheme(themeDoc) {
  // Default theme structure
  const theme = {
    colors: {},
    fonts: {
      major: 'Calibri Light',
      minor: 'Calibri'
    }
  };
  
  if (!themeDoc) {
    return theme;
  }
  
  try {
    // Parse font scheme
    const fontSchemeNode = selectSingleNode("//a:fontScheme", themeDoc);
    if (fontSchemeNode) {
      const majorNode = selectSingleNode(".//a:majorFont/a:latin", fontSchemeNode);
      const minorNode = selectSingleNode(".//a:minorFont/a:latin", fontSchemeNode);
      
      if (majorNode && majorNode.getAttribute('typeface')) {
        theme.fonts.major = majorNode.getAttribute('typeface');
      }
      
      if (minorNode && minorNode.getAttribute('typeface')) {
        theme.fonts.minor = minorNode.getAttribute('typeface');
      }
    }
    
    // Parse color scheme
    const clrSchemeNode = selectSingleNode("//a:clrScheme", themeDoc);
    if (clrSchemeNode) {
      // Get main theme colors
      const colorNodes = selectNodes("./a:*", clrSchemeNode);
      colorNodes.forEach(node => {
        const name = node.nodeName.split(':')[1];
        const colorValue = getColorValue(node);
        if (colorValue) {
          theme.colors[name] = colorValue;
        }
      });
    }
  } catch (error) {
    console.error('Error parsing theme:', error);
  }
  
  return theme;
}

/**
 * Get color value from theme color node
 * @param {Node} colorNode - Color definition node
 * @returns {string} - Color value
 */
function getColorValue(colorNode) {
  try {
    // Check srgbClr (standard RGB)
    const srgbNode = selectSingleNode("./a:srgbClr", colorNode);
    if (srgbNode && srgbNode.getAttribute('val')) {
      return '#' + srgbNode.getAttribute('val');
    }
    
    // Check system color
    const sysClrNode = selectSingleNode("./a:sysClr", colorNode);
    if (sysClrNode && sysClrNode.getAttribute('lastClr')) {
      return '#' + sysClrNode.getAttribute('lastClr');
    }
  } catch (error) {
    console.error('Error getting color value:', error);
  }
  
  return null;
}

/**
 * Parse document defaults
 * @param {Document} styleDoc - Style XML document
 * @returns {Object} - Document defaults
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
      defaults.paragraph = parseParagraphProperties(pPrDefaultNode);
    }
    
    // Get character defaults
    const rPrDefaultNode = selectSingleNode(".//w:rPrDefault/w:rPr", docDefaultsNode);
    if (rPrDefaultNode) {
      defaults.character = parseRunningProperties(rPrDefaultNode);
    }
  } catch (error) {
    console.error('Error parsing document defaults:', error);
  }
  
  return defaults;
}

/**
 * Parse document settings
 * @param {Document} settingsDoc - Settings XML document
 * @returns {Object} - Document settings
 */
function parseSettings(settingsDoc) {
  const settings = {
    defaultTabStop: '720', // Default value (0.5 inch in twentieths of a point)
    characterSpacing: 'normal',
    doNotHyphenateCaps: false,
    rtlGutter: false
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
 * IMPROVED: Generate CSS from extracted style information with better TOC and list handling
 * @param {Object} styleInfo - Style information
 * @returns {string} - CSS stylesheet
 */
function generateCssFromStyleInfo(styleInfo) {
  let css = '';
  
  try {
    // Add document defaults
    css += `
/* Document defaults */
body {
  font-family: "${styleInfo.theme.fonts.minor}", sans-serif;
  font-size: ${styleInfo.documentDefaults.character.fontSize || '11pt'};
  line-height: 1.15;
  margin: 20px;
  padding: 0;
}
`;

    // Add paragraph styles
    Object.entries(styleInfo.styles.paragraph || {}).forEach(([id, style]) => {
      const className = `docx-p-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`;
      
      css += `
.${className} {
  ${style.font ? `font-family: "${getFontFamily(style, styleInfo)}", sans-serif;` : ''}
  ${style.fontSize ? `font-size: ${style.fontSize};` : ''}
  ${style.bold ? 'font-weight: bold;' : ''}
  ${style.italic ? 'font-style: italic;' : ''}
  ${style.color ? `color: ${style.color};` : ''}
  ${style.alignment ? `text-align: ${style.alignment};` : ''}
  ${style.spacing?.before ? `margin-top: ${convertTwipToPt(style.spacing.before)}pt;` : ''}
  ${style.spacing?.after ? `margin-bottom: ${convertTwipToPt(style.spacing.after)}pt;` : ''}
  ${style.indentation?.left ? `margin-left: ${convertTwipToPt(style.indentation.left)}pt;` : ''}
  ${style.indentation?.right ? `margin-right: ${convertTwipToPt(style.indentation.right)}pt;` : ''}
  ${style.indentation?.firstLine ? `text-indent: ${convertTwipToPt(style.indentation.firstLine)}pt;` : ''}
}
`;
    });

    // Add character styles
    Object.entries(styleInfo.styles.character || {}).forEach(([id, style]) => {
      const className = `docx-c-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`;
      
      css += `
.${className} {
  ${style.font ? `font-family: "${getFontFamily(style, styleInfo)}", sans-serif;` : ''}
  ${style.fontSize ? `font-size: ${style.fontSize};` : ''}
  ${style.bold ? 'font-weight: bold;' : ''}
  ${style.italic ? 'font-style: italic;' : ''}
  ${style.color ? `color: ${style.color};` : ''}
  ${style.underline ? `text-decoration: ${style.underline.type === 'none' ? 'none' : 'underline'};` : ''}
}
`;
    });

    // Add table styles
    Object.entries(styleInfo.styles.table || {}).forEach(([id, style]) => {
      const className = `docx-t-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`;
      
      css += `
.${className} {
  border-collapse: collapse;
  width: 100%;
}
.${className} td, .${className} th {
  padding: 5pt;
  ${getBorderStyle(style, 'top')}
  ${getBorderStyle(style, 'bottom')}
  ${getBorderStyle(style, 'left')}
  ${getBorderStyle(style, 'right')}
}
`;
    });

    // IMPROVED: Enhanced TOC styles based on extracted information
    if (styleInfo.tocStyles) {
      const tocHeaderStyle = styleInfo.tocStyles.tocHeadingStyle;
      const tocEntryStyles = styleInfo.tocStyles.tocEntryStyles;
      const leaderStyle = styleInfo.tocStyles.leaderStyle;
      
      // Main TOC container
      css += `
/* Table of Contents Styles */
.docx-toc {
  margin: 20px 0;
  width: 100%;
}

.docx-toc-heading {
  ${tocHeaderStyle.fontFamily ? `font-family: "${tocHeaderStyle.fontFamily}", sans-serif;` : ''}
  ${tocHeaderStyle.fontSize ? `font-size: ${tocHeaderStyle.fontSize};` : 'font-size: 14pt;'}
  font-weight: bold;
  margin-bottom: 12pt;
  text-align: center;
}
`;

      // TOC entries with leader lines
      css += `
.docx-toc-entry {
  display: flex;
  flex-wrap: nowrap;
  align-items: baseline;
  width: 100%;
  margin-bottom: 4pt;
  position: relative;
  overflow: hidden;
  text-align: left;
}

.docx-toc-text {
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  flex-grow: 0;
  flex-shrink: 1;
}

.docx-toc-dots {
  flex-grow: 1;
  margin: 0 4pt;
  height: 1em;
  position: relative;
  display: inline-block;
  border-bottom: 1px dotted #999;
  min-width: 2em;
}

.docx-toc-pagenum {
  flex-shrink: 0;
  text-align: right;
  white-space: nowrap;
  min-width: 2em;
}
`;

      // Specific styles for different TOC levels
      tocEntryStyles.forEach(style => {
        const level = style.level || 1;
        const leftIndent = style.indentation?.left ? 
                         convertTwipToPt(style.indentation.left) : 
                         (level - 1) * 20;
        
        css += `
.docx-toc-level-${level} {
  ${style.fontFamily ? `font-family: "${style.fontFamily}", sans-serif;` : ''}
  ${style.fontSize ? `font-size: ${style.fontSize};` : ''}
  margin-left: ${leftIndent}pt;
}
`;
      });
    }

    // IMPROVED: Enhanced list styling based on extracted numbering definitions
    if (styleInfo.numberingDefs && Object.keys(styleInfo.numberingDefs.abstractNums).length > 0) {
      // Basic list containers
      css += `
/* List and Numbering Styles */
ol.docx-numbered-list,
ol.docx-alpha-list {
  list-style-type: none;
  padding-left: 0;
  margin-top: 0.5em;
  margin-bottom: 0.5em;
}

ol.docx-numbered-list > li,
ol.docx-alpha-list > li {
  position: relative;
  padding-left: 2.5em;
  margin-bottom: 0.5em;
  min-height: 1.5em;
}

ol.docx-numbered-list > li::before,
ol.docx-alpha-list > li::before {
  position: absolute;
  left: 0;
  content: attr(data-prefix);
  font-weight: bold;
  display: inline-block;
  min-width: 2em;
}
`;

      // Generate specific styles for each numbering definition
      const abstractNums = styleInfo.numberingDefs.abstractNums;
      
      Object.entries(abstractNums).forEach(([id, abstractNum]) => {
        Object.entries(abstractNum.levels).forEach(([level, levelDef]) => {
          const levelNum = parseInt(level, 10);
          const className = `docx-num-${id}-${level}`;
          
          // Calculate indentation
          const leftIndent = levelDef.indentation?.left ? 
                           convertTwipToPt(levelDef.indentation.left) : 
                           levelNum * 36; // Default 36pt per level
                           
          const hangingIndent = levelDef.indentation?.hanging ? 
                              convertTwipToPt(levelDef.indentation.hanging) : 
                              24; // Default 24pt
          
          css += `
li.${className} {
  padding-left: ${leftIndent}pt;
  text-indent: -${hangingIndent}pt;
}

li.${className}::before {
  min-width: ${hangingIndent}pt;
  text-align: ${levelDef.alignment || 'left'};
}
`;
        });
      });
      
      // Special handling for different list formats
      css += `
/* Specific list type formatting */
.docx-decimal-list > li::before {
  content: attr(data-prefix) ".";
}

.docx-alpha-list > li::before {
  content: attr(data-prefix) ".";
}

.docx-roman-list > li::before {
  content: attr(data-prefix) ".";
}

/* Special paragraph types within lists */
.docx-list-item-special {
  font-style: italic;
  margin-left: 2.5em;
  margin-top: 0.5em;
  margin-bottom: 0.5em;
}

.docx-rationale {
  font-style: italic;
  color: #555;
  margin-left: 2.5em;
  margin-top: 0.5em;
  margin-bottom: 0.7em;
  border-left: 2px solid #ddd;
  padding-left: 0.7em;
}

.docx-whereas {
  font-weight: normal;
}

.docx-resolved {
  font-weight: bold;
}
`;
    }

    // Add utility styles
    css += `
/* Utility styles */
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }
.docx-tab { display: inline-block; width: ${convertTwipToPt(styleInfo.settings.defaultTabStop)}pt; }
.docx-rtl { direction: rtl; }
.docx-image { max-width: 100%; height: auto; display: block; margin: 10px 0; }
.docx-table-default { width: 100%; border-collapse: collapse; margin: 10px 0; }
.docx-table-default td, .docx-table-default th { border: 1px solid #ddd; padding: 5pt; }
.docx-placeholder { 
  font-weight: bold; 
  text-align: center; 
  padding: 15px; 
  margin: 20px 0; 
  background-color: #f0f0f0; 
  border: 1px dashed #999; 
  border-radius: 5px; 
}

/* Special paragraph styles */
.docx-heading1, .docx-heading2, .docx-heading3, .docx-heading4, .docx-heading5, .docx-heading6 {
  font-family: "${styleInfo.theme.fonts.major}", sans-serif;
  color: #2F5496;
  margin-top: 1.2em;
  margin-bottom: 0.6em;
}

.docx-heading1 { font-size: 16pt; }
.docx-heading2 { font-size: 14pt; }
.docx-heading3 { font-size: 13pt; font-style: italic; }
.docx-heading4 { font-size: 12pt; }
.docx-heading5 { font-size: 11pt; font-style: italic; }
.docx-heading6 { font-size: 11pt; }

/* Document footer */
.docx-footer {
  text-align: center;
  margin-top: 2em;
  font-size: 0.9em;
  color: #666;
}

/* Improved table styles */
.table-responsive {
  overflow-x: auto;
  margin: 1em 0;
}
`;

  } catch (error) {
    console.error('Error generating CSS:', error);
    
    // Provide fallback CSS
    css = `
body { font-family: Calibri, sans-serif; font-size: 11pt; line-height: 1.15; margin: 20px; }
h1, h2, h3, h4, h5, h6 { font-family: "Calibri Light", sans-serif; color: #2F5496; }
h1 { font-size: 16pt; }
h2 { font-size: 13pt; }
h3 { font-size: 12pt; }
p { margin: 10pt 0; }
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }
.docx-rtl { direction: rtl; }
.docx-image { max-width: 100%; height: auto; display: block; margin: 10px 0; }
table { width: 100%; border-collapse: collapse; margin: 10px 0; }
td, th { border: 1px solid #ddd; padding: 5pt; }

/* Basic TOC styles */
.docx-toc-entry {
  display: flex;
  flex-wrap: nowrap;
  align-items: baseline;
  width: 100%;
  margin-bottom: 4pt;
}
.docx-toc-text {
  white-space: nowrap;
  overflow: hidden;
  flex-grow: 0;
}
.docx-toc-dots {
  flex-grow: 1;
  margin: 0 4pt;
  height: 1em;
  border-bottom: 1px dotted #999;
}
.docx-toc-pagenum {
  flex-shrink: 0;
  text-align: right;
}

/* Basic list styles */
ol.docx-numbered-list { counter-reset: item; list-style-type: none; padding-left: 0; }
ol.docx-numbered-list li { counter-increment: item; position: relative; padding-left: 2em; margin-bottom: 0.5em; }
ol.docx-numbered-list li::before { position: absolute; left: 0; content: attr(data-prefix); font-weight: bold; }
ol.docx-alpha-list { list-style-type: none; padding-left: 0; }
ol.docx-alpha-list li { position: relative; padding-left: 2em; margin-bottom: 0.5em; }
ol.docx-alpha-list li::before { position: absolute; left: 0; content: attr(data-prefix) "."; font-weight: bold; }

/* Special styles */
.docx-rationale { font-style: italic; margin-left: 2em; }
`;
  }

  return css;
}

/**
 * Get font family based on style and theme
 * @param {Object} style - Style information
 * @param {Object} styleInfo - Style information including theme
 * @returns {string} - Font family
 */
function getFontFamily(style, styleInfo) {
  if (!style.font) {
    return styleInfo.theme.fonts.minor;
  }
  
  // Use ASCII font if available, fall back to theme
  return style.font.ascii || style.font.hAnsi || styleInfo.theme.fonts.minor;
}

/**
 * Get border style
 * @param {Object} style - Style information
 * @param {string} side - Border side (top, bottom, left, right)
 * @returns {string} - CSS border property
 */
function getBorderStyle(style, side) {
  if (!style.borders || !style.borders[side]) {
    return '';
  }
  
  const border = style.borders[side];
  if (!border || !border.value || border.value === 'nil' || border.value === 'none') {
    return `border-${side}: none;`;
  }
  
  const width = border.size ? convertBorderSizeToPt(border.size) : 1;
  const color = border.color ? `#${border.color}` : 'black';
  const borderStyle = getBorderTypeValue(border.value);
  
  return `border-${side}: ${width}pt ${borderStyle} ${color};`;
}

/**
 * Convert border size to point value
 * @param {string} size - Border size value
 * @returns {number} - Size in points
 */
function convertBorderSizeToPt(size) {
  const sizeNum = parseInt(size, 10) || 1;
  // Border size is in 1/8th points
  return sizeNum / 8;
}

/**
 * Get CSS border style from Word border value
 * @param {string} value - Word border value
 * @returns {string} - CSS border style
 */
function getBorderTypeValue(value) {
  const borderTypes = {
    'single': 'solid',
    'double': 'double',
    'dotted': 'dotted',
    'dashed': 'dashed',
    'wavy': 'wavy',
    'dashSmallGap': 'dashed',
    'dotDash': 'dashed',
    'dotDotDash': 'dashed',
    'triple': 'double',
    'thinThickSmallGap': 'double',
    'thickThinSmallGap': 'double'
  };
  
  return borderTypes[value] || 'solid';
}

/**
 * Convert twip value to points
 * @param {string} twip - Value in twentieths of a point
 * @returns {number} - Value in points
 */
function convertTwipToPt(twip) {
  const twipNum = parseInt(twip, 10) || 0;
  // 20 twips = 1 point
  return twipNum / 20;
}

module.exports = {
  parseDocxStyles,
  generateCssFromStyleInfo,
  createXPathSelector,
  selectNodes,
  selectSingleNode,
  convertTwipToPt
};
