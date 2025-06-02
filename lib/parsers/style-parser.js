// lib/parsers/style-parser.js - Enhanced style parsing with numbering context integration
const JSZip = require('jszip');
const fs = require('fs');
const { JSDOM } = require('jsdom');
const { DOMParser } = require('xmldom');
const { selectNodes, selectSingleNode } = require('../xml/xpath-utils');
const { convertTwipToPt } = require('../utils/unit-converter');
const { getLeaderChar } = require('../utils/common-utils');
const { extractParagraphNumberingContext, resolveNumberingForParagraphs } = require('./numbering-resolver');

/**
 * Parse a DOCX file to extract detailed style information with numbering context
 * Enhanced to include paragraph numbering context from DOCX introspection
 * 
 * @param {string} docxPath - Path to the DOCX file
 * @returns {Promise<Object>} - Detailed style information with numbering context
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
    
    // Parse XML content using xmldom's DOMParser
    const styleDoc = new DOMParser().parseFromString(styleXml);
    const documentDoc = new DOMParser().parseFromString(documentXml);
    const themeDoc = themeXml ? new DOMParser().parseFromString(themeXml) : null;
    const settingsDoc = settingsXml ? new DOMParser().parseFromString(settingsXml) : null;
    const numberingDoc = numberingXml ? new DOMParser().parseFromString(numberingXml) : null;
    
    // Import these modules here to avoid circular dependencies
    const { parseTheme } = require('./theme-parser');
    const { parseDocumentDefaults, parseSettingsWithMargins, analyzeDocumentStructure, getDefaultStyleInfo } = require('./document-parser');
    const { parseTocStyles } = require('./toc-parser');
    const { parseNumberingDefinitions } = require('./numbering-parser');
    
    // Parse styles from XML
    const styles = parseStyles(styleDoc);
    
    // Parse theme (colors, fonts)
    const theme = parseTheme(themeDoc);
    
    // Parse document defaults
    const documentDefaults = parseDocumentDefaults(styleDoc);
    
    // Parse document settings with page margins
    const settings = parseSettingsWithMargins(settingsDoc, documentDoc);
    
    // Parse TOC styles
    const tocStyles = parseTocStyles(documentDoc, styleDoc);
    
    // Parse numbering definitions with enhanced detail
    const numberingDefs = parseNumberingDefinitions(numberingDoc);
    
    // Extract document structure
    const documentStructure = analyzeDocumentStructure(documentDoc, numberingDoc);
    
    // NEW: Extract paragraph numbering context using DOCX introspection
    const paragraphNumberingContext = extractParagraphNumberingContext(
      documentDoc, 
      numberingDoc, 
      styleDoc
    );
    
    // NEW: Resolve numbering for all paragraphs
    const resolvedNumberingContext = resolveNumberingForParagraphs(
      paragraphNumberingContext, 
      numberingDefs
    );
    
    // Return combined style information with numbering context
    return {
      styles,
      theme,
      documentDefaults,
      settings,
      tocStyles,
      numberingDefs,
      documentStructure,
      // NEW: Add resolved numbering context
      numberingContext: resolvedNumberingContext
    };
  } catch (error) {
    console.error('Error parsing DOCX styles:', error);
    // Import this function to avoid circular dependencies
    const { getDefaultStyleInfo } = require('./document-parser');
    // Return default style info if parsing fails
    const defaultInfo = getDefaultStyleInfo();
    // Add empty numbering context for consistency
    defaultInfo.numberingContext = [];
    return defaultInfo;
  }
}

/**
 * Parse styles.xml to extract style definitions
 * Extracts paragraph, character, table, and numbering styles
 * 
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
    
    // Process paragraph styles with enhanced numbering detection
    paragraphStyleNodes.forEach(node => {
      const style = parseStyleNodeEnhanced(node);
      styles.paragraph[style.id] = style;
    });
    
    // Process character styles
    characterStyleNodes.forEach(node => {
      const style = parseStyleNodeEnhanced(node);
      styles.character[style.id] = style;
    });
    
    // Process table styles
    tableStyleNodes.forEach(node => {
      const style = parseStyleNodeEnhanced(node);
      styles.table[style.id] = style;
    });
    
    // Process numbering styles
    numberingStyleNodes.forEach(node => {
      const style = parseStyleNodeEnhanced(node);
      styles.numbering[style.id] = style;
    });
  } catch (error) {
    console.error('Error parsing styles:', error);
  }
  
  return styles;
}

/**
 * Enhanced style node parsing with numbering detection
 * Extracts style properties including embedded numbering information
 * 
 * @param {Node} node - XML node for a style
 * @returns {Object} - Enhanced style properties
 */
function parseStyleNodeEnhanced(node) {
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
    
    // Parse paragraph properties with enhanced numbering detection
    const pPrNode = selectSingleNode("w:pPr", node);
    const paragraphProps = pPrNode ? parseEnhancedParagraphProperties(pPrNode) : {};
    
    // Detect if this is a heading style
    const isHeadingStyle = detectHeadingStyle(name, styleId, type);
    
    // Combine all properties
    return {
      id: styleId,
      type,
      name,
      basedOn,
      isDefault,
      isHeading: isHeadingStyle,
      ...runningProps,
      ...paragraphProps
    };
  } catch (error) {
    console.error('Error parsing style node:', error);
    return { id: 'default', name: 'Default', type: 'paragraph' };
  }
}

/**
 * Enhanced paragraph properties parsing with comprehensive layout support
 * Extracts paragraph properties including spacing, indentation, alignment, and formatting
 * 
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Enhanced paragraph properties with full fidelity
 */
function parseEnhancedParagraphProperties(pPrNode) {
  const props = {};
  
  try {
    // Enhanced alignment extraction
    props.alignment = extractAlignment(pPrNode);
    
    // Enhanced indentation extraction with all properties
    props.indentation = extractEnhancedIndentation(pPrNode);
    
    // Enhanced spacing extraction with line spacing support
    props.spacing = extractEnhancedSpacing(pPrNode);
    
    // Tab stops extraction
    props.tabs = extractTabStops(pPrNode);
    
    // Enhanced borders extraction
    props.borders = extractBorders(pPrNode);
    
    // Enhanced shading/background extraction
    props.shading = extractEnhancedShading(pPrNode);
    
    // Enhanced numbering properties extraction
    props.numbering = extractEnhancedNumberingProperties(pPrNode);
    
    // Enhanced outline level detection
    props.outlineLevel = extractOutlineLevel(pPrNode);
    
    // Page break and keep properties for layout fidelity
    props.pageProperties = extractPageProperties(pPrNode);
    
    // Widow/orphan control
    props.widowControl = extractWidowControl(pPrNode);
    
    // Context and frame properties
    props.frameProperties = extractFrameProperties(pPrNode);
    
  } catch (error) {
    console.error('Error parsing enhanced paragraph properties:', error);
  }
  
  return props;
}

/**
 * Extract enhanced numbering properties from paragraph properties
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object|null} - Numbering properties or null
 */
function extractEnhancedNumberingProperties(pPrNode) {
  try {
    const numPrNode = selectSingleNode("w:numPr", pPrNode);
    if (numPrNode) {
      const numId = selectSingleNode("w:numId", numPrNode);
      const ilvl = selectSingleNode("w:ilvl", numPrNode);
      
      if (numId) {
        return {
          id: numId.getAttribute('w:val'),
          level: ilvl ? ilvl.getAttribute('w:val') : '0',
          hasNumbering: true
        };
      }
    }
  } catch (error) {
    console.error('Error extracting enhanced numbering properties:', error);
  }
  
  return null;
}

/**
 * Extract outline level from paragraph properties
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {number|null} - Outline level or null
 */
function extractOutlineLevel(pPrNode) {
  try {
    const outlineLvlNode = selectSingleNode("w:outlineLvl", pPrNode);
    if (outlineLvlNode) {
      return parseInt(outlineLvlNode.getAttribute('w:val'), 10);
    }
  } catch (error) {
    console.error('Error extracting outline level:', error);
  }
  
  return null;
}

/**
 * Detect if style is a heading style
 * @param {string} name - Style name
 * @param {string} styleId - Style ID
 * @param {string} type - Style type
 * @returns {boolean} - True if heading style
 */
function detectHeadingStyle(name, styleId, type) {
  if (type !== 'paragraph') return false;
  
  // Check common heading patterns
  const headingPatterns = [
    /^heading\s*\d*$/i,
    /^h\d+$/i,
    /^title$/i,
    /^subtitle$/i,
    /^caption$/i
  ];
  
  return headingPatterns.some(pattern => 
    pattern.test(name) || pattern.test(styleId)
  );
}

/**
 * Extract alignment from paragraph properties
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {string|null} - Alignment value
 */
function extractAlignment(pPrNode) {
  try {
    const jcNode = selectSingleNode("w:jc", pPrNode);
    return jcNode ? jcNode.getAttribute('w:val') : null;
  } catch (error) {
    console.error('Error extracting alignment:', error);
    return null;
  }
}

/**
 * Extract indentation from paragraph properties
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Indentation properties
 */
function extractIndentation(pPrNode) {
  const indentation = {};
  
  try {
    const indNode = selectSingleNode("w:ind", pPrNode);
    if (indNode) {
      const left = indNode.getAttribute('w:left');
      const right = indNode.getAttribute('w:right');
      const firstLine = indNode.getAttribute('w:firstLine');
      const hanging = indNode.getAttribute('w:hanging');
      const start = indNode.getAttribute('w:start');
      const end = indNode.getAttribute('w:end');
      
      if (left) indentation.left = left;
      if (right) indentation.right = right;
      if (firstLine) indentation.firstLine = firstLine;
      if (hanging) indentation.hanging = hanging;
      if (start) indentation.start = start;
      if (end) indentation.end = end;
    }
  } catch (error) {
    console.error('Error extracting indentation:', error);
  }
  
  return indentation;
}

/**
 * Extract spacing from paragraph properties
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Spacing properties
 */
function extractSpacing(pPrNode) {
  const spacing = {};
  
  try {
    const spacingNode = selectSingleNode("w:spacing", pPrNode);
    if (spacingNode) {
      const before = spacingNode.getAttribute('w:before');
      const after = spacingNode.getAttribute('w:after');
      const line = spacingNode.getAttribute('w:line');
      const lineRule = spacingNode.getAttribute('w:lineRule');
      
      if (before) spacing.before = before;
      if (after) spacing.after = after;
      if (line) spacing.line = line;
      if (lineRule) spacing.lineRule = lineRule;
    }
  } catch (error) {
    console.error('Error extracting spacing:', error);
  }
  
  return spacing;
}

/**
 * Extract tab stops from paragraph properties
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Array} - Array of tab stop definitions
 */
function extractTabStops(pPrNode) {
  const tabs = [];
  
  try {
    const tabsNode = selectSingleNode("w:tabs", pPrNode);
    if (tabsNode) {
      const tabNodes = selectNodes("w:tab", tabsNode);
      
      tabNodes.forEach(tabNode => {
        const pos = tabNode.getAttribute('w:pos');
        const val = tabNode.getAttribute('w:val');
        const leader = tabNode.getAttribute('w:leader');
        
        if (pos && val) {
          tabs.push({
            position: pos,
            positionPt: convertTwipToPt(pos),
            type: val,
            leader: leader || 'none',
            leaderChar: getLeaderChar(leader)
          });
        }
      });
    }
  } catch (error) {
    console.error('Error extracting tab stops:', error);
  }
  
  return tabs;
}

/**
 * Extract borders from paragraph properties
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Border properties
 */
function extractBorders(pPrNode) {
  const borders = {};
  
  try {
    const pBdrNode = selectSingleNode("w:pBdr", pPrNode);
    if (pBdrNode) {
      borders.top = extractBorderSide(pBdrNode, 'top');
      borders.bottom = extractBorderSide(pBdrNode, 'bottom');
      borders.left = extractBorderSide(pBdrNode, 'left');
      borders.right = extractBorderSide(pBdrNode, 'right');
    }
  } catch (error) {
    console.error('Error extracting borders:', error);
  }
  
  return borders;
}

/**
 * Extract border side properties
 * @param {Node} borderNode - Border container node
 * @param {string} side - Border side (top, bottom, left, right)
 * @returns {Object|null} - Border side properties
 */
function extractBorderSide(borderNode, side) {
  try {
    const sideNode = selectSingleNode(`w:${side}`, borderNode);
    if (sideNode) {
      return {
        value: sideNode.getAttribute('w:val'),
        size: sideNode.getAttribute('w:sz'),
        color: sideNode.getAttribute('w:color')
      };
    }
  } catch (error) {
    console.error(`Error extracting ${side} border:`, error);
  }
  
  return null;
}

/**
 * Extract shading from paragraph properties
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Shading properties
 */
function extractShading(pPrNode) {
  const shading = {};
  
  try {
    const shadingNode = selectSingleNode("w:shd", pPrNode);
    if (shadingNode) {
      const value = shadingNode.getAttribute('w:val');
      const color = shadingNode.getAttribute('w:color');
      const fill = shadingNode.getAttribute('w:fill');
      
      if (value) shading.value = value;
      if (color) shading.color = color;
      if (fill) shading.fill = fill;
    }
  } catch (error) {
    console.error('Error extracting shading:', error);
  }
  
  return shading;
}

/**
 * Parse running properties (text formatting) - Enhanced for comprehensive fidelity
 * Extracts font, size, color, and other text formatting properties with full DOCX support
 * 
 * @param {Node} rPrNode - Running properties node
 * @returns {Object} - Enhanced running properties
 */
function parseRunningProperties(rPrNode) {
  const props = {};
  
  try {
    // Enhanced font extraction with theme support
    const fontNode = selectSingleNode("w:rFonts", rPrNode);
    if (fontNode) {
      props.font = {
        ascii: fontNode.getAttribute('w:ascii'),
        hAnsi: fontNode.getAttribute('w:hAnsi'),
        eastAsia: fontNode.getAttribute('w:eastAsia'),
        cs: fontNode.getAttribute('w:cs'),
        // Add font theme references for better fidelity
        asciiTheme: fontNode.getAttribute('w:asciiTheme'),
        hAnsiTheme: fontNode.getAttribute('w:hAnsiTheme'),
        eastAsiaTheme: fontNode.getAttribute('w:eastAsiaTheme'),
        csTheme: fontNode.getAttribute('w:csTheme')
      };
    }
    
    // Enhanced size extraction with complex script support
    const szNode = selectSingleNode("w:sz", rPrNode);
    const szCsNode = selectSingleNode("w:szCs", rPrNode);
    if (szNode || szCsNode) {
      props.fontSize = {};
      if (szNode) {
        const sizeHalfPoints = parseInt(szNode.getAttribute('w:val'), 10) || 22;
        props.fontSize.ascii = (sizeHalfPoints / 2) + 'pt';
      }
      if (szCsNode) {
        const sizeCsHalfPoints = parseInt(szCsNode.getAttribute('w:val'), 10) || 22;
        props.fontSize.cs = (sizeCsHalfPoints / 2) + 'pt';
      }
      // For backward compatibility, set single fontSize if only ascii
      if (props.fontSize.ascii && !props.fontSize.cs) {
        props.fontSize = props.fontSize.ascii;
      }
    }
    
    // Bold
    const bNode = selectSingleNode("w:b", rPrNode);
    props.bold = bNode !== null;
    
    // Italic
    const iNode = selectSingleNode("w:i", rPrNode);
    props.italic = iNode !== null;
    
    // Enhanced underline with style support
    const uNode = selectSingleNode("w:u", rPrNode);
    if (uNode) {
      props.underline = {
        type: uNode.getAttribute('w:val') || 'single',
        color: uNode.getAttribute('w:color'),
        themeColor: uNode.getAttribute('w:themeColor')
      };
    }
    
    // Enhanced color extraction with theme support
    const colorNode = selectSingleNode("w:color", rPrNode);
    if (colorNode) {
      props.color = {
        val: colorNode.getAttribute('w:val'),
        theme: colorNode.getAttribute('w:themeColor'),
        tint: colorNode.getAttribute('w:themeTint'),
        shade: colorNode.getAttribute('w:themeShade')
      };
      
      // For backward compatibility, set simple color if available
      if (props.color.val && props.color.val !== 'auto') {
        props.colorValue = '#' + props.color.val;
      }
    }
    
    // Enhanced background/highlight extraction
    const highlightNode = selectSingleNode("w:highlight", rPrNode);
    const shdNode = selectSingleNode("w:shd", rPrNode);
    if (highlightNode || shdNode) {
      props.background = {};
      if (highlightNode) {
        props.background.highlight = highlightNode.getAttribute('w:val');
      }
      if (shdNode) {
        props.background.fill = shdNode.getAttribute('w:fill');
        props.background.color = shdNode.getAttribute('w:color');
        props.background.val = shdNode.getAttribute('w:val');
      }
      
      // For backward compatibility
      if (props.background.highlight) {
        props.highlight = props.background.highlight;
      }
    }
    
    // Text effects and formatting
    const strikeNode = selectSingleNode("w:strike", rPrNode);
    const dstrikeNode = selectSingleNode("w:dstrike", rPrNode);
    if (strikeNode || dstrikeNode) {
      props.strikethrough = {
        single: strikeNode !== null,
        double: dstrikeNode !== null
      };
    }
    
    // Superscript and subscript
    const vertAlignNode = selectSingleNode("w:vertAlign", rPrNode);
    if (vertAlignNode) {
      const val = vertAlignNode.getAttribute('w:val');
      if (val === 'superscript') {
        props.superscript = true;
      } else if (val === 'subscript') {
        props.subscript = true;
      }
    }
    
    // Small caps and all caps
    const smallCapsNode = selectSingleNode("w:smallCaps", rPrNode);
    const capsNode = selectSingleNode("w:caps", rPrNode);
    if (smallCapsNode) {
      props.smallCaps = true;
    }
    if (capsNode) {
      props.allCaps = true;
    }
    
    // Text outline and shadow effects
    const outlineNode = selectSingleNode("w:outline", rPrNode);
    const shadowNode = selectSingleNode("w:shadow", rPrNode);
    if (outlineNode || shadowNode) {
      props.textEffects = {};
      if (outlineNode) {
        props.textEffects.outline = true;
      }
      if (shadowNode) {
        props.textEffects.shadow = true;
      }
    }
    
    // Character spacing (tracking)
    const spacingNode = selectSingleNode("w:spacing", rPrNode);
    if (spacingNode) {
      const spacingVal = spacingNode.getAttribute('w:val');
      if (spacingVal) {
        // Convert from twentieths of a point to CSS units
        props.characterSpacing = (parseInt(spacingVal, 10) / 20) + 'pt';
      }
    }
    
    // Character scale (width)
    const wNode = selectSingleNode("w:w", rPrNode);
    if (wNode) {
      const scaleVal = wNode.getAttribute('w:val');
      if (scaleVal && scaleVal !== '100') {
        props.characterScale = scaleVal + '%';
      }
    }
    
  } catch (error) {
    console.error('Error parsing enhanced running properties:', error);
  }
  
  return props;
}

/**
 * Enhanced indentation extraction with all DOCX properties
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Complete indentation properties
 */
function extractEnhancedIndentation(pPrNode) {
  const indentation = {};
  
  try {
    const indNode = selectSingleNode("w:ind", pPrNode);
    if (indNode) {
      // Left indentation
      const left = indNode.getAttribute('w:left');
      const leftChars = indNode.getAttribute('w:leftChars');
      const start = indNode.getAttribute('w:start');
      
      if (left) indentation.left = left;
      if (leftChars) indentation.leftChars = leftChars;
      if (start) indentation.start = start;
      
      // Right indentation
      const right = indNode.getAttribute('w:right');
      const rightChars = indNode.getAttribute('w:rightChars');
      const end = indNode.getAttribute('w:end');
      
      if (right) indentation.right = right;
      if (rightChars) indentation.rightChars = rightChars;
      if (end) indentation.end = end;
      
      // First line and hanging indentation
      const firstLine = indNode.getAttribute('w:firstLine');
      const firstLineChars = indNode.getAttribute('w:firstLineChars');
      const hanging = indNode.getAttribute('w:hanging');
      const hangingChars = indNode.getAttribute('w:hangingChars');
      
      if (firstLine) indentation.firstLine = firstLine;
      if (firstLineChars) indentation.firstLineChars = firstLineChars;
      if (hanging) indentation.hanging = hanging;
      if (hangingChars) indentation.hangingChars = hangingChars;
    }
  } catch (error) {
    console.error('Error extracting enhanced indentation:', error);
  }
  
  return indentation;
}

/**
 * Enhanced spacing extraction with comprehensive line spacing support
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Complete spacing properties
 */
function extractEnhancedSpacing(pPrNode) {
  const spacing = {};
  
  try {
    const spacingNode = selectSingleNode("w:spacing", pPrNode);
    if (spacingNode) {
      // Before and after spacing
      const before = spacingNode.getAttribute('w:before');
      const beforeLines = spacingNode.getAttribute('w:beforeLines');
      const beforeAutospacing = spacingNode.getAttribute('w:beforeAutospacing');
      
      if (before) spacing.before = before;
      if (beforeLines) spacing.beforeLines = beforeLines;
      if (beforeAutospacing) spacing.beforeAutospacing = beforeAutospacing === '1';
      
      const after = spacingNode.getAttribute('w:after');
      const afterLines = spacingNode.getAttribute('w:afterLines');
      const afterAutospacing = spacingNode.getAttribute('w:afterAutospacing');
      
      if (after) spacing.after = after;
      if (afterLines) spacing.afterLines = afterLines;
      if (afterAutospacing) spacing.afterAutospacing = afterAutospacing === '1';
      
      // Line spacing
      const line = spacingNode.getAttribute('w:line');
      const lineRule = spacingNode.getAttribute('w:lineRule');
      
      if (line) spacing.line = line;
      if (lineRule) spacing.lineRule = lineRule;
    }
  } catch (error) {
    console.error('Error extracting enhanced spacing:', error);
  }
  
  return spacing;
}

/**
 * Enhanced shading extraction with theme support
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Complete shading properties
 */
function extractEnhancedShading(pPrNode) {
  const shading = {};
  
  try {
    const shadingNode = selectSingleNode("w:shd", pPrNode);
    if (shadingNode) {
      const value = shadingNode.getAttribute('w:val');
      const color = shadingNode.getAttribute('w:color');
      const fill = shadingNode.getAttribute('w:fill');
      const themeFill = shadingNode.getAttribute('w:themeFill');
      const themeFillTint = shadingNode.getAttribute('w:themeFillTint');
      const themeFillShade = shadingNode.getAttribute('w:themeFillShade');
      
      if (value) shading.value = value;
      if (color) shading.color = color;
      if (fill) shading.fill = fill;
      if (themeFill) shading.themeFill = themeFill;
      if (themeFillTint) shading.themeFillTint = themeFillTint;
      if (themeFillShade) shading.themeFillShade = themeFillShade;
    }
  } catch (error) {
    console.error('Error extracting enhanced shading:', error);
  }
  
  return shading;
}

/**
 * Extract page properties for layout fidelity
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Page properties
 */
function extractPageProperties(pPrNode) {
  const pageProps = {};
  
  try {
    // Page break before
    const pageBreakNode = selectSingleNode("w:pageBreakBefore", pPrNode);
    if (pageBreakNode) {
      pageProps.pageBreakBefore = pageBreakNode.getAttribute('w:val') !== '0';
    }
    
    // Keep with next
    const keepNextNode = selectSingleNode("w:keepNext", pPrNode);
    if (keepNextNode) {
      pageProps.keepNext = keepNextNode.getAttribute('w:val') !== '0';
    }
    
    // Keep lines together
    const keepLinesNode = selectSingleNode("w:keepLines", pPrNode);
    if (keepLinesNode) {
      pageProps.keepLines = keepLinesNode.getAttribute('w:val') !== '0';
    }
  } catch (error) {
    console.error('Error extracting page properties:', error);
  }
  
  return pageProps;
}

/**
 * Extract widow/orphan control
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Widow control properties
 */
function extractWidowControl(pPrNode) {
  const widowControl = {};
  
  try {
    const widowControlNode = selectSingleNode("w:widowControl", pPrNode);
    if (widowControlNode) {
      widowControl.enabled = widowControlNode.getAttribute('w:val') !== '0';
    }
  } catch (error) {
    console.error('Error extracting widow control:', error);
  }
  
  return widowControl;
}

/**
 * Extract frame properties for positioning
 * @param {Node} pPrNode - Paragraph properties node
 * @returns {Object} - Frame properties
 */
function extractFrameProperties(pPrNode) {
  const frameProps = {};
  
  try {
    const framePrNode = selectSingleNode("w:framePr", pPrNode);
    if (framePrNode) {
      const dropCap = framePrNode.getAttribute('w:dropCap');
      const lines = framePrNode.getAttribute('w:lines');
      const w = framePrNode.getAttribute('w:w');
      const h = framePrNode.getAttribute('w:h');
      const vSpace = framePrNode.getAttribute('w:vSpace');
      const hSpace = framePrNode.getAttribute('w:hSpace');
      const wrap = framePrNode.getAttribute('w:wrap');
      const hAnchor = framePrNode.getAttribute('w:hAnchor');
      const vAnchor = framePrNode.getAttribute('w:vAnchor');
      const x = framePrNode.getAttribute('w:x');
      const y = framePrNode.getAttribute('w:y');
      
      if (dropCap) frameProps.dropCap = dropCap;
      if (lines) frameProps.lines = lines;
      if (w) frameProps.width = w;
      if (h) frameProps.height = h;
      if (vSpace) frameProps.vSpace = vSpace;
      if (hSpace) frameProps.hSpace = hSpace;
      if (wrap) frameProps.wrap = wrap;
      if (hAnchor) frameProps.hAnchor = hAnchor;
      if (vAnchor) frameProps.vAnchor = vAnchor;
      if (x) frameProps.x = x;
      if (y) frameProps.y = y;
    }
  } catch (error) {
    console.error('Error extracting frame properties:', error);
  }
  
  return frameProps;
}

module.exports = {
  parseDocxStyles,
  parseStyles,
  parseStyleNodeEnhanced,
  parseEnhancedParagraphProperties,
  parseRunningProperties,
  extractEnhancedNumberingProperties,
  detectHeadingStyle,
  extractEnhancedIndentation,
  extractEnhancedSpacing,
  extractEnhancedShading,
  extractPageProperties,
  extractWidowControl,
  extractFrameProperties
};
