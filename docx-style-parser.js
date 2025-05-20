// docx-style-parser.js - Extract styles directly from DOCX documents
const JSZip = require('jszip');
const fs = require('fs');
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
    
    // Extract all relevant files
    const styleXml = await zip.file('word/styles.xml')?.async('string');
    const documentXml = await zip.file('word/document.xml')?.async('string');
    const themeXml = await zip.file('word/theme/theme1.xml')?.async('string');
    const settingsXml = await zip.file('word/settings.xml')?.async('string');
    const numberingXml = await zip.file('word/numbering.xml')?.async('string');
    const fontTableXml = await zip.file('word/fontTable.xml')?.async('string');
    
    if (!documentXml) {
      throw new Error('Invalid DOCX file: missing document.xml');
    }
    
    // Parse XML content
    const documentDoc = new dom().parseFromString(documentXml);
    const styleDoc = styleXml ? new dom().parseFromString(styleXml) : null;
    const themeDoc = themeXml ? new dom().parseFromString(themeXml) : null;
    const settingsDoc = settingsXml ? new dom().parseFromString(settingsXml) : null;
    const numberingDoc = numberingXml ? new dom().parseFromString(numberingXml) : null;
    const fontTableDoc = fontTableXml ? new dom().parseFromString(fontTableXml) : null;
    
    // Extract document information
    const documentInfo = {
      styles: styleDoc ? parseStyles(styleDoc) : {},
      theme: themeDoc ? parseTheme(themeDoc) : { colors: {}, fonts: {} },
      settings: settingsDoc ? parseSettings(settingsDoc) : {},
      numbering: numberingDoc ? parseNumbering(numberingDoc) : {},
      fonts: fontTableDoc ? parseFontTable(fontTableDoc) : {},
      paragraphs: parseParagraphs(documentDoc),
      sections: parseSections(documentDoc),
      tables: parseTables(documentDoc),
      lists: parseListsFromDocument(documentDoc, numberingDoc),
      hasRTL: checkForRTL(documentDoc)
    };
    
    return documentInfo;
  } catch (error) {
    console.error('Error parsing DOCX styles:', error);
    return { 
      styles: {}, 
      theme: { colors: {}, fonts: {} }, 
      settings: {},
      numbering: {},
      fonts: {},
      paragraphs: [],
      sections: [],
      tables: [],
      lists: [],
      hasRTL: false
    };
  }
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
    numbering: {},
    defaults: {
      paragraph: {},
      character: {}
    }
  };
  
  try {
    // Get document defaults
    const docDefaultsNode = selectSingleNode("//w:docDefaults", styleDoc);
    if (docDefaultsNode) {
      // Get paragraph defaults
      const pPrDefaultNode = selectSingleNode(".//w:pPrDefault/w:pPr", docDefaultsNode);
      if (pPrDefaultNode) {
        styles.defaults.paragraph = parseParagraphProperties(pPrDefaultNode);
      }
      
      // Get character defaults
      const rPrDefaultNode = selectSingleNode(".//w:rPrDefault/w:rPr", docDefaultsNode);
      if (rPrDefaultNode) {
        styles.defaults.character = parseRunProperties(rPrDefaultNode);
      }
    }
    
    // Parse all style definitions
    const styleNodes = selectNodes("//w:style", styleDoc);
    
    styleNodes.forEach(node => {
      const styleType = node.getAttribute('w:type');
      const styleId = node.getAttribute('w:styleId');
      
      if (!styleType || !styleId) return;
      
      const style = {
        id: styleId,
        type: styleType,
        name: '',
        basedOn: null,
        next: null,
        isDefault: false,
        properties: {}
      };
      
      // Get style name
      const nameNode = selectSingleNode("w:name", node);
      if (nameNode) {
        style.name = nameNode.getAttribute('w:val') || styleId;
      }
      
      // Get based on style
      const basedOnNode = selectSingleNode("w:basedOn", node);
      if (basedOnNode) {
        style.basedOn = basedOnNode.getAttribute('w:val');
      }
      
      // Get next style
      const nextNode = selectSingleNode("w:next", node);
      if (nextNode) {
        style.next = nextNode.getAttribute('w:val');
      }
      
      // Check if default style
      const defaultNode = node.getAttribute('w:default');
      style.isDefault = defaultNode === '1';
      
      // Get style properties based on type
      if (styleType === 'paragraph') {
        // Get paragraph properties
        const pPrNode = selectSingleNode("w:pPr", node);
        if (pPrNode) {
          style.properties.paragraph = parseParagraphProperties(pPrNode);
        }
        
        // Get run properties within paragraph style
        const rPrNode = selectSingleNode("w:rPr", node);
        if (rPrNode) {
          style.properties.character = parseRunProperties(rPrNode);
        }
        
        styles.paragraph[styleId] = style;
      } 
      else if (styleType === 'character') {
        // Get run properties
        const rPrNode = selectSingleNode("w:rPr", node);
        if (rPrNode) {
          style.properties = parseRunProperties(rPrNode);
        }
        
        styles.character[styleId] = style;
      } 
      else if (styleType === 'table') {
        // Get table properties
        const tblPrNode = selectSingleNode("w:tblPr", node);
        if (tblPrNode) {
          style.properties.table = parseTableProperties(tblPrNode);
        }
        
        // Get conditional formatting for table parts
        const tblStylePrNodes = selectNodes("w:tblStylePr", node);
        if (tblStylePrNodes.length > 0) {
          style.properties.conditionalFormatting = {};
          
          tblStylePrNodes.forEach(tblStylePrNode => {
            const type = tblStylePrNode.getAttribute('w:type');
            if (type) {
              style.properties.conditionalFormatting[type] = {
                paragraph: null,
                character: null,
                table: null
              };
              
              // Get paragraph properties
              const pPrNode = selectSingleNode("w:pPr", tblStylePrNode);
              if (pPrNode) {
                style.properties.conditionalFormatting[type].paragraph = parseParagraphProperties(pPrNode);
              }
              
              // Get run properties
              const rPrNode = selectSingleNode("w:rPr", tblStylePrNode);
              if (rPrNode) {
                style.properties.conditionalFormatting[type].character = parseRunProperties(rPrNode);
              }
              
              // Get table properties
              const tcPrNode = selectSingleNode("w:tcPr", tblStylePrNode);
              if (tcPrNode) {
                style.properties.conditionalFormatting[type].table = parseTableCellProperties(tcPrNode);
              }
            }
          });
        }
        
        styles.table[styleId] = style;
      }
      else if (styleType === 'numbering') {
        // Get numbering properties
        const numPrNode = selectSingleNode("w:pPr/w:numPr", node);
        if (numPrNode) {
          const numId = selectSingleNode("w:numId", numPrNode);
          const ilvl = selectSingleNode("w:ilvl", numPrNode);
          
          if (numId && ilvl) {
            style.properties.numbering = {
              id: numId.getAttribute('w:val'),
              level: ilvl.getAttribute('w:val')
            };
          }
        }
        
        styles.numbering[styleId] = style;
      }
    });
  } catch (error) {
    console.error('Error parsing styles:', error);
  }
  
  return styles;
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
      props.indentation = {};
      
      const left = indNode.getAttribute('w:left');
      const right = indNode.getAttribute('w:right');
      const firstLine = indNode.getAttribute('w:firstLine');
      const hanging = indNode.getAttribute('w:hanging');
      
      if (left) props.indentation.left = parseInt(left, 10);
      if (right) props.indentation.right = parseInt(right, 10);
      if (firstLine) props.indentation.firstLine = parseInt(firstLine, 10);
      if (hanging) props.indentation.hanging = parseInt(hanging, 10);
    }
    
    // Spacing
    const spacingNode = selectSingleNode("w:spacing", pPrNode);
    if (spacingNode) {
      props.spacing = {};
      
      const before = spacingNode.getAttribute('w:before');
      const after = spacingNode.getAttribute('w:after');
      const line = spacingNode.getAttribute('w:line');
      const lineRule = spacingNode.getAttribute('w:lineRule');
      const beforeAutospacing = spacingNode.getAttribute('w:beforeAutospacing');
      const afterAutospacing = spacingNode.getAttribute('w:afterAutospacing');
      
      if (before) props.spacing.before = parseInt(before, 10);
      if (after) props.spacing.after = parseInt(after, 10);
      if (line) props.spacing.line = parseInt(line, 10);
      if (lineRule) props.spacing.lineRule = lineRule;
      if (beforeAutospacing === '1') props.spacing.beforeAutospacing = true;
      if (afterAutospacing === '1') props.spacing.afterAutospacing = true;
    }
    
    // Bidirectional settings
    const bidiNode = selectSingleNode("w:bidi", pPrNode);
    if (bidiNode) {
      props.bidi = bidiNode.getAttribute('w:val') !== '0';
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
    
    // Numbering properties
    const numPrNode = selectSingleNode("w:numPr", pPrNode);
    if (numPrNode) {
      const numId = selectSingleNode("w:numId", numPrNode);
      const ilvl = selectSingleNode("w:ilvl", numPrNode);
      
      if (numId || ilvl) {
        props.numbering = {};
        if (numId) props.numbering.id = numId.getAttribute('w:val');
        if (ilvl) props.numbering.level = ilvl.getAttribute('w:val');
      }
    }
    
    // Tabs
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
            position: parseInt(pos, 10),
            type: val,
            leader: leader || 'none'
          });
        }
      });
    }
    
    // Outline level
    const outlineLvlNode = selectSingleNode("w:outlineLvl", pPrNode);
    if (outlineLvlNode) {
      props.outlineLevel = parseInt(outlineLvlNode.getAttribute('w:val'), 10);
    }
    
    // Style reference
    const pStyleNode = selectSingleNode("w:pStyle", pPrNode);
    if (pStyleNode) {
      props.style = pStyleNode.getAttribute('w:val');
    }
  } catch (error) {
    console.error('Error parsing paragraph properties:', error);
  }
  
  return props;
}

/**
 * Parse run properties (text formatting)
 * @param {Node} rPrNode - Run properties node
 * @returns {Object} - Run properties
 */
function parseRunProperties(rPrNode) {
  const props = {};
  
  try {
    // Font
    const fontNode = selectSingleNode("w:rFonts", rPrNode);
    if (fontNode) {
      props.fonts = {};
      
      const ascii = fontNode.getAttribute('w:ascii');
      const hAnsi = fontNode.getAttribute('w:hAnsi');
      const eastAsia = fontNode.getAttribute('w:eastAsia');
      const cs = fontNode.getAttribute('w:cs');
      
      if (ascii) props.fonts.ascii = ascii;
      if (hAnsi) props.fonts.hAnsi = hAnsi;
      if (eastAsia) props.fonts.eastAsia = eastAsia;
      if (cs) props.fonts.cs = cs;
    }
    
    // Size
    const szNode = selectSingleNode("w:sz", rPrNode);
    if (szNode) {
      const size = parseInt(szNode.getAttribute('w:val'), 10);
      if (!isNaN(size)) {
        props.size = size; // Size in half-points
      }
    }
    
    // Bold
    const bNode = selectSingleNode("w:b", rPrNode);
    if (bNode) {
      props.bold = bNode.getAttribute('w:val') !== '0';
    }
    
    // Italic
    const iNode = selectSingleNode("w:i", rPrNode);
    if (iNode) {
      props.italic = iNode.getAttribute('w:val') !== '0';
    }
    
    // Underline
    const uNode = selectSingleNode("w:u", rPrNode);
    if (uNode) {
      props.underline = {
        type: uNode.getAttribute('w:val') || 'single',
        color: uNode.getAttribute('w:color')
      };
    }
    
    // Strikethrough
    const strikeNode = selectSingleNode("w:strike", rPrNode);
    if (strikeNode) {
      props.strike = strikeNode.getAttribute('w:val') !== '0';
    }
    
    // Color
    const colorNode = selectSingleNode("w:color", rPrNode);
    if (colorNode) {
      props.color = colorNode.getAttribute('w:val');
    }
    
    // Highlight
    const highlightNode = selectSingleNode("w:highlight", rPrNode);
    if (highlightNode) {
      props.highlight = highlightNode.getAttribute('w:val');
    }
    
    // All caps
    const capsNode = selectSingleNode("w:caps", rPrNode);
    if (capsNode) {
      props.caps = capsNode.getAttribute('w:val') !== '0';
    }
    
    // Small caps
    const smallCapsNode = selectSingleNode("w:smallCaps", rPrNode);
    if (smallCapsNode) {
      props.smallCaps = smallCapsNode.getAttribute('w:val') !== '0';
    }
    
    // Vertical alignment (superscript/subscript)
    const vertAlignNode = selectSingleNode("w:vertAlign", rPrNode);
    if (vertAlignNode) {
      props.verticalAlignment = vertAlignNode.getAttribute('w:val');
    }
    
    // Style reference
    const rStyleNode = selectSingleNode("w:rStyle", rPrNode);
    if (rStyleNode) {
      props.style = rStyleNode.getAttribute('w:val');
    }
  } catch (error) {
    console.error('Error parsing run properties:', error);
  }
  
  return props;
}

/**
 * Parse borders
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
        type: topNode.getAttribute('w:val'),
        size: parseInt(topNode.getAttribute('w:sz') || '0', 10),
        color: topNode.getAttribute('w:color')
      };
    }
    
    // Bottom border
    const bottomNode = selectSingleNode("w:bottom", borderNode);
    if (bottomNode) {
      borders.bottom = {
        type: bottomNode.getAttribute('w:val'),
        size: parseInt(bottomNode.getAttribute('w:sz') || '0', 10),
        color: bottomNode.getAttribute('w:color')
      };
    }
    
    // Left border
    const leftNode = selectSingleNode("w:left", borderNode);
    if (leftNode) {
      borders.left = {
        type: leftNode.getAttribute('w:val'),
        size: parseInt(leftNode.getAttribute('w:sz') || '0', 10),
        color: leftNode.getAttribute('w:color')
      };
    }
    
    // Right border
    const rightNode = selectSingleNode("w:right", borderNode);
    if (rightNode) {
      borders.right = {
        type: rightNode.getAttribute('w:val'),
        size: parseInt(rightNode.getAttribute('w:sz') || '0', 10),
        color: rightNode.getAttribute('w:color')
      };
    }
  } catch (error) {
    console.error('Error parsing borders:', error);
  }
  
  return borders;
}

/**
 * Parse table properties
 * @param {Node} tblPrNode - Table properties node
 * @returns {Object} - Table properties
 */
function parseTableProperties(tblPrNode) {
  const props = {};
  
  try {
    // Table style
    const tblStyleNode = selectSingleNode("w:tblStyle", tblPrNode);
    if (tblStyleNode) {
      props.style = tblStyleNode.getAttribute('w:val');
    }
    
    // Table width
    const tblWNode = selectSingleNode("w:tblW", tblPrNode);
    if (tblWNode) {
      props.width = {
        width: parseInt(tblWNode.getAttribute('w:w') || '0', 10),
        type: tblWNode.getAttribute('w:type')
      };
    }
    
    // Table cell margins
    const tblCellMarNode = selectSingleNode("w:tblCellMar", tblPrNode);
    if (tblCellMarNode) {
      props.cellMargins = {};
      
      const topNode = selectSingleNode("w:top", tblCellMarNode);
      const leftNode = selectSingleNode("w:left", tblCellMarNode);
      const bottomNode = selectSingleNode("w:bottom", tblCellMarNode);
      const rightNode = selectSingleNode("w:right", tblCellMarNode);
      
      if (topNode) {
        props.cellMargins.top = {
          width: parseInt(topNode.getAttribute('w:w') || '0', 10),
          type: topNode.getAttribute('w:type')
        };
      }
      
      if (leftNode) {
        props.cellMargins.left = {
          width: parseInt(leftNode.getAttribute('w:w') || '0', 10),
          type: leftNode.getAttribute('w:type')
        };
      }
      
      if (bottomNode) {
        props.cellMargins.bottom = {
          width: parseInt(bottomNode.getAttribute('w:w') || '0', 10),
          type: bottomNode.getAttribute('w:type')
        };
      }
      
      if (rightNode) {
        props.cellMargins.right = {
          width: parseInt(rightNode.getAttribute('w:w') || '0', 10),
          type: rightNode.getAttribute('w:type')
        };
      }
    }
    
    // Table borders
    const tblBordersNode = selectSingleNode("w:tblBorders", tblPrNode);
    if (tblBordersNode) {
      props.borders = parseBorders(tblBordersNode);
    }
    
    // Table layout
    const tblLayoutNode = selectSingleNode("w:tblLayout", tblPrNode);
    if (tblLayoutNode) {
      props.layout = tblLayoutNode.getAttribute('w:type');
    }
    
    // Table look
    const tblLookNode = selectSingleNode("w:tblLook", tblPrNode);
    if (tblLookNode) {
      props.look = {
        firstRow: (parseInt(tblLookNode.getAttribute('w:firstRow') || '0', 16) & 0x0020) !== 0,
        lastRow: (parseInt(tblLookNode.getAttribute('w:lastRow') || '0', 16) & 0x0040) !== 0,
        firstColumn: (parseInt(tblLookNode.getAttribute('w:firstColumn') || '0', 16) & 0x0080) !== 0,
        lastColumn: (parseInt(tblLookNode.getAttribute('w:lastColumn') || '0', 16) & 0x0100) !== 0,
        noHBand: (parseInt(tblLookNode.getAttribute('w:noHBand') || '0', 16) & 0x0200) !== 0,
        noVBand: (parseInt(tblLookNode.getAttribute('w:noVBand') || '0', 16) & 0x0400) !== 0
      };
    }
  } catch (error) {
    console.error('Error parsing table properties:', error);
  }
  
  return props;
}

/**
 * Parse table cell properties
 * @param {Node} tcPrNode - Table cell properties node
 * @returns {Object} - Table cell properties
 */
function parseTableCellProperties(tcPrNode) {
  const props = {};
  
  try {
    // Cell width
    const tcWNode = selectSingleNode("w:tcW", tcPrNode);
    if (tcWNode) {
      props.width = {
        width: parseInt(tcWNode.getAttribute('w:w') || '0', 10),
        type: tcWNode.getAttribute('w:type')
      };
    }
    
    // Cell borders
    const tcBordersNode = selectSingleNode("w:tcBorders", tcPrNode);
    if (tcBordersNode) {
      props.borders = parseBorders(tcBordersNode);
    }
    
    // Cell shading
    const shadingNode = selectSingleNode("w:shd", tcPrNode);
    if (shadingNode) {
      props.shading = {
        value: shadingNode.getAttribute('w:val'),
        color: shadingNode.getAttribute('w:color'),
        fill: shadingNode.getAttribute('w:fill')
      };
    }
    
    // Vertical alignment
    const vAlignNode = selectSingleNode("w:vAlign", tcPrNode);
    if (vAlignNode) {
      props.verticalAlignment = vAlignNode.getAttribute('w:val');
    }
  } catch (error) {
    console.error('Error parsing table cell properties:', error);
  }
  
  return props;
}

/**
 * Parse theme information
 * @param {Document} themeDoc - Theme XML document
 * @returns {Object} - Theme properties
 */
function parseTheme(themeDoc) {
  const theme = {
    colors: {},
    fonts: {
      major: null,
      minor: null
    }
  };
  
  try {
    // Parse font scheme
    const fontSchemeNode = selectSingleNode("//a:fontScheme", themeDoc);
    if (fontSchemeNode) {
      // Major fonts
      const majorFontNode = selectSingleNode(".//a:majorFont/a:latin", fontSchemeNode);
      if (majorFontNode) {
        theme.fonts.major = majorFontNode.getAttribute('typeface');
      }
      
      // Minor fonts
      const minorFontNode = selectSingleNode(".//a:minorFont/a:latin", fontSchemeNode);
      if (minorFontNode) {
        theme.fonts.minor = minorFontNode.getAttribute('typeface');
      }
    }
    
    // Parse color scheme
    const clrSchemeNode = selectSingleNode("//a:clrScheme", themeDoc);
    if (clrSchemeNode) {
      const colorNodes = selectNodes("./a:*", clrSchemeNode);
      
      colorNodes.forEach(node => {
        const name = node.nodeName.split(':')[1];
        
        // Get color value (RGB or system color)
        const srgbClrNode = selectSingleNode("./a:srgbClr", node);
        const sysClrNode = selectSingleNode("./a:sysClr", node);
        
        if (srgbClrNode) {
          theme.colors[name] = srgbClrNode.getAttribute('val');
        } else if (sysClrNode) {
          theme.colors[name] = sysClrNode.getAttribute('lastClr') || sysClrNode.getAttribute('val');
        }
      });
    }
  } catch (error) {
    console.error('Error parsing theme:', error);
  }
  
  return theme;
}

/**
 * Parse document settings
 * @param {Document} settingsDoc - Settings XML document
 * @returns {Object} - Document settings
 */
function parseSettings(settingsDoc) {
  const settings = {
    defaultTabStop: 720, // Default value (0.5 inch in twentieths of a point)
    characterSpacing: 'normal',
    doNotHyphenateCaps: false,
    rtlGutter: false,
    mirrorMargins: false,
    displayBackgroundShape: false,
    zoom: 100,
    evenAndOddHeaders: false
  };
  
  try {
    // Default tab stop
    const defaultTabStopNode = selectSingleNode("//w:defaultTabStop", settingsDoc);
    if (defaultTabStopNode) {
      const val = defaultTabStopNode.getAttribute('w:val');
      if (val) {
        settings.defaultTabStop = parseInt(val, 10);
      }
    }
    
    // Character spacing control
    const characterSpacingControlNode = selectSingleNode("//w:characterSpacingControl", settingsDoc);
    if (characterSpacingControlNode) {
      settings.characterSpacing = characterSpacingControlNode.getAttribute('w:val');
    }
    
    // RTL gutter
    const rtlGutterNode = selectSingleNode("//w:rtlGutter", settingsDoc);
    if (rtlGutterNode) {
      settings.rtlGutter = rtlGutterNode.getAttribute('w:val') === '1';
    }
    
    // Mirror margins
    const mirrorMarginsNode = selectSingleNode("//w:mirrorMargins", settingsDoc);
    if (mirrorMarginsNode) {
      settings.mirrorMargins = mirrorMarginsNode.getAttribute('w:val') === '1';
    }
    
    // Display background shape
    const displayBackgroundShapeNode = selectSingleNode("//w:displayBackgroundShape", settingsDoc);
    if (displayBackgroundShapeNode) {
      settings.displayBackgroundShape = displayBackgroundShapeNode.getAttribute('w:val') === '1';
    }
    
    // Zoom
    const zoomNode = selectSingleNode("//w:zoom", settingsDoc);
    if (zoomNode) {
      const percent = zoomNode.getAttribute('w:percent');
      if (percent) {
        settings.zoom = parseInt(percent, 10);
      }
    }
    
    // Even and odd headers
    const evenAndOddHeadersNode = selectSingleNode("//w:evenAndOddHeaders", settingsDoc);
    if (evenAndOddHeadersNode) {
      settings.evenAndOddHeaders = evenAndOddHeadersNode.getAttribute('w:val') === '1';
    }
  } catch (error) {
    console.error('Error parsing settings:', error);
  }
  
  return settings;
}

/**
 * Parse numbering information
 * @param {Document} numberingDoc - Numbering XML document
 * @returns {Object} - Numbering information
 */
function parseNumbering(numberingDoc) {
  const numbering = {
    abstractNums: {},
    nums: {}
  };
  
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
      
      // Parse level definitions
      const levelNodes = selectNodes("w:lvl", node);
      levelNodes.forEach(lvlNode => {
        const ilvl = lvlNode.getAttribute('w:ilvl');
        if (ilvl === null) return;
        
        const level = {
          level: parseInt(ilvl, 10),
          start: 1,
          format: 'decimal',
          text: '%1.',
          alignment: 'left',
          properties: {
            paragraph: {},
            character: {}
          }
        };
        
        // Get start value
        const startNode = selectSingleNode("w:start", lvlNode);
        if (startNode) {
          level.start = parseInt(startNode.getAttribute('w:val'), 10);
        }
        
        // Get numbering format
        const numFmtNode = selectSingleNode("w:numFmt", lvlNode);
        if (numFmtNode) {
          level.format = numFmtNode.getAttribute('w:val') || 'decimal';
        }
        
        // Get level text
        const lvlTextNode = selectSingleNode("w:lvlText", lvlNode);
        if (lvlTextNode) {
          level.text = lvlTextNode.getAttribute('w:val') || '%1.';
        }
        
        // Get alignment
        const alignmentNode = selectSingleNode("w:lvlJc", lvlNode);
        if (alignmentNode) {
          level.alignment = alignmentNode.getAttribute('w:val') || 'left';
        }
        
        // Get paragraph properties
        const pPrNode = selectSingleNode("w:pPr", lvlNode);
        if (pPrNode) {
          level.properties.paragraph = parseParagraphProperties(pPrNode);
        }
        
        // Get run properties
        const rPrNode = selectSingleNode("w:rPr", lvlNode);
        if (rPrNode) {
          level.properties.character = parseRunProperties(rPrNode);
        }
        
        abstractNum.levels[ilvl] = level;
      });
      
      numbering.abstractNums[id] = abstractNum;
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
      
      // Check for level overrides
      const lvlOverrideNodes = selectNodes("w:lvlOverride", node);
      let levelOverrides = null;
      
      if (lvlOverrideNodes.length > 0) {
        levelOverrides = {};
        
        lvlOverrideNodes.forEach(lvlOverrideNode => {
          const ilvl = lvlOverrideNode.getAttribute('w:ilvl');
          if (ilvl === null) return;
          
          // Check if this level has a complete override or just property overrides
          const lvlNode = selectSingleNode("w:lvl", lvlOverrideNode);
          if (lvlNode) {
            // Complete level override - parse as a regular level
            const level = {
              level: parseInt(ilvl, 10),
              start: 1,
              format: 'decimal',
              text: '%1.',
              alignment: 'left',
              properties: {
                paragraph: {},
                character: {}
              }
            };
            
            // Get start value
            const startNode = selectSingleNode("w:start", lvlNode);
            if (startNode) {
              level.start = parseInt(startNode.getAttribute('w:val'), 10);
            }
            
            // Get numbering format
            const numFmtNode = selectSingleNode("w:numFmt", lvlNode);
            if (numFmtNode) {
              level.format = numFmtNode.getAttribute('w:val') || 'decimal';
            }
            
            // Get level text
            const lvlTextNode = selectSingleNode("w:lvlText", lvlNode);
            if (lvlTextNode) {
              level.text = lvlTextNode.getAttribute('w:val') || '%1.';
            }
            
            // Get alignment
            const alignmentNode = selectSingleNode("w:lvlJc", lvlNode);
            if (alignmentNode) {
              level.alignment = alignmentNode.getAttribute('w:val') || 'left';
            }
            
            // Get paragraph properties
            const pPrNode = selectSingleNode("w:pPr", lvlNode);
            if (pPrNode) {
              level.properties.paragraph = parseParagraphProperties(pPrNode);
            }
            
            // Get run properties
            const rPrNode = selectSingleNode("w:rPr", lvlNode);
            if (rPrNode) {
              level.properties.character = parseRunProperties(rPrNode);
            }
            
            levelOverrides[ilvl] = { complete: true, level };
          } else {
            // Just a start override
            const startOverrideNode = selectSingleNode("w:startOverride", lvlOverrideNode);
            if (startOverrideNode) {
              const start = parseInt(startOverrideNode.getAttribute('w:val'), 10);
              levelOverrides[ilvl] = { complete: false, startOverride: start };
            }
          }
        });
      }
      
      numbering.nums[id] = {
        id,
        abstractNumId,
        levelOverrides
      };
    });
  } catch (error) {
    console.error('Error parsing numbering:', error);
  }
  
  return numbering;
}

/**
 * Parse font table
 * @param {Document} fontTableDoc - Font table XML document
 * @returns {Object} - Font information
 */
function parseFontTable(fontTableDoc) {
  const fonts = {};
  
  try {
    const fontNodes = selectNodes("//w:font", fontTableDoc);
    
    fontNodes.forEach(node => {
      const name = node.getAttribute('w:name');
      if (!name) return;
      
      const font = {
        name,
        altName: null,
        family: null,
        pitch: null,
      };
      
      // Alternative name
      const altNameNode = selectSingleNode("w:altName", node);
      if (altNameNode) {
        font.altName = altNameNode.getAttribute('w:val');
      }
      
      // Family
      const familyNode = selectSingleNode("w:family", node);
      if (familyNode) {
        font.family = familyNode.getAttribute('w:val');
      }
      
      // Pitch
      const pitchNode = selectSingleNode("w:pitch", node);
      if (pitchNode) {
        font.pitch = pitchNode.getAttribute('w:val');
      }
      
      fonts[name] = font;
    });
  } catch (error) {
    console.error('Error parsing font table:', error);
  }
  
  return fonts;
}

/**
 * Parse paragraphs from document
 * @param {Document} documentDoc - Document XML
 * @returns {Array} - Paragraph information
 */
function parseParagraphs(documentDoc) {
  const paragraphs = [];
  
  try {
    const paragraphNodes = selectNodes("//w:p", documentDoc);
    
    paragraphNodes.forEach((pNode, index) => {
      const paragraph = {
        id: `p${index}`,
        styleId: null,
        properties: {},
        text: '',
        runs: []
      };
      
      // Get paragraph properties
      const pPrNode = selectSingleNode("w:pPr", pNode);
      if (pPrNode) {
        paragraph.properties = parseParagraphProperties(pPrNode);
        
        // Get style ID
        const pStyleNode = selectSingleNode("w:pStyle", pPrNode);
        if (pStyleNode) {
          paragraph.styleId = pStyleNode.getAttribute('w:val');
        }
      }
      
      // Get runs
      const runNodes = selectNodes("w:r", pNode);
      runNodes.forEach((rNode, runIndex) => {
        const run = {
          id: `${paragraph.id}_r${runIndex}`,
          properties: {},
          text: ''
        };
        
        // Get run properties
        const rPrNode = selectSingleNode("w:rPr", rNode);
        if (rPrNode) {
          run.properties = parseRunProperties(rPrNode);
        }
        
        // Get text
        const textNodes = selectNodes("w:t", rNode);
        textNodes.forEach(tNode => {
          run.text += tNode.textContent;
        });
        
        // Get tabs
        const tabNodes = selectNodes("w:tab", rNode);
        if (tabNodes.length > 0) {
          run.hasTabs = true;
          run.tabCount = tabNodes.length;
        }
        
        // Get breaks
        const breakNodes = selectNodes("w:br", rNode);
        if (breakNodes.length > 0) {
          run.hasBreaks = true;
          run.breakCount = breakNodes.length;
          
          // Check break types
          run.breaks = breakNodes.map(brNode => {
            return {
              type: brNode.getAttribute('w:type') || 'textWrapping',
              clear: brNode.getAttribute('w:clear')
            };
          });
        }
        
        paragraph.runs.push(run);
        paragraph.text += run.text;
      });
      
      paragraphs.push(paragraph);
    });
  } catch (error) {
    console.error('Error parsing paragraphs:', error);
  }
  
  return paragraphs;
}

/**
 * Parse sections from document
 * @param {Document} documentDoc - Document XML
 * @returns {Array} - Section information
 */
function parseSections(documentDoc) {
  const sections = [];
  
  try {
    // Get section properties from the document body
    const sectPrNodes = selectNodes("//w:sectPr", documentDoc);
    
    sectPrNodes.forEach((sectPrNode, index) => {
      const section = {
        id: `section${index}`,
        properties: {}
      };
      
      // Page size
      const pgSzNode = selectSingleNode("w:pgSz", sectPrNode);
      if (pgSzNode) {
        section.properties.pageSize = {
          width: parseInt(pgSzNode.getAttribute('w:w') || '12240', 10), // Default is 8.5"
          height: parseInt(pgSzNode.getAttribute('w:h') || '15840', 10), // Default is 11"
          orientation: pgSzNode.getAttribute('w:orient') || 'portrait'
        };
      }
      
      // Page margins
      const pgMarNode = selectSingleNode("w:pgMar", sectPrNode);
      if (pgMarNode) {
        section.properties.margins = {
          top: parseInt(pgMarNode.getAttribute('w:top') || '1440', 10), // Default is 1"
          right: parseInt(pgMarNode.getAttribute('w:right') || '1440', 10),
          bottom: parseInt(pgMarNode.getAttribute('w:bottom') || '1440', 10),
          left: parseInt(pgMarNode.getAttribute('w:left') || '1440', 10),
          header: parseInt(pgMarNode.getAttribute('w:header') || '720', 10), // Default is 0.5"
          footer: parseInt(pgMarNode.getAttribute('w:footer') || '720', 10),
          gutter: parseInt(pgMarNode.getAttribute('w:gutter') || '0', 10)
        };
      }
      
      // Columns
      const colsNode = selectSingleNode("w:cols", sectPrNode);
      if (colsNode) {
        section.properties.columns = {
          count: parseInt(colsNode.getAttribute('w:num') || '1', 10),
          spacing: parseInt(colsNode.getAttribute('w:space') || '720', 10) // Default is 0.5"
        };
      }
      
      // Headers and footers
      const headerReferenceNodes = selectNodes("w:headerReference", sectPrNode);
      const footerReferenceNodes = selectNodes("w:footerReference", sectPrNode);
      
      if (headerReferenceNodes.length > 0) {
        section.properties.headers = {};
        
        headerReferenceNodes.forEach(headerNode => {
          const type = headerNode.getAttribute('w:type');
          const id = headerNode.getAttribute('r:id');
          
          if (type && id) {
            section.properties.headers[type] = id;
          }
        });
      }
      
      if (footerReferenceNodes.length > 0) {
        section.properties.footers = {};
        
        footerReferenceNodes.forEach(footerNode => {
          const type = footerNode.getAttribute('w:type');
          const id = footerNode.getAttribute('r:id');
          
          if (type && id) {
            section.properties.footers[type] = id;
          }
        });
      }
      
      sections.push(section);
    });
  } catch (error) {
    console.error('Error parsing sections:', error);
  }
  
  return sections;
}

/**
 * Parse tables from document
 * @param {Document} documentDoc - Document XML
 * @returns {Array} - Table information
 */
function parseTables(documentDoc) {
  const tables = [];
  
  try {
    const tableNodes = selectNodes("//w:tbl", documentDoc);
    
    tableNodes.forEach((tblNode, index) => {
      const table = {
        id: `table${index}`,
        properties: {},
        rows: []
      };
      
      // Get table properties
      const tblPrNode = selectSingleNode("w:tblPr", tblNode);
      if (tblPrNode) {
        table.properties = parseTableProperties(tblPrNode);
      }
      
      // Get table grid
      const tblGridNode = selectSingleNode("w:tblGrid", tblNode);
      if (tblGridNode) {
        table.grid = [];
        
        const gridColNodes = selectNodes("w:gridCol", tblGridNode);
        gridColNodes.forEach(gridColNode => {
          const width = parseInt(gridColNode.getAttribute('w:w') || '0', 10);
          table.grid.push(width);
        });
      }
      
      // Get rows
      const rowNodes = selectNodes("w:tr", tblNode);
      rowNodes.forEach((trNode, rowIndex) => {
        const row = {
          id: `${table.id}_r${rowIndex}`,
          properties: {},
          cells: []
        };
        
        // Get row properties
        const trPrNode = selectSingleNode("w:trPr", trNode);
        if (trPrNode) {
          // Row height
          const trHeightNode = selectSingleNode("w:trHeight", trPrNode);
          if (trHeightNode) {
            row.properties.height = {
              value: parseInt(trHeightNode.getAttribute('w:val') || '0', 10),
              rule: trHeightNode.getAttribute('w:hRule') || 'auto'
            };
          }
          
          // Header row
          const tblHeaderNode = selectSingleNode("w:tblHeader", trPrNode);
          if (tblHeaderNode) {
            row.properties.header = true;
          }
        }
        
        // Get cells
        const cellNodes = selectNodes("w:tc", trNode);
        cellNodes.forEach((tcNode, cellIndex) => {
          const cell = {
            id: `${row.id}_c${cellIndex}`,
            properties: {},
            content: []
          };
          
          // Get cell properties
          const tcPrNode = selectSingleNode("w:tcPr", tcNode);
          if (tcPrNode) {
            cell.properties = parseTableCellProperties(tcPrNode);
            
            // Grid span
            const gridSpanNode = selectSingleNode("w:gridSpan", tcPrNode);
            if (gridSpanNode) {
              cell.properties.gridSpan = parseInt(gridSpanNode.getAttribute('w:val') || '1', 10);
            }
            
            // Vertical merge
            const vMergeNode = selectSingleNode("w:vMerge", tcPrNode);
            if (vMergeNode) {
              const val = vMergeNode.getAttribute('w:val');
              cell.properties.vMerge = val || 'continue';
            }
          }
          
          // Get paragraphs in cell
          const paragraphNodes = selectNodes("w:p", tcNode);
          paragraphNodes.forEach((pNode, pIndex) => {
            const paragraph = {
              id: `${cell.id}_p${pIndex}`,
              styleId: null,
              properties: {},
              text: '',
              runs: []
            };
            
            // Get paragraph properties
            const pPrNode = selectSingleNode("w:pPr", pNode);
            if (pPrNode) {
              paragraph.properties = parseParagraphProperties(pPrNode);
              
              // Get style ID
              const pStyleNode = selectSingleNode("w:pStyle", pPrNode);
              if (pStyleNode) {
                paragraph.styleId = pStyleNode.getAttribute('w:val');
              }
            }
            
            // Get runs
            const runNodes = selectNodes("w:r", pNode);
            runNodes.forEach((rNode, runIndex) => {
              const run = {
                id: `${paragraph.id}_r${runIndex}`,
                properties: {},
                text: ''
              };
              
              // Get run properties
              const rPrNode = selectSingleNode("w:rPr", rNode);
              if (rPrNode) {
                run.properties = parseRunProperties(rPrNode);
              }
              
              // Get text
              const textNodes = selectNodes("w:t", rNode);
              textNodes.forEach(tNode => {
                run.text += tNode.textContent;
              });
              
              paragraph.runs.push(run);
              paragraph.text += run.text;
            });
            
            cell.content.push(paragraph);
          });
          
          row.cells.push(cell);
        });
        
        table.rows.push(row);
      });
      
      tables.push(table);
    });
  } catch (error) {
    console.error('Error parsing tables:', error);
  }
  
  return tables;
}

/**
 * Parse lists from document
 * @param {Document} documentDoc - Document XML
 * @param {Document} numberingDoc - Numbering XML document
 * @returns {Array} - List information
 */
function parseListsFromDocument(documentDoc, numberingDoc) {
  const lists = [];
  
  try {
    if (!numberingDoc) return lists;
    
    // First, get all paragraphs with numbering properties
    const numberedParaNodes = selectNodes("//w:p[.//w:numPr]", documentDoc);
    
    // Group paragraphs by numId to determine lists
    const listsByNumId = {};
    
    numberedParaNodes.forEach((pNode, index) => {
      // Get numbering properties
      const numPrNode = selectSingleNode(".//w:numPr", pNode);
      if (!numPrNode) return;
      
      const numIdNode = selectSingleNode("w:numId", numPrNode);
      const ilvlNode = selectSingleNode("w:ilvl", numPrNode);
      
      if (!numIdNode) return;
      
      const numId = numIdNode.getAttribute('w:val');
      const level = ilvlNode ? parseInt(ilvlNode.getAttribute('w:val'), 10) : 0;
      
      if (!numId) return;
      
      // Create or update list entry
      if (!listsByNumId[numId]) {
        listsByNumId[numId] = {
          id: `list_${numId}`,
          numId,
          items: []
        };
      }
      
      // Get paragraph text
      let text = '';
      const textNodes = selectNodes(".//w:t", pNode);
      textNodes.forEach(tNode => {
        text += tNode.textContent;
      });
      
      // Add item to list
      listsByNumId[numId].items.push({
        id: `list_${numId}_item${index}`,
        level,
        text,
        node: pNode
      });
    });
    
    // Convert to arrays
    Object.values(listsByNumId).forEach(list => {
      // Sort items by their order in the document
      list.items.sort((a, b) => {
        const posA = getNodePosition(a.node);
        const posB = getNodePosition(b.node);
        return posA - posB;
      });
      
      // Add list to result
      lists.push(list);
    });
  } catch (error) {
    console.error('Error parsing lists:', error);
  }
  
  return lists;
}

/**
 * Get node position in document
 * @param {Node} node - XML node
 * @returns {number} - Position index
 */
function getNodePosition(node) {
  let count = 0;
  let current = node;
  
  // Count previous siblings
  while (current.previousSibling) {
    current = current.previousSibling;
    count++;
  }
  
  return count;
}

/**
 * Check for right-to-left text in document
 * @param {Document} documentDoc - Document XML
 * @returns {boolean} - True if document contains RTL text
 */
function checkForRTL(documentDoc) {
  try {
    // Check for RTL paragraph alignment
    const rtlParagraphs = selectNodes("//w:p/w:pPr/w:bidi[@w:val='1']", documentDoc);
    if (rtlParagraphs.length > 0) {
      return true;
    }
    
    // Check for RTL text runs
    const rtlRuns = selectNodes("//w:r/w:rPr/w:rtl[@w:val='1']", documentDoc);
    if (rtlRuns.length > 0) {
      return true;
    }
    
    // Check for RTL document settings
    const rtlSettings = selectNodes("//w:rtlGutter[@w:val='1']", documentDoc);
    if (rtlSettings.length > 0) {
      return true;
    }
    
    return false;
  } catch (error) {
    console.error('Error checking for RTL:', error);
    return false;
  }
}

/**
 * Generate CSS from extracted document information
 * @param {Object} documentInfo - Document information
 * @returns {string} - CSS stylesheet
 */
function generateCssFromStyleInfo(documentInfo) {
  let css = '';
  
  try {
    // Generate default CSS
    css += generateDefaultCss(documentInfo);
    
    // Generate style-specific CSS
    css += generateStyleCss(documentInfo);
    
    // Generate table CSS
    css += generateTableCss(documentInfo);
    
    // Generate list CSS
    css += generateListCss(documentInfo);
    
    // Generate RTL support
    if (documentInfo.hasRTL) {
      css += generateRtlCss();
    }
    
    // Generate print styles
    css += generatePrintCss();
    
    // Generate utility styles
    css += generateUtilityCss(documentInfo);
  } catch (error) {
    console.error('Error generating CSS:', error);
    css = generateFallbackCss();
  }
  
  return css;
}

/**
 * Generate default CSS rules
 * @param {Object} documentInfo - Document information
 * @returns {string} - CSS rules for default styles
 */
function generateDefaultCss(documentInfo) {
  // Get default font family
  let fontFamily = 'Arial, sans-serif';
  
  if (documentInfo.theme && documentInfo.theme.fonts) {
    const defaultFont = documentInfo.theme.fonts.minor || documentInfo.theme.fonts.major;
    if (defaultFont) {
      fontFamily = `"${defaultFont}", Arial, sans-serif`;
    }
  }
  
  // Get default font size
  let fontSize = '12pt';
  
  if (documentInfo.styles && 
      documentInfo.styles.defaults && 
      documentInfo.styles.defaults.character && 
      documentInfo.styles.defaults.character.size) {
    fontSize = `${documentInfo.styles.defaults.character.size / 2}pt`;
  }
  
  return `
/* Document defaults */
body {
  font-family: ${fontFamily};
  font-size: ${fontSize};
  line-height: 1.5;
  margin: 20px;
  padding: 0;
  color: #333;
  background-color: #fff;
}

/* Responsive design */
@media (max-width: 768px) {
  body {
    margin: 10px;
    font-size: ${parseInt(fontSize, 10) + 2}pt;
  }
}

/* Default heading styles */
h1, h2, h3, h4, h5, h6 {
  font-family: ${fontFamily};
  margin-top: 1.5em;
  margin-bottom: 0.5em;
  font-weight: bold;
}

h1 { font-size: 1.8em; }
h2 { font-size: 1.5em; }
h3 { font-size: 1.3em; }
h4 { font-size: 1.2em; }
h5 { font-size: 1.1em; }
h6 { font-size: 1em; }

/* Default paragraph */
p {
  margin-top: 0;
  margin-bottom: 1em;
}

/* Default list styles */
ul, ol {
  margin-top: 0;
  margin-bottom: 1em;
  padding-left: 2em;
}

/* Default table styles */
table {
  border-collapse: collapse;
  width: 100%;
  margin-bottom: 1em;
}

th, td {
  border: 1px solid #ddd;
  padding: 8px;
  vertical-align: top;
}

th {
  background-color: #f2f2f2;
  font-weight: bold;
  text-align: left;
}

/* Default image styles */
img {
  max-width: 100%;
  height: auto;
}
`;
}

/**
 * Generate CSS for specific styles
 * @param {Object} documentInfo - Document information
 * @returns {string} - CSS for styles
 */
function generateStyleCss(documentInfo) {
  let css = '';
  
  if (!documentInfo.styles || !documentInfo.styles.paragraph) {
    return css;
  }
  
  // Process paragraph styles
  Object.entries(documentInfo.styles.paragraph).forEach(([id, style]) => {
    if (!style.name || !style.properties) return;
    
    // Create class name from style ID
    const className = `.doc-p-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`;
    
    css += `${className} {\n`;
    
    // Add paragraph properties
    if (style.properties.paragraph) {
      // Alignment
      if (style.properties.paragraph.alignment) {
        css += `  text-align: ${convertAlignment(style.properties.paragraph.alignment)};\n`;
      }
      
      // Indentation
      if (style.properties.paragraph.indentation) {
        const indentation = style.properties.paragraph.indentation;
        
        if (indentation.left) {
          css += `  margin-left: ${twipToPt(indentation.left)}pt;\n`;
        }
        
        if (indentation.right) {
          css += `  margin-right: ${twipToPt(indentation.right)}pt;\n`;
        }
        
        if (indentation.firstLine) {
          css += `  text-indent: ${twipToPt(indentation.firstLine)}pt;\n`;
        }
        
        if (indentation.hanging) {
          css += `  text-indent: -${twipToPt(indentation.hanging)}pt;\n`;
        }
      }
      
      // Spacing
      if (style.properties.paragraph.spacing) {
        const spacing = style.properties.paragraph.spacing;
        
        if (spacing.before) {
          css += `  margin-top: ${twipToPt(spacing.before)}pt;\n`;
        }
        
        if (spacing.after) {
          css += `  margin-bottom: ${twipToPt(spacing.after)}pt;\n`;
        }
        
        if (spacing.line) {
          if (spacing.lineRule === 'auto') {
            // Line spacing as multiple
            const lineSpacing = spacing.line / 240; // 240 = 100%
            css += `  line-height: ${lineSpacing.toFixed(2)};\n`;
          } else if (spacing.lineRule === 'exact' || spacing.lineRule === 'atLeast') {
            // Line spacing as absolute value
            css += `  line-height: ${twipToPt(spacing.line)}pt;\n`;
          }
        }
      }
      
      // Borders
      if (style.properties.paragraph.borders) {
        const borders = style.properties.paragraph.borders;
        
        if (borders.top) {
          css += `  border-top: ${generateBorderStyle(borders.top)};\n`;
        }
        
        if (borders.right) {
          css += `  border-right: ${generateBorderStyle(borders.right)};\n`;
        }
        
        if (borders.bottom) {
          css += `  border-bottom: ${generateBorderStyle(borders.bottom)};\n`;
        }
        
        if (borders.left) {
          css += `  border-left: ${generateBorderStyle(borders.left)};\n`;
        }
      }
      
      // Shading
      if (style.properties.paragraph.shading) {
        const shading = style.properties.paragraph.shading;
        
        if (shading.fill && shading.fill !== 'auto') {
          css += `  background-color: #${shading.fill};\n`;
        }
        
        if (shading.color && shading.color !== 'auto') {
          css += `  color: #${shading.color};\n`;
        }
      }
      
      // RTL
      if (style.properties.paragraph.bidi) {
        css += `  direction: rtl;\n`;
        css += `  text-align: right;\n`;
      }
    }
    
    // Add character properties
    if (style.properties.character) {
      // Font
      if (style.properties.character.fonts) {
        const fonts = style.properties.character.fonts;
        const fontFamily = fonts.ascii || fonts.hAnsi || 'inherit';
        css += `  font-family: "${fontFamily}", Arial, sans-serif;\n`;
      }
      
      // Size
      if (style.properties.character.size) {
        css += `  font-size: ${style.properties.character.size / 2}pt;\n`;
      }
      
      // Bold
      if (style.properties.character.bold) {
        css += `  font-weight: bold;\n`;
      }
      
      // Italic
      if (style.properties.character.italic) {
        css += `  font-style: italic;\n`;
      }
      
      // Underline
      if (style.properties.character.underline) {
        css += `  text-decoration: underline;\n`;
      }
      
      // Strike
      if (style.properties.character.strike) {
        css += `  text-decoration: line-through;\n`;
      }
      
      // Color
      if (style.properties.character.color) {
        css += `  color: #${style.properties.character.color};\n`;
      }
      
      // All caps
      if (style.properties.character.caps) {
        css += `  text-transform: uppercase;\n`;
      }
      
      // Small caps
      if (style.properties.character.smallCaps) {
        css += `  font-variant: small-caps;\n`;
      }
      
      // Vertical alignment
      if (style.properties.character.verticalAlignment) {
        if (style.properties.character.verticalAlignment === 'superscript') {
          css += `  vertical-align: super;\n`;
          css += `  font-size: smaller;\n`;
        } else if (style.properties.character.verticalAlignment === 'subscript') {
          css += `  vertical-align: sub;\n`;
          css += `  font-size: smaller;\n`;
        }
      }
    }
    
    css += `}\n\n`;
  });
  
  // Process character styles
  Object.entries(documentInfo.styles.character || {}).forEach(([id, style]) => {
    if (!style.name || !style.properties) return;
    
    // Create class name from style ID
    const className = `.doc-c-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`;
    
    css += `${className} {\n`;
    
    // Font
    if (style.properties.fonts) {
      const fonts = style.properties.fonts;
      const fontFamily = fonts.ascii || fonts.hAnsi || 'inherit';
      css += `  font-family: "${fontFamily}", Arial, sans-serif;\n`;
    }
    
    // Size
    if (style.properties.size) {
      css += `  font-size: ${style.properties.size / 2}pt;\n`;
    }
    
    // Bold
    if (style.properties.bold) {
      css += `  font-weight: bold;\n`;
    }
    
    // Italic
    if (style.properties.italic) {
      css += `  font-style: italic;\n`;
    }
    
    // Underline
    if (style.properties.underline) {
      css += `  text-decoration: underline;\n`;
    }
    
    // Strike
    if (style.properties.strike) {
      css += `  text-decoration: line-through;\n`;
    }
    
    // Color
    if (style.properties.color) {
      css += `  color: #${style.properties.color};\n`;
    }
    
    // All caps
    if (style.properties.caps) {
      css += `  text-transform: uppercase;\n`;
    }
    
    // Small caps
    if (style.properties.smallCaps) {
      css += `  font-variant: small-caps;\n`;
    }
    
    // Vertical alignment
    if (style.properties.verticalAlignment) {
      if (style.properties.verticalAlignment === 'superscript') {
        css += `  vertical-align: super;\n`;
        css += `  font-size: smaller;\n`;
      } else if (style.properties.verticalAlignment === 'subscript') {
        css += `  vertical-align: sub;\n`;
        css += `  font-size: smaller;\n`;
      }
    }
    
    css += `}\n\n`;
  });
  
  // Generate heading styles based on document paragraphs
  const headingStyles = [
    { name: 'heading 1', selector: 'h1' },
    { name: 'heading 2', selector: 'h2' },
    { name: 'heading 3', selector: 'h3' },
    { name: 'heading 4', selector: 'h4' },
    { name: 'heading 5', selector: 'h5' },
    { name: 'heading 6', selector: 'h6' }
  ];
  
  headingStyles.forEach(heading => {
    // Find style with matching name (case-insensitive)
    const matchingStyle = Object.values(documentInfo.styles.paragraph || {}).find(style => 
      style.name.toLowerCase() === heading.name
    );
    
    if (matchingStyle && matchingStyle.properties) {
      css += `${heading.selector} {\n`;
      
      // Add character properties
      if (matchingStyle.properties.character) {
        // Font
        if (matchingStyle.properties.character.fonts) {
          const fonts = matchingStyle.properties.character.fonts;
          const fontFamily = fonts.ascii || fonts.hAnsi || 'inherit';
          css += `  font-family: "${fontFamily}", Arial, sans-serif;\n`;
        }
        
        // Size
        if (matchingStyle.properties.character.size) {
          css += `  font-size: ${matchingStyle.properties.character.size / 2}pt;\n`;
        }
        
        // Bold - headings are usually bold by default, so we only need to set it to normal if it's explicitly false
        if (matchingStyle.properties.character.bold === false) {
          css += `  font-weight: normal;\n`;
        }
        
        // Italic
        if (matchingStyle.properties.character.italic) {
          css += `  font-style: italic;\n`;
        }
        
        // Color
        if (matchingStyle.properties.character.color) {
          css += `  color: #${matchingStyle.properties.character.color};\n`;
        }
      }
      
      // Add paragraph properties
      if (matchingStyle.properties.paragraph) {
        // Spacing
        if (matchingStyle.properties.paragraph.spacing) {
          const spacing = matchingStyle.properties.paragraph.spacing;
          
          if (spacing.before) {
            css += `  margin-top: ${twipToPt(spacing.before)}pt;\n`;
          }
          
          if (spacing.after) {
            css += `  margin-bottom: ${twipToPt(spacing.after)}pt;\n`;
          }
        }
      }
      
      css += `}\n\n`;
    }
  });
  
  return css;
}

/**
 * Generate CSS for tables
 * @param {Object} documentInfo - Document information
 * @returns {string} - CSS for tables
 */
function generateTableCss(documentInfo) {
  let css = '';
  
  if (!documentInfo.tables || !documentInfo.styles || !documentInfo.styles.table) {
    return css;
  }
  
  // Add table wrapper for responsiveness
  css += `
/* Table wrapper for responsiveness */
.doc-table-container {
  overflow-x: auto;
  margin-bottom: 1em;
}
`;
  
  // Process table styles
  Object.entries(documentInfo.styles.table).forEach(([id, style]) => {
    if (!style.name || !style.properties) return;
    
    // Create class name from style ID
    const className = `.doc-t-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`;
    
    css += `${className} {\n`;
    css += `  border-collapse: collapse;\n`;
    css += `  width: 100%;\n`;
    
    // Add table properties
    if (style.properties.table) {
      // Table width
      if (style.properties.table.width) {
        if (style.properties.table.width.type === 'dxa') {
          css += `  width: ${twipToPt(style.properties.table.width.width)}pt;\n`;
        } else if (style.properties.table.width.type === 'pct') {
          css += `  width: ${style.properties.table.width.width / 50}%;\n`;
        }
      }
      
      // Table layout
      if (style.properties.table.layout === 'fixed') {
        css += `  table-layout: fixed;\n`;
      }
    }
    
    css += `}\n\n`;
    
    // Cell styles
    css += `${className} td, ${className} th {\n`;
    css += `  border: 1px solid #ddd;\n`;
    css += `  padding: 8px;\n`;
    css += `  vertical-align: top;\n`;
    
    // Add cell margins
    if (style.properties.table && style.properties.table.cellMargins) {
      const margins = style.properties.table.cellMargins;
      
      if (margins.top) {
        css += `  padding-top: ${twipToPt(margins.top.width)}pt;\n`;
      }
      
      if (margins.right) {
        css += `  padding-right: ${twipToPt(margins.right.width)}pt;\n`;
      }
      
      if (margins.bottom) {
        css += `  padding-bottom: ${twipToPt(margins.bottom.width)}pt;\n`;
      }
      
      if (margins.left) {
        css += `  padding-left: ${twipToPt(margins.left.width)}pt;\n`;
      }
    }
    
    // Add borders
    if (style.properties.table && style.properties.table.borders) {
      const borders = style.properties.table.borders;
      
      if (borders.top) {
        css += `  border-top: ${generateBorderStyle(borders.top)};\n`;
      }
      
      if (borders.right) {
        css += `  border-right: ${generateBorderStyle(borders.right)};\n`;
      }
      
      if (borders.bottom) {
        css += `  border-bottom: ${generateBorderStyle(borders.bottom)};\n`;
      }
      
      if (borders.left) {
        css += `  border-left: ${generateBorderStyle(borders.left)};\n`;
      }
    }
    
    css += `}\n\n`;
    
    // Add conditional formatting if available
    if (style.properties.conditionalFormatting) {
      // First row formatting
      if (style.properties.conditionalFormatting.firstRow) {
        css += `${className} tr:first-child td, ${className} tr:first-child th {\n`;
        
        const formatting = style.properties.conditionalFormatting.firstRow;
        
        // Add character properties
        if (formatting.character) {
          if (formatting.character.bold) {
            css += `  font-weight: bold;\n`;
          }
          
          if (formatting.character.color) {
            css += `  color: #${formatting.character.color};\n`;
          }
        }
        
        // Add cell properties
        if (formatting.table) {
          if (formatting.table.shading) {
            css += `  background-color: #${formatting.table.shading.fill};\n`;
          }
          
          if (formatting.table.borders) {
            const borders = formatting.table.borders;
            
            if (borders.top) {
              css += `  border-top: ${generateBorderStyle(borders.top)};\n`;
            }
            
            if (borders.bottom) {
              css += `  border-bottom: ${generateBorderStyle(borders.bottom)};\n`;
            }
          }
        }
        
        css += `}\n\n`;
      }
      
      // Banded row formatting
      if (style.properties.conditionalFormatting.band1Horz) {
        css += `${className} tr:nth-child(odd) {\n`;
        
        const formatting = style.properties.conditionalFormatting.band1Horz;
        
        if (formatting.table && formatting.table.shading) {
          css += `  background-color: #${formatting.table.shading.fill};\n`;
        }
        
        css += `}\n\n`;
      }
      
      if (style.properties.conditionalFormatting.band2Horz) {
        css += `${className} tr:nth-child(even) {\n`;
        
        const formatting = style.properties.conditionalFormatting.band2Horz;
        
        if (formatting.table && formatting.table.shading) {
          css += `  background-color: #${formatting.table.shading.fill};\n`;
        }
        
        css += `}\n\n`;
      }
    }
  });
  
  return css;
}

/**
 * Generate CSS for lists
 * @param {Object} documentInfo - Document information
 * @returns {string} - CSS for lists
 */
function generateListCss(documentInfo) {
  let css = '';
  
  if (!documentInfo.numbering || !documentInfo.numbering.abstractNums) {
    return css;
  }
  
  // Basic list styles
  css += `
/* List container */
.doc-list {
  margin: 0;
  padding: 0;
  list-style-position: outside;
}

/* List item */
.doc-list-item {
  position: relative;
}
`;
  
  // Process abstract numbering definitions
  Object.entries(documentInfo.numbering.abstractNums).forEach(([abstractId, abstractNum]) => {
    if (!abstractNum.levels) return;
    
    // Process each level
    Object.entries(abstractNum.levels).forEach(([level, levelDef]) => {
      // Create class names
      const listClass = `.doc-list-${abstractId}`;
      const levelClass = `.doc-list-${abstractId}-${level}`;
      
      // Determine list style type
      let listStyleType = 'decimal';
      
      switch (levelDef.format) {
        case 'decimal': listStyleType = 'decimal'; break;
        case 'lowerLetter': listStyleType = 'lower-alpha'; break;
        case 'upperLetter': listStyleType = 'upper-alpha'; break;
        case 'lowerRoman': listStyleType = 'lower-roman'; break;
        case 'upperRoman': listStyleType = 'upper-roman'; break;
        case 'bullet': listStyleType = 'disc'; break;
        default: listStyleType = 'decimal';
      }
      
      // Add list style
      css += `${listClass} {\n`;
      css += `  list-style-type: ${listStyleType};\n`;
      
      // If there's a level 0, assume it's an ordered list
      if (abstractNum.levels['0']) {
        css += `  counter-reset: item;\n`;
      }
      
      css += `}\n\n`;
      
      // Add level-specific styles
      css += `${levelClass} {\n`;
      
      // Add indentation
      const leftIndent = levelDef.properties.paragraph && levelDef.properties.paragraph.indentation ?
                        (levelDef.properties.paragraph.indentation.left || 0) : 
                        (parseInt(level, 10) + 1) * 720; // Default is 0.5" per level
      
      css += `  margin-left: ${twipToPt(leftIndent)}pt;\n`;
      
      // Add text properties from paragraph style
      if (levelDef.properties.paragraph) {
        if (levelDef.properties.paragraph.indentation && levelDef.properties.paragraph.indentation.hanging) {
          const hanging = levelDef.properties.paragraph.indentation.hanging;
          css += `  text-indent: -${twipToPt(hanging)}pt;\n`;
        }
      }
      
      // Add text properties from character style
      if (levelDef.properties.character) {
        if (levelDef.properties.character.bold) {
          css += `  font-weight: bold;\n`;
        }
        
        if (levelDef.properties.character.italic) {
          css += `  font-style: italic;\n`;
        }
        
        if (levelDef.properties.character.color) {
          css += `  color: #${levelDef.properties.character.color};\n`;
        }
      }
      
      css += `}\n\n`;
    });
  });
  
  return css;
}

/**
 * Generate CSS for RTL support
 * @returns {string} - CSS for RTL
 */
function generateRtlCss() {
  return `
/* RTL support */
.doc-rtl {
  direction: rtl;
  text-align: right;
}

[dir="rtl"] {
  direction: rtl;
  text-align: right;
}

/* Language-specific styling */
[lang="ar"], [lang="he"], [lang="fa"], [lang="ur"] {
  direction: rtl;
  text-align: right;
}
`;
}

/**
 * Generate CSS for print
 * @returns {string} - CSS for print
 */
function generatePrintCss() {
  return `
/* Print styles */
@media print {
  body {
    font-size: 10pt;
    margin: 0;
  }
  
  h1, h2, h3, h4, h5, h6 {
    page-break-after: avoid;
    page-break-inside: avoid;
  }
  
  p, li {
    orphans: 2;
    widows: 2;
  }
  
  img, table, figure {
    page-break-inside: avoid;
  }
  
  .doc-table-container {
    overflow-x: visible;
  }
}
`;
}

/**
 * Generate utility CSS
 * @param {Object} documentInfo - Document information
 * @returns {string} - CSS for utilities
 */
function generateUtilityCss(documentInfo) {
  // Calculate default tab width
  const defaultTabStop = documentInfo.settings && documentInfo.settings.defaultTabStop ? 
                        twipToPt(documentInfo.settings.defaultTabStop) : 
                        36; // Default is 0.5" (36pt)
  
  return `
/* Utility styles */
.doc-underline { text-decoration: underline; }
.doc-strike { text-decoration: line-through; }
.doc-tab { display: inline-block; width: ${defaultTabStop}pt; }
.doc-break { display: block; margin-bottom: 1em; }
.doc-superscript { vertical-align: super; font-size: smaller; }
.doc-subscript { vertical-align: sub; font-size: smaller; }
.doc-smallcaps { font-variant: small-caps; }
.doc-caps { text-transform: uppercase; }
.doc-hidden { display: none; }

/* Language-specific fonts */
[lang="zh-CN"], [lang="zh-SG"] {
  font-family: "SimSun", "NSimSun", "Microsoft YaHei", sans-serif;
}

[lang="zh-TW"], [lang="zh-HK"] {
  font-family: "MingLiU", "PMingLiU", "Microsoft JhengHei", sans-serif;
}

[lang="ja"] {
  font-family: "MS Mincho", "MS PMincho", "Meiryo", sans-serif;
}

[lang="ko"] {
  font-family: "Batang", "Gulim", "Malgun Gothic", sans-serif;
}

[lang="ru"] {
  font-family: "Times New Roman", "Arial", sans-serif;
}

[lang="ar"], [lang="he"], [lang="fa"], [lang="ur"] {
  font-family: "Tahoma", "Arial", sans-serif;
}
`;
}

/**
 * Generate fallback CSS
 * @returns {string} - Fallback CSS
 */
function generateFallbackCss() {
  return `
/* Fallback styles */
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  font-size: 12pt;
  line-height: 1.5;
  margin: 20px;
  padding: 0;
  color: #333;
  background-color: #fff;
}

h1, h2, h3, h4, h5, h6 {
  margin-top: 1em;
  margin-bottom: 0.5em;
  line-height: 1.3;
}

h1 { font-size: 1.8em; }
h2 { font-size: 1.5em; }
h3 { font-size: 1.3em; }
h4 { font-size: 1.1em; }
h5 { font-size: 1em; font-weight: bold; }
h6 { font-size: 1em; font-style: italic; }

p { margin-bottom: 1em; }

ul, ol { margin-bottom: 1em; padding-left: 2em; }

table {
  border-collapse: collapse;
  width: 100%;
  margin-bottom: 1em;
}

th, td {
  border: 1px solid #ddd;
  padding: 8px;
  vertical-align: top;
}

th {
  background-color: #f2f2f2;
  font-weight: bold;
}

img { max-width: 100%; height: auto; margin-bottom: 1em; }

@media (max-width: 768px) {
  body { margin: 10px; font-size: 14pt; }
}

@media print {
  body { margin: 0; font-size: 10pt; }
  h1, h2, h3, h4, h5, h6 { page-break-after: avoid; }
  img, table { page-break-inside: avoid; }
}
`;
}

/**
 * Convert twip value to points
 * @param {number} twip - Value in twentieths of a point
 * @returns {number} - Value in points
 */
function twipToPt(twip) {
  return Math.round((parseInt(twip, 10) || 0) / 20 * 100) / 100;
}

/**
 * Convert Word alignment to CSS alignment
 * @param {string} alignment - Word alignment value
 * @returns {string} - CSS alignment value
 */
function convertAlignment(alignment) {
  switch (alignment) {
    case 'left': return 'left';
    case 'right': return 'right';
    case 'center': return 'center';
    case 'both': return 'justify';
    default: return 'left';
  }
}

/**
 * Generate CSS border style
 * @param {Object} border - Border definition
 * @returns {string} - CSS border style
 */
function generateBorderStyle(border) {
  if (!border || !border.type || border.type === 'nil' || border.type === 'none') {
    return 'none';
  }
  
  const width = border.size / 8; // Border size is in 1/8th points
  const color = border.color ? `#${border.color}` : '#000000';
  
  let style = 'solid';
  switch (border.type) {
    case 'single': style = 'solid'; break;
    case 'double': style = 'double'; break;
    case 'dotted': style = 'dotted'; break;
    case 'dashed': style = 'dashed'; break;
    case 'dotDash': style = 'dashed'; break;
    case 'dotDotDash': style = 'dashed'; break;
    case 'triple': style = 'double'; break;
    case 'wave': style = 'wavy'; break;
    default: style = 'solid';
  }
  
  return `${width}pt ${style} ${color}`;
}

module.exports = {
  parseDocxStyles,
  generateCssFromStyleInfo,
  selectNodes,
  selectSingleNode,
  twipToPt
};
