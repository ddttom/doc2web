// docx-style-parser.js
const JSZip = require('jszip');
const fs = require('fs');
const { DOMParser } = require('jsdom').JSDOM;
const xpath = require('xpath');
const dom = require('xmldom').DOMParser;

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
    
    if (!styleXml || !documentXml) {
      throw new Error('Invalid DOCX file: missing core XML files');
    }
    
    // Parse XML content
    const styleDoc = new dom().parseFromString(styleXml);
    const documentDoc = new dom().parseFromString(documentXml);
    const themeDoc = themeXml ? new dom().parseFromString(themeXml) : null;
    const settingsDoc = settingsXml ? new dom().parseFromString(settingsXml) : null;
    
    // Parse styles from XML
    const styles = parseStyles(styleDoc);
    
    // Parse theme (colors, fonts)
    const theme = parseTheme(themeDoc);
    
    // Parse document defaults
    const documentDefaults = parseDocumentDefaults(styleDoc);
    
    // Parse document settings
    const settings = parseSettings(settingsDoc);
    
    // Return combined style information
    return {
      styles,
      theme,
      documentDefaults,
      settings
    };
  } catch (error) {
    console.error('Error parsing DOCX styles:', error);
    // Return default style info if parsing fails
    return getDefaultStyleInfo();
  }
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
    
    // Select nodes
    const paragraphStyleNodes = xpath.select(paragraphStylesXPath, styleDoc) || [];
    const characterStyleNodes = xpath.select(characterStylesXPath, styleDoc) || [];
    const tableStyleNodes = xpath.select(tableStylesXPath, styleDoc) || [];
    const numberingStyleNodes = xpath.select(numberingStylesXPath, styleDoc) || [];
    
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
    const nameNode = xpath.select1("w:name", node);
    const name = nameNode ? nameNode.getAttribute('w:val') : styleId;
    
    // Get based on style
    const basedOnNode = xpath.select1("w:basedOn", node);
    const basedOn = basedOnNode ? basedOnNode.getAttribute('w:val') : null;
    
    // Check if default style
    const defaultNode = xpath.select1("@w:default", node);
    const isDefault = defaultNode && defaultNode.value === '1';
    
    // Parse running properties (text formatting)
    const rPrNode = xpath.select1("w:rPr", node);
    const runningProps = rPrNode ? parseRunningProperties(rPrNode) : {};
    
    // Parse paragraph properties
    const pPrNode = xpath.select1("w:pPr", node);
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
    const fontNode = xpath.select1("w:rFonts", rPrNode);
    if (fontNode) {
      props.font = {
        ascii: fontNode.getAttribute('w:ascii'),
        hAnsi: fontNode.getAttribute('w:hAnsi'),
        eastAsia: fontNode.getAttribute('w:eastAsia'),
        cs: fontNode.getAttribute('w:cs')
      };
    }
    
    // Size
    const szNode = xpath.select1("w:sz", rPrNode);
    if (szNode) {
      // Size is in half-points, convert to points
      const sizeHalfPoints = parseInt(szNode.getAttribute('w:val'), 10) || 22; // Default 11pt
      props.fontSize = (sizeHalfPoints / 2) + 'pt';
    }
    
    // Bold
    const bNode = xpath.select1("w:b", rPrNode);
    props.bold = bNode !== null;
    
    // Italic
    const iNode = xpath.select1("w:i", rPrNode);
    props.italic = iNode !== null;
    
    // Underline
    const uNode = xpath.select1("w:u", rPrNode);
    if (uNode) {
      props.underline = {
        type: uNode.getAttribute('w:val') || 'single'
      };
    }
    
    // Color
    const colorNode = xpath.select1("w:color", rPrNode);
    if (colorNode) {
      const colorVal = colorNode.getAttribute('w:val');
      if (colorVal) {
        props.color = '#' + colorVal;
      }
    }
    
    // Highlight
    const highlightNode = xpath.select1("w:highlight", rPrNode);
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
    const jcNode = xpath.select1("w:jc", pPrNode);
    if (jcNode) {
      props.alignment = jcNode.getAttribute('w:val');
    }
    
    // Indentation
    const indNode = xpath.select1("w:ind", pPrNode);
    if (indNode) {
      props.indentation = {
        left: indNode.getAttribute('w:left'),
        right: indNode.getAttribute('w:right'),
        firstLine: indNode.getAttribute('w:firstLine'),
        hanging: indNode.getAttribute('w:hanging')
      };
    }
    
    // Spacing
    const spacingNode = xpath.select1("w:spacing", pPrNode);
    if (spacingNode) {
      props.spacing = {
        before: spacingNode.getAttribute('w:before'),
        after: spacingNode.getAttribute('w:after'),
        line: spacingNode.getAttribute('w:line'),
        lineRule: spacingNode.getAttribute('w:lineRule')
      };
    }
    
    // Borders
    const pBdrNode = xpath.select1("w:pBdr", pPrNode);
    if (pBdrNode) {
      props.borders = parseBorders(pBdrNode);
    }
    
    // Shading
    const shadingNode = xpath.select1("w:shd", pPrNode);
    if (shadingNode) {
      props.shading = {
        value: shadingNode.getAttribute('w:val'),
        color: shadingNode.getAttribute('w:color'),
        fill: shadingNode.getAttribute('w:fill')
      };
    }
  } catch (error) {
    console.error('Error parsing paragraph properties:', error);
  }
  
  return props;
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
    const topNode = xpath.select1("w:top", borderNode);
    if (topNode) {
      borders.top = {
        value: topNode.getAttribute('w:val'),
        size: topNode.getAttribute('w:sz'),
        color: topNode.getAttribute('w:color')
      };
    }
    
    // Bottom border
    const bottomNode = xpath.select1("w:bottom", borderNode);
    if (bottomNode) {
      borders.bottom = {
        value: bottomNode.getAttribute('w:val'),
        size: bottomNode.getAttribute('w:sz'),
        color: bottomNode.getAttribute('w:color')
      };
    }
    
    // Left border
    const leftNode = xpath.select1("w:left", borderNode);
    if (leftNode) {
      borders.left = {
        value: leftNode.getAttribute('w:val'),
        size: leftNode.getAttribute('w:sz'),
        color: leftNode.getAttribute('w:color')
      };
    }
    
    // Right border
    const rightNode = xpath.select1("w:right", borderNode);
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
    const fontSchemeNode = xpath.select1("//a:fontScheme", themeDoc);
    if (fontSchemeNode) {
      const majorNode = xpath.select1(".//a:majorFont/a:latin", fontSchemeNode);
      const minorNode = xpath.select1(".//a:minorFont/a:latin", fontSchemeNode);
      
      if (majorNode && majorNode.getAttribute('typeface')) {
        theme.fonts.major = majorNode.getAttribute('typeface');
      }
      
      if (minorNode && minorNode.getAttribute('typeface')) {
        theme.fonts.minor = minorNode.getAttribute('typeface');
      }
    }
    
    // Parse color scheme
    const clrSchemeNode = xpath.select1("//a:clrScheme", themeDoc);
    if (clrSchemeNode) {
      // Get main theme colors
      const colorNodes = xpath.select("./a:*", clrSchemeNode) || [];
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
    const srgbNode = xpath.select1("./a:srgbClr", colorNode);
    if (srgbNode && srgbNode.getAttribute('val')) {
      return '#' + srgbNode.getAttribute('val');
    }
    
    // Check system color
    const sysClrNode = xpath.select1("./a:sysClr", colorNode);
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
    const docDefaultsNode = xpath.select1("//w:docDefaults", styleDoc);
    if (!docDefaultsNode) {
      return defaults;
    }
    
    // Get paragraph defaults
    const pPrDefaultNode = xpath.select1(".//w:pPrDefault/w:pPr", docDefaultsNode);
    if (pPrDefaultNode) {
      defaults.paragraph = parseParagraphProperties(pPrDefaultNode);
    }
    
    // Get character defaults
    const rPrDefaultNode = xpath.select1(".//w:rPrDefault/w:rPr", docDefaultsNode);
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
    const defaultTabStopNode = xpath.select1("//w:defaultTabStop", settingsDoc);
    if (defaultTabStopNode) {
      const val = defaultTabStopNode.getAttribute('w:val');
      if (val) {
        settings.defaultTabStop = val;
      }
    }
    
    // Character spacing control
    const characterSpacingControlNode = xpath.select1("//w:characterSpacingControl", settingsDoc);
    if (characterSpacingControlNode) {
      const val = characterSpacingControlNode.getAttribute('w:val');
      if (val) {
        settings.characterSpacing = val;
      }
    }
    
    // Do not hyphenate capital letters
    const doNotHyphenateCapsNode = xpath.select1("//w:doNotHyphenateCaps", settingsDoc);
    if (doNotHyphenateCapsNode) {
      const val = doNotHyphenateCapsNode.getAttribute('w:val');
      settings.doNotHyphenateCaps = val === 'true' || val === '1';
    }
    
    // Right-to-left document gutter
    const rtlGutterNode = xpath.select1("//w:rtlGutter", settingsDoc);
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
 * Generate CSS from extracted style information
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
  margin: 0;
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
`;
  } catch (error) {
    console.error('Error generating CSS:', error);
    
    // Provide fallback CSS
    css = `
body { font-family: Calibri, sans-serif; font-size: 11pt; line-height: 1.15; }
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
  if (border.value === 'nil' || border.value === 'none') {
    return `border-${side}: none;`;
  }
  
  const width = border.size ? convertBorderSizeToPt(border.size) : 1;
  const color = border.color ? `#${border.color}` : 'black';
  const style = getBorderTypeValue(border.value);
  
  return `border-${side}: ${width}pt ${style} ${color};`;
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
  generateCssFromStyleInfo
};
