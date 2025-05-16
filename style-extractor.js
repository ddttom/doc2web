// style-extractor.js
const mammoth = require('mammoth');
const { JSDOM } = require('jsdom');
const { parseDocxStyles, generateCssFromStyleInfo } = require('./docx-style-parser');

/**
 * Extract and apply styles from a DOCX document to HTML
 * @param {string} docxPath - Path to the DOCX file
 * @returns {Promise<{html: string, styles: string}>} - HTML with embedded styles
 */
async function extractAndApplyStyles(docxPath) {
  try {
    // First, extract the raw styles from the document
    const styleInfo = await parseDocxStyles(docxPath);
    
    // Generate CSS from the extracted styles
    const css = generateCssFromStyleInfo(styleInfo);
    
    // Convert DOCX to HTML with style preservation
    const htmlResult = await convertToStyledHtml(docxPath, styleInfo);
    
    // Combine HTML and CSS
    const styledHtml = applyStylesToHtml(htmlResult.value, css, styleInfo);
    
    return {
      html: styledHtml,
      styles: css,
      messages: htmlResult.messages
    };
  } catch (error) {
    console.error('Error extracting styles:', error);
    throw error;
  }
}

/**
 * Convert DOCX to HTML while preserving style information
 * @param {string} docxPath - Path to the DOCX file 
 * @param {Object} styleInfo - Extracted style information
 * @returns {Promise<{value: string, messages: Array}>} - HTML with style information
 */
async function convertToStyledHtml(docxPath, styleInfo) {
  // Create a custom style map based on extracted styles
  const styleMap = createStyleMap(styleInfo);
  
  // Custom document transformer to enhance style preservation
  const transformDocument = createDocumentTransformer(styleInfo);
  
  // Configure image handling
  const imageOptions = {
    convertImage: mammoth.images.imgElement(function(image) {
      return image.read("base64").then(function(imageBuffer) {
        const extension = image.contentType.split('/')[1];
        const filename = `image-${image.altText || Date.now()}.${extension}`;
        
        return {
          src: `./images/${filename}`,
          alt: image.altText || 'Image',
          className: 'docx-image'
        };
      });
    })
  };
  
  // Use mammoth to convert with the custom style map
  const result = await mammoth.convertToHtml({
    path: docxPath,
    styleMap: styleMap,
    transformDocument: transformDocument,
    includeDefaultStyleMap: true,
    ...imageOptions
  });
  
  return result;
}

/**
 * Create a custom style map based on extracted styles
 * @param {Object} styleInfo - Extracted style information
 * @returns {Array<string>} - Style map entries
 */
function createStyleMap(styleInfo) {
  const styleMap = [];
  
  // Map paragraph styles
  Object.entries(styleInfo.styles.paragraph || {}).forEach(([id, style]) => {
    styleMap.push(`p[style-name='${style.name}'] => p.docx-p-${id.toLowerCase()}`);
  });
  
  // Map character styles
  Object.entries(styleInfo.styles.character || {}).forEach(([id, style]) => {
    styleMap.push(`r[style-name='${style.name}'] => span.docx-c-${id.toLowerCase()}`);
  });
  
  // Map table styles
  Object.entries(styleInfo.styles.table || {}).forEach(([id, style]) => {
    styleMap.push(`table[style-name='${style.name}'] => table.docx-t-${id.toLowerCase()}`);
  });
  
  // Additional custom mappings for specific elements
  styleMap.push("p:fresh => p");
  styleMap.push("r[bold] => strong");
  styleMap.push("r[italic] => em");
  styleMap.push("r[underline] => span.docx-underline");
  styleMap.push("r[strikethrough] => span.docx-strike");
  styleMap.push("r[subscript] => sub");
  styleMap.push("r[superscript] => sup");

  return styleMap;
}

/**
 * Create document transformer function for style preservation
 * @param {Object} styleInfo - Style information
 * @returns {Function} - Document transformer
 */
function createDocumentTransformer(styleInfo) {
  return function(document) {
    // In a complete implementation, this would walk the document tree
    // and enhance elements with style attributes based on the styleInfo
    
    // Process paragraphs, runs, tables and other elements
    // Add class names and style attributes as needed
    
    return document;
  };
}

/**
 * Apply generated CSS to HTML
 * @param {string} html - HTML content
 * @param {string} css - CSS styles
 * @param {Object} styleInfo - Style information for reference
 * @returns {string} - HTML with embedded styles
 */
function applyStylesToHtml(html, css, styleInfo) {
  // Create a DOM to manipulate the HTML
  const dom = new JSDOM(html);
  const document = dom.window.document;
  
  // Create a style element
  const styleElement = document.createElement('style');
  styleElement.textContent = css;
  
  // Ensure we have a proper HTML structure
  ensureHtmlStructure(document);
  
  // Add to document head
  document.head.appendChild(styleElement);
  
  // Add metadata
  addDocumentMetadata(document, styleInfo);
  
  // Add body class for RTL if needed
  if (styleInfo.settings?.rtlGutter) {
    document.body.classList.add('docx-rtl');
  }
  
  // Process table elements to match word styling better
  processTables(document);
  
  // Process images to maintain aspect ratio and positioning
  processImages(document);
  
  // Handle language-specific elements
  processLanguageElements(document);
  
  // Serialize back to HTML string
  return dom.serialize();
}

/**
 * Ensure HTML has proper structure (html, head, body)
 * @param {Document} document - DOM document
 */
function ensureHtmlStructure(document) {
  // Create html element if not exists
  if (!document.documentElement || document.documentElement.nodeName.toLowerCase() !== 'html') {
    const html = document.createElement('html');
    
    // Move existing content
    while (document.childNodes.length > 0) {
      html.appendChild(document.childNodes[0]);
    }
    
    document.appendChild(html);
  }
  
  // Create head if not exists
  if (!document.head) {
    const head = document.createElement('head');
    document.documentElement.insertBefore(head, document.documentElement.firstChild);
  }
  
  // Add meta charset
  if (!document.querySelector('meta[charset]')) {
    const meta = document.createElement('meta');
    meta.setAttribute('charset', 'utf-8');
    document.head.appendChild(meta);
  }
  
  // Create body if not exists
  if (!document.body) {
    const body = document.createElement('body');
    
    // Move content to body
    Array.from(document.documentElement.childNodes).forEach(node => {
      if (node !== document.head && node.nodeName.toLowerCase() !== 'body') {
        body.appendChild(node);
      }
    });
    
    document.documentElement.appendChild(body);
  }
}

/**
 * Add document metadata based on style information
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function addDocumentMetadata(document, styleInfo) {
  // Add title
  const title = document.createElement('title');
  title.textContent = 'DOCX Document';
  document.head.appendChild(title);
  
  // Add viewport meta
  const viewport = document.createElement('meta');
  viewport.setAttribute('name', 'viewport');
  viewport.setAttribute('content', 'width=device-width, initial-scale=1');
  document.head.appendChild(viewport);
}

/**
 * Process tables for better styling
 * @param {Document} document - DOM document
 */
function processTables(document) {
  const tables = document.querySelectorAll('table');
  tables.forEach(table => {
    // Add default class if no class is present
    if (!table.classList.length) {
      table.classList.add('docx-table-default');
    }
    
    // Make sure tables have tbody
    if (!table.querySelector('tbody')) {
      const tbody = document.createElement('tbody');
      
      // Move rows to tbody
      Array.from(table.querySelectorAll('tr')).forEach(row => {
        tbody.appendChild(row);
      });
      
      table.appendChild(tbody);
    }
  });
}

/**
 * Process images for better styling
 * @param {Document} document - DOM document
 */
function processImages(document) {
  const images = document.querySelectorAll('img');
  images.forEach(img => {
    // Add default class if no class is present
    if (!img.classList.contains('docx-image')) {
      img.classList.add('docx-image');
    }
    
    // Ensure max-width for responsive images
    img.style.maxWidth = '100%';
    
    // Make sure images have alt text
    if (!img.hasAttribute('alt')) {
      img.setAttribute('alt', 'Document image');
    }
  });
}

/**
 * Process language-specific elements
 * @param {Document} document - DOM document
 */
function processLanguageElements(document) {
  // Find elements with dir="rtl" and add class
  const rtlElements = document.querySelectorAll('[dir="rtl"]');
  rtlElements.forEach(el => {
    el.classList.add('docx-rtl');
  });
}

module.exports = {
  extractAndApplyStyles
};
