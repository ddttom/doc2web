// style-extractor.js
const mammoth = require('mammoth');
const { JSDOM } = require('jsdom');
const path = require('path');
const { parseDocxStyles, generateCssFromStyleInfo } = require('./docx-style-parser');

/**
 * Extract and apply styles from a DOCX document to HTML
 * @param {string} docxPath - Path to the DOCX file
 * @param {string} cssFilename - Filename for the CSS file (without path)
 * @returns {Promise<{html: string, styles: string}>} - HTML with embedded styles
 */
async function extractAndApplyStyles(docxPath, cssFilename = null) {
  try {
    // First, extract the raw styles from the document
    const styleInfo = await parseDocxStyles(docxPath);
    
    // Generate CSS from the extracted styles
    const css = generateCssFromStyleInfo(styleInfo);
    
    // Convert DOCX to HTML with style preservation
    const htmlResult = await convertToStyledHtml(docxPath, styleInfo);
    
    // Get the CSS filename
    const cssFile = cssFilename || path.basename(docxPath, path.extname(docxPath)) + '.css';
    
    // Combine HTML and CSS
    const styledHtml = applyStylesToHtml(htmlResult.value, css, styleInfo, cssFile);
    
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
 * @param {string} cssFilename - Filename for the CSS file
 * @returns {string} - HTML with link to external CSS
 */
function applyStylesToHtml(html, css, styleInfo, cssFilename) {
  // Create a DOM to manipulate the HTML
  const dom = new JSDOM(html);
  const document = dom.window.document;
  
  // Ensure we have a proper HTML structure
  ensureHtmlStructure(document);
  
  // Instead of embedding CSS, add a link to the external CSS file
  const linkElement = document.createElement('link');
  linkElement.rel = 'stylesheet';
  linkElement.href = `./${cssFilename}`; // Use relative path to the CSS file
  
  // Add to document head
  document.head.appendChild(linkElement);
  
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
  
  // Process numbered paragraphs
  processNumberedParagraphs(document);
  
  // Detect and replace TOC and index elements
  detectAndReplaceTocAndIndex(document);
  
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

/**
 * Process numbered paragraphs for better styling
 * @param {Document} document - DOM document
 */
function processNumberedParagraphs(document) {
  const paragraphs = document.querySelectorAll('p');
  
  // Patterns for identifying different types of paragraph numbering
  const numberPattern = /^\s*(\d+)\.(.+)$/;
  const alphaPattern = /^\s*([a-z])\.(.+)$/;
  const romanPattern = /^\s*([ivx]+)\.(.+)$/;
  
  // Check for sequential paragraphs that might form a list
  let currentListType = null;
  let currentList = null;
  let currentItem = null;
  
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const text = p.textContent;
    
    // Check for different numbering patterns
    let match = null;
    let listType = null;
    
    if (numberPattern.test(text)) {
      match = text.match(numberPattern);
      listType = 'numbered';
    } else if (alphaPattern.test(text)) {
      match = text.match(alphaPattern);
      listType = 'alpha';
    } else if (romanPattern.test(text)) {
      match = text.match(romanPattern);
      listType = 'roman';
    }
    
    if (match) {
      // Extract the number/letter and content
      const prefix = match[1];
      const content = match[2].trim();
      
      // If this is a continuation of a list or a new list
      if (listType !== currentListType || !currentList) {
        // Create a new list
        currentList = document.createElement('ol');
        currentList.className = `docx-${listType}-list`;
        p.parentNode.insertBefore(currentList, p);
        currentListType = listType;
      }
      
      // Create list item
      currentItem = document.createElement('li');
      currentItem.textContent = content;
      
      // Add a data attribute to store the original number/letter
      currentItem.setAttribute('data-prefix', prefix);
      
      // Replace the paragraph with the list item
      currentList.appendChild(currentItem);
      p.parentNode.removeChild(p);
      
      // Adjust the counter for the loop since we've removed an element
      i--;
    } else {
      // Reset list tracking when encountering a non-numbered paragraph
      currentListType = null;
      currentList = null;
      currentItem = null;
      
      // Special case for paragraphs that have numbering but don't match the patterns above
      // Often these are manually formatted numbers
      if (/^\s*\d+\.\s+/.test(text) || /^\s*[a-z]\.\s+/.test(text)) {
        const parts = text.split(/^(\s*\S+\.\s+)/);
        if (parts.length >= 3) {
          const numPrefix = parts[1];
          const content = parts.slice(2).join('');
          
          // Wrap the number in a span for styling
          p.innerHTML = `<span class="docx-num">${numPrefix.trim()}</span> ${content}`;
        }
      }
    }
  }
}

/**
 * Detect and replace table of contents and index elements with a placeholder
 * @param {Document} document - DOM document
 */
function detectAndReplaceTocAndIndex(document) {
  // Common patterns for TOC elements
  const tocPatterns = [
    // Look for TOC field codes or TOC headings
    { selector: 'p.TOC, p.TOCHeading, div.TOC', type: 'TOC' },
    // Look for elements with TOC-specific classes
    { selector: '[class*="toc"]', type: 'TOC' },
    // Look for elements with index-specific classes
    { selector: '[class*="index"]', type: 'INDEX' },
    // Look for common TOC structures (lists following a TOC heading)
    { selector: 'p:contains("Table of Contents"), p:contains("Contents"), h1:contains("Contents"), h2:contains("Contents")', type: 'TOC' },
    // Look for common Index structures
    { selector: 'p:contains("Index"), h1:contains("Index"), h2:contains("Index")', type: 'INDEX' }
  ];

  // Custom implementation of :contains selector since JSDOM doesn't support it natively
  const findElementsContainingText = (selector, text) => {
    const elements = document.querySelectorAll(selector);
    return Array.from(elements).filter(el => 
      el.textContent.toLowerCase().includes(text.toLowerCase())
    );
  };

  // Process each pattern
  tocPatterns.forEach(pattern => {
    try {
      let elements = [];
      
      // Handle the custom :contains selector
      if (pattern.selector.includes(':contains(')) {
        const [baseSelector, containsText] = pattern.selector.split(':contains(');
        const text = containsText.replace(/[")]/g, '');
        elements = findElementsContainingText(baseSelector, text);
      } else {
        // Standard selector
        elements = document.querySelectorAll(pattern.selector);
      }

      // Process found elements
      elements.forEach(el => {
        // Check if this is likely a TOC or Index
        if (isTocOrIndexElement(el, pattern.type)) {
          // Replace the element with a placeholder
          replaceTocOrIndexElement(el, pattern.type);
        }
      });
    } catch (error) {
      console.error(`Error processing pattern ${pattern.selector}:`, error);
    }
  });

  // Additional heuristic detection for TOC and Index
  detectTocByStructure(document);
}

/**
 * Determine if an element is likely a TOC or Index element
 * @param {Element} element - DOM element
 * @param {string} type - 'TOC' or 'INDEX'
 * @returns {boolean} - True if element is likely a TOC or Index
 */
function isTocOrIndexElement(element, type) {
  // Check element text content
  const text = element.textContent.toLowerCase();
  
  if (type === 'TOC') {
    // Check for TOC indicators
    if (text.includes('table of contents') || text.includes('contents')) {
      return true;
    }
    
    // Check for TOC structure (list of items with page numbers)
    const nextSibling = element.nextElementSibling;
    if (nextSibling && (nextSibling.nodeName === 'UL' || nextSibling.nodeName === 'OL')) {
      return true;
    }
  } else if (type === 'INDEX') {
    // Check for Index indicators
    if (text.includes('index')) {
      return true;
    }
  }
  
  return false;
}

/**
 * Replace a TOC or Index element with a placeholder
 * @param {Element} element - DOM element to replace
 * @param {string} type - 'TOC' or 'INDEX'
 */
function replaceTocOrIndexElement(element, type) {
  // Create placeholder element
  const placeholder = element.ownerDocument.createElement('p');
  placeholder.classList.add('docx-placeholder');
  placeholder.textContent = `** ${type} HERE **`;
  
  // Replace the element with the placeholder
  element.parentNode.replaceChild(placeholder, element);
  
  // If this is a TOC heading, also remove the following list if present
  if (type === 'TOC') {
    const nextSibling = placeholder.nextElementSibling;
    if (nextSibling && (nextSibling.nodeName === 'UL' || nextSibling.nodeName === 'OL')) {
      nextSibling.parentNode.removeChild(nextSibling);
    }
  }
}

/**
 * Detect TOC by analyzing document structure
 * @param {Document} document - DOM document
 */
function detectTocByStructure(document) {
  // Look for sequences of paragraphs with tab characters and page numbers
  // which is a common pattern in TOCs
  const paragraphs = document.querySelectorAll('p');
  let consecutiveTocLikeParagraphs = 0;
  let tocStartElement = null;
  let tocEndElement = null;
  
  // First, try to find an explicit TOC heading
  let tocHeading = null;
  for (let i = 0; i < paragraphs.length; i++) {
    const text = paragraphs[i].textContent.toLowerCase().trim();
    if (text === 'table of contents' || text === 'contents' || text.includes('table of contents:')) {
      tocHeading = paragraphs[i];
      break;
    }
  }
  
  // Define patterns for TOC entries
  const pageNumberPattern = /\[\d+\]$/;  // Pattern for page numbers in brackets [123]
  const sectionPattern = /^(\d+\.|[a-z]\.|[ivx]+\.)|\d+\s/; // Pattern for section numbers
  
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const text = p.textContent;
    
    // Check if paragraph looks like a TOC entry using various patterns
    const hasSectionNumber = sectionPattern.test(text.trim());
    const hasPageNumber = pageNumberPattern.test(text.trim()) || /\d+\s*$/.test(text.trim());
    const hasTabOrMultipleSpaces = text.includes('\t') || /\s{3,}/.test(text);
    
    // If it's near a TOC heading, be more lenient with the pattern matching
    const nearTocHeading = tocHeading && Math.abs(i - Array.from(paragraphs).indexOf(tocHeading)) < 20;
    
    if ((hasSectionNumber && hasPageNumber) || 
        (hasTabOrMultipleSpaces && hasPageNumber) ||
        (nearTocHeading && (hasSectionNumber || hasPageNumber))) {
      
      consecutiveTocLikeParagraphs++;
      
      // Remember the first element in the sequence
      if (consecutiveTocLikeParagraphs === 1) {
        tocStartElement = p;
      }
      
      // Remember the last element
      tocEndElement = p;
      
      // If we've found several consecutive TOC-like paragraphs, it's likely a TOC
      if (consecutiveTocLikeParagraphs >= 3) {
        // We have enough evidence this is a TOC
        if (tocStartElement) {
          // Create a placeholder
          const placeholder = document.createElement('p');
          placeholder.classList.add('docx-placeholder');
          placeholder.textContent = '** TOC HERE **';
          
          // Insert placeholder before the TOC
          if (tocHeading) {
            // If we found a TOC heading, insert after it
            tocHeading.parentNode.insertBefore(placeholder, tocHeading.nextSibling);
            
            // Remove the TOC entries but keep the heading
            let current = placeholder.nextSibling;
            while (current && current !== tocEndElement.nextSibling) {
              const next = current.nextSibling;
              current.parentNode.removeChild(current);
              current = next;
            }
          } else {
            // No heading found, insert before the first TOC entry
            tocStartElement.parentNode.insertBefore(placeholder, tocStartElement);
            
            // Remove all the TOC paragraphs
            let current = placeholder.nextSibling;
            let removedCount = 0;
            while (current && removedCount < consecutiveTocLikeParagraphs) {
              const next = current.nextSibling;
              current.parentNode.removeChild(current);
              current = next;
              removedCount++;
            }
          }
          
          // Reset counter and break the loop
          consecutiveTocLikeParagraphs = 0;
          break;
        }
      }
    } else {
      // If we have some TOC-like paragraphs but hit a non-TOC paragraph, 
      // check if we have enough to consider it a complete TOC
      if (consecutiveTocLikeParagraphs >= 3) {
        // We have enough evidence this is a TOC
        if (tocStartElement) {
          // Create a placeholder
          const placeholder = document.createElement('p');
          placeholder.classList.add('docx-placeholder');
          placeholder.textContent = '** TOC HERE **';
          
          // Insert placeholder before the TOC
          tocStartElement.parentNode.insertBefore(placeholder, tocStartElement);
          
          // Remove all the TOC paragraphs
          for (let j = 0; j < consecutiveTocLikeParagraphs; j++) {
            if (i - j - 1 >= 0 && paragraphs[i - j - 1]) {
              paragraphs[i - j - 1].parentNode.removeChild(paragraphs[i - j - 1]);
            }
          }
          
          // Reset counter and break the loop
          consecutiveTocLikeParagraphs = 0;
          break;
        }
      }
      
      // Reset counter if we encounter a non-TOC-like paragraph
      consecutiveTocLikeParagraphs = 0;
      tocStartElement = null;
    }
  }
}

module.exports = {
  extractAndApplyStyles
};
