// style-extractor.js - Improved version for better TOC and list handling
const mammoth = require("mammoth");
const { JSDOM } = require("jsdom");
const path = require("path");
const {
  parseDocxStyles,
  generateCssFromStyleInfo,
} = require("./docx-style-parser");

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
    const cssFile =
      cssFilename || path.basename(docxPath, path.extname(docxPath)) + ".css";

    // Combine HTML and CSS
    const styledHtml = applyStylesToHtml(
      htmlResult.value,
      css,
      styleInfo,
      cssFile
    );

    return {
      html: styledHtml,
      styles: css + generateAdditionalCss(styleInfo),
      messages: htmlResult.messages,
    };
  } catch (error) {
    console.error("Error extracting styles:", error);
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
    convertImage: mammoth.images.imgElement(function (image) {
      return image.read("base64").then(function (imageBuffer) {
        const extension = image.contentType.split("/")[1];
        const filename = `image-${image.altText || Date.now()}.${extension}`;

        return {
          src: `./images/${filename}`,
          alt: image.altText || "Image",
          className: "docx-image",
        };
      });
    }),
  };

  // Use mammoth to convert with the custom style map
  const result = await mammoth.convertToHtml({
    path: docxPath,
    styleMap: styleMap,
    transformDocument: transformDocument,
    includeDefaultStyleMap: true,
    ...imageOptions,
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
    styleMap.push(
      `p[style-name='${style.name}'] => p.docx-p-${id.toLowerCase()}`
    );
  });

  // Map character styles
  Object.entries(styleInfo.styles.character || {}).forEach(([id, style]) => {
    styleMap.push(
      `r[style-name='${style.name}'] => span.docx-c-${id.toLowerCase()}`
    );
  });

  // Map table styles
  Object.entries(styleInfo.styles.table || {}).forEach(([id, style]) => {
    styleMap.push(
      `table[style-name='${style.name}'] => table.docx-t-${id.toLowerCase()}`
    );
  });

  // Improved heading mappings
  styleMap.push("p[style-name='heading 1'] => h1.docx-heading1");
  styleMap.push("p[style-name='heading 2'] => h2.docx-heading2");
  styleMap.push("p[style-name='heading 3'] => h3.docx-heading3");
  styleMap.push("p[style-name='heading 4'] => h4.docx-heading4");
  styleMap.push("p[style-name='heading 5'] => h5.docx-heading5");
  styleMap.push("p[style-name='heading 6'] => h6.docx-heading6");

  // Handle headings with styles like "Heading1"
  styleMap.push("p[style-name='Heading1'] => h1.docx-heading1");
  styleMap.push("p[style-name='Heading2'] => h2.docx-heading2");
  styleMap.push("p[style-name='Heading3'] => h3.docx-heading3");

  // Handle TOC specific styles
  styleMap.push("p[style-name='TOC Heading'] => h2.docx-toc-heading");
  styleMap.push("p[style-name='TOC1'] => p.docx-toc-entry.docx-toc-level-1");
  styleMap.push("p[style-name='TOC2'] => p.docx-toc-entry.docx-toc-level-2");
  styleMap.push("p[style-name='TOC3'] => p.docx-toc-entry.docx-toc-level-3");

  // Additional custom mappings for specific elements
  styleMap.push("p:fresh => p");
  styleMap.push("r[bold] => strong");
  styleMap.push("r[italic] => em");
  styleMap.push("r[underline] => span.docx-underline");
  styleMap.push("r[strikethrough] => span.docx-strike");
  styleMap.push("r[subscript] => sub");
  styleMap.push("r[superscript] => sup");

  // Handle specific paragraph types that might be common in documents
  styleMap.push("p[style-name='Normal Web'] => p.docx-normalweb"); 
  styleMap.push("p[style-name='Body Text'] => p.docx-bodytext");
  
  // Catch any paragraphs with "Rationale" in the style name
  styleMap.push("p[style-name*='Rationale'] => p.docx-rationale");
  
  return styleMap;
}

/**
 * Create document transformer function for style preservation
 * @param {Object} styleInfo - Style information
 * @returns {Function} - Document transformer
 */
function createDocumentTransformer(styleInfo) {
  return function (document) {
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
  try {
    // Create a DOM to manipulate the HTML
    const dom = new JSDOM(html);
    const document = dom.window.document;

    // Ensure we have a proper HTML structure
    ensureHtmlStructure(document);

    // Instead of embedding CSS, add a link to the external CSS file
    const linkElement = document.createElement("link");
    linkElement.rel = "stylesheet";
    linkElement.href = `./${cssFilename}`; // Use relative path to the CSS file

    // Add to document head
    document.head.appendChild(linkElement);

    // Add metadata
    addDocumentMetadata(document, styleInfo);

    // Add body class for RTL if needed
    if (styleInfo.settings?.rtlGutter) {
      document.body.classList.add("docx-rtl");
    }

    // Process table elements to match word styling better
    try {
      processTables(document);
    } catch (error) {
      console.error("Error processing tables:", error.message);
    }

    // Process images to maintain aspect ratio and positioning
    try {
      processImages(document);
    } catch (error) {
      console.error("Error processing images:", error.message);
    }

    // Handle language-specific elements
    try {
      processLanguageElements(document);
    } catch (error) {
      console.error("Error processing language elements:", error.message);
    }

    // Process heading structure to match TOC
    try {
      processHeadings(document, styleInfo);
    } catch (error) {
      console.error("Error processing headings:", error.message);
    }

    // IMPROVED: Process the Table of Contents with much better structure and styling
    try {
      processTOC(document, styleInfo);
    } catch (error) {
      console.error("Error processing TOC:", error.message);
    }

    // IMPROVED: Process numbered paragraphs with proper nesting for hierarchical lists
    try {
      processNestedNumberedParagraphs(document, styleInfo);
    } catch (error) {
      console.error("Error processing numbered paragraphs:", error.message);
    }

    // IMPROVED: Process special paragraph types like "Rationale for Resolution"
    try {
      processSpecialParagraphs(document, styleInfo);
    } catch (error) {
      console.error("Error processing special paragraphs:", error.message);
    }

    // Serialize back to HTML string
    return dom.serialize();
  } catch (error) {
    console.error("Error in applyStylesToHtml:", error.message);
    // Return the original HTML if there was an error
    return html;
  }
}

/**
 * IMPROVED: Enhanced heading processing with better numbering
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processHeadings(document, styleInfo) {
  // Get all headings
  const headings = document.querySelectorAll('h1, h2, h3, h4, h5, h6');
  
  // Current section counters
  const counters = {
    h1: 0,
    h2: 0,
    h3: 0,
    h4: 0,
    h5: 0,
    h6: 0
  };
  
  // Track current level to reset lower levels
  let currentLevel = 0;
  
  for (let i = 0; i < headings.length; i++) {
    const heading = headings[i];
    const level = parseInt(heading.tagName.substring(1), 10);
    
    // Skip TOC headings
    if (heading.classList.contains('docx-toc-heading')) {
      continue;
    }
    
    // Reset all counters below the current level
    if (level <= currentLevel) {
      for (let j = level + 1; j <= 6; j++) {
        counters[`h${j}`] = 0;
      }
    }
    
    // Increment counter for this level
    counters[`h${level}`]++;
    currentLevel = level;
    
    // Create numbered prefix
    let prefix = '';
    if (level === 1) {
      prefix = `${counters.h1}.`;
    } else if (level === 2) {
      // Use letter for h2 (a., b., c.)
      if (heading.textContent.trim().toLowerCase().startsWith('rationale')) {
        // Skip numbering for rationale headings
        continue;
      }
      
      const letter = String.fromCharCode(96 + counters.h2); // 'a', 'b', 'c', etc.
      prefix = `${counters.h1}.${letter}.`;
    } else if (level === 3) {
      prefix = `${counters.h1}.${counters.h2}.${counters.h3}.`;
    }
    
    // Text already starts with the prefix? Skip adding it again
    if (heading.textContent.startsWith(prefix)) {
      continue;
    }
    
    // Store the prefix for potential use in CSS
    heading.setAttribute('data-number', prefix);
    
    // Create a wrapper for the number
    const numberSpan = document.createElement('span');
    numberSpan.className = 'heading-number';
    numberSpan.textContent = prefix + ' ';
    
    // Insert the number span at the beginning of the heading
    if (heading.firstChild) {
      heading.insertBefore(numberSpan, heading.firstChild);
    } else {
      heading.appendChild(numberSpan);
    }
  }
}

/**
 * IMPROVED: Enhanced TOC processing for better structure and styling
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processTOC(document, styleInfo) {
  // Find TOC entries which might be paragraphs with TOC classes
  const tocEntries = document.querySelectorAll('.docx-toc-entry, p[class*="toc"]');
  
  if (tocEntries.length === 0) {
    // Try to find any paragraphs that might be TOC entries based on content pattern
    const paragraphs = document.querySelectorAll('p');
    let tocCandidates = [];
    
    // Look for paragraphs with page numbers at the end
    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const text = p.textContent.trim();
      
      // Check for the pattern: text followed by digits at the end (page number)
      const pageNumMatch = text.match(/^(.*?)[.\s]+(\d+)$/);
      
      if (pageNumMatch) {
        p.classList.add('docx-toc-entry');
        tocCandidates.push(p);
      }
    }
    
    // If we found potential TOC entries, process them
    if (tocCandidates.length > 0) {
      tocEntries = tocCandidates;
    }
  }
  
  // Create TOC container if we have entries
  if (tocEntries.length > 0) {
    // Check if there's already a TOC container
    let tocContainer = document.querySelector('.docx-toc');
    
    if (!tocContainer) {
      // Create a container for all TOC entries
      tocContainer = document.createElement('div');
      tocContainer.className = 'docx-toc';
      
      // Check if there's an existing TOC heading
      const tocHeading = document.querySelector('.docx-toc-heading');
      
      if (!tocHeading) {
        // Create a TOC heading if needed
        const heading = document.createElement('h2');
        heading.className = 'docx-toc-heading';
        heading.textContent = 'Table of Contents';
        tocContainer.appendChild(heading);
      } else {
        // Move existing TOC heading
        tocContainer.appendChild(tocHeading);
      }
      
      // Insert the TOC container before the first TOC entry
      const firstEntry = tocEntries[0];
      firstEntry.parentNode.insertBefore(tocContainer, firstEntry);
    }
    
    // Process each TOC entry
    tocEntries.forEach((entry, index) => {
      // Skip if already processed
      if (entry.classList.contains('processed-toc-entry')) {
        return;
      }
      
      // Mark as processed
      entry.classList.add('processed-toc-entry');
      
      // Determine TOC level from class name or calculate from indentation
      let level = 1;
      
      // Try to get level from class
      for (const cls of entry.classList) {
        if (cls.includes('toc') && cls.match(/\d+$/)) {
          level = parseInt(cls.match(/\d+$/)[0], 10);
          break;
        }
      }
      
      // Add level class if not present
      if (!entry.classList.contains(`docx-toc-level-${level}`)) {
        entry.classList.add(`docx-toc-level-${level}`);
      }
      
      // Get the text content
      const text = entry.textContent.trim();
      
      // Check for the pattern: text followed by digits at the end (page number)
      const pageNumMatch = text.match(/^(.*?)[\.\s]+(\d+)$/);
      
      if (pageNumMatch) {
        // Clear existing content
        entry.textContent = '';
        
        // Create text span
        const textSpan = document.createElement('span');
        textSpan.className = 'docx-toc-text';
        textSpan.textContent = pageNumMatch[1].trim();
        entry.appendChild(textSpan);
        
        // Create dots span
        const dotsSpan = document.createElement('span');
        dotsSpan.className = 'docx-toc-dots';
        entry.appendChild(dotsSpan);
        
        // Create page number span
        const pageSpan = document.createElement('span');
        pageSpan.className = 'docx-toc-pagenum';
        pageSpan.textContent = pageNumMatch[2];
        entry.appendChild(pageSpan);
      }
      
      // Move this entry to the TOC container if it's not already there
      if (entry.parentNode !== tocContainer) {
        tocContainer.appendChild(entry);
      }
    });
  }
}

/**
 * IMPROVED: Enhanced numbered paragraph processing for better hierarchical structure
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processNestedNumberedParagraphs(document, styleInfo) {
  // Process paragraphs that look like they should be part of a numbered list
  const paragraphs = document.querySelectorAll('p');
  
  // Regex patterns for different types of numbering
  const mainNumberPattern = /^\s*(\d+)\.\s+(.+)$/;
  const alphaPattern = /^\s*([a-z])\.\s+(.+)$/;
  const subNumberPattern = /^\s*(\d+\.\d+)\.\s+(.+)$/;
  
  // Current list structures
  let currentMainList = null;
  let currentMainItem = null;
  let currentSubList = null;
  let lastMainNumber = 0;
  
  // Process paragraphs
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const text = p.textContent.trim();
    
    // Skip if this paragraph is already part of a list
    if (p.closest('ol, ul')) continue;
    
    // Skip any special paragraphs (handled separately)
    if (p.classList.contains('docx-rationale') || 
        p.classList.contains('docx-toc-entry') ||
        text.startsWith('Rationale for Resolution')) continue;

    // Check for main numbered items (1., 2., etc.)
    let mainMatch = text.match(mainNumberPattern);
    if (mainMatch) {
      const number = parseInt(mainMatch[1], 10);
      const content = mainMatch[2];
      
      // If this is a new list or a restart
      if (!currentMainList || number === 1 || (number > 1 && number <= lastMainNumber + 1)) {
        if (!currentMainList || number === 1) {
          currentMainList = document.createElement('ol');
          currentMainList.className = 'docx-numbered-list';
          p.parentNode.insertBefore(currentMainList, p);
        }
        
        // Create the list item
        currentMainItem = document.createElement('li');
        currentMainItem.textContent = content;
        currentMainItem.setAttribute('data-prefix', number);
        
        // Add the item to the list
        currentMainList.appendChild(currentMainItem);
        
        // Remember this number
        lastMainNumber = number;
        
        // Reset sublist tracking
        currentSubList = null;
        
        // Remove the original paragraph
        p.parentNode.removeChild(p);
        i--; // Adjust counter
        
        continue;
      }
    }
    
    // Check for alpha subitems (a., b., etc.)
    let alphaMatch = text.match(alphaPattern);
    if (alphaMatch && currentMainItem) {
      const letter = alphaMatch[1];
      const content = alphaMatch[2];
      
      // If we don't have a sublist for this main item yet
      if (!currentSubList) {
        currentSubList = document.createElement('ol');
        currentSubList.className = 'docx-alpha-list';
        currentMainItem.appendChild(currentSubList);
      }
      
      // Create the sub-item
      const subItem = document.createElement('li');
      subItem.textContent = content;
      subItem.setAttribute('data-prefix', letter);
      
      // Add the sub-item to the sublist
      currentSubList.appendChild(subItem);
      
      // Remove the original paragraph
      p.parentNode.removeChild(p);
      i--; // Adjust counter
      
      continue;
    }
    
    // Check if this is a special paragraph that belongs in the current list
    if (text.startsWith('Rationale for') && currentMainItem) {
      const specialPara = document.createElement('p');
      specialPara.className = 'docx-rationale';
      specialPara.textContent = text;
      
      // Add it after the current main item
      if (currentMainItem.nextSibling) {
        currentMainList.insertBefore(specialPara, currentMainItem.nextSibling);
      } else {
        currentMainList.appendChild(specialPara);
      }
      
      // Remove the original paragraph
      p.parentNode.removeChild(p);
      i--; // Adjust counter
      
      continue;
    }
  }
  
  // Now look for any paragraphs with "Resolved" pattern and convert to list items
  const resolvedPattern = /^\s*Resolved\s+\(([^)]+)\)(.+)$/;
  
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    
    // Skip if already processed
    if (p.closest('ol, ul')) continue;
    
    const text = p.textContent.trim();
    const resolvedMatch = text.match(resolvedPattern);
    
    if (resolvedMatch) {
      const id = resolvedMatch[1].trim();
      const content = resolvedMatch[2].trim();
      
      // Check if we need to start a new list
      let resolvedList = p.previousElementSibling;
      if (!resolvedList || !resolvedList.classList.contains('docx-resolved-list')) {
        resolvedList = document.createElement('ol');
        resolvedList.className = 'docx-numbered-list docx-resolved-list';
        p.parentNode.insertBefore(resolvedList, p);
      }
      
      // Create the list item
      const listItem = document.createElement('li');
      listItem.innerHTML = `<strong>Resolved (${id})</strong>${content}`;
      listItem.setAttribute('data-prefix', id);
      resolvedList.appendChild(listItem);
      
      // Remove the original paragraph
      p.parentNode.removeChild(p);
      i--; // Adjust counter
    }
  }
}

/**
 * IMPROVED: Process special paragraph types based on document analysis
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processSpecialParagraphs(document, styleInfo) {
  // Check if document has rationale paragraphs
  const hasRationale = styleInfo.documentStructure?.hasRationale;
  const specialPatterns = styleInfo.documentStructure?.specialParagraphPatterns || [];
  
  // Process paragraphs for special formatting
  const paragraphs = document.querySelectorAll('p');
  
  paragraphs.forEach(para => {
    const text = para.textContent.trim();
    
    // Check for rationale paragraphs
    if (text.startsWith('Rationale for Resolution') || 
        para.classList.contains('docx-rationale')) {
      
      // Apply standard rationale formatting
      para.classList.add('docx-rationale');
      
      // If it doesn't have italic formatting already, add it
      if (!para.querySelector('em') && !para.style.fontStyle) {
        // Wrap contents in italic
        const content = para.innerHTML;
        para.innerHTML = `<em>${content}</em>`;
      }
    }
    
    // Check for "Whereas" paragraphs
    if (text.startsWith('Whereas,')) {
      para.classList.add('docx-whereas');
    }
    
    // Check for "Resolved" paragraphs
    if (text.match(/^Resolved\s+\([\d\.]+\)/)) {
      para.classList.add('docx-resolved');
      
      // Make the ID part bold if it's not already
      const parts = text.match(/^(Resolved\s+\([^)]+\))(.+)$/);
      if (parts) {
        para.innerHTML = `<strong>${parts[1]}</strong>${parts[2]}`;
      }
    }
    
    // Apply special formatting based on detected patterns
    specialPatterns.forEach(pattern => {
      // Convert string pattern back to RegExp
      let patternRegex;
      try {
        // Remove "/" from start and end and any flags
        const patternStr = pattern.pattern.replace(/^\/|\/[gim]*$/g, '');
        patternRegex = new RegExp(patternStr);
      } catch (error) {
        console.error(`Invalid pattern: ${pattern.pattern}`);
        return;
      }
      
      if (patternRegex.test(text)) {
        para.classList.add(`docx-${pattern.type}`);
      }
    });
  });
}

/**
 * Ensure HTML has proper structure (html, head, body)
 * @param {Document} document - DOM document
 */
function ensureHtmlStructure(document) {
  // Create html element if not exists
  if (
    !document.documentElement ||
    document.documentElement.nodeName.toLowerCase() !== "html"
  ) {
    const html = document.createElement("html");

    // Move existing content
    while (document.childNodes.length > 0) {
      html.appendChild(document.childNodes[0]);
    }

    document.appendChild(html);
  }

  // Create head if not exists
  if (!document.head) {
    const head = document.createElement("head");
    document.documentElement.insertBefore(
      head,
      document.documentElement.firstChild
    );
  }

  // Add meta charset
  if (!document.querySelector("meta[charset]")) {
    const meta = document.createElement("meta");
    meta.setAttribute("charset", "utf-8");
    document.head.appendChild(meta);
  }

  // Create body if not exists
  if (!document.body) {
    const body = document.createElement("body");

    // Move content to body
    Array.from(document.documentElement.childNodes).forEach((node) => {
      if (node !== document.head && node.nodeName.toLowerCase() !== "body") {
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
  const title = document.createElement("title");
  title.textContent = "Document";
  document.head.appendChild(title);

  // Add viewport meta
  const viewport = document.createElement("meta");
  viewport.setAttribute("name", "viewport");
  viewport.setAttribute("content", "width=device-width, initial-scale=1");
  document.head.appendChild(viewport);

  // Add class to body 
  document.body.classList.add('doc2web-document');
}

/**
 * Process tables for better styling
 * @param {Document} document - DOM document
 */
function processTables(document) {
  const tables = document.querySelectorAll("table");
  tables.forEach((table) => {
    // Add default class if no class is present
    if (!table.classList.length) {
      table.classList.add("docx-table-default");
    }

    // Make sure tables have tbody
    if (!table.querySelector("tbody")) {
      const tbody = document.createElement("tbody");

      // Move rows to tbody
      Array.from(table.querySelectorAll("tr")).forEach((row) => {
        tbody.appendChild(row);
      });

      table.appendChild(tbody);
    }
    
    // Add responsive table wrapper
    const wrapper = document.createElement('div');
    wrapper.className = 'table-responsive';
    table.parentNode.insertBefore(wrapper, table);
    wrapper.appendChild(table);
  });
}

/**
 * Process images for better styling
 * @param {Document} document - DOM document
 */
function processImages(document) {
  const images = document.querySelectorAll("img");
  images.forEach((img) => {
    // Add default class if no class is present
    if (!img.classList.contains("docx-image")) {
      img.classList.add("docx-image");
    }

    // Ensure max-width for responsive images
    img.style.maxWidth = "100%";

    // Make sure images have alt text
    if (!img.hasAttribute("alt")) {
      img.setAttribute("alt", "Document image");
    }
    
    // Add figure wrapper for better semantics
    const figure = document.createElement('figure');
    img.parentNode.insertBefore(figure, img);
    figure.appendChild(img);
    
    // If there's text immediately after, it might be a caption
    if (img.nextSibling && img.nextSibling.nodeType === 3) {
      const caption = document.createElement('figcaption');
      caption.textContent = img.nextSibling.textContent.trim();
      figure.appendChild(caption);
      img.nextSibling.remove();
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
  rtlElements.forEach((el) => {
    el.classList.add("docx-rtl");
  });
  
  // Look for non-Latin scripts and add appropriate language attributes
  const allTextNodes = [];
  
  function collectTextNodes(node) {
    if (node.nodeType === 3) { // TEXT_NODE
      allTextNodes.push(node);
    } else {
      for (let i = 0; i < node.childNodes.length; i++) {
        collectTextNodes(node.childNodes[i]);
      }
    }
  }
  
  collectTextNodes(document.body);
  
  // Check for scripts and add language attributes
  const cyrillicRegex = /[\u0400-\u04FF]/;
  const chineseRegex = /[\u4E00-\u9FFF]/;
  const arabicRegex = /[\u0600-\u06FF]/;
  
  allTextNodes.forEach(textNode => {
    const text = textNode.nodeValue;
    const parent = textNode.parentNode;
    
    if (cyrillicRegex.test(text) && !parent.hasAttribute('lang')) {
      parent.setAttribute('lang', 'ru');
    } else if (chineseRegex.test(text) && !parent.hasAttribute('lang')) {
      parent.setAttribute('lang', 'zh');
    } else if (arabicRegex.test(text) && !parent.hasAttribute('lang')) {
      parent.setAttribute('lang', 'ar');
      if (!parent.hasAttribute('dir')) {
        parent.setAttribute('dir', 'rtl');
      }
    }
  });
}

/**
 * IMPROVED: Generate additional CSS for TOC and list styling
 * @param {Object} styleInfo - Style information
 * @returns {string} - Additional CSS
 */
function generateAdditionalCss(styleInfo) {
  // Get TOC styles
  const tocStyles = styleInfo.tocStyles || {};
  const leaderStyle = tocStyles.leaderStyle || { character: '.', position: 6 * 72 };
  
  // Get numbering styles
  const numberingDefs = styleInfo.numberingDefs || {};
  
  // Base CSS
  let css = `
/* Document defaults */
body {
  font-family: "Cambria", sans-serif;
  font-size: 12pt;
  line-height: 1.5;
  margin: 20px;
  padding: 0;
  color: #333;
  background-color: #fff;
}

/* For better mobile readability */
@media (max-width: 768px) {
  body {
    margin: 10px;
    font-size: 14pt;
  }
}
`;

  // Improved TOC styles
  css += `
/* Enhanced TOC styling */
.docx-toc {
  margin: 1.5em 0;
  width: 100%;
}

.docx-toc-heading {
  font-size: 14pt;
  font-weight: bold;
  margin-top: 1em;
  margin-bottom: 0.8em;
  text-align: center;
}

.docx-toc-entry {
  display: flex !important;
  flex-wrap: nowrap;
  justify-content: space-between;
  align-items: baseline;
  width: 100%;
  margin-bottom: 0.4em;
  line-height: 1.5;
  position: relative;
}

.docx-toc-text {
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  margin-right: 0.5em;
  flex-grow: 0;
  flex-shrink: 1;
}

.docx-toc-dots {
  flex-grow: 1;
  min-width: 2em;
  margin: 0 0.5em;
  border-bottom: 1px dotted #666;
  position: relative;
  height: 0.7em;
}

.docx-toc-pagenum {
  flex-shrink: 0;
  text-align: right;
  min-width: 1.5em;
  font-weight: normal;
}

/* TOC level-specific indentation */
.docx-toc-level-1 {
  font-weight: bold;
}

.docx-toc-level-2 {
  margin-left: 1.5em;
}

.docx-toc-level-3 {
  margin-left: 3em;
  font-style: italic;
}
`;

  // Enhanced list styles
  css += `
/* Improved list styling */
ol.docx-numbered-list,
ol.docx-alpha-list,
ol.docx-resolved-list {
  counter-reset: item;
  list-style-type: none;
  padding-left: 0;
  margin: 1em 0;
}

ol.docx-numbered-list > li,
ol.docx-alpha-list > li,
ol.docx-resolved-list > li {
  position: relative;
  padding-left: 2.5em;
  margin-bottom: 0.5em;
  min-height: 1.5em;
}

ol.docx-numbered-list > li::before,
ol.docx-alpha-list > li::before,
ol.docx-resolved-list > li::before {
  position: absolute;
  left: 0;
  content: attr(data-prefix) ".";
  font-weight: bold;
  min-width: 2em;
  display: inline-block;
}

/* Nested list styles */
ol.docx-alpha-list {
  margin: 0.5em 0 0.5em 1em;
}

ol.docx-alpha-list > li {
  padding-left: 2em;
}

/* Special paragraph styles */
.docx-rationale {
  font-style: italic;
  margin: 0.5em 0 1em 2.5em;
  color: #555;
  border-left: 3px solid #ddd;
  padding-left: 1em;
  line-height: 1.4;
}

.docx-whereas {
  margin-bottom: 0.5em;
}

.docx-resolved {
  font-weight: bold;
  margin-bottom: 0.7em;
}
`;

  // Table improvements
  css += `
/* Table improvements */
.table-responsive {
  overflow-x: auto;
  margin: 1.25em 0;
  border-radius: 4px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
}

table {
  width: 100%;
  border-collapse: collapse;
  margin: 0;
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

tr:nth-child(even) {
  background-color: #f9f9f9;
}

/* Better link styling */
a {
  color: #0066cc;
  text-decoration: none;
}

a:hover {
  text-decoration: underline;
}

/* Print styling */
@media print {
  body {
    font-size: 10pt;
    margin: 0;
  }
  
  .docx-toc {
    border: none;
    background: none;
    page-break-after: always;
  }
  
  h1, h2, h3, h4, h5, h6 {
    page-break-after: avoid;
    page-break-inside: avoid;
  }
  
  img, table, figure {
    page-break-inside: avoid;
  }
  
  .table-responsive {
    overflow-x: visible;
    box-shadow: none;
  }
}

/* Utility styles */
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }
.docx-tab { display: inline-block; width: 36pt; }
.docx-placeholder { 
  font-weight: bold; 
  text-align: center; 
  padding: 15px; 
  margin: 20px 0; 
  background-color: #f0f0f0; 
  border: 1px dashed #999; 
  border-radius: 5px; 
}
`;

  return css;
}

module.exports = {
  extractAndApplyStyles,
  generateAdditionalCss,
};
