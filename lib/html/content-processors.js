// lib/html/content-processors.js - Enhanced content element processors using DOCX introspection
const { selectNodes, selectSingleNode } = require('../xml/xpath-utils');
const { extractParagraphNumberingContext, resolveNumberingForParagraphs } = require('../parsers/numbering-resolver');

/**
 * Enhanced heading processing using DOCX introspection
 * Extracts numbering from DOCX XML structure and applies to headings
 * 
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 * @param {Array} numberingContext - Resolved paragraph numbering context
 */
function processHeadings(document, styleInfo, numberingContext = null) {
  const headings = document.querySelectorAll('h1, h2, h3, h4, h5, h6');
  
  // If no numbering context provided, skip numbering processing
  if (!numberingContext) {
    console.warn('No numbering context provided for heading processing');
    return;
  }
  
  // Process each heading
  headings.forEach(heading => {
    // Skip TOC headings
    if (heading.classList.contains('docx-toc-heading')) {
      return;
    }
    
    const headingText = heading.textContent.trim();
    
    // Find corresponding paragraph context by matching text content
    const context = findParagraphContextByText(headingText, numberingContext);
    
    if (context && context.resolvedNumbering) {
      // Apply DOCX-derived numbering to heading
      applyNumberingToHeading(heading, context.resolvedNumbering, context);
    } else {
      // Try to find context by partial text match for truncated headings
      const partialContext = findParagraphContextByPartialText(headingText, numberingContext);
      if (partialContext && partialContext.resolvedNumbering) {
        applyNumberingToHeading(heading, partialContext.resolvedNumbering, partialContext);
      }
    }
    
    // Ensure headings have proper ARIA and accessibility attributes
    ensureHeadingAccessibility(heading);
  });
}

/**
 * Find paragraph context by exact text match
 * @param {string} text - Text to match
 * @param {Array} numberingContext - Numbering context array
 * @returns {Object|null} - Found context or null
 */
function findParagraphContextByText(text, numberingContext) {
  return numberingContext.find(context => {
    if (!context.textContent) return false;
    
    // Try exact match first
    if (context.textContent.trim() === text) {
      return true;
    }
    
    // Try match without trailing punctuation
    const cleanContextText = context.textContent.trim().replace(/[:\.\s]*$/, '');
    const cleanSearchText = text.replace(/[:\.\s]*$/, '');
    
    return cleanContextText === cleanSearchText;
  }) || null;
}

/**
 * Find paragraph context by partial text match
 * @param {string} text - Text to match
 * @param {Array} numberingContext - Numbering context array
 * @returns {Object|null} - Found context or null
 */
function findParagraphContextByPartialText(text, numberingContext) {
  const cleanSearchText = text.toLowerCase().trim().replace(/[:\.\s]*$/, '');
  
  return numberingContext.find(context => {
    if (!context.textContent) return false;
    
    const cleanContextText = context.textContent.toLowerCase().trim().replace(/[:\.\s]*$/, '');
    
    // Check if either text contains the other (for truncated headings)
    return cleanContextText.includes(cleanSearchText) || 
           cleanSearchText.includes(cleanContextText);
  }) || null;
}

/**
 * Apply DOCX-derived numbering to heading
 * @param {HTMLElement} heading - The heading element
 * @param {Object} numbering - Resolved numbering information
 * @param {Object} context - Paragraph context
 */
function applyNumberingToHeading(heading, numbering, context) {
  // Store existing anchors and other child elements
  const childElements = Array.from(heading.childNodes);
  
  // Clear the heading content
  heading.textContent = '';
  
  // Create the number span with DOCX-derived numbering
  const numberSpan = document.createElement('span');
  numberSpan.className = 'heading-number';
  
  // Use the full numbering format from DOCX
  let numberText = numbering.fullNumbering;
  
  // Ensure proper spacing after number
  if (!numberText.endsWith(' ') && !numberText.endsWith('\t')) {
    numberText += ' ';
  }
  
  numberSpan.textContent = numberText;
  heading.appendChild(numberSpan);
  
  // Restore child elements (like links)
  childElements.forEach(child => {
    if (child.nodeType === Node.TEXT_NODE) {
      // For text nodes, add the text content without the numbering part
      let textContent = child.textContent || '';
      
      // Remove any existing numbering pattern from the beginning
      textContent = textContent.replace(/^\s*\d+\.?\s*/, '');
      textContent = textContent.replace(/^\s*[a-zA-Z]\.?\s*/, '');
      textContent = textContent.replace(/^\s*[ivxlcdmIVXLCDM]+\.?\s*/, '');
      
      if (textContent.trim()) {
        heading.appendChild(document.createTextNode(textContent));
      }
    } else {
      // For element nodes, append as-is
      heading.appendChild(child);
    }
  });
  
  // Store metadata for CSS and other processing
  heading.setAttribute('data-number', numbering.rawNumber.toString());
  heading.setAttribute('data-numbering-id', context.numberingId || '');
  heading.setAttribute('data-numbering-level', context.numberingLevel?.toString() || '');
  heading.setAttribute('data-format', numbering.format || '');
}

/**
 * Ensure heading has proper accessibility attributes
 * @param {HTMLElement} heading - The heading element
 */
function ensureHeadingAccessibility(heading) {
  // Generate a clean ID from the heading text if not present
  if (!heading.id) {
    const cleanText = heading.textContent
      .replace(/[^\w\s]/g, '')
      .replace(/\s+/g, '-')
      .toLowerCase()
      .substring(0, 50);
    heading.id = 'heading-' + cleanText;
  }
  
  // Add keyboard focusable class for accessibility
  heading.classList.add('keyboard-focusable');
  
  // Add tabindex for keyboard navigation
  if (!heading.hasAttribute('tabindex')) {
    heading.setAttribute('tabindex', '0');
  }
}

/**
 * Enhanced TOC processing for better structure and styling
 * Creates a properly structured Table of Contents with leader lines and page numbers
 * 
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processTOC(document, styleInfo) {
  // Find TOC entries which might be paragraphs with TOC classes
  let tocEntries = document.querySelectorAll('.docx-toc-entry, p[class*="toc"]');
  
  if (tocEntries.length === 0) {
    // Try to find any paragraphs that might be TOC entries based on content pattern
    const paragraphs = document.querySelectorAll('p');
    let tocCandidates = [];
    
    // Look for paragraphs with page numbers at the end
    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const text = p.textContent.trim();
      
      // Check for the pattern: text followed by dots/spaces, then digits at the end
      const pageNumMatch = text.match(/^(.*?)[\.\s]+(\d+)$/);
      
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
    createTOCContainer(document, tocEntries, styleInfo);
  }
}

/**
 * Create and populate TOC container
 * @param {Document} document - DOM document
 * @param {NodeList} tocEntries - TOC entry elements
 * @param {Object} styleInfo - Style information
 */
function createTOCContainer(document, tocEntries, styleInfo) {
  // Check if there's already a TOC container
  let tocContainer = document.querySelector('.docx-toc');
  
  if (!tocContainer) {
    // Create a container for all TOC entries
    tocContainer = document.createElement('nav');
    tocContainer.className = 'docx-toc';
    tocContainer.setAttribute('role', 'navigation');
    tocContainer.setAttribute('aria-label', 'Table of Contents');
    
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
      tocContainer.appendChild(tocHeading.cloneNode(true));
      tocHeading.parentNode.removeChild(tocHeading);
    }
    
    // Insert the TOC container before the first TOC entry
    const firstEntry = tocEntries[0];
    firstEntry.parentNode.insertBefore(tocContainer, firstEntry);
  }
  
  // Process each TOC entry
  Array.from(tocEntries).forEach((entry, index) => {
    processTOCEntry(entry, index, tocContainer, styleInfo);
  });
}

/**
 * Process individual TOC entry
 * @param {HTMLElement} entry - TOC entry element
 * @param {number} index - Entry index
 * @param {HTMLElement} tocContainer - TOC container
 * @param {Object} styleInfo - Style information
 */
function processTOCEntry(entry, index, tocContainer, styleInfo) {
  // Skip if already processed
  if (entry.classList.contains('processed-toc-entry')) {
    return;
  }
  
  // Mark as processed
  entry.classList.add('processed-toc-entry');
  
  // Determine TOC level from class name or calculate from indentation
  let level = determineTOCLevel(entry, styleInfo);
  
  // Add level class if not present
  const levelClass = `docx-toc-level-${level}`;
  if (!entry.classList.contains(levelClass)) {
    entry.classList.add(levelClass);
  }
  
  // Get the text content
  const text = entry.textContent.trim();
  
  // Structure the TOC entry with proper spans
  structureTOCEntry(entry, text, level);
  
  // Move this entry to the TOC container if it's not already there
  if (entry.parentNode !== tocContainer) {
    const clonedEntry = entry.cloneNode(true);
    tocContainer.appendChild(clonedEntry);
    entry.parentNode.removeChild(entry);
  }
  
  // Add accessibility attributes
  entry.setAttribute('tabindex', '0');
  entry.setAttribute('role', 'link');
}

/**
 * Determine TOC level from style information or class
 * @param {HTMLElement} entry - TOC entry element
 * @param {Object} styleInfo - Style information
 * @returns {number} - TOC level
 */
function determineTOCLevel(entry, styleInfo) {
  let level = 1;
  
  // Try to get level from class
  const classArray = Array.from(entry.classList);
  for (const cls of classArray) {
    if (cls.includes('toc') && cls.match(/\d+$/)) {
      level = parseInt(cls.match(/\d+$/)[0], 10);
      break;
    }
  }
  
  // If no level found from class, try to determine from style or indentation
  if (level === 1 && styleInfo.tocStyles) {
    // Use TOC style information to determine level
    const computedStyle = window.getComputedStyle ? window.getComputedStyle(entry) : null;
    if (computedStyle) {
      const marginLeft = parseInt(computedStyle.marginLeft, 10) || 0;
      // Estimate level based on indentation (every 20pt = one level)
      level = Math.floor(marginLeft / 20) + 1;
    }
  }
  
  return Math.max(1, Math.min(6, level)); // Clamp between 1 and 6
}

/**
 * Structure TOC entry with proper spans for text, dots, and page number
 * @param {HTMLElement} entry - TOC entry element
 * @param {string} text - Entry text content
 * @param {number} level - TOC level
 */
function structureTOCEntry(entry, text, level) {
  // Check for the pattern: text followed by dots/spaces, then digits at the end
  const pageNumMatch = text.match(/^(.*?)[\.\s]*(\d+)$/);
  
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
    dotsSpan.setAttribute('aria-hidden', 'true'); // Hide from screen readers
    entry.appendChild(dotsSpan);
    
    // Create page number span
    const pageSpan = document.createElement('span');
    pageSpan.className = 'docx-toc-pagenum';
    pageSpan.textContent = pageNumMatch[2];
    pageSpan.setAttribute('aria-label', `Page ${pageNumMatch[2]}`);
    entry.appendChild(pageSpan);
  } else {
    // No page number found, just structure the text
    entry.textContent = '';
    
    const textSpan = document.createElement('span');
    textSpan.className = 'docx-toc-text';
    textSpan.textContent = text;
    entry.appendChild(textSpan);
  }
}

/**
 * Enhanced numbered paragraph processing using DOCX introspection
 * Converts paragraphs with numbering into properly structured lists based on DOCX structure
 * 
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 * @param {Array} numberingContext - Resolved paragraph numbering context
 */
function processNestedNumberedParagraphs(document, styleInfo, numberingContext = null) {
  if (!numberingContext) {
    console.warn('No numbering context provided for numbered paragraph processing');
    return;
  }
  
  // Group numbering contexts by numbering ID to create separate lists
  const numberingGroups = groupNumberingContexts(numberingContext);
  
  // Process each numbering group
  Object.values(numberingGroups).forEach(group => {
    processNumberingGroup(document, group, styleInfo);
  });
}

/**
 * Group numbering contexts by numbering ID
 * @param {Array} numberingContext - Numbering context array
 * @returns {Object} - Grouped contexts by numbering ID
 */
function groupNumberingContexts(numberingContext) {
  const groups = {};
  
  numberingContext.forEach(context => {
    if (context.numberingId && context.resolvedNumbering) {
      if (!groups[context.numberingId]) {
        groups[context.numberingId] = [];
      }
      groups[context.numberingId].push(context);
    }
  });
  
  return groups;
}

/**
 * Process a single numbering group
 * @param {Document} document - DOM document
 * @param {Array} group - Group of contexts with same numbering ID
 * @param {Object} styleInfo - Style information
 */
function processNumberingGroup(document, group, styleInfo) {
  // Sort by paragraph index to maintain document order
  group.sort((a, b) => a.paragraphIndex - b.paragraphIndex);
  
  let currentList = null;
  let listItems = {}; // Track list items by level for nesting
  
  group.forEach(context => {
    // Find the corresponding paragraph in the HTML document
    const paragraph = findParagraphByContext(document, context);
    if (!paragraph) return;
    
    // Skip if already processed
    if (paragraph.classList.contains('processed-list-item')) {
      return;
    }
    
    const level = context.numberingLevel;
    
    // Create or get list container
    if (!currentList || level === 0) {
      currentList = createListContainer(document, paragraph, context);
    }
    
    // Create list item
    const listItem = createListItem(document, paragraph, context);
    
    // Handle nesting
    if (level === 0) {
      // Top level item
      currentList.appendChild(listItem);
      listItems[level] = listItem;
      
      // Clear deeper level references
      Object.keys(listItems).forEach(lvl => {
        if (parseInt(lvl, 10) > level) {
          delete listItems[lvl];
        }
      });
    } else {
      // Nested item
      const parentItem = listItems[level - 1];
      if (parentItem) {
        // Find or create nested list
        let nestedList = parentItem.querySelector('ol, ul');
        if (!nestedList) {
          nestedList = createNestedList(context);
          parentItem.appendChild(nestedList);
        }
        
        nestedList.appendChild(listItem);
        listItems[level] = listItem;
      } else {
        // No parent found, add to main list
        currentList.appendChild(listItem);
        listItems[level] = listItem;
      }
    }
    
    // Mark as processed and remove original paragraph
    paragraph.classList.add('processed-list-item');
    paragraph.parentNode.removeChild(paragraph);
  });
}

/**
 * Find paragraph element by context
 * @param {Document} document - DOM document
 * @param {Object} context - Paragraph context
 * @returns {HTMLElement|null} - Found paragraph element
 */
function findParagraphByContext(document, context) {
  const paragraphs = document.querySelectorAll('p');
  
  // Try to find by text content match
  for (const p of paragraphs) {
    const pText = p.textContent.trim();
    if (pText === context.textContent) {
      return p;
    }
  }
  
  // Try partial match for truncated content
  for (const p of paragraphs) {
    const pText = p.textContent.trim();
    if (pText.includes(context.textContent) || context.textContent.includes(pText)) {
      return p;
    }
  }
  
  return null;
}

/**
 * Create list container
 * @param {Document} document - DOM document
 * @param {HTMLElement} paragraph - Reference paragraph
 * @param {Object} context - Paragraph context
 * @returns {HTMLElement} - Created list element
 */
function createListContainer(document, paragraph, context) {
  const list = document.createElement('ol');
  list.className = 'docx-numbered-list';
  list.setAttribute('data-numbering-id', context.numberingId);
  
  // Insert before the paragraph
  paragraph.parentNode.insertBefore(list, paragraph);
  
  return list;
}

/**
 * Create nested list
 * @param {Object} context - Paragraph context
 * @returns {HTMLElement} - Created nested list element
 */
function createNestedList(context) {
  const list = document.createElement('ol');
  list.className = getListClassName(context);
  list.setAttribute('data-numbering-id', context.numberingId);
  list.setAttribute('data-numbering-level', context.numberingLevel.toString());
  
  return list;
}

/**
 * Get appropriate list class name based on context
 * @param {Object} context - Paragraph context
 * @returns {string} - CSS class name
 */
function getListClassName(context) {
  if (!context.resolvedNumbering) {
    return 'docx-numbered-list';
  }
  
  const format = context.resolvedNumbering.format;
  
  switch (format) {
    case 'lowerLetter':
    case 'upperLetter':
      return 'docx-alpha-list';
    case 'lowerRoman':
    case 'upperRoman':
      return 'docx-roman-list';
    case 'bullet':
      return 'docx-bulleted-list';
    default:
      return 'docx-numbered-list';
  }
}

/**
 * Create list item
 * @param {Document} document - DOM document
 * @param {HTMLElement} paragraph - Original paragraph
 * @param {Object} context - Paragraph context
 * @returns {HTMLElement} - Created list item
 */
function createListItem(document, paragraph, context) {
  const listItem = document.createElement('li');
  
  // Copy content from paragraph, removing any existing numbering
  let content = paragraph.textContent.trim();
  
  // Remove numbering prefix if present (this should have been handled by mammoth)
  content = content.replace(/^\s*\d+\.?\s*/, '');
  content = content.replace(/^\s*[a-zA-Z]\.?\s*/, '');
  content = content.replace(/^\s*[ivxlcdmIVXLCDM]+\.?\s*/, '');
  
  listItem.textContent = content;
  
  // Copy classes and attributes from original paragraph
  paragraph.classList.forEach(cls => {
    if (!cls.startsWith('docx-p-')) {
      listItem.classList.add(cls);
    }
  });
  
  // Add data attributes for styling
  listItem.setAttribute('data-numbering-id', context.numberingId);
  listItem.setAttribute('data-numbering-level', context.numberingLevel.toString());
  listItem.setAttribute('data-number', context.resolvedNumbering.rawNumber.toString());
  listItem.setAttribute('data-format', context.resolvedNumbering.format);
  
  return listItem;
}

/**
 * Process special paragraph types based on document structure analysis
 * This function remains unchanged as it's already structure-based
 */
function processSpecialParagraphs(document, styleInfo) {
  // Get special paragraph patterns from document structure analysis
  const specialPatterns = styleInfo.documentStructure?.specialParagraphPatterns || [];
  
  // Process paragraphs for special formatting based on structure
  const paragraphs = document.querySelectorAll('p');
  
  paragraphs.forEach(para => {
    // Skip paragraphs that are already processed as list items or TOC entries
    if (para.classList.contains('processed-list-item') || 
        para.classList.contains('processed-special-para') ||
        para.classList.contains('docx-toc-entry')) {
      return;
    }
    
    const text = para.textContent.trim();
    
    // Apply structural pattern classes based on document analysis
    specialPatterns.forEach(pattern => {
      if (pattern.type && isStructuralMatch(text, pattern)) {
        para.classList.add(`docx-${pattern.type}`);
        para.classList.add('docx-structural-pattern');
      }
    });
    
    // Mark as processed
    para.classList.add('processed-special-para');
  });
}

/**
 * Check if text matches structural pattern
 * @param {string} text - Text to check
 * @param {Object} pattern - Pattern definition
 * @returns {boolean} - True if matches
 */
function isStructuralMatch(text, pattern) {
  if (!pattern.examples || !pattern.examples.length) {
    return false;
  }
  
  // Check if text matches the structural pattern based on examples
  return pattern.examples.some(example => {
    // Extract structural pattern from example
    const examplePattern = extractStructuralPattern(example);
    const textPattern = extractStructuralPattern(text);
    
    return examplePattern === textPattern;
  });
}

/**
 * Extract structural pattern from text
 * @param {string} text - Text to analyze
 * @returns {string} - Structural pattern
 */
function extractStructuralPattern(text) {
  // Extract pattern based on punctuation and structure, not content
  if (text.match(/^\w+:\s/)) return 'word_colon';
  if (text.match(/^\w+,\s/)) return 'word_comma';
  if (text.match(/^\w+\s+\([^)]+\)/)) return 'word_parenthesis';
  
  return 'unknown';
}

/**
 * Legacy function maintained for backward compatibility
 */
function identifyListPatterns(paragraphs) {
  // This function is no longer needed with DOCX introspection
  // but maintained for compatibility
  return {
    numberedItems: [],
    alphaItems: [],
    romanItems: [],
    specialParagraphs: []
  };
}

/**
 * Legacy function maintained for backward compatibility
 */
function isSpecialParagraph(text, specialPatterns) {
  // Use the new structural matching approach
  const pattern = extractStructuralPattern(text);
  if (pattern !== 'unknown') {
    return { type: pattern, prefix: text.match(/^(\w+)/)?.[1] || '' };
  }
  return null;
}

module.exports = {
  processHeadings,
  processTOC,
  processNestedNumberedParagraphs,
  processSpecialParagraphs,
  identifyListPatterns,
  isSpecialParagraph,
  findParagraphContextByText,
  findParagraphContextByPartialText
};