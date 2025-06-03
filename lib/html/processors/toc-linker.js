// File: lib/html/processors/toc-linker.js
// TOC linking functionality - adds navigation links from TOC entries to document sections

/**
 * Main entry point for TOC linking functionality
 * Processes all TOC entries and adds clickable links to corresponding sections
 * 
 * @param {Document} document - The HTML document
 * @param {Object} styleInfo - Style information including numbering context
 */
function linkTOCEntries(document, styleInfo) {
  if (!document) {
    console.warn('No document provided for TOC linking');
    return;
  }

  console.log('Starting TOC linking process...');
  
  try {
    // Find all TOC entries
    const tocEntries = findTOCEntries(document);

    let linkedCount = 0;
    let skippedCount = 0;
    
    // Track hierarchical context for section relationships
    const hierarchicalContext = {
      currentMainSection: null,
      sectionStack: []
    };

    // Process each TOC entry
    tocEntries.forEach((tocEntry, index) => {
      try {
        const targetSection = findTargetSection(tocEntry, document, styleInfo, hierarchicalContext);
        
        if (targetSection) {
          createTOCLink(tocEntry, targetSection);
          linkedCount++;
          
          // Update hierarchical context based on the found section
          updateHierarchicalContext(tocEntry, targetSection, hierarchicalContext);
        } else {
          console.warn(`Could not find target section for TOC entry: "${getTOCEntryText(tocEntry)}"`);
          skippedCount++;
        }
      } catch (error) {
        console.error(`Error processing TOC entry ${index}:`, error);
        skippedCount++;
      }
    });

    // Validate all created links
    const validationResults = validateTOCLinks(document);
    
    console.log(`TOC linking complete: ${linkedCount} linked, ${skippedCount} skipped`);
    console.log(`Link validation: ${validationResults.valid} valid, ${validationResults.invalid} invalid`);

    // Add accessibility enhancements
    addTOCAccessibilityFeatures(document);

  } catch (error) {
    console.error('Error in TOC linking process:', error);
  }
}

/**
 * Find all TOC entries in the document using semantic markers
 * 
 * @param {Document} document - The HTML document
 * @returns {Array} Array of TOC entry elements
 */
function findTOCEntries(document) {
  // Find all elements marked as TOC entries by the TOC processor
  const tocEntries = document.querySelectorAll('.docx-toc-entry[data-toc-entry="true"]');
  
  return Array.from(tocEntries);
}

/**
 * Find the target section for a TOC entry using multiple strategies
 * 
 * @param {Element} tocEntry - The TOC entry element
 * @param {Document} document - The HTML document
 * @param {Object} styleInfo - Style information including numbering context
 * @param {Object} context - Hierarchical context for tracking section relationships
 * @returns {Element|null} The target section element or null if not found
 */
function findTargetSection(tocEntry, document, styleInfo, context = {}) {
  const tocText = getTOCEntryText(tocEntry);
  
  // Strategy 1: Hierarchical section ID matching
  let targetSection = findByHierarchicalSectionId(tocEntry, tocText, document, context);
  if (targetSection) {
    return targetSection;
  }

  // Strategy 2: Direct section ID matching
  targetSection = findBySectionId(tocEntry, tocText, document);
  if (targetSection) {
    return targetSection;
  }

  // Strategy 3: Text content matching
  targetSection = findByTextContent(tocText, document);
  if (targetSection) {
    return targetSection;
  }

  // Strategy 4: Numbering pattern matching
  targetSection = findByNumberingPattern(tocText, document);
  if (targetSection) {
    return targetSection;
  }

  // Strategy 5: Fallback ID generation
  targetSection = findWithFallbackId(tocEntry, tocText, document);
  if (targetSection) {
    return targetSection;
  }

  return null;
}

/**
 * Strategy 1: Find target section by hierarchical section ID matching
 * 
 * @param {Element} tocEntry - The TOC entry element
 * @param {string} tocText - The text content of the TOC entry
 * @param {Document} document - The HTML document
 * @param {Object} context - Hierarchical context for tracking section relationships
 * @returns {Element|null} The target section element or null if not found
 */
function findByHierarchicalSectionId(tocEntry, tocText, document, context) {
  // Extract numbering from TOC text
  const numberMatch = tocText.match(/^(\d+)\./);
  const letterMatch = tocText.match(/^([a-z])\./i);
  
  if (numberMatch) {
    // This is a main section (e.g., "1. Title", "2. Title")
    const sectionNumber = numberMatch[1];
    const expectedId = `section-${sectionNumber}`;
    
    const targetElement = document.getElementById(expectedId);
    if (targetElement) {
      // Update context for future sub-sections
      context.currentMainSection = sectionNumber;
      context.sectionStack = [sectionNumber];
      return targetElement;
    }
  } else if (letterMatch && context.currentMainSection) {
    // This is a sub-section (e.g., "a. Title", "b. Title")
    const subLetter = letterMatch[1].toLowerCase();
    const expectedId = `section-${context.currentMainSection}-${subLetter}`;
    
    const targetElement = document.getElementById(expectedId);
    if (targetElement) {
      return targetElement;
    }
  }
  
  return null;
}

/**
 * Update hierarchical context based on processed TOC entry
 * 
 * @param {Element} tocEntry - The TOC entry element
 * @param {Element} targetSection - The target section element that was found
 * @param {Object} context - Hierarchical context to update
 */
function updateHierarchicalContext(tocEntry, targetSection, context) {
  const tocText = getTOCEntryText(tocEntry);
  const numberMatch = tocText.match(/^(\d+)\./);
  
  if (numberMatch) {
    // This was a main section, context is already updated in findByHierarchicalSectionId
    return;
  }
  
  // For other types of entries, maintain the current context
}

/**
 * Strategy 2: Find target section by matching section IDs
 * 
 * @param {Element} tocEntry - The TOC entry element
 * @param {string} tocText - The text content of the TOC entry
 * @param {Document} document - The HTML document
 * @returns {Element|null} The target section element or null if not found
 */
function findBySectionId(tocEntry, tocText, document) {
  // Extract numbering from TOC text (e.g., "1.2.a Some Title" -> "1.2.a", "a. Some Title" -> "a")
  const numberingMatch = tocText.match(/^([a-z]+\.|\d+(?:\.\d+)*(?:\.[a-z]+)*)/i);
  
  if (numberingMatch) {
    const numbering = numberingMatch[1];
    
    // Generate expected section ID
    const expectedId = `section-${numbering.toLowerCase().replace(/\./g, '-')}`;
    
    // Look for element with this ID
    let targetElement = document.getElementById(expectedId);
    if (targetElement) {
      return targetElement;
    }

    // Also check data-section-id attributes
    targetElement = document.querySelector(`[data-section-id="${expectedId}"]`);
    if (targetElement) {
      return targetElement;
    }
  }

  return null;
}

/**
 * Strategy 2: Find target section by text content matching
 * 
 * @param {string} tocText - The text content of the TOC entry
 * @param {Document} document - The HTML document
 * @returns {Element|null} The target section element or null if not found
 */
function findByTextContent(tocText, document) {
  // Clean the TOC text for comparison
  const cleanTocText = tocText.toLowerCase()
    .replace(/^[\d\.\s]+/, '') // Remove leading numbering
    .replace(/[^\w\s]/g, '') // Remove special characters
    .trim();

  if (!cleanTocText) {
    return null;
  }

  // Search in headings first
  const headings = document.querySelectorAll('h1, h2, h3, h4, h5, h6');
  for (const heading of headings) {
    const headingText = heading.textContent.toLowerCase()
      .replace(/^[\d\.\s]+/, '') // Remove leading numbering
      .replace(/[^\w\s]/g, '') // Remove special characters
      .trim();

    if (headingText === cleanTocText || 
        headingText.includes(cleanTocText) || 
        cleanTocText.includes(headingText)) {
      return heading;
    }
  }

  // Search in numbered paragraphs
  const numberedParagraphs = document.querySelectorAll('p[data-abstract-num]');
  for (const paragraph of numberedParagraphs) {
    const paragraphText = paragraph.textContent.toLowerCase()
      .replace(/^[\d\.\s]+/, '') // Remove leading numbering
      .replace(/[^\w\s]/g, '') // Remove special characters
      .trim();

    if (paragraphText === cleanTocText || 
        paragraphText.includes(cleanTocText) || 
        cleanTocText.includes(paragraphText)) {
      return paragraph;
    }
  }

  return null;
}

/**
 * Strategy 3: Find target section by numbering pattern matching
 * 
 * @param {string} tocText - The text content of the TOC entry
 * @param {Document} document - The HTML document
 * @returns {Element|null} The target section element or null if not found
 */
function findByNumberingPattern(tocText, document) {
  // Extract numbering pattern from TOC text
  const numberingMatch = tocText.match(/^([\d\.\w]+)/);
  
  if (!numberingMatch) {
    return null;
  }

  const numbering = numberingMatch[1];

  // Look for elements with matching numbering in their content
  const allElements = document.querySelectorAll('h1, h2, h3, h4, h5, h6, p[data-abstract-num]');
  
  for (const element of allElements) {
    const elementText = element.textContent.trim();
    
    // Check if element starts with the same numbering
    if (elementText.startsWith(numbering)) {
      return element;
    }
  }

  return null;
}

/**
 * Strategy 4: Find target section with fallback ID generation
 * 
 * @param {Element} tocEntry - The TOC entry element
 * @param {string} tocText - The text content of the TOC entry
 * @param {Document} document - The HTML document
 * @returns {Element|null} The target section element or null if not found
 */
function findWithFallbackId(tocEntry, tocText, document) {
  // Generate a fallback ID from the text content
  const fallbackId = generateIdFromText(tocText);
  
  if (!fallbackId) {
    return null;
  }

  // Look for existing element with this ID
  let targetElement = document.getElementById(fallbackId);
  if (targetElement) {
    return targetElement;
  }

  // Try to find a suitable element to assign this ID to
  const cleanText = tocText.replace(/^[\d\.\s]+/, '').trim();
  
  if (cleanText) {
    const allElements = document.querySelectorAll('h1, h2, h3, h4, h5, h6, p');
    
    for (const element of allElements) {
      const elementText = element.textContent.replace(/^[\d\.\s]+/, '').trim();
      
      if (elementText.toLowerCase().includes(cleanText.toLowerCase()) && !element.id) {
        // Assign the fallback ID to this element
        element.id = fallbackId;
        return element;
      }
    }
  }

  return null;
}

/**
 * Create a clickable link for a TOC entry
 * 
 * @param {Element} tocEntry - The TOC entry element
 * @param {Element} targetSection - The target section element
 */
function createTOCLink(tocEntry, targetSection) {
  // Ensure target has an ID
  if (!targetSection.id) {
    const tocText = getTOCEntryText(tocEntry);
    targetSection.id = generateIdFromText(tocText) || `section-${Date.now()}`;
  }

  // Find the text content within the TOC entry
  const tocTextElement = tocEntry.querySelector('.docx-toc-text') || tocEntry;
  
  // Create the link element
  const link = tocEntry.ownerDocument.createElement('a');
  link.href = `#${targetSection.id}`;
  link.className = 'docx-toc-link';
  link.setAttribute('aria-describedby', 'toc-nav-desc');
  
  // Move the text content into the link
  if (tocTextElement === tocEntry) {
    // If the entire entry is the text, wrap it
    const textContent = tocEntry.innerHTML;
    tocEntry.innerHTML = '';
    link.innerHTML = textContent;
    tocEntry.appendChild(link);
  } else {
    // If there's a specific text element, wrap that
    const textContent = tocTextElement.innerHTML;
    tocTextElement.innerHTML = '';
    link.innerHTML = textContent;
    tocTextElement.appendChild(link);
  }

  // Add data attributes for tracking
  tocEntry.setAttribute('data-toc-linked', 'true');
  tocEntry.setAttribute('data-target-id', targetSection.id);
}

/**
 * Validate all TOC links in the document
 * 
 * @param {Document} document - The HTML document
 * @returns {Object} Validation results with counts of valid and invalid links
 */
function validateTOCLinks(document) {
  const tocLinks = document.querySelectorAll('.docx-toc-link');
  let validCount = 0;
  let invalidCount = 0;

  tocLinks.forEach(link => {
    const href = link.getAttribute('href');
    
    if (href && href.startsWith('#')) {
      const targetId = href.substring(1);
      const targetElement = document.getElementById(targetId);
      
      if (targetElement) {
        validCount++;
      } else {
        console.warn(`Invalid TOC link: ${href} - target element not found`);
        invalidCount++;
      }
    } else {
      console.warn(`Invalid TOC link href: ${href}`);
      invalidCount++;
    }
  });

  return {
    valid: validCount,
    invalid: invalidCount
  };
}

/**
 * Add accessibility features for TOC navigation
 * 
 * @param {Document} document - The HTML document
 */
function addTOCAccessibilityFeatures(document) {
  // Add navigation description for screen readers
  const tocContainers = document.querySelectorAll('.docx-toc, nav[role="navigation"]');
  
  tocContainers.forEach(container => {
    if (!container.querySelector('#toc-nav-desc')) {
      const navDesc = document.createElement('div');
      navDesc.id = 'toc-nav-desc';
      navDesc.className = 'sr-only';
      navDesc.textContent = 'Table of Contents navigation - click to jump to section';
      container.insertBefore(navDesc, container.firstChild);
    }
  });

  // Ensure TOC links are keyboard accessible
  const tocLinks = document.querySelectorAll('.docx-toc-link');
  tocLinks.forEach(link => {
    if (!link.getAttribute('tabindex')) {
      link.setAttribute('tabindex', '0');
    }
  });
}

/**
 * Extract text content from a TOC entry
 * 
 * @param {Element} tocEntry - The TOC entry element
 * @returns {string} The text content of the TOC entry
 */
function getTOCEntryText(tocEntry) {
  return tocEntry.textContent.trim();
}

/**
 * Generate a valid HTML ID from text content
 * 
 * @param {string} text - The text to convert to an ID
 * @returns {string|null} A valid HTML ID or null if text is empty
 */
function generateIdFromText(text) {
  if (!text) return null;
  
  let id = text
    .toLowerCase()
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');

  // Ensure ID starts with a letter
  if (id && !/^[a-z]/.test(id)) {
    id = `section-${id}`;
  }

  return id || null;
}

module.exports = {
  linkTOCEntries,
  findTOCEntries,
  findTargetSection,
  findByHierarchicalSectionId,
  updateHierarchicalContext,
  findBySectionId,
  findByTextContent,
  findByNumberingPattern,
  findWithFallbackId,
  createTOCLink,
  validateTOCLinks,
  addTOCAccessibilityFeatures,
  getTOCEntryText,
  generateIdFromText
};
