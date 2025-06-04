// File: lib/html/processors/toc-processor.js
// TOC processing utilities - preserves existing link structure

const { convertTwipToPt } = require('../../utils/unit-converter');

/**
 * Process TOC entries - minimal processing to preserve existing links
 */
function processTOC(document, styleInfo, numberingContext) {
  if (!document) return;

  // Find TOC containers
  const tocContainers = document.querySelectorAll('.docx-toc, nav[role="navigation"]');
  
  tocContainers.forEach(container => {
    processTOCContainer(container, styleInfo, numberingContext);
  });

  // Also process any standalone TOC entries
  const standaloneEntries = document.querySelectorAll('.docx-toc-entry');
  standaloneEntries.forEach((entry, index) => {
    if (!entry.closest('.docx-toc') && !entry.closest('nav[role="navigation"]')) {
      processTOCEntry(entry, index, null, styleInfo, null);
    }
  });
}

/**
 * Process TOC container - minimal enhancement without breaking existing structure
 */
function processTOCContainer(container, styleInfo, numberingContext) {
  if (!container) return;

  // Ensure proper accessibility attributes
  if (!container.hasAttribute('role')) {
    container.setAttribute('role', 'navigation');
  }
  if (!container.hasAttribute('aria-label')) {
    container.setAttribute('aria-label', 'Table of Contents');
  }

  // Process TOC entries within this container
  const entries = container.querySelectorAll('.docx-toc-entry');
  entries.forEach((entry, index) => {
    processTOCEntry(entry, index, container, styleInfo, numberingContext);
  });
}

/**
 * Process individual TOC entry - preserve existing links and structure
 */
function processTOCEntry(entry, index, container, styleInfo, numberingContext) {
  if (!entry) return;

  // Skip if already processed
  if (entry.hasAttribute('data-toc-processed')) {
    return;
  }
  entry.setAttribute('data-toc-processed', 'true');

  // Ensure proper accessibility
  if (!entry.hasAttribute('role')) {
    entry.setAttribute('role', 'listitem');
  }
  if (!entry.hasAttribute('tabindex')) {
    entry.setAttribute('tabindex', '0');
  }

  // Preserve existing links - just enhance them
  const existingLinks = entry.querySelectorAll('a');
  existingLinks.forEach(link => {
    if (!link.classList.contains('keyboard-focusable')) {
      link.classList.add('keyboard-focusable');
    }
    if (!link.hasAttribute('tabindex')) {
      link.setAttribute('tabindex', '0');
    }
  });

  // Apply enhanced TOC processing if enhanced data is available
  const entryText = entry.textContent.trim();
  const enhancedTOCEntries = styleInfo.tocStyles?.enhancedTOCEntries || [];
  const tocEntryData = enhancedTOCEntries.find(tocData => {
    const tocText = tocData.runProperties?.map(rp => rp.text).join('') || '';
    const cleanEntryText = entryText.replace(/\s+/g, ' ').trim();
    const cleanTocText = tocText.replace(/\s+/g, ' ').trim();
    
    return cleanEntryText.includes(cleanTocText) || 
           cleanTocText.includes(cleanEntryText) ||
           cleanEntryText === cleanTocText;
  });

  // Apply character formatting if available (without breaking links)
  if (tocEntryData && tocEntryData.runProperties) {
    applyCharacterFormattingToTOCEntry(entry, tocEntryData.runProperties);
  }

  // Apply precise indentation if available
  if (tocEntryData) {
    applyPreciseIndentationToTOCEntry(entry, tocEntryData);
  }
}

/**
 * Apply character formatting while preserving link structure
 */
function applyCharacterFormattingToTOCEntry(entry, runProperties) {
  // Check if this entry has existing links that we need to preserve
  const existingLinks = entry.querySelectorAll('a');
  
  if (existingLinks.length > 0) {
    // Apply formatting to link text instead of replacing the entire entry
    existingLinks.forEach(link => {
      applyFormattingToLinkText(link, runProperties);
    });
    return;
  }

  // If no existing links, we can safely apply formatting to the entire entry
  let hasItalic = false;
  let hasBold = false;
  
  runProperties.forEach(run => {
    if (run.formatting) {
      if (run.formatting.italic) hasItalic = true;
      if (run.formatting.bold) hasBold = true;
    }
  });

  if (hasItalic || hasBold) {
    const textContent = entry.textContent;
    let formattedText = textContent;
    
    if (hasItalic) {
      formattedText = `<em class="toc-italic">${formattedText}</em>`;
    }
    if (hasBold) {
      formattedText = `<strong class="toc-bold">${formattedText}</strong>`;
    }
    
    entry.innerHTML = formattedText;
  }
}

/**
 * Apply formatting to link text while preserving link structure
 */
function applyFormattingToLinkText(link, runProperties) {
  let hasItalic = false;
  let hasBold = false;
  let hasColor = false;
  
  runProperties.forEach(run => {
    if (run.formatting) {
      if (run.formatting.italic) hasItalic = true;
      if (run.formatting.bold) hasBold = true;
      if (run.formatting.color) hasColor = true;
    }
  });

  if (hasItalic || hasBold || hasColor) {
    const linkText = link.textContent;
    let formattedText = linkText;
    
    if (hasItalic) {
      formattedText = `<em class="toc-italic">${formattedText}</em>`;
    }
    if (hasBold) {
      formattedText = `<strong class="toc-bold">${formattedText}</strong>`;
    }
    
    link.innerHTML = formattedText;
    
    if (hasColor) {
      const colorRun = runProperties.find(run => run.formatting?.color);
      if (colorRun) {
        link.style.color = colorRun.formatting.color;
      }
    }
  }
}

/**
 * Apply precise indentation to TOC entry based on DOCX measurements
 */
function applyPreciseIndentationToTOCEntry(entry, tocEntryData) {
  const leftIndentPt = convertTwipToPt(tocEntryData.leftIndent || 0);
  const hangingIndentPt = convertTwipToPt(tocEntryData.hangingIndent || 0);
  
  if (leftIndentPt > 0) {
    entry.style.marginLeft = `${leftIndentPt}pt`;
  }
  
  if (hangingIndentPt > 0) {
    entry.style.paddingLeft = `${hangingIndentPt}pt`;
    entry.style.textIndent = `-${hangingIndentPt}pt`;
  }
  
  // Add data attributes for CSS targeting
  entry.setAttribute('data-left-indent', leftIndentPt);
  entry.setAttribute('data-hanging-indent', hangingIndentPt);
  
  // Add enhanced formatting flags
  if (tocEntryData.hasItalic) {
    entry.setAttribute('data-has-italic', 'true');
  }
  if (tocEntryData.hasBold) {
    entry.setAttribute('data-has-bold', 'true');
  }
}

/**
 * Remove page numbers from TOC entries
 * Only processes elements that are part of the Table of Contents
 */
function removeTOCPageNumbers(document) {
  if (!document) return;

  console.log('Starting TOC page number removal...');

  // Simple patterns to detect trailing page numbers
  const pageNumberPatterns = [
    /\t+\d+$/,              // Tab + any digits at end
    /\s{2,}\d+$/,           // 2+ spaces + any digits at end  
    /:\s*\d+$/,             // Colon + optional spaces + any digits at end
    /\s+[ivxlcdm]+$/i       // Roman numerals (case insensitive)
  ];

  let processedCount = 0;

  // First, try to find explicit TOC containers
  const tocContainers = document.querySelectorAll('.docx-toc, nav[role="navigation"]');
  
  if (tocContainers.length > 0) {
    // Process paragraphs within TOC containers
    tocContainers.forEach(container => {
      const tocParagraphs = container.querySelectorAll('p');
      tocParagraphs.forEach(element => {
        processedCount += processElementForPageNumbers(element, pageNumberPatterns);
        
        // Add semantic marker for TOC linking
        element.classList.add('docx-toc-entry');
        element.setAttribute('data-toc-entry', 'true');
      });
    });
  } else {
    // Fallback: Look for the beginning of the document where TOC typically appears
    // Process only the first section of paragraphs (likely TOC area)
    const allParagraphs = document.querySelectorAll('p');
    let tocSectionFound = false;
    let consecutiveNonTocCount = 0;
    
    for (let i = 0; i < allParagraphs.length && i < 50; i++) { // Limit to first 50 paragraphs
      const element = allParagraphs[i];
      const text = element.textContent.trim();
      
      // Check if this looks like a TOC entry (has trailing numbers with specific patterns)
      const looksLikeTocEntry = pageNumberPatterns.some(pattern => pattern.test(text));
      
      if (looksLikeTocEntry) {
        tocSectionFound = true;
        consecutiveNonTocCount = 0;
        processedCount += processElementForPageNumbers(element, pageNumberPatterns);
        
        // Add semantic marker for TOC linking
        element.classList.add('docx-toc-entry');
        element.setAttribute('data-toc-entry', 'true');
      } else if (tocSectionFound) {
        consecutiveNonTocCount++;
        // If we've found 5 consecutive non-TOC entries after finding TOC entries, stop
        if (consecutiveNonTocCount >= 5) {
          break;
        }
      }
    }
  }

  console.log(`TOC page number removal complete. Processed ${processedCount} entries.`);
}

/**
 * Process a single element for page number removal
 */
function processElementForPageNumbers(element, pageNumberPatterns) {
  const originalText = element.textContent;
  if (!originalText) return 0;

  // Check if this has any trailing number patterns
  let hasTrailingNumbers = false;
  pageNumberPatterns.forEach(pattern => {
    if (pattern.test(originalText)) {
      hasTrailingNumbers = true;
    }
  });

  if (!hasTrailingNumbers) return 0;

  // Apply page number removal patterns - remove ALL trailing numbers
  let cleanedText = originalText;
  let wasModified = false;

  pageNumberPatterns.forEach(pattern => {
    const newText = cleanedText.replace(pattern, '');
    if (newText !== cleanedText) {
      cleanedText = newText;
      wasModified = true;
    }
  });

  // Update the element if we removed something
  if (wasModified && cleanedText.trim() !== originalText.trim()) {
    element.textContent = cleanedText.trim();
    // console.log(`Removed trailing number: "${originalText}" → "${cleanedText.trim()}"`);
    return 1;
  }

  return 0;
}

/**
 * Create TOC container - for backward compatibility
 */
function createTOCContainer(document) {
  const container = document.createElement('nav');
  container.className = 'docx-toc';
  container.setAttribute('role', 'navigation');
  container.setAttribute('aria-label', 'Table of Contents');
  return container;
}

/**
 * Determine TOC level - for backward compatibility
 */
function determineTOCLevel(entry, index) {
  // Simple level determination based on indentation or index
  const indentPt = parseFloat(entry.style.marginLeft) || 0;
  if (indentPt === 0) return 1;
  if (indentPt <= 20) return 1;
  if (indentPt <= 40) return 2;
  if (indentPt <= 60) return 3;
  return Math.min(Math.floor(indentPt / 20) + 1, 6);
}

/**
 * Structure TOC entry - for backward compatibility
 */
function structureTOCEntry(entry, level) {
  entry.classList.add(`docx-toc-level-${level}`);
  return entry;
}

/**
 * Validate TOC entry structure - for backward compatibility
 */
function validateTOCEntryStructure(entry) {
  return entry && entry.nodeType === 1; // Element node
}

/**
 * Find matching heading ID - for backward compatibility
 */
function findMatchingHeadingId(text, document) {
  const cleanText = text.toLowerCase().replace(/[^\w\s]/g, '').trim();
  const headings = document.querySelectorAll('h1, h2, h3, h4, h5, h6');
  
  for (const heading of headings) {
    const headingText = heading.textContent.toLowerCase().replace(/[^\w\s]/g, '').trim();
    if (headingText.includes(cleanText) || cleanText.includes(headingText)) {
      return heading.id || null;
    }
  }
  return null;
}

/**
 * Generate ID from text - for backward compatibility
 */
function generateIdFromText(text, context) {
  if (!text) return null;
  
  let id = text
    .toLowerCase()
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');

  if (context?.resolvedNumbering?.displayText) {
    const numberPrefix = context.resolvedNumbering.displayText
      .replace(/[^\w]/g, '')
      .toLowerCase();
    if (numberPrefix) {
      id = `section-${numberPrefix}-${id}`;
    }
  } else {
    id = `section-${id}`;
  }

  return id;
}

module.exports = {
  processTOC,
  createTOCContainer,
  processTOCEntry,
  determineTOCLevel,
  structureTOCEntry,
  validateTOCEntryStructure,
  findMatchingHeadingId,
  generateIdFromText,
  applyCharacterFormattingToTOCEntry,
  applyFormattingToLinkText,
  applyPreciseIndentationToTOCEntry,
  removeTOCPageNumbers
};
