// File: lib/html/processors/toc-processor.js
// Table of Contents processing utilities

const { findParagraphContextByText, findParagraphContextByPartialText } = require("./heading-processor");

/**
 * Process TOC in the document
 */
function processTOC(doc, styleInfo, numberingContext = null) {
  try {
    // First, look for elements already marked as TOC entries
    let tocEntries = Array.from(
      doc.querySelectorAll(
        'p[class*="docx-toc-level-"], p[class*="TOC"], .docx-toc-entry'
      )
    );

    // If we don't find any TOC entries with standard classes, use broader detection
    if (tocEntries.length === 0) {
      // First pass: look for paragraphs with tab stops followed by numbers
      const paragraphs = doc.querySelectorAll("p");
      let tocCandidates = [];

      for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        const text = p.textContent.trim();

        // Check for text followed by a page number
        const pageNumMatch = text.match(/^(.*?)[\.\s_]+(\d+)$/);

        if (pageNumMatch) {
          // Found a potential TOC entry
          p.classList.add("docx-toc-entry");
          tocCandidates.push(p);
          continue;
        }

        // Special case: also detect entries that might be part of TOC but don't have page numbers
        // Like "Rationale for Resolution..." lines in the screenshot
        if (tocCandidates.length > 0) {
          // We already have some TOC candidates, check if this could be a related entry

          // If it's indented similar to other TOC entries or has TOC-like styling
          if (p.style.marginLeft && parseFloat(p.style.marginLeft) > 0) {
            p.classList.add("docx-toc-entry");
            tocCandidates.push(p);
          }
          // Check if it's near other TOC entries in the DOM structure
          else if (tocCandidates.includes(p.previousElementSibling)) {
            // This paragraph follows a TOC entry, check its structure
            // Look for lines that might be TOC-related but don't have page numbers yet
            // This is a structural pattern, not content pattern matching
            if (text.length > 0 && !text.match(/^[\d\.]+\s/)) {
              // Non-numbered paragraph after TOC entry, likely part of TOC
              p.classList.add("docx-toc-entry");
              tocCandidates.push(p);
            }
          }
        }
      }

      if (tocCandidates.length > 0) tocEntries = tocCandidates;
    }

    // Filter out already processed entries
    tocEntries = tocEntries.filter((entry) => !entry.closest("nav.docx-toc"));

    if (tocEntries.length > 0) {
      // Create TOC container
      const tocContainer = createTOCContainer(doc, tocEntries[0], styleInfo);

      // Process each TOC entry
      tocEntries.forEach((entry, index) => {
        // Try to find matching numbering context for this TOC entry
        const entryText = entry.textContent
          .trim()
          .replace(/[\.\s_]+\d+$/, "") // Remove page number if present
          .trim();

        let context = null;
        if (numberingContext) {
          context =
            findParagraphContextByText(entryText, numberingContext) ||
            findParagraphContextByPartialText(entryText, numberingContext);
        }

        processTOCEntry(entry, index, tocContainer, styleInfo, context);
      });

      // Additional pass to ensure all entries have right-aligned page numbers with leader dots
      const entriesInContainer = tocContainer.querySelectorAll(".docx-toc-entry");
      entriesInContainer.forEach((entry) => {
        // Check if this entry has the dots and page number structure
        validateTOCEntryStructure(entry);
      });
      
      // Debug output to help diagnose issues
      console.log(`Processed ${tocEntries.length} TOC entries`);
      const finalTOCEntries = doc.querySelectorAll('.docx-toc-entry');
      console.log(`Final TOC entry count: ${finalTOCEntries.length}`);
      console.log(`TOC container HTML structure: ${doc.querySelector('.docx-toc')?.innerHTML.substring(0, 200)}...`);
    }
  } catch (error) {
    console.error("Error processing TOC:", error.message, error.stack);
  }
}

/**
 * Create TOC container
 */
function createTOCContainer(doc, firstEntryElement, styleInfo) {
  let tocContainer = doc.querySelector("nav.docx-toc");
  if (!tocContainer) {
    tocContainer = doc.createElement("nav");
    tocContainer.className = "docx-toc";
    tocContainer.setAttribute("role", "navigation");
    tocContainer.setAttribute("aria-label", "Table of Contents");
    let tocHeading = doc.querySelector("h2.docx-toc-heading"); // Assume TOC heading is H2
    if (!tocHeading) {
      tocHeading = doc.createElement("h2");
      tocHeading.className = "docx-toc-heading";
      tocHeading.textContent = "Table of Contents";
      tocContainer.appendChild(tocHeading);
    } else {
      if (tocHeading.parentNode !== tocContainer) {
        // Avoid moving if already in a new container
        tocContainer.appendChild(tocHeading.cloneNode(true));
        if (tocHeading.parentNode)
          tocHeading.parentNode.removeChild(tocHeading);
      }
    }
    if (firstEntryElement && firstEntryElement.parentNode) {
      firstEntryElement.parentNode.insertBefore(
        tocContainer,
        firstEntryElement
      );
    } else {
      doc.body.insertBefore(tocContainer, doc.body.firstChild); // Fallback
    }
  }
  return tocContainer;
}

/**
 * Process individual TOC entry
 */
function processTOCEntry(entry, index, tocContainer, styleInfo, numberingCtx) {
  try {
    // Get the document object from the entry element
    const doc = entry.ownerDocument;
    
    if (
      entry.closest("nav.docx-toc") &&
      entry.classList.contains("processed-toc-entry")
    )
      return; // Already processed and moved

    entry.classList.add("processed-toc-entry");
    let level = determineTOCLevel(entry, styleInfo);
    entry.classList.add(`docx-toc-level-${level}`);

    const textContent = entry.textContent.trim();
    structureTOCEntry(entry, textContent, level); // This will create spans for text, dots, pagenum

    // If this TOC entry should be numbered (e.g., "1. Introduction ... page 5")
    if (
      numberingCtx &&
      numberingCtx.resolvedNumbering &&
      numberingCtx.abstractNumId
    ) {
      entry.setAttribute("data-num-id", numberingCtx.numberingId || "");
      entry.setAttribute("data-abstract-num", numberingCtx.abstractNumId);
      entry.setAttribute(
        "data-num-level",
        numberingCtx.numberingLevel?.toString() || ""
      );
      entry.setAttribute(
        "data-format",
        numberingCtx.resolvedNumbering.format || ""
      );
      
      // Set a unique TOC entry ID (different from the heading ID it links to)
      if (numberingCtx.resolvedNumbering.sectionId) {
        entry.id = `toc-${numberingCtx.resolvedNumbering.sectionId}`;
        entry.setAttribute("data-section-id", numberingCtx.resolvedNumbering.sectionId);
      }
    }

    if (entry.parentNode !== tocContainer) {
      tocContainer.appendChild(entry); // Move the original entry
    }

    entry.setAttribute("role", "listitem"); // TOC entries are items in a list (the NAV)
    const textSpan = entry.querySelector(".docx-toc-text");
    if (textSpan && !textSpan.hasAttribute("role")) {
      // Make the text part a link if it's not already
      const anchor = doc.createElement("a");
      
      // Get the text content without any numbering prefix
      let linkText = textSpan.textContent;
      let hrefTarget = "";
      
      // Priority 1: Use section ID from numbering context if available and exists in document
      if (numberingCtx && numberingCtx.resolvedNumbering && numberingCtx.resolvedNumbering.sectionId) {
        const targetElement = doc.getElementById(numberingCtx.resolvedNumbering.sectionId);
        if (targetElement) {
          hrefTarget = numberingCtx.resolvedNumbering.sectionId;
        }
      }
      
      // Priority 2: Find the matching heading in the document body
      if (!hrefTarget) {
        hrefTarget = findMatchingHeadingId(doc, linkText, numberingCtx);
      }
      
      // Priority 3: Fallback to a generated ID if nothing found
      if (!hrefTarget) {
        hrefTarget = generateIdFromText(linkText);
      }
      
      anchor.href = `#${hrefTarget}`; // Link to the actual heading ID
      anchor.textContent = linkText;
      anchor.className = "keyboard-focusable";
      textSpan.textContent = "";
      textSpan.appendChild(anchor);
      textSpan.setAttribute("role", "link");
      textSpan.setAttribute("tabindex", "0");
    }
    
    // Validate the structure after processing
    validateTOCEntryStructure(entry);
  } catch (error) {
    console.error(
      "Error processing TOC entry:",
      entry.textContent?.substring(0, 50),
      error.message,
      error.stack
    );
  }
}

/**
 * Determine TOC level from entry
 */
function determineTOCLevel(entry, styleInfo) {
  let level = 1;
  // Try from class like docx-toc-level-2
  for (const cls of entry.classList) {
    const match = cls.match(/^docx-toc-level-(\d+)$/);
    if (match) return parseInt(match[1], 10);
  }
  // Try from style name like TOC1, toc 2
  const styleNameMatch = Array.from(entry.classList).find((cls) =>
    /^docx-p-toc\d+/.test(cls)
  );
  if (styleNameMatch) {
    const levelMatch = styleNameMatch.match(/toc(\d+)/i);
    if (levelMatch && levelMatch[1]) return parseInt(levelMatch[1], 10);
  }
  return level;
}

/**
 * Structure TOC entry with spans for text, dots, and page number
 */
function structureTOCEntry(entry, text, level) {
  try {
    const pageNumMatch = text.match(/^(.*?)[\s._]+(\d+)$/); // More flexible leader match
    const doc = entry.ownerDocument;
    entry.textContent = "";

    const textSpan = doc.createElement("span");
    textSpan.className = "docx-toc-text";
    textSpan.textContent = pageNumMatch ? pageNumMatch[1].trim() : text.trim();
    entry.appendChild(textSpan);

    if (pageNumMatch) {
      const dotsSpan = doc.createElement("span");
      dotsSpan.className = "docx-toc-dots";
      dotsSpan.setAttribute("aria-hidden", "true");
      entry.appendChild(dotsSpan);

      const pageSpan = doc.createElement("span");
      pageSpan.className = "docx-toc-pagenum";
      pageSpan.textContent = pageNumMatch[2];
      pageSpan.setAttribute("aria-label", `Page ${pageNumMatch[2]}`);
      entry.appendChild(pageSpan);
    } else {
      // Even if no page number, add the dots and empty page number span for consistent layout
      const dotsSpan = doc.createElement("span");
      dotsSpan.className = "docx-toc-dots";
      dotsSpan.setAttribute("aria-hidden", "true");
      entry.appendChild(dotsSpan);
      
      const pageSpan = doc.createElement("span");
      pageSpan.className = "docx-toc-pagenum";
      pageSpan.textContent = "";
      entry.appendChild(pageSpan);
    }
  } catch (error) {
    console.error("Error structuring TOC entry:", error.message);
  }
}

/**
 * Validate and fix TOC entry structure
 */
function validateTOCEntryStructure(entry) {
  // Check if entry has all required spans
  if (!entry.querySelector('.docx-toc-text') || 
      !entry.querySelector('.docx-toc-dots') || 
      !entry.querySelector('.docx-toc-pagenum')) {
    
    // If missing, restructure completely
    const text = entry.textContent.trim();
    entry.textContent = '';
    
    // Create text span
    const textSpan = entry.ownerDocument.createElement('span');
    textSpan.className = 'docx-toc-text';
    // Look for page number pattern
    const pageNumMatch = text.match(/^(.*?)[\.\s_]+(\d+)$/);
    textSpan.textContent = pageNumMatch ? pageNumMatch[1].trim() : text;
    entry.appendChild(textSpan);
    
    // Add dots span
    const dotsSpan = entry.ownerDocument.createElement('span');
    dotsSpan.className = 'docx-toc-dots';
    dotsSpan.setAttribute("aria-hidden", "true");
    entry.appendChild(dotsSpan);
    
    // Add page number span
    const pageSpan = entry.ownerDocument.createElement('span');
    pageSpan.className = 'docx-toc-pagenum';
    pageSpan.textContent = pageNumMatch ? pageNumMatch[2] : '';
    if (pageNumMatch) {
      pageSpan.setAttribute("aria-label", `Page ${pageNumMatch[2]}`);
    }
    entry.appendChild(pageSpan);
    
    return true; // Fixed structure
  }
  return false; // No fix needed
}

/**
 * Find matching heading ID in document
 */
function findMatchingHeadingId(doc, tocText, numberingCtx) {
  // Clean the TOC text for matching
  const cleanTocText = tocText.trim().toLowerCase();
  
  // First, try to use the section ID from numbering context if available
  if (numberingCtx && numberingCtx.resolvedNumbering && numberingCtx.resolvedNumbering.sectionId) {
    // Verify this section ID actually exists in the document
    const targetElement = doc.getElementById(numberingCtx.resolvedNumbering.sectionId);
    if (targetElement) {
      return numberingCtx.resolvedNumbering.sectionId;
    }
  }
  
  // Find all headings and anchors in the document body
  const headings = doc.querySelectorAll('h1, h2, h3, h4, h5, h6');
  const anchors = doc.querySelectorAll('a[id]');
  
  // Try to match with headings first - prioritize exact text matches
  for (const heading of headings) {
    if (heading.closest('nav.docx-toc')) continue; // Skip TOC headings
    
    // Get the heading content without any numbering prefix
    const headingContentSpan = heading.querySelector('.heading-content');
    const headingText = headingContentSpan ? 
      headingContentSpan.textContent.trim().toLowerCase() : 
      heading.textContent.trim().toLowerCase();
    
    // Remove any numbering prefix from TOC text for comparison
    const tocTextWithoutNumber = cleanTocText.replace(/^\d+\.\s*/, '').replace(/^[a-z]\.\s*/, '');
    
    // Try exact match first
    if (headingText === tocTextWithoutNumber || headingText === cleanTocText) {
      return heading.id;
    }
    
    // Try partial match (TOC text contains heading text or vice versa)
    if (tocTextWithoutNumber.includes(headingText) || headingText.includes(tocTextWithoutNumber)) {
      return heading.id;
    }
    
    // Try matching without punctuation and extra spaces
    const cleanHeadingText = headingText.replace(/[^\w\s]/g, '').replace(/\s+/g, ' ');
    const cleanTocTextNoPunct = tocTextWithoutNumber.replace(/[^\w\s]/g, '').replace(/\s+/g, ' ');
    
    if (cleanHeadingText === cleanTocTextNoPunct) {
      return heading.id;
    }
  }
  
  // Try to match with existing anchors
  for (const anchor of anchors) {
    if (anchor.closest('nav.docx-toc')) continue; // Skip TOC anchors
    
    // Check if the anchor is near text that matches
    const parentText = anchor.parentElement ? anchor.parentElement.textContent.trim().toLowerCase() : '';
    const tocTextWithoutNumber = cleanTocText.replace(/^\d+\.\s*/, '').replace(/^[a-z]\.\s*/, '');
    
    if (parentText.includes(tocTextWithoutNumber) || tocTextWithoutNumber.includes(parentText)) {
      return anchor.id;
    }
  }
  
  // Fallback: generate a reasonable ID from the TOC text
  return generateIdFromText(cleanTocText);
}

/**
 * Generate clean ID from text
 */
function generateIdFromText(text) {
  return text
    .toLowerCase()
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '-')
    .replace(/^-+|-+$/g, '')
    .substring(0, 50) || 'toc-entry';
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
};