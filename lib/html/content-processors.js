// File: ddttom/doc2web/lib/html/content-processors.js
// Modify the processHeadings and applyNumberingToHeading functions

const { selectNodes, selectSingleNode } = require("../xml/xpath-utils");
const {
  extractParagraphNumberingContext,
  resolveNumberingForParagraphs,
} = require("../parsers/numbering-resolver");

// Define Node constants for Node.js environment
const Node = {
  TEXT_NODE: 3,
  ELEMENT_NODE: 1,
};

/**
 * Enhanced heading processing using DOCX introspection
 * Extracts numbering from DOCX XML structure and applies to headings
 * @param {Document} doc - DOM document
 * @param {Object} styleInfo - Style information
 * @param {Array} numberingContext - Resolved paragraph numbering context
 */
function processHeadings(doc, styleInfo, numberingContext = null) {
  try {
    const headings = doc.querySelectorAll("h1, h2, h3, h4, h5, h6");

    if (!numberingContext) {
      console.warn(
        "No numbering context provided for heading processing. Headings will not be autonumbered from DOCX defs."
      );
      // Still ensure accessibility for non-numbered headings
      Array.from(headings).forEach((heading) => {
        if (!heading.classList.contains("docx-toc-heading")) {
          ensureHeadingAccessibility(heading);
        }
      });
      return;
    }

    Array.from(headings).forEach((heading) => {
      if (heading.classList.contains("docx-toc-heading")) {
        return;
      }

      const headingText = heading.textContent.trim();
      const context =
        findParagraphContextByText(headingText, numberingContext) ||
        findParagraphContextByPartialText(headingText, numberingContext);

      if (context && context.resolvedNumbering) {
        // If numbering comes from DOCX definitions, rely on CSS ::before for the number.
        // Just set the data attributes.
        heading.setAttribute("data-numbering-id", context.numberingId || "");
        heading.setAttribute(
          "data-numbering-level",
          context.numberingLevel?.toString() || ""
        );
        heading.setAttribute(
          "data-format",
          context.resolvedNumbering.format || ""
        );
        // The original text content of the heading remains, CSS will add the number.
      }
      // Ensure accessibility attributes are set regardless of numbering
      ensureHeadingAccessibility(heading);
    });
  } catch (error) {
    console.error("Error processing headings:", error.message);
  }
}

/**
 * Find paragraph context by exact text match
 * @param {string} text - Text to match
 * @param {Array} numberingContext - Numbering context array
 * @returns {Object|null} - Found context or null
 */
function findParagraphContextByText(text, numberingContext) {
  return (
    numberingContext.find((context) => {
      if (!context.textContent) return false;

      // Try exact match first
      if (context.textContent.trim() === text) {
        return true;
      }

      // Try match without trailing punctuation
      const cleanContextText = context.textContent
        .trim()
        .replace(/[:\.\s]*$/, "");
      const cleanSearchText = text.replace(/[:\.\s]*$/, "");

      return cleanContextText === cleanSearchText;
    }) || null
  );
}

/**
 * Find paragraph context by partial text match
 * @param {string} text - Text to match
 * @param {Array} numberingContext - Numbering context array
 * @returns {Object|null} - Found context or null
 */
function findParagraphContextByPartialText(text, numberingContext) {
  const cleanSearchText = text
    .toLowerCase()
    .trim()
    .replace(/[:\.\s]*$/, "");

  return (
    numberingContext.find((context) => {
      if (!context.textContent) return false;

      const cleanContextText = context.textContent
        .toLowerCase()
        .trim()
        .replace(/[:\.\s]*$/, "");

      // Check if either text contains the other (for truncated headings)
      return (
        cleanContextText.includes(cleanSearchText) ||
        cleanSearchText.includes(cleanContextText)
      );
    }) || null
  );
}

/**
 * Apply DOCX-derived numbering to heading (REMOVED - functionality moved to CSS)
 * This function is now simplified to only set data attributes if needed,
 * or can be removed if processHeadings directly sets attributes.
 * For clarity, we'll leave ensureHeadingAccessibility and rely on processHeadings to set data attributes.
 */
function applyNumberingToHeading(heading, numbering, context) {
  // This function's core logic of inserting a <span> for the number is removed.
  // processHeadings will now directly set data-numbering-id and data-numbering-level
  // attributes, and CSS will handle the display of the number.
  // We can keep metadata attributes setting here if desired, or move to processHeadings.
  try {
    heading.setAttribute("data-number", numbering.rawNumber.toString());
    // data-numbering-id and data-numbering-level should be set in processHeadings
    // heading.setAttribute('data-numbering-id', context.numberingId || '');
    // heading.setAttribute('data-numbering-level', context.numberingLevel?.toString() || '');
    heading.setAttribute("data-format", numbering.format || "");
  } catch (error) {
    console.error(
      "Error applying numbering data attributes to heading:",
      error.message
    );
  }
}

/**
 * Ensure heading has proper accessibility attributes
 * @param {HTMLElement} heading - The heading element
 */
function ensureHeadingAccessibility(heading) {
  try {
    if (!heading.id) {
      const cleanText =
        (heading.textContent || "")
          .replace(/[^\w\s-]/g, "") // Allow hyphens
          .replace(/\s+/g, "-")
          .toLowerCase()
          .substring(0, 50) ||
        `gen-heading-${Math.random().toString(36).substring(2, 7)}`;
      heading.id = "heading-" + cleanText;
    }

    heading.classList.add("keyboard-focusable");

    if (!heading.hasAttribute("tabindex")) {
      heading.setAttribute("tabindex", "0");
    }
  } catch (error) {
    console.error("Error ensuring heading accessibility:", error.message);
  }
}

/**
 * Enhanced TOC processing for better structure and styling
 * Creates a properly structured Table of Contents with leader lines and page numbers
 * @param {Document} doc - DOM document
 * @param {Object} styleInfo - Style information
 */
function processTOC(doc, styleInfo) {
  try {
    let tocEntries = doc.querySelectorAll('.docx-toc-entry, p[class*="toc"]');

    if (tocEntries.length === 0) {
      const paragraphs = doc.querySelectorAll("p");
      let tocCandidates = [];
      for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        const text = p.textContent.trim();
        const pageNumMatch = text.match(/^(.*?)[\.\s]+(\d+)$/);
        if (pageNumMatch) {
          p.classList.add("docx-toc-entry");
          tocCandidates.push(p);
        }
      }
      if (tocCandidates.length > 0) {
        tocEntries = tocCandidates;
      }
    }

    if (tocEntries.length > 0) {
      createTOCContainer(doc, tocEntries, styleInfo);
    }
  } catch (error) {
    console.error("Error processing TOC:", error.message);
  }
}

function createTOCContainer(doc, tocEntries, styleInfo) {
  try {
    let tocContainer = doc.querySelector(".docx-toc");
    if (!tocContainer) {
      tocContainer = doc.createElement("nav");
      tocContainer.className = "docx-toc";
      tocContainer.setAttribute("role", "navigation");
      tocContainer.setAttribute("aria-label", "Table of Contents");

      const tocHeading = doc.querySelector(".docx-toc-heading");
      if (!tocHeading) {
        const heading = doc.createElement("h2");
        heading.className = "docx-toc-heading";
        heading.textContent = "Table of Contents";
        tocContainer.appendChild(heading);
      } else {
        tocContainer.appendChild(tocHeading.cloneNode(true));
        if (tocHeading.parentNode) {
          tocHeading.parentNode.removeChild(tocHeading);
        }
      }
      const firstEntry = tocEntries[0];
      if (firstEntry.parentNode) {
        firstEntry.parentNode.insertBefore(tocContainer, firstEntry);
      } else {
        // If firstEntry has no parent, append tocContainer to body or a default element.
        // This case should ideally not happen if HTML is well-formed from Mammoth.
        doc.body.insertBefore(tocContainer, doc.body.firstChild);
      }
    }

    Array.from(tocEntries).forEach((entry, index) => {
      processTOCEntry(entry, index, tocContainer, styleInfo);
    });
  } catch (error) {
    console.error("Error creating TOC container:", error.message);
  }
}

function processTOCEntry(entry, index, tocContainer, styleInfo) {
  try {
    if (entry.classList.contains("processed-toc-entry")) {
      return;
    }
    entry.classList.add("processed-toc-entry");

    let level = determineTOCLevel(entry, styleInfo);
    const levelClass = `docx-toc-level-${level}`;
    if (!entry.classList.contains(levelClass)) {
      entry.classList.add(levelClass);
    }

    const text = entry.textContent.trim();
    structureTOCEntry(entry, text, level);

    if (entry.parentNode !== tocContainer) {
      const clonedEntry = entry.cloneNode(true); // Clone to avoid issues if entry was already moved
      tocContainer.appendChild(clonedEntry);
      if (entry.parentNode) {
        // Check if entry still has a parent before removing
        entry.parentNode.removeChild(entry);
      }
    }

    entry.setAttribute("tabindex", "0");
    entry.setAttribute("role", "link");
  } catch (error) {
    console.error("Error processing TOC entry:", error.message);
  }
}

function determineTOCLevel(entry, styleInfo) {
  try {
    let level = 1;
    const classArray = Array.from(entry.classList);
    for (const cls of classArray) {
      if (cls.includes("toc") && cls.match(/\d+$/)) {
        level = parseInt(cls.match(/\d+$/)[0], 10);
        break;
      }
    }
    if (
      level === 1 &&
      styleInfo.tocStyles &&
      styleInfo.tocStyles.tocEntryStyles
    ) {
      const entryStyleInfo = styleInfo.tocStyles.tocEntryStyles.find((s) =>
        entry.classList.contains(
          `docx-p-${s.id.toLowerCase().replace(/[^a-z0-9]/g, "-")}`
        )
      );
      if (
        entryStyleInfo &&
        entryStyleInfo.indentation &&
        entryStyleInfo.indentation.left
      ) {
        const leftIndentPt =
          typeof entryStyleInfo.indentation.left === "string"
            ? parseFloat(entryStyleInfo.indentation.left)
            : entryStyleInfo.indentation.left;
        if (leftIndentPt > 0) {
          level = Math.floor(leftIndentPt / 18) + 1; // Assuming ~18pt per level, adjust as needed
        }
      } else if (entryStyleInfo && entryStyleInfo.level) {
        level = entryStyleInfo.level;
      }
    }
    return Math.max(1, Math.min(6, level));
  } catch (error) {
    console.error("Error determining TOC level:", error.message);
    return 1;
  }
}

function structureTOCEntry(entry, text, level) {
  try {
    const pageNumMatch = text.match(/^(.*?)[\.\s]*(\d+)$/);
    const doc = entry.ownerDocument;

    entry.textContent = ""; // Clear existing content before restructuring

    if (pageNumMatch) {
      const textSpan = doc.createElement("span");
      textSpan.className = "docx-toc-text";
      textSpan.textContent = pageNumMatch[1].trim();
      entry.appendChild(textSpan);

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
      const textSpan = doc.createElement("span");
      textSpan.className = "docx-toc-text";
      textSpan.textContent = text;
      entry.appendChild(textSpan);
    }
  } catch (error) {
    console.error("Error structuring TOC entry:", error.message);
  }
}

function processNestedNumberedParagraphs(
  doc,
  styleInfo,
  numberingContext = null
) {
  try {
    if (!numberingContext) {
      console.warn(
        "No numbering context provided for numbered paragraph processing"
      );
      return;
    }
    processAllParagraphsForNumbering(doc, numberingContext, styleInfo);
    const numberingGroups = groupNumberingContexts(numberingContext);
    Object.values(numberingGroups).forEach((group) => {
      processNumberingGroup(doc, group, styleInfo);
    });
    processIndentedParagraphs(doc, styleInfo, numberingContext);
  } catch (error) {
    console.error("Error processing numbered paragraphs:", error.message);
  }
}

function processAllParagraphsForNumbering(doc, numberingContext, styleInfo) {
  try {
    const paragraphs = doc.querySelectorAll("p");
    paragraphs.forEach((p) => {
      const pText = p.textContent.trim();
      const context = numberingContext.find((ctx) => {
        if (!ctx.textContent) return false;
        const ctxText = ctx.textContent.trim();
        return (
          ctxText === pText ||
          ctxText.includes(pText) ||
          pText.includes(ctxText)
        );
      });
      if (context && context.numberingId) {
        p.setAttribute("data-num-id", context.numberingId);
        p.setAttribute(
          "data-num-level",
          context.numberingLevel?.toString() || "0"
        );
        if (context.resolvedNumbering) {
          p.setAttribute(
            "data-number",
            context.resolvedNumbering.rawNumber.toString()
          );
          p.setAttribute("data-format", context.resolvedNumbering.format);
        }
      }
    });
  } catch (error) {
    console.error(
      "Error processing paragraphs for numbering attributes:",
      error.message
    );
  }
}

function processIndentedParagraphs(doc, styleInfo, numberingContext) {
  try {
    const paragraphs = doc.querySelectorAll("p");
    let lastNumberedLevel = -1;
    let lastNumberedId = null;
    paragraphs.forEach((p) => {
      if (p.classList.contains("processed-list-item")) return;
      const numId = p.getAttribute("data-num-id");
      const numLevelAttr = p.getAttribute("data-num-level");

      if (numId && numLevelAttr !== null) {
        lastNumberedId = numId;
        lastNumberedLevel = parseInt(numLevelAttr, 10);
      } else if (lastNumberedId !== null) {
        const pText = p.textContent.trim().toLowerCase();
        const continuationPatterns = [
          "whereas",
          "resolved",
          "therefore",
          "provided",
          "however",
        ];
        if (continuationPatterns.some((pattern) => pText.startsWith(pattern))) {
          p.classList.add("docx-continues-numbering");
          p.classList.add(`docx-indent-level-${lastNumberedLevel + 1}`);
          p.setAttribute("data-follows-numbered", "true");
        } else {
          // Reset if it's not a continuation pattern
          // lastNumberedId = null;
          // lastNumberedLevel = -1;
        }
      }
    });
  } catch (error) {
    console.error("Error processing indented paragraphs:", error.message);
  }
}

function groupNumberingContexts(numberingContext) {
  const groups = {};
  numberingContext.forEach((context) => {
    if (context.numberingId && context.resolvedNumbering) {
      if (!groups[context.numberingId]) {
        groups[context.numberingId] = [];
      }
      groups[context.numberingId].push(context);
    }
  });
  return groups;
}

function processNumberingGroup(doc, group, styleInfo) {
  try {
    group.sort((a, b) => a.paragraphIndex - b.paragraphIndex);
    let currentList = null;
    let listItemsByLevel = {};
    let listContainerStack = [];

    group.forEach((context) => {
      const paragraph = findParagraphByContext(doc, context);
      if (!paragraph || paragraph.classList.contains("processed-list-item"))
        return;

      const level = context.numberingLevel;
      let listItem = createListItem(doc, paragraph, context);

      // Determine the correct parent list for this item
      while (
        listContainerStack.length > 0 &&
        level <
          parseInt(
            listContainerStack[listContainerStack.length - 1].getAttribute(
              "data-level"
            ) || "0",
            10
          )
      ) {
        listContainerStack.pop(); // Pop higher or equal level lists
      }

      let parentList;
      if (
        listContainerStack.length > 0 &&
        level ===
          parseInt(
            listContainerStack[listContainerStack.length - 1].getAttribute(
              "data-level"
            ) || "0",
            10
          )
      ) {
        // Same level, use current list in stack
        parentList = listContainerStack[listContainerStack.length - 1];
      } else if (
        listContainerStack.length > 0 &&
        level >
          parseInt(
            listContainerStack[listContainerStack.length - 1].getAttribute(
              "data-level"
            ) || "0",
            10
          )
      ) {
        // Deeper level, create nested list under last item of parent list
        let parentListItem =
          listContainerStack[listContainerStack.length - 1].lastChild;
        if (!parentListItem || parentListItem.nodeName.toLowerCase() !== "li") {
          // Ensure last child is an LI
          parentListItem = doc.createElement("li"); // Create a dummy li if needed
          listContainerStack[listContainerStack.length - 1].appendChild(
            parentListItem
          );
        }
        parentList = parentListItem.querySelector("ol, ul");
        if (!parentList) {
          parentList = createNestedListElement(doc, context);
          parentListItem.appendChild(parentList);
          listContainerStack.push(parentList);
        }
      } else {
        // level == 0 or stack is empty
        parentList = createListContainer(doc, paragraph, context); // Top-level list
        listContainerStack = [parentList]; // Reset stack with this new top-level list
      }

      parentList.appendChild(listItem);
      listItemsByLevel[level] = listItem; // Track last item at this level

      paragraph.classList.add("processed-list-item");
      if (paragraph.parentNode) {
        paragraph.parentNode.removeChild(paragraph);
      }
    });
  } catch (error) {
    console.error("Error processing numbering group:", error.message);
  }
}

function createNestedListElement(doc, context) {
  const list = doc.createElement("ol"); // Or 'ul' based on context.resolvedNumbering.format
  list.className = getListClassName(context);
  list.setAttribute("data-numbering-id", context.numberingId);
  list.setAttribute("data-level", context.numberingLevel.toString());
  return list;
}

function findParagraphByContext(doc, context) {
  try {
    const matchingParagraphs = Array.from(doc.querySelectorAll("p")).filter(
      (p) => {
        const pNumId = p.getAttribute("data-num-id");
        const pNumLevel = p.getAttribute("data-num-level");
        const pText = p.textContent.trim();
        const ctxText = context.textContent.trim();

        const idMatch = pNumId === context.numberingId;
        const levelMatch =
          pNumLevel === (context.numberingLevel?.toString() || "0");
        const textMatch =
          pText === ctxText ||
          pText.includes(ctxText) ||
          ctxText.includes(pText);

        return (
          idMatch &&
          levelMatch &&
          textMatch &&
          !p.classList.contains("processed-list-item")
        );
      }
    );
    return matchingParagraphs[0] || null; // Return the first match
  } catch (error) {
    console.error("Error finding paragraph by context:", error.message);
    return null;
  }
}

function createListContainer(doc, referenceParagraph, context) {
  try {
    const list = doc.createElement("ol");
    list.className = getListClassName(context);
    list.setAttribute("data-numbering-id", context.numberingId);
    list.setAttribute("data-level", "0"); // Top-level lists are effectively level 0 container for level 0 items

    if (referenceParagraph.parentNode) {
      referenceParagraph.parentNode.insertBefore(list, referenceParagraph);
    } else {
      doc.body.appendChild(list); // Fallback if no parent
    }
    return list;
  } catch (error) {
    console.error("Error creating list container:", error.message);
    const fallbackList = doc.createElement("ol"); // Basic fallback
    doc.body.appendChild(fallbackList);
    return fallbackList;
  }
}

function createListItem(doc, paragraph, context) {
  try {
    const listItem = doc.createElement("li");

    // Transfer inner HTML (including links, formatting) instead of just textContent
    listItem.innerHTML = paragraph.innerHTML;

    // Remove any prepended number by finding the first text node and stripping number pattern
    const firstTextNode = Array.from(listItem.childNodes).find(
      (node) => node.nodeType === Node.TEXT_NODE && node.nodeValue.trim() !== ""
    );
    if (firstTextNode) {
      let textContent = firstTextNode.nodeValue;
      const originalLength = textContent.length;

      // More specific regex to only remove at the beginning of the text node
      textContent = textContent.replace(
        /^\s*(\d+\.|[a-zA-Z]\.|[ivxlcdmIVXLCDM]+\.)\s*/,
        ""
      );

      if (textContent.length < originalLength) {
        firstTextNode.nodeValue = textContent;
      }
    }

    paragraph.classList.forEach((cls) => {
      if (!cls.startsWith("docx-p-") && cls !== "processed-list-item") {
        // Avoid adding paragraph-specific class
        listItem.classList.add(cls);
      }
    });

    listItem.setAttribute("data-numbering-id", context.numberingId);
    listItem.setAttribute(
      "data-numbering-level",
      context.numberingLevel.toString()
    );
    listItem.setAttribute(
      "data-number",
      context.resolvedNumbering.rawNumber.toString()
    );
    listItem.setAttribute("data-format", context.resolvedNumbering.format);
    // data-level is redundant if data-numbering-level is set, but can keep for CSS consistency
    listItem.setAttribute("data-level", context.numberingLevel.toString());

    if (context.resolvedNumbering.fullNumbering) {
      // This prefix is used by CSS ::before typically, but can store for reference
      listItem.setAttribute(
        "data-prefix",
        context.resolvedNumbering.fullNumbering
      );
    }

    return listItem;
  } catch (error) {
    console.error("Error creating list item:", error.message);
    const fallbackItem = doc.createElement("li");
    fallbackItem.textContent = paragraph.textContent.trim(); // Basic fallback
    return fallbackItem;
  }
}

function getListClassName(context) {
  if (!context.resolvedNumbering || !context.resolvedNumbering.format) {
    return "docx-numbered-list"; // Default
  }
  const format = context.resolvedNumbering.format;
  switch (format) {
    case "lowerLetter":
    case "upperLetter":
      return "docx-alpha-list";
    case "lowerRoman":
    case "upperRoman":
      return "docx-roman-list";
    case "bullet":
      return "docx-bulleted-list";
    default:
      return "docx-numbered-list";
  }
}

function processSpecialParagraphs(doc, styleInfo) {
  try {
    const specialPatterns =
      styleInfo.documentStructure?.specialParagraphPatterns || [];
    const paragraphs = doc.querySelectorAll("p");
    Array.from(paragraphs).forEach((para) => {
      if (
        para.classList.contains("processed-list-item") ||
        para.classList.contains("processed-special-para") ||
        para.classList.contains("docx-toc-entry")
      ) {
        return;
      }
      const text = para.textContent.trim();
      specialPatterns.forEach((pattern) => {
        if (pattern.type && isStructuralMatch(text, pattern)) {
          para.classList.add(`docx-${pattern.type.toLowerCase()}`); // Ensure class is lowercase
          para.classList.add("docx-structural-pattern");
        }
      });
      para.classList.add("processed-special-para");
    });
  } catch (error) {
    console.error("Error processing special paragraphs:", error.message);
  }
}

function isStructuralMatch(text, pattern) {
  if (!pattern.examples || !pattern.examples.length) return false;
  return pattern.examples.some((example) => {
    const examplePattern = extractStructuralPattern(example);
    const textPattern = extractStructuralPattern(text);
    return examplePattern === textPattern && examplePattern !== "unknown";
  });
}

function extractStructuralPattern(text) {
  if (/^\w+:\s/.test(text)) return "word_colon";
  if (/^\w+,\s/.test(text)) return "word_comma";
  if (/^\w+\s+\([^)]+\)/.test(text)) return "word_parenthesis";
  return "unknown";
}

function identifyListPatterns(paragraphs) {
  return {
    numberedItems: [],
    alphaItems: [],
    romanItems: [],
    specialParagraphs: [],
  };
}
function isSpecialParagraph(text, specialPatterns) {
  const pattern = extractStructuralPattern(text);
  if (pattern !== "unknown") {
    return { type: pattern, prefix: text.match(/^(\w+)/)?.[1] || "" };
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
  findParagraphContextByPartialText,
};
