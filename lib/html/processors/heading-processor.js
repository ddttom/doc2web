// File: lib/html/processors/heading-processor.js
// Heading processing and accessibility utilities

/**
 * Process headings in the document
 */
function processHeadings(doc, styleInfo, numberingContext = null) {
  try {
    const headings = doc.querySelectorAll("h1, h2, h3, h4, h5, h6");
    if (!numberingContext) {
      console.warn("No numbering context for headings.");
      headings.forEach((h) => {
        if (!h.classList.contains("docx-toc-heading")) {
          ensureHeadingAccessibility(h);
          // Apply paragraph style class if available
          applyParagraphStyleToHeading(h, styleInfo);
        }
      });
      return;
    }
    headings.forEach((heading) => {
      // Skip TOC headings - don't modify them
      if (heading.classList.contains("docx-toc-heading")) return;

      // Apply paragraph style class if available
      applyParagraphStyleToHeading(heading, styleInfo);

      const headingText = heading.textContent.trim();
      const context =
        findParagraphContextByText(headingText, numberingContext) ||
        findParagraphContextByPartialText(headingText, numberingContext);

      if (context && context.resolvedNumbering && context.abstractNumId) {
        // Apply existing data attributes
        heading.setAttribute("data-num-id", context.numberingId || "");
        heading.setAttribute("data-abstract-num", context.abstractNumId); // Essential for CSS targeting
        heading.setAttribute(
          "data-num-level",
          context.numberingLevel?.toString() || ""
        );

        // Add format attribute for CSS targeting
        const format = context.resolvedNumbering.format || "";
        heading.setAttribute("data-format", format);

        // Add box model fixes to prevent character overlap
        heading.style.overflowWrap = "break-word";
        heading.style.wordWrap = "break-word";

        // Special handling for Roman numerals
        if (format === "upperRoman" || format === "lowerRoman") {
          heading.classList.add("roman-numeral-heading");

          // Add extra padding for Roman numerals
          const currentPadding = heading.style.paddingLeft || "0";
          const paddingValue = parseInt(currentPadding) + 8;
          heading.style.paddingLeft = `${paddingValue}px`;
        }

        // Apply section ID from resolved numbering
        if (context.resolvedNumbering.sectionId) {
          heading.id = context.resolvedNumbering.sectionId;
          heading.setAttribute(
            "data-section-id",
            context.resolvedNumbering.sectionId
          );

          // Create a single text node with numbering and content together
          // This prevents line breaks between the number and heading text
          const originalContent = heading.textContent.trim();
          const numberingText = context.resolvedNumbering.fullNumbering || "";

          // Clear the heading and set the combined content
          heading.innerHTML = "";

          // Add the numbering as a non-breaking span followed immediately by content
          if (numberingText) {
            const wrapper = doc.createElement("span");
            wrapper.className = "heading-number";
            wrapper.setAttribute("aria-hidden", "true");
            wrapper.textContent = numberingText;
            wrapper.style.whiteSpace = "nowrap";
            heading.appendChild(wrapper);

            // Add a non-breaking space to ensure proper spacing
            heading.appendChild(doc.createTextNode("\u00A0")); // Non-breaking space
          }

          // Add the content directly as text with normal wrapping
          const contentSpan = doc.createElement("span");
          contentSpan.className = "heading-content";
          contentSpan.textContent = originalContent;
          heading.appendChild(contentSpan);
        }
      }

      ensureHeadingAccessibility(heading);
    });
  } catch (error) {
    console.error("Error processing headings:", error.message, error.stack);
  }
}

/**
 * Apply paragraph style class to heading if it has associated style with indentation
 * @param {Element} heading - Heading element
 * @param {Object} styleInfo - Style information
 */
function applyParagraphStyleToHeading(heading, styleInfo) {
  try {
    // Check if heading already has a paragraph style class
    const existingClasses = Array.from(heading.classList);
    let hasParagraphStyle = false;

    existingClasses.forEach((className) => {
      if (className.startsWith("docx-p-")) {
        // Check if this style has indentation
        const styleId = className.replace("docx-p-", "").replace(/-/g, "");
        Object.entries(styleInfo.styles?.paragraph || {}).forEach(
          ([id, style]) => {
            const safeId = id.toLowerCase().replace(/[^a-z0-9-_]/g, "-");
            if (safeId === styleId && style.indentation) {
              hasParagraphStyle = true;
            }
          }
        );
      }
    });

    // If no paragraph style with indentation, check if the heading level has a default style with indentation
    if (!hasParagraphStyle) {
      const headingLevel = heading.tagName.toLowerCase().charAt(1);
      Object.entries(styleInfo.styles?.paragraph || {}).forEach(
        ([id, style]) => {
          if (
            style.isHeading &&
            style.indentation &&
            (style.indentation.hanging ||
              style.indentation.firstLine ||
              style.indentation.left) &&
            (style.name?.toLowerCase().includes(`heading ${headingLevel}`) ||
              style.name?.toLowerCase().includes(`heading${headingLevel}`))
          ) {
            const safeClassName = `docx-p-${id
              .toLowerCase()
              .replace(/[^a-z0-9-_]/g, "-")}`;
            heading.classList.add(safeClassName);
          }
        }
      );
    }
  } catch (error) {
    console.error("Error applying paragraph style to heading:", error.message);
  }
}

/**
 * Find paragraph context by exact text match
 */
function findParagraphContextByText(text, numberingContext) {
  return (
    numberingContext.find((context) => {
      if (!context.textContent) return false;
      if (context.textContent.trim() === text) return true;
      const cleanContextText = context.textContent
        .trim()
        .replace(/[:.\s]*$/, "");
      const cleanSearchText = text.replace(/[:.\s]*$/, "");
      return cleanContextText === cleanSearchText;
    }) || null
  );
}

/**
 * Find paragraph context by partial text match
 */
function findParagraphContextByPartialText(text, numberingContext) {
  const cleanSearchText = text
    .toLowerCase()
    .trim()
    .replace(/[:.\s]*$/, "");
  return (
    numberingContext.find((context) => {
      if (!context.textContent) return false;
      const cleanContextText = context.textContent
        .toLowerCase()
        .trim()
        .replace(/[:.\s]*$/, "");
      return (
        cleanContextText.includes(cleanSearchText) ||
        cleanSearchText.includes(cleanContextText)
      );
    }) || null
  );
}

/**
 * Ensure heading accessibility
 */
function ensureHeadingAccessibility(heading) {
  try {
    // Only generate ID if one doesn't already exist (from section numbering)
    if (!heading.id) {
      const cleanText =
        (
          heading.textContent ||
          `gen-h-${Math.random().toString(36).substring(2, 7)}`
        )
          .replace(/[^\w\s-]/g, "")
          .replace(/\s+/g, "-")
          .toLowerCase()
          .substring(0, 50) ||
        `gen-heading-${Math.random().toString(36).substring(2, 7)}`;
      heading.id = "heading-" + cleanText;
    }
    
    heading.classList.add("keyboard-focusable");
    if (!heading.hasAttribute("tabindex"))
      heading.setAttribute("tabindex", "0");
  } catch (error) {
    console.error("Error ensuring heading accessibility:", error.message);
  }
}

module.exports = {
  processHeadings,
  applyParagraphStyleToHeading,
  findParagraphContextByText,
  findParagraphContextByPartialText,
  ensureHeadingAccessibility,
};