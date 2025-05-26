// File: lib/css/generators/toc-styles.js
// Table of Contents style generation

/**
 * Generate TOC styles
 */
function generateTOCStyles(styleInfo) {
  const tocStyles = styleInfo.tocStyles || {};
  const leaderStyle = tocStyles.leaderStyle || {
    character: ".",
    position: 468, // default 6.5 inches in points
    spacesBetween: 3,
  };

  let css = `
/* Table of Contents Styles */
.docx-toc {
  margin: 1em 0 2em 0;
  width: 100%;
  padding: 0;
  /* Force single-column layout */
  column-count: 1 !important;
  -webkit-column-count: 1 !important;
  -moz-column-count: 1 !important;
  columns: 1 !important; /* Additional property */
  column-width: auto !important; /* Prevent auto width calculation */
}

/* Ensure parent containers don't interfere */
.docx-toc, .docx-toc * {
  box-sizing: border-box;
}

.docx-toc-heading {
  font-family: "${
    tocStyles.tocHeadingStyle?.fontFamily ||
    styleInfo.theme?.fonts?.major ||
    "Calibri Light"
  }", sans-serif;
  font-size: ${tocStyles.tocHeadingStyle?.fontSize || "16pt"}; 
  font-weight: bold;
  margin-bottom: 1em;
  text-align: ${tocStyles.tocHeadingStyle?.alignment || "left"};
  column-span: all;
  -webkit-column-span: all;
  color: ${
    tocStyles.tocHeadingStyle?.color ||
    styleInfo.theme?.colors?.dk1 ||
    "#2F5496"
  };
}

/* Structure TOC entries as a flex container with proper alignment */
.docx-toc-entry {
  display: flex !important;
  flex-wrap: nowrap !important;
  align-items: baseline !important;
  width: 100% !important;
  max-width: 100%; /* Prevent overflow */
  margin-bottom: 0.5em !important;
  line-height: 1.4 !important;
  position: relative !important;
  overflow: hidden !important;
  page-break-inside: avoid; /* Keep entries together on page breaks */
}

/* TOC text part - flex-grow: 0 to prevent it from expanding */
.docx-toc-text {
  flex: 0 0 auto !important;
  white-space: nowrap !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  padding-right: 0.5em !important;
  max-width: 80% !important;
}

/* For clickable TOC entries */
.docx-toc-text a {
  text-decoration: none !important;
  color: inherit !important;
}

/* TOC dots - flex-grow: 1 to fill available space with dots */
.docx-toc-dots {
  flex: 1 1 auto !important;
  position: relative !important;
  height: 1.2em !important;
  margin: 0 0.5em !important;
  overflow: hidden !important;
}

/* Create dots using background image with proper spacing based on leaderStyle */
.docx-toc-dots::after {
  content: "";
  position: absolute;
  bottom: 0.5em; /* Position closer to text baseline */
  left: 0;
  right: 0;
  height: 1px; /* Use 1px for consistent rendering */
  background-image: radial-gradient(circle, currentColor 1px, transparent 1.5px);
  background-position: bottom;
  background-size: 6px 1px; /* Fixed size for consistency */
  background-repeat: repeat-x;
}

/* TOC page number - flex-grow: 0 to maintain size and right-align */
.docx-toc-pagenum {
  flex: 0 0 auto !important;
  text-align: right !important;
  white-space: nowrap !important;
  font-variant-numeric: tabular-nums !important;
}

/* Print-specific styles for TOC */
@media print {
  .docx-toc-dots::after {
    background-image: radial-gradient(circle, black 1px, transparent 1.5px);
  }
}

/* Force TOC entry paragraphs to use flex */
p.docx-toc-entry {
  display: flex !important;
  margin: 0.5em 0 !important; /* Override paragraph margins */
}

/* Browser Compatibility */
@supports not (background-image: radial-gradient(circle, currentColor 1px, transparent 1.5px)) {
  .docx-toc-dots::after {
    background-image: none;
    border-bottom: 1px dotted currentColor;
  }
}
`;

  // Add specific styles for each TOC level
  (tocStyles.tocEntryStyles || []).forEach((style, index) => {
    const level = style.level || index + 1;
    const baseIndent = (level - 1) * 1.5; // Base indent in em (e.g., 1.5em per level)
    const leftIndentPt = style.indentation?.left || 0; // Use points from DOCX if available

    css += `
.docx-toc-level-${level} {
  font-family: "${
    style.fontFamily || styleInfo.theme?.fonts?.minor || "Calibri"
  }", sans-serif;
  font-size: ${style.fontSize || "10pt"};
  font-weight: ${
    level === 1
      ? style.bold === false
        ? "normal"
        : "bold"
      : style.bold
      ? "bold"
      : "normal"
  };
  font-style: ${style.italic ? "italic" : "normal"};
  color: ${style.color || styleInfo.theme?.colors?.tx1 || "inherit"};
  padding-left: ${leftIndentPt > 0 ? leftIndentPt + "pt" : baseIndent + "em"};
}

/* Handle TOC entries that are numbered by the main numbering system */
.docx-toc-level-${level}[data-abstract-num] {
  padding-left: 0;
  margin-left: ${leftIndentPt > 0 ? leftIndentPt + "pt" : baseIndent + "em"};
}
`;
  });

  // Special handling for potential "Rationale" lines or similar entries
  // that might not have regular TOC styling but need consistent layout
  css += `
/* Special handling for TOC entries without explicit TOC styles */
.docx-toc-entry:not([class*="docx-toc-level-"]) {
  font-family: "${styleInfo.theme?.fonts?.minor || "Calibri"}", sans-serif;
  font-size: 10pt;
  padding-left: 0;
  margin-top: 0.5em;
  margin-bottom: 0.5em;
}

/* Force all TOC entries to use flex layout for consistent alignment */
p.docx-toc-entry {
  display: flex !important;
}

/* Ensure all TOC entries have dots and page numbers aligned consistently */
.docx-toc p, .docx-toc-entry {
  display: flex !important;
  flex-wrap: nowrap !important;
  align-items: baseline !important;
  width: 100% !important;
  margin-bottom: 0.5em !important;
  line-height: 1.4 !important;
  position: relative !important;
  overflow: hidden !important;
}

/* Special handling for TOC entries with numbering - adjust ::before positioning for flex layout */
.docx-toc-entry[data-abstract-num][data-num-level]::before {
  position: relative !important;
  left: 0 !important;
  margin-right: 0.5em !important;
  width: auto !important;
  display: inline !important;
}
`;

  return css;
}

module.exports = {
  generateTOCStyles,
};