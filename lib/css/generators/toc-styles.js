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

/* Structure TOC entries with hanging indent and proper page number alignment */
.docx-toc-entry {
  display: block !important;
  width: 100% !important;
  max-width: 100%; /* Prevent overflow */
  margin-bottom: 0.5em !important;
  line-height: 1.4 !important;
  position: relative !important;
  overflow: visible !important;
  page-break-inside: avoid; /* Keep entries together on page breaks */
  /* Hanging indent for TOC entries */
  text-indent: -1.5em !important;
  padding-left: 1.5em !important;
  overflow-wrap: break-word !important;
  word-wrap: break-word !important;
  hyphens: auto !important;
}

/* TOC text part - inline with proper wrapping */
.docx-toc-text {
  display: inline !important;
  white-space: normal !important;
  overflow-wrap: break-word !important;
  word-wrap: break-word !important;
  hyphens: auto !important;
}

/* For clickable TOC entries */
.docx-toc-text a {
  text-decoration: none !important;
  color: inherit !important;
}

/* TOC dots and page numbers are hidden */
.docx-toc-dots {
  display: none !important;
}

/* TOC page number - hidden */
.docx-toc-pagenum {
  display: none !important;
}



/* Force TOC entry paragraphs to use block layout with hanging indent */
p.docx-toc-entry {
  display: block !important;
  margin: 0.5em 0 !important; /* Override paragraph margins */
  text-indent: -1.5em !important;
  padding-left: 1.5em !important;
  overflow-wrap: break-word !important;
  word-wrap: break-word !important;
}


`;

  // Add specific styles for each TOC level
  (tocStyles.tocEntryStyles || []).forEach((style, index) => {
    const level = style.level || index + 1;
    const baseIndent = (level - 1) * 1.5; // Base indent in em (e.g., 1.5em per level)
    const leftIndentPt = style.indentation?.left || 0; // Use points from DOCX if available
    const totalIndent = leftIndentPt > 0 ? leftIndentPt + "pt" : (baseIndent + 1.5) + "em";
    const hangingIndent = leftIndentPt > 0 ? "-" + leftIndentPt + "pt" : "-1.5em";

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
  /* Use hanging indent with proper level indentation */
  padding-left: ${totalIndent};
  text-indent: ${hangingIndent};
  overflow-wrap: break-word;
  word-wrap: break-word;
  hyphens: auto;
}

/* Handle TOC entries that are numbered by the main numbering system */
.docx-toc-level-${level}[data-abstract-num] {
  padding-left: ${totalIndent};
  text-indent: ${hangingIndent};
  margin-left: 0;
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

/* Force all TOC entries to use block layout */
p.docx-toc-entry {
  display: block !important;
}

/* Ensure all TOC entries use block layout with hanging indent */
.docx-toc p, .docx-toc-entry {
  display: block !important;
  width: 100% !important;
  margin-bottom: 0.5em !important;
  line-height: 1.4 !important;
  position: relative !important;
  overflow: visible !important;
  text-indent: -1.5em !important;
  padding-left: 1.5em !important;
  overflow-wrap: break-word !important;
  word-wrap: break-word !important;
  hyphens: auto !important;
}

/* Special handling for TOC entries with numbering - adjust ::before positioning for hanging indent */
.docx-toc-entry[data-abstract-num][data-num-level]::before {
  position: static !important;
  left: auto !important;
  margin-right: 0.5em !important;
  width: auto !important;
  display: inline !important;
  /* Ensure numbering doesn't interfere with hanging indent */
  text-indent: 0 !important;
}
`;

  return css;
}

module.exports = {
  generateTOCStyles,
};
