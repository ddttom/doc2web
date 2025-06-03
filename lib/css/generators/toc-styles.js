// File: lib/css/generators/toc-styles.js
// Table of Contents style generation

const { convertTwipToPt } = require("../../utils/unit-converter");

/**
 * Determine TOC level based on entry position and indentation
 */
function determineTOCLevel(entry, index) {
  // Simple level determination based on indentation
  // Could be enhanced with more sophisticated logic
  const indentPt = convertTwipToPt(entry.leftIndent || 0);
  if (indentPt === 0) return 1;
  if (indentPt <= 20) return 1;
  if (indentPt <= 40) return 2;
  if (indentPt <= 60) return 3;
  return Math.min(Math.floor(indentPt / 20) + 1, 6); // Max 6 levels
}

/**
 * Generate enhanced TOC level styles based on actual DOCX formatting
 */
function generateEnhancedTOCLevelStyles(enhancedTOCEntries, styleInfo) {
  if (!enhancedTOCEntries || enhancedTOCEntries.length === 0) {
    return '';
  }

  let css = '\n/* Enhanced TOC Level Styles - Based on DOCX Formatting */\n';
  
  // Group TOC entries by level and extract their actual formatting
  const levelStyles = new Map();
  
  enhancedTOCEntries.forEach((entry, index) => {
    const level = determineTOCLevel(entry, index);
    
    if (!levelStyles.has(level)) {
      levelStyles.set(level, {
        entries: [],
        commonFormatting: {}
      });
    }
    
    levelStyles.get(level).entries.push(entry);
  });
  
  // Generate CSS for each level based on actual DOCX formatting
  levelStyles.forEach((levelData, level) => {
    const entries = levelData.entries;
    
    // Analyze common formatting across entries at this level
    const hasItalic = entries.some(entry => entry.hasItalic);
    const hasBold = entries.some(entry => entry.hasBold);
    const avgLeftIndent = entries.reduce((sum, entry) => sum + (entry.leftIndent || 0), 0) / entries.length;
    const avgHangingIndent = entries.reduce((sum, entry) => sum + (entry.hangingIndent || 0), 0) / entries.length;
    
    // Convert DOCX indentation to CSS with precision
    const leftIndentPt = convertTwipToPt(avgLeftIndent);
    const hangingIndentPt = convertTwipToPt(avgHangingIndent);
    const paddingLeft = hangingIndentPt > 0 ? `${hangingIndentPt}pt` : `${(level * 1.5)}em`;
    const textIndent = hangingIndentPt > 0 ? `-${hangingIndentPt}pt` : `-${level * 1.5}em`;
    
    // Extract font information from first entry with font data
    const fontEntry = entries.find(entry => 
      entry.runProperties && entry.runProperties.some(rp => rp.formatting.font)
    );
    const fontSize = fontEntry?.runProperties?.[0]?.formatting?.fontSize || '10pt';
    const fontFamily = fontEntry?.runProperties?.[0]?.formatting?.font?.ascii || 'Calibri';
    
    css += `
.docx-toc-level-${level} {
  font-family: "${fontFamily}", sans-serif;
  font-size: ${fontSize};
  font-weight: ${hasBold ? 'bold' : 'normal'};
  font-style: ${hasItalic ? 'italic' : 'normal'}; /* CRITICAL: Preserve DOCX italic */
  margin-left: ${leftIndentPt}pt;
  padding-left: ${paddingLeft};
  text-indent: ${textIndent};
  /* Enhanced positioning for Word-like fidelity */
  position: relative;
  overflow-wrap: break-word;
  word-wrap: break-word;
  hyphens: auto;
}

/* Handle mixed formatting within TOC entries */
.docx-toc-level-${level} .toc-italic {
  font-style: italic;
}

.docx-toc-level-${level} .toc-bold {
  font-weight: bold;
}

/* Handle TOC entries that are numbered by the main numbering system */
.docx-toc-level-${level}[data-abstract-num] {
  padding-left: ${paddingLeft};
  text-indent: ${textIndent};
  margin-left: 0;
}
`;
  });
  
  return css;
}

/**
 * Generate TOC styles
 */
function generateTOCStyles(styleInfo) {
  const tocStyles = styleInfo.tocStyles || {};
  const leaderStyle = tocStyles.leaderStyle || {
    character: null, // No leader character for web-based TOC
    position: 0, // No position for web-based TOC
    spacesBetween: 0, // No spacing for web-based TOC
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

/* Structure TOC entries with hanging indent (page numbers removed for web format) */
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

/* TOC Navigation Links - Enhanced for clickable functionality */
.docx-toc-link {
  color: inherit !important;
  text-decoration: none !important;
  display: block !important;
  transition: background-color 0.2s ease, outline 0.2s ease;
  border-radius: 2px;
  padding: 2px 4px;
  margin: -2px -4px;
}

.docx-toc-link:hover,
.docx-toc-link:focus {
  background-color: #f0f8ff !important;
  outline: 2px solid #007bff !important;
  outline-offset: 2px;
}

.docx-toc-link:active {
  background-color: #e6f3ff !important;
}

/* For backward compatibility with existing clickable TOC entries */
.docx-toc-text a {
  text-decoration: none !important;
  color: inherit !important;
}

/* TOC dots and page numbers are not generated for web format */
.docx-toc-dots {
  display: none !important;
}

/* TOC page numbers are not generated for web format */
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

  // NEW: Use enhanced TOC level styles if available
  if (tocStyles.enhancedTOCEntries && tocStyles.enhancedTOCEntries.length > 0) {
    css += generateEnhancedTOCLevelStyles(tocStyles.enhancedTOCEntries, styleInfo);
  } else {
    // Fallback to original TOC level generation
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
  }

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

/* TOC Character Formatting - Critical for DOCX fidelity */
.toc-italic {
  font-style: italic !important;
}

.toc-bold {
  font-weight: bold !important;
}

/* Ensure TOC formatting doesn't interfere with layout */
.docx-toc-entry .toc-italic,
.docx-toc-entry .toc-bold {
  display: inline;
  margin: 0;
  padding: 0;
}
`;

  return css;
}

module.exports = {
  generateTOCStyles,
  generateEnhancedTOCLevelStyles,
};
