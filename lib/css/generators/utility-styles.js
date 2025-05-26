// File: lib/css/generators/utility-styles.js
// Utility and general purpose styles

const { convertTwipToPt } = require("../../utils/unit-converter");

/**
 * Generate utility styles
 */
function generateUtilityStyles(styleInfo) {
  // Calculate default paragraph margin based on DOCX introspection
  let defaultParagraphMargin = "0pt";
  
  // Method 1: Check document defaults for paragraph indentation
  if (styleInfo.documentDefaults?.paragraph?.indentation?.left) {
    defaultParagraphMargin = convertTwipToPt(styleInfo.documentDefaults.paragraph.indentation.left) + "pt";
  }
  // Method 2: Check the most common paragraph style indentation
  else if (styleInfo.styles?.paragraph) {
    const paragraphStyles = Object.values(styleInfo.styles.paragraph);
    const indentations = paragraphStyles
      .map(style => style.indentation?.left)
      .filter(indent => indent && parseInt(indent) > 0)
      .map(indent => convertTwipToPt(indent));
    
    if (indentations.length > 0) {
      // Use the most common indentation, or the first one if all are different
      const indentCounts = {};
      indentations.forEach(indent => {
        indentCounts[indent] = (indentCounts[indent] || 0) + 1;
      });
      
      const mostCommonIndent = Object.keys(indentCounts)
        .reduce((a, b) => indentCounts[a] > indentCounts[b] ? a : b);
      
      defaultParagraphMargin = mostCommonIndent + "pt";
    }
  }
  // Method 3: Check if numbered paragraphs have consistent indentation
  if (defaultParagraphMargin === "0pt" && styleInfo.numberingDefs?.abstractNums) {
    const abstractNums = Object.values(styleInfo.numberingDefs.abstractNums);
    const levelIndentations = [];
    
    abstractNums.forEach(abstractNum => {
      if (abstractNum.levels) {
        Object.values(abstractNum.levels).forEach(level => {
          if (level.indentation?.left) {
            levelIndentations.push(convertTwipToPt(level.indentation.left));
          }
        });
      }
    });
    
    if (levelIndentations.length > 0) {
      // Use the minimum indentation from numbered lists as the base paragraph margin
      const minIndent = Math.min(...levelIndentations);
      if (minIndent > 0) {
        defaultParagraphMargin = minIndent + "pt";
      }
    }
  }
  
  // Fallback: Use a reasonable default based on typical Word documents
  if (defaultParagraphMargin === "0pt") {
    defaultParagraphMargin = "0pt"; // Keep 0 if we can't determine from DOCX
  }

  return `
/* Utility Styles */
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }

/* Section navigation styling */
[id^="section-"]:target {
  background-color: #fffbcc;
  transition: background-color 0.3s ease;
}

/* Improved spacing for section IDs */
[id^="section-"] {
  margin-top: 0.5em;
  clear: both;
}

/* Fix for Roman numeral sections */
[id^="section-v"], 
[id^="section-vi"],
[id^="section-vii"],
[id^="section-viii"],
[id^="section-ix"],
[id^="section-x"] {
  padding-left: 0.5em;
  margin-left: 0.5em;
}

/* Enhanced styling for Roman numeral sections */
.roman-numeral-section,
.roman-numeral-heading,
[data-format="upperRoman"],
[data-format="lowerRoman"] {
  padding-left: 0.5em !important;
  margin-left: 0.5em !important;
  position: relative;
}

/* Ensure proper spacing for Roman numeral headings */
.roman-numeral-heading::before,
[data-format="upperRoman"]::before,
[data-format="lowerRoman"]::before {
  margin-right: 0.5em !important;
  white-space: nowrap !important;
}

/* Add spacing between heading number and content */
.heading-number {
  white-space: nowrap;
  display: inline;
  margin-right: 0;
}

/* Ensure heading content can wrap but stays connected to numbering */
.heading-content {
  display: inline;
  white-space: normal;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

/* Ensure headings allow full text display with proper wrapping */
h1:has(.heading-number), h2:has(.heading-number), h3:has(.heading-number),
h4:has(.heading-number), h5:has(.heading-number), h6:has(.heading-number) {
  white-space: normal;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

/* Prevent line break specifically between number and content start */
.heading-number + .heading-content {
  margin-left: 0;
}

/* Ensure proper spacing between numbering and content */
[data-num-id][data-abstract-num] {
  overflow-wrap: break-word;
  word-wrap: break-word;
  word-break: normal;
  hyphens: auto;
}

/* Default paragraph styling with DOCX-derived margins */
p {
  overflow-wrap: break-word;
  word-wrap: break-word;
  max-width: 100%;
  box-sizing: border-box;
  /* Apply consistent left margin to all paragraphs based on DOCX introspection */
  margin-left: ${defaultParagraphMargin};
}

/* Override for paragraphs that already have specific styling */
p[class*="docx-p-"] {
  /* Styled paragraphs use their own margin-left from generateParagraphStyles */
  margin-left: unset;
}

/* Override for TOC entries and other special paragraphs */
p.docx-toc-entry,
p[class*="docx-toc-"],
nav.docx-toc p {
  /* TOC entries have their own indentation logic */
  margin-left: 0;
}

/* Apply default page margins to TOC link text */
.docx-toc-text a {
  /* Inherit the default paragraph margin for consistent spacing */
  margin-left: ${defaultParagraphMargin};
  display: inline-block;
}

h1,h2,h3,h4,h5,h6 { 
  font-family:"${styleInfo.theme?.fonts?.major || "Calibri Light"}", Arial, sans-serif; 
  color: ${styleInfo.theme?.colors?.dk1 || "#2E74B5"}; 
  overflow-wrap: break-word;
  word-wrap: break-word;
  /* Apply default paragraph margin to all headings */
  margin-left: ${defaultParagraphMargin};
  /* Restore default browser margins for headings */
  margin-top: 0.83em;
  margin-bottom: 0.83em;
}

/* Specific margins for each heading level to match browser defaults */
h1 { 
  margin-top: 0.67em; 
  margin-bottom: 0.67em; 
  font-size: 2em; 
}
h2 { 
  margin-top: 0.83em; 
  margin-bottom: 0.83em; 
  font-size: 1.5em; 
}
h3 { 
  margin-top: 1em; 
  margin-bottom: 1em; 
  font-size: 1.17em; 
}
h4 { 
  margin-top: 1.33em; 
  margin-bottom: 1.33em; 
  font-size: 1em; 
}
h5 { 
  margin-top: 1.67em; 
  margin-bottom: 1.67em; 
  font-size: 0.83em; 
}
h6 { 
  margin-top: 2.33em; 
  margin-bottom: 2.33em; 
  font-size: 0.67em; 
}
/* ... more utility styles ... */
`;
}

module.exports = {
  generateUtilityStyles,
};