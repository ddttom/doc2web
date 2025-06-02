// File: lib/css/generators/numbering-styles.js
// Numbering and list style generation

const { convertTwipToPt } = require("../../utils/unit-converter");
const {
  getCSSCounterContent,
  getCSSCounterFormat,
} = require("../../parsers/numbering-parser");

/**
 * Generate DOCX numbering styles
 */
function generateDOCXNumberingStyles(numberingDefs, styleInfo) {
  let css = `\n/* DOCX Numbering Styles */\n`;
  
  // Add global box model fixes for all numbered elements
  css += `
/* Global box model fixes for numbered elements */
[data-abstract-num][data-num-level] {
  box-sizing: border-box;
  overflow-wrap: break-word;
  word-break: normal;
  hyphens: auto;
}

/* Fix for section IDs with Roman numerals */
[id^="section-"][data-format="upperRoman"],
[id^="section-"][data-format="lowerRoman"] {
  padding-left: 2em; /* Extra space for Roman numerals */
}
`;
  
  // Process each abstract numbering definition
  Object.values(numberingDefs.abstractNums || {}).forEach((abstractNum) => {
    // Set up counter resets for all levels in this abstract numbering
    const allLevelCountersForAbstractNum = Object.keys(abstractNum.levels || {})
      .map((levelIndex) => `docx-counter-${abstractNum.id}-${levelIndex} 0`)
      .join(" ");
    
    if (allLevelCountersForAbstractNum) {
      css += `
/* Counter resets for abstract numbering ID ${abstractNum.id} */
body { 
  counter-reset: ${allLevelCountersForAbstractNum};
}

/* More targeted counter reset for specific containers with this abstract numbering */
ol[data-abstract-num="${abstractNum.id}"], 
ul[data-abstract-num="${abstractNum.id}"], 
div[data-abstract-num="${abstractNum.id}"],
nav[data-abstract-num="${abstractNum.id}"] {
  counter-reset: ${allLevelCountersForAbstractNum};
}
`;
    }

    // Process each level in the abstract numbering
    Object.entries(abstractNum.levels || {}).forEach(([levelIndexStr, levelDef]) => {
      const levelIndex = parseInt(levelIndexStr, 10);
      const itemSelector = `[data-abstract-num="${abstractNum.id}"][data-num-level="${levelIndex}"]`;
      
      // Calculate indentation and spacing based on DOCX definition with enhanced precision
      let pIndentLeft = 0;
      let pTextIndent = 0;
      let numRegionWidth = 24; // Base width for numbering region
      
      // Use consistent base margin approach for document fidelity
      // Instead of varying margin-left, use consistent base + level-based hanging indents
      
      // Calculate additional indentation for this level (beyond base document margin)
      let additionalIndent = levelIndex * 18; // 0.25 inch per level beyond base
      
      // Extract hanging indent from DOCX for precise formatting
      if (levelDef.indentation) {
        // Handle hanging indent (most common for numbered lists)
        if (levelDef.indentation.hanging !== undefined) {
          const hangingIndent = convertTwipToPt(levelDef.indentation.hanging);
          numRegionWidth = hangingIndent;
          pTextIndent = -hangingIndent; // Create true hanging indent
        } 
        // Handle first line indent
        else if (levelDef.indentation.firstLine !== undefined) {
          pTextIndent = convertTwipToPt(levelDef.indentation.firstLine);
          numRegionWidth = Math.max(24, Math.abs(pTextIndent)); // Ensure enough space
        }
        
        // Use additional indentation for nested levels, but keep base margin consistent
        if (levelDef.indentation.left !== undefined) {
          const totalIndent = convertTwipToPt(levelDef.indentation.left);
          // Calculate additional indent beyond the document's base margin
          // This will be applied as padding-left to maintain consistent base margins
          additionalIndent = Math.max(0, totalIndent - 36); // Assume 36pt base, adjust as needed
        }
      } else {
        // Default indentation based on level
        additionalIndent = levelIndex * 18; // 0.25 inch per level
        numRegionWidth = 24; // Default numbering width
      }
      
      // Set the left margin to 0 - let the base paragraph margin handle document consistency
      pIndentLeft = additionalIndent;
      
      // Adjust numbering width based on format
      if (levelDef.format === "upperRoman" || levelDef.format === "lowerRoman") {
        numRegionWidth = Math.max(numRegionWidth, 48); // Wider for Roman numerals
      } else if (levelDef.format === "bullet") {
        numRegionWidth = Math.max(numRegionWidth, 18); // Bullets need less space
      } else {
        numRegionWidth = Math.max(numRegionWidth, 36); // Numbers need moderate space
      }
      
      // Ensure we always have a proper hanging indent for numbered elements
      // If the hanging indent is too small (less than half the number region), use the full region width
      if (numRegionWidth > 0 && Math.abs(pTextIndent) < numRegionWidth / 2) {
        pTextIndent = -numRegionWidth; // Create hanging indent equal to number region width
      }
      
      // Create CSS for the numbered element with consistent base margin approach
      css += `
/* Styling for level ${levelIndex} of abstract numbering ${abstractNum.id} */
${itemSelector} {
  display: block;
  position: relative;
  margin-left: 0; /* Use base document margin for consistency */
  padding-left: ${pIndentLeft + numRegionWidth}pt; /* Additional level indent + numbering space */
  text-indent: ${pTextIndent}pt;
  counter-increment: docx-counter-${abstractNum.id}-${levelIndex};
  box-sizing: border-box;
  overflow-wrap: break-word;
  word-wrap: break-word;
  hyphens: auto;
  /* Match Word's paragraph spacing */
  margin-top: 0;
  margin-bottom: 0;
}

/* Section ID based styling for easy targeting */
${itemSelector}[id^="section-"] {
  scroll-margin-top: 20px; /* For smooth scrolling to sections */
}

/* Number or bullet display using ::before pseudo-element with precise positioning */
${itemSelector}::before {
  content: ${getCSSCounterContent(levelDef, abstractNum.id, numberingDefs)};
  position: absolute;
  left: ${pTextIndent !== 0 ? pTextIndent : -numRegionWidth}pt;
  top: 0;
  width: ${numRegionWidth}pt;
  text-align: ${levelDef.alignment || "left"};
  box-sizing: border-box;
  white-space: nowrap;
  overflow: hidden;
  /* Spacing after number based on DOCX suffix */
  padding-right: ${
    levelDef.suffix === "tab" ? "0.5em" : 
    levelDef.suffix === "space" ? "0.25em" : 
    "0.125em"
  };
  /* Apply DOCX formatting to numbering */
  ${levelDef.runProps?.bold ? "font-weight: bold;" : ""}
  ${levelDef.runProps?.italic ? "font-style: italic;" : ""}
  ${levelDef.runProps?.fontSize ? `font-size: ${levelDef.runProps.fontSize};` : ""}
  ${levelDef.runProps?.color ? `color: ${levelDef.runProps.color};` : ""}
  font-family: ${levelDef.runProps?.font?.ascii ? `"${levelDef.runProps.font.ascii}", ` : ""}${levelDef.runProps?.font?.hAnsi ? `"${levelDef.runProps.font.hAnsi}", ` : ""}sans-serif;
}
`;
      
      // Reset deeper level counters when this level increments
      const deeperLevelsToReset = Object.keys(abstractNum.levels)
        .map((l) => parseInt(l, 10))
        .filter((l) => l > levelIndex)
        .map((l) => `docx-counter-${abstractNum.id}-${l} 0`)
        .join(" ");
      
      if (deeperLevelsToReset) {
        css += `
/* Reset deeper level counters when level ${levelIndex} increments */
${itemSelector} {
  counter-reset: ${deeperLevelsToReset};
}
`;
      }
    });
  });
  
  return css;
}

/**
 * Generate enhanced list styles
 */
function generateEnhancedListStyles(styleInfo) {
  let css = `\n/* List Container Styles */\n`;
  css += `
ol[data-abstract-num], ul[data-abstract-num] {
  list-style-type: none; 
  padding-left: 0; 
  margin-top: 0.5em; 
  margin-bottom: 0.5em;
}
li[data-abstract-num] { /* Targeting <li> if they are the numbered items */
  display: block; 
  /* Indentation and numbering are handled by the [data-abstract-num][data-num-level] styles */
}
`;
  return css;
}

module.exports = {
  generateDOCXNumberingStyles,
  generateEnhancedListStyles,
};
