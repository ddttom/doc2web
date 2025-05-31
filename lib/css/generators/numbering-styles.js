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
      
      // Extract precise indentation from DOCX
      if (levelDef.indentation) {
        // Left indentation - where the paragraph starts
        if (levelDef.indentation.left !== undefined) {
          pIndentLeft = convertTwipToPt(levelDef.indentation.left);
        } else {
          // Fallback to level-based indentation
          pIndentLeft = levelIndex * 36; // 0.5 inch per level
        }
        
        // Handle hanging indent (most common for numbered lists)
        if (levelDef.indentation.hanging !== undefined) {
          const hangingIndent = convertTwipToPt(levelDef.indentation.hanging);
          numRegionWidth = hangingIndent;
          pTextIndent = -hangingIndent; // Create true hanging indent
          // The number will be positioned in the hanging indent space
        } 
        // Handle first line indent
        else if (levelDef.indentation.firstLine !== undefined) {
          pTextIndent = convertTwipToPt(levelDef.indentation.firstLine);
          numRegionWidth = Math.max(24, Math.abs(pTextIndent)); // Ensure enough space
        }
      } else {
        // Default indentation based on level
        pIndentLeft = levelIndex * 36; // 0.5 inch per level
        numRegionWidth = 24; // Default numbering width
      }
      
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
      
      // Create CSS for the numbered element with precise Word-like positioning
      css += `
/* Styling for level ${levelIndex} of abstract numbering ${abstractNum.id} */
${itemSelector} {
  display: block;
  position: relative;
  margin-left: ${pIndentLeft}pt;
  padding-left: ${numRegionWidth}pt;
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
