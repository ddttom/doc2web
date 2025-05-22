// File: ddttom/doc2web/doc2web-804421e293a28695ca8c4527a7c79f342ff0c562/lib/parsers/numbering-resolver.js

const { selectNodes, selectSingleNode } = require("../xml/xpath-utils");
// Removed getCSSCounterFormat as it's in numbering-parser.js and used there for getCSSCounterContent
// const { getCSSCounterFormat } = require('./numbering-parser');

/**
 * Extract complete paragraph numbering context from document XML
 * Analyzes document.xml to extract numbering context for each paragraph
 * * @param {Document} documentDoc - Document XML document
 * @param {Document} numberingDoc - Numbering XML document
 * @param {Document} styleDoc - Style XML document
 * @returns {Array} - Array of paragraph contexts with numbering information
 */
function extractParagraphNumberingContext(documentDoc, numberingDoc, styleDoc) {
  const paragraphContexts = [];

  try {
    const paragraphs = selectNodes("//w:p", documentDoc);

    paragraphs.forEach((p, index) => {
      const context = {
        paragraphIndex: index,
        paragraphId: extractParagraphId(p),
        numberingId: extractNumberingId(p), // Concrete numId
        numberingLevel: extractNumberingLevel(p), // 0-indexed ilvl
        styleId: extractStyleId(p),
        textContent: extractTextContent(p),
        indentation: extractIndentation(p),
        isHeading: false,
        resolvedNumbering: null,
        // abstractNumId will be populated later if numberingId exists
      };

      context.isHeading = isHeadingParagraph(context.styleId, styleDoc);
      context.paragraphNode = p; // Keep reference if needed later
      paragraphContexts.push(context);
    });
  } catch (error) {
    console.error("Error extracting paragraph numbering context:", error);
  }

  return paragraphContexts;
}

function extractParagraphId(paragraphNode) {
  try {
    const pId = selectSingleNode("w:pPr/w:pId", paragraphNode);
    return pId ? pId.getAttribute("w:val") : null;
  } catch (error) {
    console.error("Error extracting paragraph ID:", error);
    return null;
  }
}

function extractNumberingId(paragraphNode) {
  try {
    const numIdNode = selectSingleNode("w:pPr/w:numPr/w:numId", paragraphNode);
    return numIdNode ? numIdNode.getAttribute("w:val") : null;
  } catch (error) {
    console.error("Error extracting numbering ID:", error);
    return null;
  }
}

function extractNumberingLevel(paragraphNode) {
  try {
    const ilvlNode = selectSingleNode("w:pPr/w:numPr/w:ilvl", paragraphNode);
    return ilvlNode ? parseInt(ilvlNode.getAttribute("w:val"), 10) : null;
  } catch (error) {
    console.error("Error extracting numbering level:", error);
    return null;
  }
}

function extractStyleId(paragraphNode) {
  try {
    const pStyleNode = selectSingleNode("w:pPr/w:pStyle", paragraphNode);
    return pStyleNode ? pStyleNode.getAttribute("w:val") : null;
  } catch (error) {
    console.error("Error extracting style ID:", error);
    return null;
  }
}

function extractTextContent(paragraphNode) {
  try {
    const textNodes = selectNodes(".//w:t", paragraphNode);
    return textNodes
      .map((t) => t.textContent || "")
      .join("")
      .trim();
  } catch (error) {
    console.error("Error extracting text content:", error);
    return "";
  }
}

function extractIndentation(paragraphNode) {
  const indentation = {};
  try {
    const indNode = selectSingleNode("w:pPr/w:ind", paragraphNode);
    if (indNode) {
      if (indNode.hasAttribute("w:left"))
        indentation.left = parseInt(indNode.getAttribute("w:left"), 10);
      if (indNode.hasAttribute("w:hanging"))
        indentation.hanging = parseInt(indNode.getAttribute("w:hanging"), 10);
      if (indNode.hasAttribute("w:firstLine"))
        indentation.firstLine = parseInt(
          indNode.getAttribute("w:firstLine"),
          10
        );
      if (indNode.hasAttribute("w:start"))
        indentation.start = parseInt(indNode.getAttribute("w:start"), 10);
      if (indNode.hasAttribute("w:end"))
        indentation.end = parseInt(indNode.getAttribute("w:end"), 10);
    }
  } catch (error) {
    console.error("Error extracting indentation:", error);
  }
  return indentation;
}

function isHeadingParagraph(styleId, styleDoc) {
  if (!styleId || !styleDoc) return false;
  try {
    if (/^heading\d+$|^toc\d+$|^h\d+$/i.test(styleId)) return true;
    const styleNode = selectSingleNode(
      `//w:style[@w:styleId='${styleId}']`,
      styleDoc
    );
    if (styleNode) {
      const nameNode = selectSingleNode("w:name", styleNode);
      const styleName = nameNode ? nameNode.getAttribute("w:val") || "" : "";
      return /heading|title|caption/i.test(styleName);
    }
  } catch (error) {
    console.error("Error checking if heading paragraph:", error);
  }
  return false;
}

function resolveNumberingForParagraphs(paragraphContexts, numberingDefs) {
  const sequenceTrackers = {};

  Object.keys(numberingDefs.nums || {}).forEach((numId) => {
    if (!sequenceTrackers[numId]) {
      // Ensure tracker is initialized only once per numId
      const abstractNumId = numberingDefs.nums[numId].abstractNumId;
      if (abstractNumId && numberingDefs.abstractNums[abstractNumId]) {
        sequenceTrackers[numId] = new NumberingSequenceTracker(
          numId,
          abstractNumId,
          numberingDefs
        );
      } else {
        console.warn(
          `Cannot initialize tracker: AbstractNumId ${abstractNumId} not found for NumId ${numId}`
        );
      }
    }
  });

  paragraphContexts.forEach((context) => {
    if (context.numberingId && context.numberingLevel !== null) {
      const tracker = sequenceTrackers[context.numberingId];
      if (tracker) {
        context.abstractNumId = tracker.abstractNumId; // Store for CSS generation
        context.resolvedNumbering = tracker.getNextNumber(
          context.numberingLevel,
          context
        );
      } else {
        // This can happen if a numId in document.xml doesn't have a corresponding <w:num> entry in numbering.xml
        // or if the abstractNumId it points to doesn't exist.
        console.warn(
          `No sequence tracker found for numId: ${context.numberingId}. Paragraph index: ${context.paragraphIndex}`
        );
      }
    }
  });
  return paragraphContexts;
}

class NumberingSequenceTracker {
  constructor(numId, abstractNumId, numberingDefs) {
    this.numId = numId; // Concrete numId
    this.abstractNumId = abstractNumId;
    this.numberingDefs = numberingDefs;
    this.levelCounters = {}; // Stores current count for each level index, e.g., {0: 1, 1: 3}

    const abstractNum = this.numberingDefs.abstractNums[this.abstractNumId];
    if (abstractNum && abstractNum.levels) {
      Object.keys(abstractNum.levels).forEach((levelIndex) => {
        const levelDef = abstractNum.levels[levelIndex];
        // Initialize with start value -1 so first increment makes it the start value
        this.levelCounters[levelIndex] = (levelDef.start || 1) - 1;
      });
    }
  }

  getNextNumber(levelIndex, context) {
    // levelIndex is 0-based (from w:ilvl)
    const abstractNum = this.numberingDefs.abstractNums[this.abstractNumId];
    if (
      !abstractNum ||
      !abstractNum.levels ||
      !abstractNum.levels[levelIndex]
    ) {
      console.warn(
        `Numbering definition not found for abstractNumId ${this.abstractNumId}, level ${levelIndex}`
      );
      return null;
    }

    let currentLevelDef = abstractNum.levels[levelIndex];
    const numInstance = this.numberingDefs.nums[this.numId];

    // Check for instance overrides for this level
    if (
      numInstance &&
      numInstance.overrides &&
      numInstance.overrides[levelIndex]
    ) {
      const override = numInstance.overrides[levelIndex];
      if (override.completeLevel) {
        // If w:lvlOverride/w:lvl exists
        currentLevelDef = { ...currentLevelDef, ...override.completeLevel };
      }
      if (override.startOverride !== undefined) {
        // If there's a startOverride, reset counter for this level to startOverride - 1
        this.levelCounters[levelIndex] = override.startOverride - 1;
      }
    }

    // Handle restart logic: if a higher level (lower index) has incremented,
    // or if this level has a restart rule.
    // This is complex. A simple approach: if w:lvlRestart is defined for currentLevelDef.
    // The logic for resetting deeper levels is simpler:

    // Increment counter for the current level
    this.levelCounters[levelIndex] =
      (this.levelCounters[levelIndex] === undefined
        ? (currentLevelDef.start || 1) - 1
        : this.levelCounters[levelIndex]) + 1;

    // Reset counters for all deeper levels
    Object.keys(abstractNum.levels).forEach((deeperLvlIndexStr) => {
      const deeperLvlIndex = parseInt(deeperLvlIndexStr, 10);
      if (deeperLvlIndex > levelIndex) {
        const deeperLevelDef = abstractNum.levels[deeperLvlIndex];
        this.levelCounters[deeperLvlIndex] = (deeperLevelDef.start || 1) - 1;
      }
    });

    const actualNumberForThisLevel = this.levelCounters[levelIndex];
    return this.formatNumbering(
      actualNumberForThisLevel,
      levelIndex,
      currentLevelDef,
      abstractNum
    );
  }

  formatNumbering(
    actualNumberForThisLevel,
    currentLevelIndex,
    currentLevelDef,
    abstractNum
  ) {
    // buildHierarchicalNumbering constructs the full string like "1.a.i"
    // based on currentLevelDef.textFormat (w:lvlText) which refers to multiple levels.
    const fullNumberingString = this.buildHierarchicalNumbering(
      currentLevelIndex,
      currentLevelDef,
      abstractNum
    );

    return {
      rawNumber: actualNumberForThisLevel, // The number for this specific level, e.g., if "1.a.i", for level 2 (i), rawNumber is 1.
      formattedNumber: this.formatSingleNumber(
        actualNumberForThisLevel,
        currentLevelDef.format
      ), // "i"
      fullNumbering: fullNumberingString, // "1.a.i"
      levelDef: currentLevelDef, // current level's definition
      format: currentLevelDef.format, // e.g. 'lowerRoman'
      textFormat: currentLevelDef.textFormat, // e.g. "%1.%2.%3"
      suffix: currentLevelDef.suffix, // e.g. 'tab'
    };
  }

  buildHierarchicalNumbering(currentLevelIndex, currentLevelDef, abstractNum) {
    // Uses currentLevelDef.parsedFormat.segments
    if (
      !currentLevelDef.parsedFormat ||
      !currentLevelDef.parsedFormat.segments
    ) {
      // Fallback if parsedFormat is missing or malformed
      return (
        this.formatSingleNumber(
          this.levelCounters[currentLevelIndex] || 1,
          currentLevelDef.format
        ) + (currentLevelDef.format === "bullet" ? "" : ".")
      );
    }

    let result = "";
    currentLevelDef.parsedFormat.segments.forEach((segment) => {
      if (segment.type === "literal") {
        result += segment.value;
      } else if (segment.type === "level") {
        const referencedLevelOneBased = segment.value; // %N is 1-indexed
        const referencedLevelZeroBased = referencedLevelOneBased - 1;

        const countForLevel = this.levelCounters[referencedLevelZeroBased];
        const defForLevel = abstractNum.levels[referencedLevelZeroBased];

        if (countForLevel !== undefined && defForLevel) {
          result += this.formatSingleNumber(countForLevel, defForLevel.format);
        } else {
          // Fallback if a referenced level has no count or definition (should not happen in valid DOCX)
          // console.warn(`Missing count or def for level ${referencedLevelZeroBased} in buildHierarchicalNumbering`);
          // result += `%${referencedLevelOneBased}`; // Keep placeholder if problematic
        }
      }
    });
    return result;
  }

  formatSingleNumber(number, formatType) {
    const { getCSSCounterFormat } = require("./numbering-parser"); // Lazy load to avoid circular dep issues at module load time
    const cssFormat = getCSSCounterFormat(formatType); // Convert DOCX format to CSS-like for consistency

    switch (cssFormat) {
      case "decimal":
        return number.toString();
      case "lower-alpha":
        return String.fromCharCode(96 + number); // a, b, c (assuming number is 1-based for char codes)
      case "upper-alpha":
        return String.fromCharCode(64 + number); // A, B, C
      case "lower-roman":
        return this.toRoman(number).toLowerCase();
      case "upper-roman":
        return this.toRoman(number);
      case "decimal-leading-zero":
        return number.toString().padStart(2, "0");
      case "disc":
        return "•"; // Default bullet
      // Add other supported formats from getCSSCounterFormat if they need special string generation
      default:
        return number.toString(); // Fallback
    }
  }

  toRoman(num) {
    if (isNaN(num) || num < 1 || num > 3999) return num.toString(); // Fallback for invalid
    const roman = {
      M: 1000,
      CM: 900,
      D: 500,
      CD: 400,
      C: 100,
      XC: 90,
      L: 50,
      XL: 40,
      X: 10,
      IX: 9,
      V: 5,
      IV: 4,
      I: 1,
    };
    let str = "";
    for (let i of Object.keys(roman)) {
      let q = Math.floor(num / roman[i]);
      num -= q * roman[i];
      str += i.repeat(q);
    }
    return str;
  }
}

module.exports = {
  extractParagraphNumberingContext,
  resolveNumberingForParagraphs,
  NumberingSequenceTracker, // Exporting class for potential external use or testing
  // No need to export helpers if only used internally here
  extractParagraphId,
  extractNumberingId,
  extractNumberingLevel,
  extractStyleId,
  extractTextContent,
  extractIndentation,
  isHeadingParagraph,
};
