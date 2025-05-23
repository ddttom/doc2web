// File: ddttom/doc2web/doc2web-804421e293a28695ca8c4527a7c79f342ff0c562/lib/parsers/numbering-resolver.js

const { selectNodes, selectSingleNode } = require("../xml/xpath-utils");
// getCSSCounterFormat is defined in numbering-parser.js; it's used by getCSSCounterContent which is also in numbering-parser.js
// No direct import needed here if CSS content generation is fully handled by getCSSCounterContent.

/**
 * Extract complete paragraph numbering context from document XML.
 */
function extractParagraphNumberingContext(documentDoc, numberingDoc, styleDoc) {
  const paragraphContexts = [];
  try {
    const paragraphs = selectNodes("//w:p", documentDoc);
    paragraphs.forEach((p, index) => {
      const styleId = extractStyleId(p);
      const context = {
        paragraphIndex: index,
        paragraphId: extractParagraphId(p),
        numberingId: extractNumberingId(p),
        numberingLevel: extractNumberingLevel(p),
        styleId: styleId,
        textContent: extractTextContent(p),
        indentation: extractIndentation(p),
        isHeading: isHeadingParagraph(styleId, styleDoc),
        resolvedNumbering: null,
        abstractNumId: null, // Will be populated later
      };
      paragraphContexts.push(context);
    });
  } catch (error) {
    console.error(
      "Error extracting paragraph numbering context:",
      error.message,
      error.stack
    );
  }
  return paragraphContexts;
}

function extractParagraphId(paragraphNode) {
  try {
    const pId = selectSingleNode("w:pPr/w:pId", paragraphNode);
    return pId ? pId.getAttribute("w:val") : null;
  } catch (e) {
    console.error("Err extractPId:", e.message);
    return null;
  }
}

function extractNumberingId(paragraphNode) {
  try {
    const numIdNode = selectSingleNode("w:pPr/w:numPr/w:numId", paragraphNode);
    return numIdNode ? numIdNode.getAttribute("w:val") : null;
  } catch (e) {
    console.error("Err extractNumId:", e.message);
    return null;
  }
}

function extractNumberingLevel(paragraphNode) {
  try {
    const ilvlNode = selectSingleNode("w:pPr/w:numPr/w:ilvl", paragraphNode);
    return ilvlNode ? parseInt(ilvlNode.getAttribute("w:val"), 10) : null;
  } catch (e) {
    console.error("Err extractLvl:", e.message);
    return null;
  }
}

function extractStyleId(paragraphNode) {
  try {
    const pStyleNode = selectSingleNode("w:pPr/w:pStyle", paragraphNode);
    return pStyleNode ? pStyleNode.getAttribute("w:val") : null;
  } catch (e) {
    console.error("Err extractStyleId:", e.message);
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
  } catch (e) {
    console.error("Err extractTxt:", e.message);
    return "";
  }
}

function extractIndentation(paragraphNode) {
  /* ... (same as previous correct version) ... */
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
  /* ... (same as previous correct version) ... */
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

  // Initialize trackers based on concrete numId entries that map to valid abstractNums
  Object.keys(numberingDefs.nums || {}).forEach((numId) => {
    const numDef = numberingDefs.nums[numId];
    if (
      numDef &&
      numDef.abstractNumId &&
      numberingDefs.abstractNums[numDef.abstractNumId]
    ) {
      if (!sequenceTrackers[numId]) {
        sequenceTrackers[numId] = new NumberingSequenceTracker(
          numId,
          numDef.abstractNumId,
          numberingDefs
        );
      }
    } else {
      // console.warn(`Skipping tracker init for numId ${numId}: Invalid or missing abstractNumId link.`);
    }
  });

  paragraphContexts.forEach((context) => {
    if (context.numberingId && context.numberingLevel !== null) {
      const tracker = sequenceTrackers[context.numberingId];
      if (tracker) {
        context.abstractNumId = tracker.abstractNumId;
        context.resolvedNumbering = tracker.getNextNumber(
          context.numberingLevel,
          context
        );
      } else {
        // console.warn(`No sequence tracker for numId: ${context.numberingId} (paragraph index: ${context.paragraphIndex}). Check numbering.xml.`);
      }
    }
  });
  return paragraphContexts;
}

class NumberingSequenceTracker {
  constructor(numId, abstractNumId, numberingDefs) {
    this.numId = numId;
    this.abstractNumId = abstractNumId;
    this.numberingDefs = numberingDefs;
    this.levelCounters = {};

    const abstractNum = this.numberingDefs.abstractNums[this.abstractNumId];
    if (abstractNum && abstractNum.levels) {
      Object.keys(abstractNum.levels).forEach((levelIndexStr) => {
        const levelIndex = parseInt(levelIndexStr, 10);
        const levelDef = abstractNum.levels[levelIndex];
        this.levelCounters[levelIndex] = (levelDef.start || 1) - 1;
      });
    }
  }

  getNextNumber(levelIndex, context) {
    // levelIndex is 0-based
    const abstractNum = this.numberingDefs.abstractNums[this.abstractNumId];
    if (
      !abstractNum ||
      !abstractNum.levels ||
      !abstractNum.levels[levelIndex]
    ) {
      // console.warn(`Def missing for abstractNumId ${this.abstractNumId}, level ${levelIndex}`);
      return null;
    }

    let currentLevelDef = abstractNum.levels[levelIndex];
    const numInstance = this.numberingDefs.nums[this.numId];

    if (numInstance?.overrides?.[levelIndex]) {
      const override = numInstance.overrides[levelIndex];
      if (override.completeLevel) {
        currentLevelDef = { ...currentLevelDef, ...override.completeLevel };
      }
      if (override.startOverride !== undefined) {
        this.levelCounters[levelIndex] = override.startOverride - 1;
      }
    }

    // Restart logic for higher levels: If a level X-1 is restarting, all levels X, X+1... should reset.
    // This is implicitly handled by resetting deeper levels below.
    // Additionally, check currentLevelDef.restartAfterLevel
    if (
      currentLevelDef.restartAfterLevel !== undefined &&
      currentLevelDef.restartAfterLevel < levelIndex
    ) {
      // If this level restarts after a level higher than itself (e.g. level 1 restarts after level 0),
      // and that higher level has just incremented (which we don't track directly here but assume for restart condition)
      // This type of restart is complex and often depends on seeing a paragraph of the 'restartAfterLevel'.
      // For now, the main restart logic is handled by resetting deeper levels.
    }

    this.levelCounters[levelIndex] =
      (this.levelCounters[levelIndex] ?? (currentLevelDef.start || 1) - 1) + 1;

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

  // Updated formatNumbering method in NumberingSequenceTracker class
  formatNumbering(
    actualNumberForThisLevel,
    currentLevelIndex,
    currentLevelDef,
    abstractNum
  ) {
    // Build clean hierarchical numbering for section IDs
    const hierarchicalNumbering = this.buildHierarchicalNumbering(
      currentLevelIndex,
      currentLevelDef,
      abstractNum
    );

    // Build display formatting (with punctuation, etc.)
    const displayNumbering = this.buildDisplayNumbering(
      currentLevelIndex,
      currentLevelDef,
      abstractNum
    );

    // Generate section ID from the clean hierarchical numbering
    const sectionId = this.generateSectionId(
      hierarchicalNumbering,
      currentLevelIndex
    );

    return {
      rawNumber: actualNumberForThisLevel,
      formattedNumber: this.formatSingleNumber(
        actualNumberForThisLevel,
        currentLevelDef.format
      ),
      fullNumbering: displayNumbering, // Keep display format for CSS content
      hierarchicalNumbering: hierarchicalNumbering, // Clean format for section IDs
      sectionId: sectionId,
      levelDef: currentLevelDef,
      format: currentLevelDef.format,
      textFormat: currentLevelDef.textFormat,
      suffix: currentLevelDef.suffix,
    };
  }
  // Fixed generateSectionId method in NumberingSequenceTracker class
  generateSectionId(fullNumberingString, currentLevelIndex) {
    // Convert numbering string to valid HTML ID
    // Examples: "1.2.a" -> "section-1-2-a", "(1)" -> "section-1", "1)" -> "section-1"
    return (
      "section-" +
      fullNumberingString
        .replace(/[^\w.-]/g, "") // Remove non-word chars except dots and hyphens
        .replace(/\./g, "-") // Convert dots to hyphens
        .toLowerCase()
    ); // Ensure lowercase for consistency
  }
  // Enhanced buildHierarchicalNumbering method in NumberingSequenceTracker class
  buildHierarchicalNumbering(currentLevelIndex, currentLevelDef, abstractNum) {
    // For section IDs, we want a clean hierarchical structure like "1.2.a"
    // not the display format like "1.2.a." or "(1.2.a)"

    let hierarchicalParts = [];

    // Build the hierarchy from level 0 up to current level
    for (let level = 0; level <= currentLevelIndex; level++) {
      const levelDef = abstractNum.levels[level];
      const countForLevel = this.levelCounters[level];

      if (countForLevel !== undefined && levelDef) {
        const formattedNumber = this.formatSingleNumber(
          countForLevel,
          levelDef.format
        );
        hierarchicalParts.push(formattedNumber);
      }
    }

    // Join with dots for clean hierarchical structure
    return hierarchicalParts.join(".");
  }

  // Also add a separate method for display formatting that uses the original logic
  buildDisplayNumbering(currentLevelIndex, currentLevelDef, abstractNum) {
    // This method uses the parsedFormat.segments for display purposes
    // (keeping the original logic for the actual number display)

    if (
      !currentLevelDef.parsedFormat ||
      !currentLevelDef.parsedFormat.segments
    ) {
      // Fallback if parsedFormat is missing
      const currentNumber = this.levelCounters[currentLevelIndex];
      const suffix =
        currentLevelDef.format === "bullet"
          ? ""
          : currentLevelDef.textFormat &&
            currentLevelDef.textFormat.endsWith(")")
          ? ""
          : ".";
      return (
        this.formatSingleNumber(
          currentNumber === undefined ? 1 : currentNumber,
          currentLevelDef.format
        ) + suffix
      );
    }

    let result = "";
    currentLevelDef.parsedFormat.segments.forEach((segment) => {
      if (segment.type === "literal") {
        result += segment.value;
      } else if (segment.type === "level") {
        const referencedLevelOneBased = segment.value;
        const referencedLevelZeroBased = referencedLevelOneBased - 1;
        const countForLevel = this.levelCounters[referencedLevelZeroBased];
        const defForLevel = abstractNum.levels[referencedLevelZeroBased];

        if (countForLevel !== undefined && defForLevel) {
          result += this.formatSingleNumber(countForLevel, defForLevel.format);
        }
      }
    });
    return result;
  }
  formatSingleNumber(number, formatType) {
    // Lazy require getCSSCounterFormat to avoid potential circular dependencies at module load time
    const { getCSSCounterFormat } = require("./numbering-parser");
    const cssFormat = getCSSCounterFormat(formatType);

    // Adjust number for 1-based char codes if it's 0-based for calculation
    const numForChar =
      number < 1 && (cssFormat === "lower-alpha" || cssFormat === "upper-alpha")
        ? 1
        : number;

    switch (cssFormat) {
      case "decimal":
        return number.toString();
      // For letters, Word's %N often implies the Nth letter of the alphabet for that count.
      // If 'number' is 1, it's 'a' or 'A'.
      case "lower-alpha":
        return String.fromCharCode(96 + numForChar);
      case "upper-alpha":
        return String.fromCharCode(64 + numForChar);
      case "lower-roman":
        return this.toRoman(number).toLowerCase();
      case "upper-roman":
        return this.toRoman(number);
      case "decimal-leading-zero":
        return number.toString().padStart(2, "0");
      case "disc":
        return "•";
      case "none":
        return ""; // Explicitly handle 'none' format
      default:
        return number.toString();
    }
  }

  toRoman(num) {
    /* ... (same as previous correct version) ... */
    if (isNaN(num) || num < 1 || num > 3999) return num.toString();
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
  NumberingSequenceTracker,
  extractParagraphId,
  extractNumberingId,
  extractNumberingLevel,
  extractStyleId,
  extractTextContent,
  extractIndentation,
  isHeadingParagraph,
};
