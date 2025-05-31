// File: lib/html/html-generator.js
// Main HTML generator - orchestrates HTML conversion and processing

const mammoth = require("mammoth");
const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");
const { DOMParser } = require("xmldom");

const { parseDocxStyles } = require("../parsers/style-parser");
const { generateCssFromStyleInfo } = require("../css/css-generator");
const { createStyleMap, createDocumentTransformer } = require("../css/style-mapper");
const { createEnhancedStyleMap, createEnhancedDocumentTransformer } = require("./generators/style-mapping");
const { createImageOptions, extractImagesFromDocx } = require("./generators/image-processing");
const { formatHtml } = require("./generators/html-formatting");
const { applyStylesAndProcessHtml } = require("./generators/html-processing");
const {
  parseDocumentMetadata,
  applyMetadataToHtml,
} = require("../parsers/metadata-parser");
const {
  parseTrackChanges,
  processTrackChanges: applyTrackChangesToHtml,
} = require("../parsers/track-changes-parser");
const {
  extractDocumentHeader,
  processHeaderForHtml,
} = require("../parsers/header-parser");

/**
 * Extract and apply styles from DOCX file
 */
async function extractAndApplyStyles(
  docxPath,
  cssFilename = null,
  options = {},
  outputDir = null
) {
  try {
    console.log(`Starting DOCX processing for: ${docxPath}`);
    if (!fs.existsSync(docxPath))
      throw new Error(`DOCX file not found: ${docxPath}`);

    const data = fs.readFileSync(docxPath);
    const zip = await JSZip.loadAsync(data);

    const styleXml = await zip.file("word/styles.xml")?.async("string");
    const documentXml = await zip.file("word/document.xml")?.async("string");
    const themeXml = await zip.file("word/theme/theme1.xml")?.async("string");
    const settingsXml = await zip.file("word/settings.xml")?.async("string");
    const numberingXml = await zip.file("word/numbering.xml")?.async("string");
    const corePropsXml = await zip.file("docProps/core.xml")?.async("string");
    const appPropsXml = await zip.file("docProps/app.xml")?.async("string");

    if (!styleXml || !documentXml)
      throw new Error("Invalid DOCX: missing styles.xml or document.xml");

    const styleDoc = new DOMParser().parseFromString(styleXml);
    const documentDoc = new DOMParser().parseFromString(documentXml);
    const themeDoc = themeXml
      ? new DOMParser().parseFromString(themeXml)
      : null;
    const settingsDoc = settingsXml
      ? new DOMParser().parseFromString(settingsXml)
      : null;
    const numberingDoc = numberingXml
      ? new DOMParser().parseFromString(numberingXml)
      : null;
    const corePropsDoc = corePropsXml
      ? new DOMParser().parseFromString(corePropsXml)
      : null;
    const appPropsDoc = appPropsXml
      ? new DOMParser().parseFromString(appPropsXml)
      : null;

    console.log(
      "Extracting styles, numbering context, metadata, track changes, and header..."
    );
    const styleInfo = await parseDocxStyles(docxPath);
    const metadata = parseDocumentMetadata(
      corePropsDoc,
      appPropsDoc,
      documentDoc
    );
    const trackChanges = parseTrackChanges(documentDoc);
    const headerInfo = await extractDocumentHeader(zip, documentDoc, styleDoc, styleInfo);

    if (styleInfo.numberingContext && styleInfo.numberingDefs) {
      styleInfo.numberingContext.forEach((ctx) => {
        if (ctx.numberingId && styleInfo.numberingDefs.nums[ctx.numberingId]) {
          ctx.abstractNumId =
            styleInfo.numberingDefs.nums[ctx.numberingId].abstractNumId;
        }
      });
    }

    console.log("Generating CSS...");
    const css = generateCssFromStyleInfo(styleInfo);

    console.log("Converting DOCX to HTML...");
    const htmlResult = await convertToStyledHtml(docxPath, styleInfo);

    const actualCssFilename =
      cssFilename || path.basename(docxPath, path.extname(docxPath)) + ".css";

    console.log("Applying styles and processing HTML structure...");
    const processedHtml = await applyStylesAndProcessHtml(
      htmlResult.value,
      css,
      styleInfo,
      actualCssFilename,
      metadata,
      trackChanges,
      headerInfo,
      options,
      zip,
      outputDir || path.dirname(docxPath)
    );

    const finalHtml = formatHtml(processedHtml);

    return {
      html: finalHtml,
      styles: css,
      messages: htmlResult.messages,
      metadata,
      trackChanges,
      numberingContext: styleInfo.numberingContext || [],
    };
  } catch (error) {
    console.error(
      "Error in extractAndApplyStyles:",
      error.message,
      error.stack
    );
    throw error;
  }
}

/**
 * Convert DOCX to styled HTML using mammoth
 */
async function convertToStyledHtml(docxPath, styleInfo) {
  try {
    const styleMap = createEnhancedStyleMap(styleInfo);
    const transformDocument = createEnhancedDocumentTransformer(styleInfo);
    const imageOptions = createImageOptions();

    const result = await mammoth.convertToHtml({
      path: docxPath,
      styleMap: styleMap,
      transformDocument: transformDocument,
      includeDefaultStyleMap: false,
      ...imageOptions,
    });

    if (!result.value || result.value.trim().length < 10) {
      console.warn(
        "Mammoth conversion resulted in very short HTML. Trying fallback."
      );
      const fallbackResult = await mammoth.convertToHtml({
        path: docxPath,
        ...imageOptions,
      });
      return fallbackResult.value.length > result.value.length
        ? fallbackResult
        : result;
    }
    return result;
  } catch (error) {
    console.error("Error in convertToStyledHtml:", error.message, error.stack);
    return mammoth.convertToHtml({ path: docxPath });
  }
}

// Export the main functions for backward compatibility
module.exports = {
  extractAndApplyStyles,
  convertToStyledHtml,
  applyStylesAndProcessHtml,
  extractImagesFromDocx,
};
