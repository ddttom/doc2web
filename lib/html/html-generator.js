// File: ddttom/doc2web/doc2web-804421e293a28695ca8c4527a7c79f342ff0c562/lib/html/html-generator.js
// Corrected formatHtml function

const mammoth = require("mammoth");
const { JSDOM } = require("jsdom");
const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");
const { DOMParser } = require("xmldom");

const { parseDocxStyles } = require("../parsers/style-parser");
const { generateCssFromStyleInfo } = require("../css/css-generator");
const {
  createStyleMap,
  createDocumentTransformer,
} = require("../css/style-mapper");
const { ensureHtmlStructure } = require("./structure-processor");
const {
  processHeadings,
  processTOC,
  processNestedNumberedParagraphs,
  processSpecialParagraphs,
} = require("./content-processors");
const {
  processTables,
  processImages,
  processLanguageElements,
} = require("./element-processors");
const {
  parseDocumentMetadata,
  applyMetadataToHtml,
} = require("../parsers/metadata-parser");
const {
  parseTrackChanges,
  processTrackChanges: applyTrackChangesToHtml,
} = require("../parsers/track-changes-parser");

async function extractAndApplyStyles(
  docxPath,
  cssFilename = null,
  options = {}
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
      "Extracting styles, numbering context, metadata, and track changes..."
    );
    const styleInfo = await parseDocxStyles(docxPath);
    const metadata = parseDocumentMetadata(
      corePropsDoc,
      appPropsDoc,
      documentDoc
    );
    const trackChanges = parseTrackChanges(documentDoc);

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
    const finalHtml = applyStylesAndProcessHtml(
      htmlResult.value,
      css,
      styleInfo,
      actualCssFilename,
      metadata,
      trackChanges,
      options
    );

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

async function convertToStyledHtml(docxPath, styleInfo) {
  try {
    const styleMap = createEnhancedStyleMap(styleInfo);
    const transformDocument = createEnhancedDocumentTransformer(styleInfo);

    const imageOptions = {
      convertImage: mammoth.images.imgElement(function (image) {
        return image.read("base64").then(function (imageBuffer) {
          try {
            const extension = image.contentType.split("/")[1] || "png";
            const altTextSanitized = (image.altText || `image-${Date.now()}`)
              .replace(/[^\w-]/g, "_")
              .substring(0, 50);
            const filename = `${altTextSanitized}.${extension}`;
            return {
              src: `./images/${filename}`,
              alt: image.altText || "Document image",
              className: "docx-image",
            };
          } catch (imgError) {
            console.error("Error processing single image:", imgError);
            return {
              src: `./images/error-image-${Date.now()}.png`,
              alt: "Error processing image",
              className: "docx-image docx-image-error",
            };
          }
        });
      }),
    };

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

function createEnhancedStyleMap(styleInfo) {
  const styleMap = [];
  Object.entries(styleInfo.styles?.paragraph || {}).forEach(([id, style]) => {
    const safeClassName = id.toLowerCase().replace(/[^a-z0-9-_]/g, "-");
    let mapping = `p[style-name='${style.name}'] => p.docx-p-${safeClassName}`;
    if (
      style.numbering?.hasNumbering &&
      style.numbering.id &&
      style.numbering.level != null
    ) {
      const abstractNumId =
        styleInfo.numberingDefs?.nums[style.numbering.id]?.abstractNumId;
      if (abstractNumId) {
        mapping += `[data-num-id='${style.numbering.id}'][data-num-level='${style.numbering.level}'][data-abstract-num='${abstractNumId}']`;
      }
    }
    styleMap.push(mapping);
  });
  Object.entries(styleInfo.styles?.character || {}).forEach(([id, style]) => {
    const safeClassName = id.toLowerCase().replace(/[^a-z0-9-_]/g, "-");
    styleMap.push(
      `r[style-name='${style.name}'] => span.docx-c-${safeClassName}`
    );
  });
  for (let i = 1; i <= 6; i++) {
    styleMap.push(`p[style-name='heading ${i}'] => h${i}.docx-heading${i}`);
    styleMap.push(`p[style-name='Heading ${i}'] => h${i}.docx-heading${i}`);
  }
  styleMap.push("p[style-name='TOC Heading'] => h2.docx-toc-heading");
  for (let i = 1; i <= 9; i++) {
    styleMap.push(
      `p[style-name='TOC ${i}'] => p.docx-toc-entry.docx-toc-level-${i}`
    );
    styleMap.push(
      `p[style-name='toc ${i}'] => p.docx-toc-entry.docx-toc-level-${i}`
    );
  }
  styleMap.push("r[bold] => strong");
  styleMap.push("r[italic] => em");
  styleMap.push("r[underline] => span.docx-underline");
  styleMap.push("r[strikethrough] => s");
  styleMap.push("r[subscript] => sub");
  styleMap.push("r[superscript] => sup");
  return styleMap;
}

function createEnhancedDocumentTransformer(styleInfo) {
  return function (element) {
    return element;
  };
}

function applyStylesAndProcessHtml(
  html,
  cssString,
  styleInfo,
  cssFilename,
  metadata,
  trackChanges,
  options
) {
  try {
    console.log(`Applying styles. Initial HTML length: ${html.length}`);
    const dom = new JSDOM(html, {
      contentType: "text/html",
      includeNodeLocations: true,
    }); // includeNodeLocations can be memory intensive
    const document = dom.window.document;

    if (!document || !document.body)
      throw new Error("Failed to create DOM from HTML input.");

    ensureHtmlStructure(document);

    const link = document.createElement("link");
    link.setAttribute("rel", "stylesheet");
    link.setAttribute("href", `./${cssFilename}`);
    document.head.appendChild(link);

    if (options.preserveMetadata && metadata) {
      applyMetadataToHtml(document, metadata);
    }
    if (styleInfo.settings?.rtlGutter) document.body.classList.add("docx-rtl");

    const elementsToNumber = document.querySelectorAll(
      "p, h1, h2, h3, h4, h5, h6, li"
    );
    elementsToNumber.forEach((el) => {
      const numId = el.getAttribute("data-num-id");
      if (numId && styleInfo.numberingDefs?.nums[numId]) {
        el.setAttribute(
          "data-abstract-num",
          styleInfo.numberingDefs.nums[numId].abstractNumId
        );
      }
    });

    processHeadings(document, styleInfo, styleInfo.numberingContext);
    processTOC(document, styleInfo);
    processNestedNumberedParagraphs(
      document,
      styleInfo,
      styleInfo.numberingContext
    );
    processSpecialParagraphs(document, styleInfo);
    processTables(document);
    processImages(document);
    processLanguageElements(document);

    if (
      trackChanges &&
      trackChanges.hasTrackedChanges &&
      options.trackChangesMode !== "hide"
    ) {
      applyTrackChangesToHtml(document, trackChanges, {
        mode: options.trackChangesMode,
        showAuthor: options.showAuthor,
        showDate: options.showDate,
      });
    }
    if (options.enhanceAccessibility) {
      const {
        processForAccessibility,
      } = require("../accessibility/wcag-processor");
      processForAccessibility(document, styleInfo, metadata);
    }

    const finalHtml = dom.serialize();
    return formatHtml(finalHtml);
  } catch (error) {
    console.error(
      "Error in applyStylesAndProcessHtml:",
      error.message,
      error.stack
    );
    return formatHtml(
      `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Error</title></head><body><p>Error during HTML processing: ${error.message}</p><pre>${html}</pre></body></html>`
    );
  }
}

async function extractImagesFromDocx(docxPath, outputDir) {
  try {
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

    const imageOptions = {
      path: docxPath,
      convertImage: mammoth.images.imgElement(function (image) {
        return image.read("base64").then(function (imageBuffer) {
          try {
            const extension = image.contentType.split("/")[1] || "png";
            const sanitizedAltText = (image.altText || `image-${Date.now()}`)
              .replace(/[^\w-]/g, "_")
              .substring(0, 50);
            const filename = `${sanitizedAltText}.${extension}`;
            const imagePath = path.join(outputDir, filename);
            fs.writeFileSync(imagePath, Buffer.from(imageBuffer, "base64"));
            return {
              src: `./images/${filename}`,
              alt: image.altText || "Document image",
            };
          } catch (imgError) {
            console.error("Error saving image:", imgError);
            return {
              src: `./images/error-img-${Date.now()}.${extension}`,
              alt: "Error saving image",
            };
          }
        });
      }),
    };
    await mammoth.convertToHtml(imageOptions);
    console.log(`Images extracted to ${outputDir}`);
  } catch (error) {
    console.error("Error extracting images:", error.message, error.stack);
  }
}

/**
 * Format HTML with proper indentation and line breaks for better readability
 * @param {string} html - The HTML string to format
 * @returns {string} - Formatted HTML string
 */
// Enhanced formatHtml function for lib/html/html-generator.js
function formatHtml(html) {
  try {
    // Normalize self-closing tags for consistency
    html = html.replace(/<([^>]+?)([^\s/])\/>/g, "<$1$2 />");

    // Split into lines and prepare for formatting
    let lines = html.split("\n");
    let formattedLines = [];
    let indentLevel = 0;
    const indentSize = 2;
    
    // Define block elements that affect indentation
    const blockElements = new Set([
      "div", "p", "h1", "h2", "h3", "h4", "h5", "h6", 
      "ul", "ol", "li", "table", "thead", "tbody", "tfoot", "tr", "td", "th",
      "nav", "section", "article", "header", "footer", "aside", "main",
      "figure", "figcaption", "blockquote", "pre", "form", "fieldset",
      "dl", "dt", "dd"
    ]);
    
    // Define self-closing elements
    const selfClosingElements = new Set([
      "area", "base", "br", "col", "embed", "hr", "img", "input",
      "link", "meta", "param", "source", "track", "wbr"
    ]);

    // Process each line
    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      if (!line) continue;

      // Check for closing tags
      let closingTagMatch = line.match(/^<\/\s*([\w-]+)\s*>/);
      let openingTagMatch = line.match(/^<\s*([\w-]+)([^>]*)>/);

      // Decrease indent for closing tags
      if (closingTagMatch) {
        const tagName = closingTagMatch[1].toLowerCase();
        if (blockElements.has(tagName) && indentLevel > 0) {
          indentLevel--;
        }
      }

      // Add the line with proper indentation
      formattedLines.push(" ".repeat(indentLevel * indentSize) + line);

      // Increase indent for opening tags
      if (openingTagMatch) {
        const tagName = openingTagMatch[1].toLowerCase();
        const attributes = openingTagMatch[2];
        const isSelfClosingInTag = attributes.endsWith("/") || selfClosingElements.has(tagName);

        // Only increase indent for block elements with content
        if (
          blockElements.has(tagName) &&
          !isSelfClosingInTag &&
          !line.includes(`</${tagName}>`) // Don't indent if closing tag on same line
        ) {
          indentLevel++;
        }
      }
    }

    // Join formatted lines and add final newline
    return formattedLines.join("\n") + "\n";
  } catch (error) {
    console.error("Error formatting HTML:", error.message, error.stack);
    return html; // Return original on error
  }
}

// Enhanced version of applyStylesAndProcessHtml to ensure content preservation
function applyStylesAndProcessHtml(html, cssString, styleInfo, cssFilename, metadata, trackChanges, options) {
  try {
    console.log(`Applying styles. Initial HTML length: ${html.length}`);
    
    // Create DOM with content validation
    if (!html || html.trim().length < 10) {
      console.warn("HTML input is very short or empty. Using fallback.");
      html = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>Empty document</body></html>`;
    }
    
    const dom = new JSDOM(html, {
      contentType: "text/html",
      includeNodeLocations: true,
    });
    
    const document = dom.window.document;
    
    // Validate document structure
    if (!document || !document.body) {
      throw new Error("Failed to create DOM from HTML input.");
    }
    
    // Check content before processing
    const initialBodyContent = document.body.innerHTML;
    const initialBodyLength = initialBodyContent.length;
    console.log(`Initial body content length: ${initialBodyLength}`);
    
    // Ensure proper HTML structure
    ensureHtmlStructure(document);
    
    // Add CSS link
    const link = document.createElement("link");
    link.setAttribute("rel", "stylesheet");
    link.setAttribute("href", `./${cssFilename}`);
    document.head.appendChild(link);
    
    // Apply metadata if enabled
    if (options.preserveMetadata && metadata) {
      applyMetadataToHtml(document, metadata);
    }
    
    // Add RTL class if needed
    if (styleInfo.settings?.rtlGutter) {
      document.body.classList.add("docx-rtl");
    }
    
    // Apply abstract numbering IDs to elements with numbering
    const elementsToNumber = document.querySelectorAll("p, h1, h2, h3, h4, h5, h6, li");
    elementsToNumber.forEach((el) => {
      const numId = el.getAttribute("data-num-id");
      if (numId && styleInfo.numberingDefs?.nums[numId]) {
        el.setAttribute(
          "data-abstract-num",
          styleInfo.numberingDefs.nums[numId].abstractNumId
        );
      }
    });
    
    // Process document structure
    processHeadings(document, styleInfo, styleInfo.numberingContext);
    processTOC(document, styleInfo, styleInfo.numberingContext);
    processNestedNumberedParagraphs(document, styleInfo, styleInfo.numberingContext);
    processSpecialParagraphs(document, styleInfo);
    processTables(document);
    processImages(document);
    processLanguageElements(document);
    
    // Apply track changes if enabled
    if (
      trackChanges &&
      trackChanges.hasTrackedChanges &&
      options.trackChangesMode !== "hide"
    ) {
      applyTrackChangesToHtml(document, trackChanges, {
        mode: options.trackChangesMode,
        showAuthor: options.showAuthor,
        showDate: options.showDate,
      });
    }
    
    // Apply accessibility enhancements if enabled
    if (options.enhanceAccessibility) {
      const { processForAccessibility } = require("../accessibility/wcag-processor");
      processForAccessibility(document, styleInfo, metadata);
    }
    
    // Validate content after processing
    const finalBodyContent = document.body.innerHTML;
    const finalBodyLength = finalBodyContent.length;
    console.log(`Final body content length: ${finalBodyLength}`);
    
    // Check for significant content loss
    if (finalBodyLength < initialBodyLength * 0.8 && initialBodyLength > 100) {
      console.warn("Significant content loss detected during processing. Using content preservation fallback.");
      // Implement fallback strategy
      const bodyWithStructure = document.createElement('div');
      bodyWithStructure.innerHTML = initialBodyContent;
      
      // Add structure with minimal DOM manipulation
      const bodyElements = Array.from(bodyWithStructure.childNodes);
      document.body.innerHTML = '';
      bodyElements.forEach(node => {
        document.body.appendChild(node.cloneNode(true));
      });
      
      // Re-apply minimal styling
      document.querySelectorAll('p, h1, h2, h3, h4, h5, h6').forEach(el => {
        const text = el.textContent.trim();
        // Apply minimal context-specific styling
        if (el.tagName.toLowerCase().startsWith('h')) {
          el.classList.add(`docx-heading${el.tagName.slice(1)}`);
        }
      });
    }
    
    // Serialize and format HTML
    const finalHtml = dom.serialize();
    return formatHtml(finalHtml);
  } catch (error) {
    console.error(
      "Error in applyStylesAndProcessHtml:",
      error.message,
      error.stack
    );
    
    // Create fallback HTML with error message
    return formatHtml(
      `<!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <title>Error</title>
          <link rel="stylesheet" href="./${cssFilename}">
        </head>
        <body>
          <p>Error during HTML processing: ${error.message}</p>
          <pre>${html.substring(0, 1000)}...</pre>
        </body>
      </html>`
    );
  }
}
module.exports = {
  extractAndApplyStyles,
  convertToStyledHtml,
  applyStylesAndProcessHtml,
  extractImagesFromDocx,
};
