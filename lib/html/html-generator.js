// File: ddttom/doc2web/doc2web-804421e293a28695ca8c4527a7c79f342ff0c562/lib/html/html-generator.js
// No changes to this file in this iteration, but providing it for completeness of context.
// The formatHtml function is located here.

const mammoth = require("mammoth");
const { JSDOM } = require("jsdom");
const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");
const { DOMParser } = require("xmldom");

const { parseDocxStyles } = require("../parsers/style-parser");
const { generateCssFromStyleInfo } = require("../css/css-generator");
const { createStyleMap, createDocumentTransformer } = require("../css/style-mapper"); // Kept for mammoth options
const { ensureHtmlStructure /*, addDocumentMetadata - Initial metadata is added by applyMetadataToHtml */ } = require("./structure-processor");
const { processHeadings, processTOC, processNestedNumberedParagraphs, processSpecialParagraphs } = require("./content-processors");
const { processTables, processImages, processLanguageElements } = require("./element-processors");
const { parseDocumentMetadata, applyMetadataToHtml } = require("../parsers/metadata-parser"); // applyMetadataToHtml is key
const { parseTrackChanges, processTrackChanges: applyTrackChangesToHtml } = require("../parsers/track-changes-parser"); // Renamed for clarity

async function extractAndApplyStyles(docxPath, cssFilename = null, options = {}) {
  try {
    console.log(`Starting DOCX processing for: ${docxPath}`);
    if (!fs.existsSync(docxPath)) throw new Error(`DOCX file not found: ${docxPath}`);
    
    const data = fs.readFileSync(docxPath);
    const zip = await JSZip.loadAsync(data);
    
    const styleXml = await zip.file('word/styles.xml')?.async('string');
    const documentXml = await zip.file('word/document.xml')?.async('string');
    const themeXml = await zip.file('word/theme/theme1.xml')?.async('string');
    const settingsXml = await zip.file('word/settings.xml')?.async('string');
    const numberingXml = await zip.file('word/numbering.xml')?.async('string');
    const corePropsXml = await zip.file('docProps/core.xml')?.async('string');
    const appPropsXml = await zip.file('docProps/app.xml')?.async('string');
    
    if (!styleXml || !documentXml) throw new Error('Invalid DOCX: missing styles.xml or document.xml');
    
    const styleDoc = new DOMParser().parseFromString(styleXml);
    const documentDoc = new DOMParser().parseFromString(documentXml);
    const themeDoc = themeXml ? new DOMParser().parseFromString(themeXml) : null;
    const settingsDoc = settingsXml ? new DOMParser().parseFromString(settingsXml) : null;
    const numberingDoc = numberingXml ? new DOMParser().parseFromString(numberingXml) : null;
    const corePropsDoc = corePropsXml ? new DOMParser().parseFromString(corePropsXml) : null;
    const appPropsDoc = appPropsXml ? new DOMParser().parseFromString(appPropsXml) : null;

    console.log('Extracting styles, numbering context, metadata, and track changes...');
    const styleInfo = await parseDocxStyles(docxPath); // This already parses numbering, theme, etc.
    const metadata = parseDocumentMetadata(corePropsDoc, appPropsDoc, documentDoc);
    const trackChanges = parseTrackChanges(documentDoc);
    
    // Add abstractNumId to each item in numberingContext for CSS counter generation
    if (styleInfo.numberingContext && styleInfo.numberingDefs) {
        styleInfo.numberingContext.forEach(ctx => {
            if (ctx.numberingId && styleInfo.numberingDefs.nums[ctx.numberingId]) {
                ctx.abstractNumId = styleInfo.numberingDefs.nums[ctx.numberingId].abstractNumId;
            }
        });
    }


    console.log('Generating CSS...');
    const css = generateCssFromStyleInfo(styleInfo);

    console.log('Converting DOCX to HTML...');
    const htmlResult = await convertToStyledHtml(docxPath, styleInfo);

    const actualCssFilename = cssFilename || path.basename(docxPath, path.extname(docxPath)) + ".css";

    console.log('Applying styles and processing HTML structure...');
    const finalHtml = applyStylesAndProcessHtml( // Renamed for clarity
      htmlResult.value, 
      css, // CSS string passed for potential direct use, but mostly for saving
      styleInfo, 
      actualCssFilename, // CSS filename for linking
      metadata, 
      trackChanges, 
      options
    );
    
    return {
      html: finalHtml, styles: css, messages: htmlResult.messages,
      metadata, trackChanges, numberingContext: styleInfo.numberingContext || []
    };
  } catch (error) {
    console.error("Error in extractAndApplyStyles:", error.message, error.stack);
    throw error;
  }
}

async function convertToStyledHtml(docxPath, styleInfo) {
  try {
    const styleMap = createEnhancedStyleMap(styleInfo); // Use the refined style map
    const transformDocument = createEnhancedDocumentTransformer(styleInfo); // Use refined transformer

    const imageOptions = {
      convertImage: mammoth.images.imgElement(function (image) {
        return image.read("base64").then(function (imageBuffer) {
          try {
            const extension = image.contentType.split("/")[1] || 'png';
            // Alt text can be null, provide a fallback. Sanitize for filename.
            const altTextSanitized = (image.altText || `image-${Date.now()}`).replace(/[^\w-]/g, '_').substring(0, 50);
            const filename = `<span class="math-inline">\{altTextSanitized\}\.</span>{extension}`;
            return { src: `./images/${filename}`, alt: image.altText || "Document image", className: "docx-image" };
          } catch (imgError) {
            console.error('Error processing single image:', imgError);
            return { src: `./images/error-image-${Date.now()}.png`, alt: "Error processing image", className: "docx-image docx-image-error" };
          }
        });
      }),
    };

    const result = await mammoth.convertToHtml({
      path: docxPath, styleMap: styleMap,
      transformDocument: transformDocument, includeDefaultStyleMap: false, // Set to false if our CSS is comprehensive
      ...imageOptions,
    });
    
    if (!result.value || result.value.trim().length < 10) {
      console.warn('Mammoth conversion resulted in very short HTML. Trying fallback.');
      const fallbackResult = await mammoth.convertToHtml({ path: docxPath, ...imageOptions });
      return fallbackResult.value.length > result.value.length ? fallbackResult : result;
    }
    return result;
  } catch (error) {
    console.error('Error in convertToStyledHtml:', error.message, error.stack);
    return mammoth.convertToHtml({ path: docxPath }); // Basic fallback
  }
}

function createEnhancedStyleMap(styleInfo) {
  const styleMap = [];
  // Paragraph styles
  Object.entries(styleInfo.styles?.paragraph || {}).forEach(([id, style]) => {
    const safeClassName = id.toLowerCase().replace(/[^a-z0-9-_]/g, '-');
    let mapping = `p[style-name='<span class="math-inline">\{style\.name\}'\] \=\> p\.docx\-p\-</span>{safeClassName}`;
    // Add data attributes for numbering if present in the style definition
    if (style.numbering?.hasNumbering && style.numbering.id && style.numbering.level != null) {
        const abstractNumId = styleInfo.numberingDefs?.nums[style.numbering.id]?.abstractNumId;
        if (abstractNumId) {
            mapping += `[data-num-id='<span class="math-inline">\{style\.numbering\.id\}'\]\[data\-num\-level\='</span>{style.numbering.level}'][data-abstract-num='${abstractNumId}']`;
        }
    }
    styleMap.push(mapping);
  });
  // Character styles
  Object.entries(styleInfo.styles?.character || {}).forEach(([id, style]) => {
    const safeClassName = id.toLowerCase().replace(/[^a-z0-9-_]/g, '-');
    styleMap.push(`r[style-name='<span class="math-inline">\{style\.name\}'\] \=\> span\.docx\-c\-</span>{safeClassName}`);
  });
   // Common heading styles (ensure these map to hN tags)
  for (let i = 1; i <= 6; i++) {
    styleMap.push(`p[style-name='heading <span class="math-inline">\{i\}'\] \=\> h</span>{i}.docx-heading${i}`);
    styleMap.push(`p[style-name='Heading <span class="math-inline">\{i\}'\] \=\> h</span>{i}.docx-heading${i}`); // Case variation
  }
  // TOC styles
  styleMap.push("p[style-name='TOC Heading'] => h2.docx-toc-heading");
  for (let i = 1; i <= 9; i++) { // TOC levels 1-9
    styleMap.push(`p[style-name='TOC <span class="math-inline">\{i\}'\] \=\> p\.docx\-toc\-entry\.docx\-toc\-level\-</span>{i}`);
    styleMap.push(`p[style-name='toc <span class="math-inline">\{i\}'\] \=\> p\.docx\-toc\-entry\.docx\-toc\-level\-</span>{i}`);
  }
  // Basic formatting
  styleMap.push("r[bold] => strong");
  styleMap.push("r[italic] => em");
  styleMap.push("r[underline] => span.docx-underline"); // Use class for better styling control
  styleMap.push("r[strikethrough] => s"); // Use <s> for strikethrough
  styleMap.push("r[subscript] => sub");
  styleMap.push("r[superscript] => sup");
  return styleMap;
}

function createEnhancedDocumentTransformer(styleInfo) {
  return function (element) { // Mammoth passes the element to transform
    // This transformer can be used to add data attributes directly during mammoth's processing
    // For example, finding paragraphs that mammoth identifies as list items and adding data-* attributes
    // However, current numbering processing is mostly done post-mammoth on the full HTML DOM.
    // This function can be expanded if more direct manipulation during mammoth's phase is needed.
    return element;
  };
}


function applyStylesAndProcessHtml(html, cssString, styleInfo, cssFilename, metadata, trackChanges, options) {
  try {
    console.log(`Applying styles. Initial HTML length: ${html.length}`);
    const dom = new JSDOM(html, { contentType: "text/html", includeNodeLocations: true });
    const document = dom.window.document;

    if (!document || !document.body) throw new Error('Failed to create DOM from HTML input.');
    
    ensureHtmlStructure(document); // Ensures html, head, body exist

    // Link CSS
    const link = document.createElement("link");
    link.setAttribute("rel", "stylesheet");
    link.setAttribute("href", `./${cssFilename}`); // Ensure it's a relative path
    document.head.appendChild(link);

    if (options.preserveMetadata && metadata) {
      applyMetadataToHtml(document, metadata);
    }
    if (styleInfo.settings?.rtlGutter) document.body.classList.add("docx-rtl");

    // Apply data-abstract-num to paragraph/heading elements for CSS targeting
    // This helps link elements to their abstract numbering definitions for styling
    const elementsToNumber = document.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li');
    elementsToNumber.forEach(el => {
        const numId = el.getAttribute('data-num-id');
        if (numId && styleInfo.numberingDefs?.nums[numId]) {
            el.setAttribute('data-abstract-num', styleInfo.numberingDefs.nums[numId].abstractNumId);
        }
    });
    
    // Process content (headings, TOC, lists etc.)
    processHeadings(document, styleInfo, styleInfo.numberingContext);
    processTOC(document, styleInfo); // This should correctly structure TOC entries
    processNestedNumberedParagraphs(document, styleInfo, styleInfo.numberingContext); // This handles lists
    processSpecialParagraphs(document, styleInfo);
    processTables(document);
    processImages(document);
    processLanguageElements(document);

    if (trackChanges && trackChanges.hasTrackedChanges && options.trackChangesMode !== 'hide') {
      applyTrackChangesToHtml(document, trackChanges, {
        mode: options.trackChangesMode,
        showAuthor: options.showAuthor,
        showDate: options.showDate
      });
    }
    if (options.enhanceAccessibility) {
      const { processForAccessibility } = require('../accessibility/wcag-processor');
      processForAccessibility(document, styleInfo, metadata);
    }
    
    const finalHtml = dom.serialize();
    return formatHtml(finalHtml); // Format before returning

  } catch (error) {
    console.error("Error in applyStylesAndProcessHtml:", error.message, error.stack);
    return formatHtml(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Error</title></head><body><p>Error during HTML processing: <span class="math-inline">\{error\.message\}</p\><pre\></span>{html}</pre></body></html>`); // Fallback
  }
}

async function extractImagesFromDocx(docxPath, outputDir) {
  try {
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
    
    const imageOptions = {
      path: docxPath,
      convertImage: mammoth.images.imgElement(function(image) {
        return image.read("base64").then(function(imageBuffer) {
          try {
            const extension = image.contentType.split('/')[1] || 'png';
            const sanitizedAltText = (image.altText || `image-${Date.now()}`).replace(/[^\w-]/g, '_').substring(0, 50);
            const filename = `<span class="math-inline">\{sanitizedAltText\}\.</span>{extension}`;
            const imagePath = path.join(outputDir, filename);
            fs.writeFileSync(imagePath, Buffer.from(imageBuffer, 'base64'));
            return { src: `./images/${filename}`, alt: image.altText || 'Document image' };
          } catch (imgError) {
            console.error('Error saving image:', imgError);
            return { src: `./images/error-img-<span class="math-inline">\{Date\.now\(\)\}\.</span>{extension}`, alt: 'Error saving image' };
          }
        });
      })
    };
    await mammoth.convertToHtml(imageOptions); // This call is primarily for image extraction side-effect
    console.log(`Images extracted to ${outputDir}`);
  } catch (error) {
    console.error('Error extracting images:', error.message, error.stack);
    // Don't throw, allow main process to continue if image extraction fails
  }
}

function formatHtml(html) {
  try {
    html = html.replace(/<([^>]+?)([^\s])\/>/g, '<$1$2 />');
    const majorElements = ['<!DOCTYPE', '<html', '</html>', '<head', '</head>', '<body', '</body>', 
                          '<div', '</div>', '<nav', '</nav>', '<section', '</section>', 
                          '<article', '</article>', '<h1', '</h1>', '<h2', '</h2>', 
                          '<h3', '</h3>', '<h4', '</h4>', '<h5', '</h5>', '<h6', '</h6>',
                          '<p', '</p>', '<ul', '</ul>', '<ol', '</ol>', '<li', '</li>',
                          '<table', '</table>', '<tr', '</tr>', '<td', '</td>', '<th', '</th>',
                          '<thead', '</thead>', '<tbody', '</tbody>', '<tfoot', '</tfoot>',
                          '<script', '</script>', '<style', '</style>', '<link', '<meta'];
    
    majorElements.forEach(element => {
      const regex = new RegExp(`(<${element.replace(/[<>]/g, '')}(?:[^>]*)?>)`, 'gi');
      html = html.replace(regex, '\n$1');
      if (!element.startsWith('</') && !element.endsWith('/>') && element !== '<!DOCTYPE' && element !=='<meta' && element !=='<link' && element !=='<br' && element !=='<hr' && element !=='<img') {
        const endRegex = new RegExp(`(</${element.substring(1)}>)`, 'gi');
        html = html.replace(endRegex, '\n$1');
      }
    });
    
    let lines = html.split('\n');
    let formattedLines = [];
    let indentLevel = 0;
    const indentSize = 2; // Number of spaces for indentation

    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      if (!line) continue;

      const isClosingTag = line.startsWith('</');
      const isSelfClosing = line.endsWith('/>') || line.startsWith('<meta') || line.startsWith('<link') || line.startsWith('<img') || line.startsWith('<br') || line.startsWith('<hr') || line.startsWith('<!DOCTYPE');
      
      if (isClosingTag && indentLevel > 0) {
        indentLevel--;
      }
      
      formattedLines.push(' '.repeat(indentLevel * indentSize) + line);
      
      if (!isClosingTag && !isSelfClosing && line.startsWith('<') && !line.startsWith('