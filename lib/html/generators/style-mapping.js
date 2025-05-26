// File: lib/html/generators/style-mapping.js
// Style mapping and document transformation

/**
 * Create enhanced style map for mammoth conversion
 */
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

  // Enhanced heading mappings to include paragraph style classes for indentation
  for (let i = 1; i <= 6; i++) {
    // First check if there's a specific heading style in the paragraph styles
    let headingStyleId = null;
    let headingStyleClassName = null;

    Object.entries(styleInfo.styles?.paragraph || {}).forEach(([id, style]) => {
      if (
        style.name &&
        (style.name.toLowerCase() === `heading ${i}` ||
          style.name.toLowerCase() === `heading${i}`)
      ) {
        headingStyleId = id;
        headingStyleClassName = `docx-p-${id
          .toLowerCase()
          .replace(/[^a-z0-9-_]/g, "-")}`;
      }
    });

    if (headingStyleClassName) {
      // If we found a specific heading style, apply its class
      styleMap.push(
        `p[style-name='heading ${i}'] => h${i}.docx-heading${i}.${headingStyleClassName}`
      );
      styleMap.push(
        `p[style-name='Heading ${i}'] => h${i}.docx-heading${i}.${headingStyleClassName}`
      );
      styleMap.push(
        `p[style-name='Heading${i}'] => h${i}.docx-heading${i}.${headingStyleClassName}`
      );
    } else {
      // Otherwise use default heading mapping
      styleMap.push(`p[style-name='heading ${i}'] => h${i}.docx-heading${i}`);
      styleMap.push(`p[style-name='Heading ${i}'] => h${i}.docx-heading${i}`);
      styleMap.push(`p[style-name='Heading${i}'] => h${i}.docx-heading${i}`);
    }
  }

  // Handle TOC specific styles
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

/**
 * Create enhanced document transformer for mammoth conversion
 */
function createEnhancedDocumentTransformer(styleInfo) {
  return function (element) {
    return element;
  };
}

module.exports = {
  createEnhancedStyleMap,
  createEnhancedDocumentTransformer,
};