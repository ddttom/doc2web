// File: lib/html/generators/style-mapping.js
// Style mapping and document transformation

/**
 * Add bullet point and list style mappings
 */
function addListStyleMappings(styleMap, styleInfo) {
  // Map common bullet point paragraph styles
  styleMap.push("p[style-name='List Paragraph'] => li.docx-list-item");
  styleMap.push("p[style-name='ListParagraph'] => li.docx-list-item");
  styleMap.push("p[style-name='List Bullet'] => li.docx-list-item");
  styleMap.push("p[style-name='ListBullet'] => li.docx-list-item");
  
  // Map bullet-specific styles from numbering definitions
  Object.entries(styleInfo.styles?.paragraph || {}).forEach(([id, style]) => {
    if (style.numbering?.hasNumbering) {
      const safeClassName = id.toLowerCase().replace(/[^a-z0-9-_]/g, "-");
      
      // Check if this style has bullet numbering
      if (isBulletStyle(style, styleInfo)) {
        styleMap.push(`p[style-name='${style.name}'] => li.docx-bullet-${safeClassName}`);
      }
    }
  });
  
  // Handle numbered lists that contain bullets at sub-levels
  for (let level = 0; level <= 8; level++) {
    styleMap.push(`p[data-num-level='${level}'][data-format='bullet'] => li.docx-bullet-level-${level}`);
  }
}

/**
 * Check if a paragraph style represents a bullet list
 */
function isBulletStyle(style, styleInfo) {
  if (!style.numbering?.hasNumbering) return false;
  
  const numId = style.numbering.id;
  const level = style.numbering.level;
  
  if (styleInfo.numberingDefs?.nums[numId]) {
    const abstractNumId = styleInfo.numberingDefs.nums[numId].abstractNumId;
    const abstractNum = styleInfo.numberingDefs.abstractNums[abstractNumId];
    
    if (abstractNum?.levels[level]?.format === 'bullet') {
      return true;
    }
  }
  
  return false;
}

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

  // Add bullet point and list style mappings
  addListStyleMappings(styleMap, styleInfo);

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
  // Enhanced character formatting mappings
  addCharacterFormattingMappings(styleMap, styleInfo);
  
  // Enhanced paragraph formatting mappings
  addParagraphFormattingMappings(styleMap, styleInfo);
  
  // Enhanced table formatting mappings
  addTableFormattingMappings(styleMap, styleInfo);
  
  return styleMap;
}

/**
 * Create enhanced document transformer for mammoth conversion
 */
function createEnhancedDocumentTransformer(styleInfo) {
  return function (element) {
    // Transform bullet point paragraphs into list structures
    if (element.type === "paragraph") {
      const numberingInfo = getNumberingInfo(element, styleInfo);
      
      if (numberingInfo && numberingInfo.format === 'bullet') {
        // Convert paragraph to list item
        return {
          ...element,
          type: "listItem",
          listType: "unordered",
          level: numberingInfo.level || 0,
          bulletChar: numberingInfo.bulletChar || '•'
        };
      }
      
      // Add comprehensive paragraph formatting attributes
      element = addParagraphFormattingAttributes(element, styleInfo);
    }
    
    // Transform character runs with enhanced formatting
    if (element.type === "run") {
      element = addCharacterFormattingAttributes(element, styleInfo);
    }
    
    // Transform tables with enhanced formatting
    if (element.type === "table") {
      element = addTableFormattingAttributes(element, styleInfo);
    }
    
    // Transform table rows with enhanced formatting
    if (element.type === "tableRow") {
      element = addTableRowFormattingAttributes(element, styleInfo);
    }
    
    // Transform table cells with enhanced formatting
    if (element.type === "tableCell") {
      element = addTableCellFormattingAttributes(element, styleInfo);
    }
    
    return element;
  };
}

/**
 * Extract numbering information from element and styleInfo
 */
function getNumberingInfo(element, styleInfo) {
  const styleName = element.styleName;
  if (!styleName) return null;
  
  // Check if this style has bullet numbering
  const paragraphStyle = findStyleByName(styleInfo.styles?.paragraph, styleName);
  if (paragraphStyle?.numbering?.format === 'bullet') {
    return {
      format: 'bullet',
      level: paragraphStyle.numbering.level,
      bulletChar: paragraphStyle.numbering.bulletChar || '•'
    };
  }
  
  // Check if this is a bullet style based on numbering definitions
  if (paragraphStyle && isBulletStyle(paragraphStyle, styleInfo)) {
    const numId = paragraphStyle.numbering.id;
    const level = paragraphStyle.numbering.level;
    
    if (styleInfo.numberingDefs?.nums[numId]) {
      const abstractNumId = styleInfo.numberingDefs.nums[numId].abstractNumId;
      const abstractNum = styleInfo.numberingDefs.abstractNums[abstractNumId];
      const levelDef = abstractNum?.levels[level];
      
      if (levelDef?.format === 'bullet') {
        return {
          format: 'bullet',
          level: level,
          bulletChar: levelDef.bulletChar || levelDef.text || '•'
        };
      }
    }
  }
  
  return null;
}

/**
 * Find style by name in styles collection
 */
function findStyleByName(styles, styleName) {
  if (!styles || !styleName) return null;
  
  return Object.values(styles).find(style => 
    style.name && style.name.toLowerCase() === styleName.toLowerCase()
  );
}

module.exports = {
  createEnhancedStyleMap,
  createEnhancedDocumentTransformer,
};

/**
 * Add comprehensive character formatting mappings
 */
function addCharacterFormattingMappings(styleMap, styleInfo) {
  // Basic formatting mappings
  styleMap.push("r[bold] => strong");
  styleMap.push("r[italic] => em");
  styleMap.push("r[underline] => span.docx-underline");
  styleMap.push("r[strikethrough] => s");
  styleMap.push("r[subscript] => sub");
  styleMap.push("r[superscript] => sup");
  
  // Enhanced italic mappings for better coverage
  styleMap.push("r[italic=true] => em");
  styleMap.push("r[font-style='italic'] => em");
  
  // Font family mappings
  styleMap.push("r[data-font-ascii] => span.docx-font-ascii");
  styleMap.push("r[data-font-hansi] => span.docx-font-hansi");
  styleMap.push("r[data-font-eastasia] => span.docx-font-eastasia");
  styleMap.push("r[data-font-cs] => span.docx-font-cs");
  
  // Font size mappings
  styleMap.push("r[data-font-size] => span.docx-font-size");
  styleMap.push("r[data-font-size-cs] => span.docx-font-size-cs");
  
  // Color mappings
  styleMap.push("r[data-color] => span.docx-color");
  styleMap.push("r[data-color-theme] => span.docx-color-theme");
  styleMap.push("r[data-color-tint] => span.docx-color-tint");
  styleMap.push("r[data-color-shade] => span.docx-color-shade");
  
  // Background/highlight mappings
  styleMap.push("r[data-highlight] => span.docx-highlight");
  styleMap.push("r[data-shd-fill] => span.docx-shd-fill");
  styleMap.push("r[data-shd-color] => span.docx-shd-color");
  styleMap.push("r[data-shd-pattern] => span.docx-shd-pattern");
  
  // Text effects mappings
  styleMap.push("r[data-double-strikethrough] => span.docx-double-strikethrough");
  styleMap.push("r[data-caps] => span.docx-caps");
  styleMap.push("r[data-small-caps] => span.docx-small-caps");
  styleMap.push("r[data-vanish] => span.docx-vanish");
  styleMap.push("r[data-web-hidden] => span.docx-web-hidden");
  styleMap.push("r[data-emboss] => span.docx-emboss");
  styleMap.push("r[data-imprint] => span.docx-imprint");
  styleMap.push("r[data-outline] => span.docx-outline");
  styleMap.push("r[data-shadow] => span.docx-shadow");
  
  // Character spacing mappings
  styleMap.push("r[data-spacing] => span.docx-spacing");
  styleMap.push("r[data-w] => span.docx-w");
  styleMap.push("r[data-kern] => span.docx-kern");
  styleMap.push("r[data-position] => span.docx-position");
  
  // Character style mappings with enhanced properties
  Object.entries(styleInfo.styles?.character || {}).forEach(([id, style]) => {
    const safeClassName = id.toLowerCase().replace(/[^a-z0-9-_]/g, "-");
    let mapping = `r[style-name='${style.name}'] => span.docx-c-${safeClassName}`;
    
    // Add data attributes for character properties
    if (style.italic) {
      mapping = `r[style-name='${style.name}'] => em.docx-c-${safeClassName}`;
    }
    if (style.bold) {
      mapping = mapping.replace('span.', 'strong.');
    }
    
    styleMap.push(mapping);
  });
  
  // Debug logging for character mappings
  console.log(`Applied ${styleMap.filter(m => m.includes('docx-')).length} enhanced character formatting mappings`);
}

/**
 * Add comprehensive paragraph formatting mappings
 */
function addParagraphFormattingMappings(styleMap, styleInfo) {
  // Alignment mappings
  styleMap.push("p[data-jc='left'] => p.docx-align-left");
  styleMap.push("p[data-jc='center'] => p.docx-align-center");
  styleMap.push("p[data-jc='right'] => p.docx-align-right");
  styleMap.push("p[data-jc='both'] => p.docx-align-justify");
  styleMap.push("p[data-jc='distribute'] => p.docx-align-distribute");
  
  // Indentation mappings
  styleMap.push("p[data-ind-left] => p.docx-ind-left");
  styleMap.push("p[data-ind-right] => p.docx-ind-right");
  styleMap.push("p[data-ind-first-line] => p.docx-ind-first-line");
  styleMap.push("p[data-ind-hanging] => p.docx-ind-hanging");
  styleMap.push("p[data-ind-start] => p.docx-ind-start");
  styleMap.push("p[data-ind-end] => p.docx-ind-end");
  
  // Spacing mappings
  styleMap.push("p[data-spacing-before] => p.docx-spacing-before");
  styleMap.push("p[data-spacing-after] => p.docx-spacing-after");
  styleMap.push("p[data-spacing-line] => p.docx-spacing-line");
  styleMap.push("p[data-spacing-line-rule] => p.docx-spacing-line-rule");
  styleMap.push("p[data-spacing-before-auto] => p.docx-spacing-before-auto");
  styleMap.push("p[data-spacing-after-auto] => p.docx-spacing-after-auto");
  
  // Shading mappings
  styleMap.push("p[data-shd-fill] => p.docx-shd-fill");
  styleMap.push("p[data-shd-color] => p.docx-shd-color");
  styleMap.push("p[data-shd-pattern] => p.docx-shd-pattern");
  
  // Page properties mappings
  styleMap.push("p[data-page-break-before] => p.docx-page-break-before");
  styleMap.push("p[data-keep-next] => p.docx-keep-next");
  styleMap.push("p[data-keep-lines] => p.docx-keep-lines");
  styleMap.push("p[data-widow-control] => p.docx-widow-control");
  
  // Frame properties mappings
  styleMap.push("p[data-frame-w] => p.docx-frame");
  styleMap.push("p[data-frame-h] => p.docx-frame");
  styleMap.push("p[data-frame-x] => p.docx-frame");
  styleMap.push("p[data-frame-y] => p.docx-frame");
  styleMap.push("p[data-frame-anchor] => p.docx-frame");
  
  console.log('Applied comprehensive paragraph formatting mappings');
}

/**
 * Add comprehensive table formatting mappings
 */
function addTableFormattingMappings(styleMap, styleInfo) {
  // Table mappings
  styleMap.push("table[data-tbl-w] => table.docx-table");
  styleMap.push("table[data-tbl-ind] => table.docx-table-ind");
  styleMap.push("table[data-tbl-layout] => table.docx-table-layout");
  styleMap.push("table[data-tbl-overlap] => table.docx-table-overlap");
  
  // Table row mappings
  styleMap.push("tr[data-tr-height] => tr.docx-tr");
  styleMap.push("tr[data-tr-height-rule] => tr.docx-tr");
  styleMap.push("tr[data-tr-cant-split] => tr.docx-tr-cant-split");
  styleMap.push("tr[data-tr-header] => tr.docx-tr-header");
  
  // Table cell mappings
  styleMap.push("td[data-tc-w] => td.docx-tc");
  styleMap.push("td[data-tc-borders] => td.docx-tc-borders");
  styleMap.push("td[data-tc-shd] => td.docx-tc-shd");
  styleMap.push("td[data-tc-mar] => td.docx-tc-margins");
  styleMap.push("td[data-tc-valign] => td.docx-tc-valign");
  styleMap.push("td[data-tc-fit-text] => td.docx-tc-fit-text");
  styleMap.push("td[data-tc-no-wrap] => td.docx-tc-no-wrap");
  styleMap.push("td[data-tc-grid-span] => td.docx-tc-grid-span");
  styleMap.push("td[data-tc-vmerge] => td.docx-tc-vmerge");
  
  // Table style mappings
  Object.entries(styleInfo.styles?.table || {}).forEach(([id, style]) => {
    const safeClassName = id.toLowerCase().replace(/[^a-z0-9-_]/g, "-");
    styleMap.push(`table[style-name='${style.name}'] => table.docx-tbl-${safeClassName}`);
  });
  
  console.log('Applied comprehensive table formatting mappings');
}

/**
 * Add comprehensive paragraph formatting attributes to element
 */
function addParagraphFormattingAttributes(element, styleInfo) {
  if (!element.styleName) return element;
  
  const paragraphStyle = findStyleByName(styleInfo.styles?.paragraph, element.styleName);
  if (!paragraphStyle) return element;
  
  const attributes = {};
  
  // Add alignment attributes
  if (paragraphStyle.alignment) {
    attributes['data-jc'] = paragraphStyle.alignment;
  }
  
  // Add indentation attributes
  if (paragraphStyle.indentation) {
    const ind = paragraphStyle.indentation;
    if (ind.left !== undefined) attributes['data-ind-left'] = ind.left;
    if (ind.right !== undefined) attributes['data-ind-right'] = ind.right;
    if (ind.firstLine !== undefined) attributes['data-ind-first-line'] = ind.firstLine;
    if (ind.hanging !== undefined) attributes['data-ind-hanging'] = ind.hanging;
    if (ind.start !== undefined) attributes['data-ind-start'] = ind.start;
    if (ind.end !== undefined) attributes['data-ind-end'] = ind.end;
  }
  
  // Add spacing attributes
  if (paragraphStyle.spacing) {
    const spacing = paragraphStyle.spacing;
    if (spacing.before !== undefined) attributes['data-spacing-before'] = spacing.before;
    if (spacing.after !== undefined) attributes['data-spacing-after'] = spacing.after;
    if (spacing.line !== undefined) attributes['data-spacing-line'] = spacing.line;
    if (spacing.lineRule) attributes['data-spacing-line-rule'] = spacing.lineRule;
    if (spacing.beforeAuto) attributes['data-spacing-before-auto'] = 'true';
    if (spacing.afterAuto) attributes['data-spacing-after-auto'] = 'true';
  }
  
  // Add shading attributes
  if (paragraphStyle.shading) {
    const shd = paragraphStyle.shading;
    if (shd.fill) attributes['data-shd-fill'] = shd.fill;
    if (shd.color) attributes['data-shd-color'] = shd.color;
    if (shd.pattern) attributes['data-shd-pattern'] = shd.pattern;
  }
  
  // Add page properties attributes
  if (paragraphStyle.pageProperties) {
    const page = paragraphStyle.pageProperties;
    if (page.pageBreakBefore) attributes['data-page-break-before'] = 'true';
    if (page.keepNext) attributes['data-keep-next'] = 'true';
    if (page.keepLines) attributes['data-keep-lines'] = 'true';
  }
  
  // Add widow control attributes
  if (paragraphStyle.widowControl !== undefined) {
    attributes['data-widow-control'] = paragraphStyle.widowControl.toString();
  }
  
  // Add frame properties attributes
  if (paragraphStyle.frameProperties) {
    const frame = paragraphStyle.frameProperties;
    if (frame.w !== undefined) attributes['data-frame-w'] = frame.w;
    if (frame.h !== undefined) attributes['data-frame-h'] = frame.h;
    if (frame.x !== undefined) attributes['data-frame-x'] = frame.x;
    if (frame.y !== undefined) attributes['data-frame-y'] = frame.y;
    if (frame.anchor) attributes['data-frame-anchor'] = frame.anchor;
  }
  
  return {
    ...element,
    attributes: { ...element.attributes, ...attributes }
  };
}

/**
 * Add comprehensive character formatting attributes to element
 */
function addCharacterFormattingAttributes(element, styleInfo) {
  if (!element.styleName) return element;
  
  const characterStyle = findStyleByName(styleInfo.styles?.character, element.styleName);
  if (!characterStyle) return element;
  
  const attributes = {};
  
  // Add font attributes
  if (characterStyle.fonts) {
    const fonts = characterStyle.fonts;
    if (fonts.ascii) attributes['data-font-ascii'] = fonts.ascii;
    if (fonts.hAnsi) attributes['data-font-hansi'] = fonts.hAnsi;
    if (fonts.eastAsia) attributes['data-font-eastasia'] = fonts.eastAsia;
    if (fonts.cs) attributes['data-font-cs'] = fonts.cs;
  }
  
  // Add font size attributes
  if (characterStyle.fontSize !== undefined) {
    attributes['data-font-size'] = characterStyle.fontSize;
  }
  if (characterStyle.fontSizeCs !== undefined) {
    attributes['data-font-size-cs'] = characterStyle.fontSizeCs;
  }
  
  // Add color attributes
  if (characterStyle.color) {
    const color = characterStyle.color;
    if (color.val) attributes['data-color'] = color.val;
    if (color.themeColor) attributes['data-color-theme'] = color.themeColor;
    if (color.themeTint) attributes['data-color-tint'] = color.themeTint;
    if (color.themeShade) attributes['data-color-shade'] = color.themeShade;
  }
  
  // Add highlight/background attributes
  if (characterStyle.highlight) {
    attributes['data-highlight'] = characterStyle.highlight;
  }
  if (characterStyle.shading) {
    const shd = characterStyle.shading;
    if (shd.fill) attributes['data-shd-fill'] = shd.fill;
    if (shd.color) attributes['data-shd-color'] = shd.color;
    if (shd.pattern) attributes['data-shd-pattern'] = shd.pattern;
  }
  
  // Add text effects attributes
  if (characterStyle.doubleStrikethrough) attributes['data-double-strikethrough'] = 'true';
  if (characterStyle.caps) attributes['data-caps'] = 'true';
  if (characterStyle.smallCaps) attributes['data-small-caps'] = 'true';
  if (characterStyle.vanish) attributes['data-vanish'] = 'true';
  if (characterStyle.webHidden) attributes['data-web-hidden'] = 'true';
  if (characterStyle.emboss) attributes['data-emboss'] = 'true';
  if (characterStyle.imprint) attributes['data-imprint'] = 'true';
  if (characterStyle.outline) attributes['data-outline'] = 'true';
  if (characterStyle.shadow) attributes['data-shadow'] = 'true';
  
  // Add character spacing attributes
  if (characterStyle.spacing !== undefined) {
    attributes['data-spacing'] = characterStyle.spacing;
  }
  if (characterStyle.w !== undefined) {
    attributes['data-w'] = characterStyle.w;
  }
  if (characterStyle.kern !== undefined) {
    attributes['data-kern'] = characterStyle.kern;
  }
  if (characterStyle.position !== undefined) {
    attributes['data-position'] = characterStyle.position;
  }
  
  return {
    ...element,
    attributes: { ...element.attributes, ...attributes }
  };
}

/**
 * Add comprehensive table formatting attributes to element
 */
function addTableFormattingAttributes(element, styleInfo) {
  if (!element.styleName) return element;
  
  const tableStyle = findStyleByName(styleInfo.styles?.table, element.styleName);
  if (!tableStyle) return element;
  
  const attributes = {};
  
  // Add table properties attributes
  if (tableStyle.tableProperties) {
    const tblPr = tableStyle.tableProperties;
    if (tblPr.width !== undefined) attributes['data-tbl-w'] = tblPr.width;
    if (tblPr.indent !== undefined) attributes['data-tbl-ind'] = tblPr.indent;
    if (tblPr.layout) attributes['data-tbl-layout'] = tblPr.layout;
    if (tblPr.overlap !== undefined) attributes['data-tbl-overlap'] = tblPr.overlap;
  }
  
  return {
    ...element,
    attributes: { ...element.attributes, ...attributes }
  };
}

/**
 * Add comprehensive table row formatting attributes to element
 */
function addTableRowFormattingAttributes(element, styleInfo) {
  const attributes = {};
  
  // Add table row properties from element properties
  if (element.height !== undefined) {
    attributes['data-tr-height'] = element.height;
  }
  if (element.heightRule) {
    attributes['data-tr-height-rule'] = element.heightRule;
  }
  if (element.cantSplit) {
    attributes['data-tr-cant-split'] = 'true';
  }
  if (element.isHeader) {
    attributes['data-tr-header'] = 'true';
  }
  
  return {
    ...element,
    attributes: { ...element.attributes, ...attributes }
  };
}

/**
 * Add comprehensive table cell formatting attributes to element
 */
function addTableCellFormattingAttributes(element, styleInfo) {
  const attributes = {};
  
  // Add table cell properties from element properties
  if (element.width !== undefined) {
    attributes['data-tc-w'] = element.width;
  }
  if (element.borders) {
    attributes['data-tc-borders'] = JSON.stringify(element.borders);
  }
  if (element.shading) {
    attributes['data-tc-shd'] = JSON.stringify(element.shading);
  }
  if (element.margins) {
    attributes['data-tc-mar'] = JSON.stringify(element.margins);
  }
  if (element.verticalAlignment) {
    attributes['data-tc-valign'] = element.verticalAlignment;
  }
  if (element.fitText) {
    attributes['data-tc-fit-text'] = 'true';
  }
  if (element.noWrap) {
    attributes['data-tc-no-wrap'] = 'true';
  }
  if (element.gridSpan !== undefined) {
    attributes['data-tc-grid-span'] = element.gridSpan;
  }
  if (element.verticalMerge) {
    attributes['data-tc-vmerge'] = element.verticalMerge;
  }
  
  return {
    ...element,
    attributes: { ...element.attributes, ...attributes }
  };
}
