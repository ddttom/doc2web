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
  
  // Character style italic mappings
  Object.entries(styleInfo.styles?.character || {}).forEach(([id, style]) => {
    if (style.italic) {
      const safeClassName = id.toLowerCase().replace(/[^a-z0-9-_]/g, "-");
      styleMap.push(`r[style-name='${style.name}'] => em.docx-c-${safeClassName}`);
    }
  });
  
  // Debug logging for italic mappings
  const italicMappings = styleMap.filter(map => 
    map.includes('italic') || map.includes('=> em')
  );
  if (italicMappings.length > 0) {
    console.log('Applied italic style mappings:', italicMappings);
  }
  
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
