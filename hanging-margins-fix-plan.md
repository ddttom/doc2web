# Hanging Margins Fix Plan for doc2web

## Problem Statement

The doc2web application has multiple hanging margin and text alignment issues:

1. **Main Content**: Numbered headings like "c. Redelegation of the .TG domain..." don't display with proper hanging indents
2. **Table of Contents**: TOC entries are truncated and not properly aligned with hanging indents
3. **Text Wrapping**: Long text should wrap with proper hanging margin alignment, not truncate

## Root Cause Analysis

From code analysis, the issues stem from:
- Incomplete DOCX hanging indent extraction
- CSS generation not properly handling hanging margins for all element types
- TOC-specific hanging indent logic missing
- Text overflow and wrapping not configured for hanging indents

## Comprehensive Fix Plan

### Phase 1: Enhanced DOCX Parsing for Hanging Indents

**Files to modify:**
- `lib/parsers/style-parser.js` - Enhance indentation extraction
- `lib/parsers/numbering-parser.js` - Improve hanging indent parsing
- `lib/parsers/toc-parser.js` - Add TOC hanging indent detection
- `lib/html/processors/heading-processor.js` - Fix heading indentation detection

**Tasks:**
1. **Improve hanging indent detection** in style parsing to ensure all hanging indent values are properly extracted from DOCX XML
2. **Enhance numbering definition parsing** to capture precise hanging indent measurements
3. **Add TOC-specific hanging indent parsing** to extract TOC entry indentation from DOCX
4. **Fix heading processor** to properly identify and apply hanging indents to numbered headings

### Phase 2: CSS Generation Overhaul

**Files to modify:**
- `lib/css/generators/paragraph-styles.js` - Fix hanging indent CSS generation
- `lib/css/generators/numbering-styles.js` - Improve numbered element hanging margins
- `lib/css/generators/toc-styles.js` - Fix TOC hanging indents and text wrapping
- `lib/css/generators/utility-styles.js` - Ensure proper default margins

**Tasks:**
1. **Fix CSS generation logic** to properly convert DOCX hanging indent values to CSS `text-indent` and `padding-left`
2. **Implement TOC hanging indent CSS** with proper text wrapping and overflow handling
3. **Add text overflow prevention** using `word-wrap: break-word` and `overflow-wrap: break-word`
4. **Ensure proper CSS specificity** so hanging indents override default paragraph styles
5. **Fix TOC alignment issues** with proper flex layout and hanging indent support

### Phase 3: HTML Structure Enhancement

**Files to modify:**
- `lib/html/processors/numbering-processor.js` - Ensure proper HTML structure for hanging indents
- `lib/html/processors/toc-processor.js` - Fix TOC HTML structure for hanging indents
- `lib/html/structure-processor.js` - Apply hanging indent classes correctly

**Tasks:**
1. **Ensure proper HTML attributes** are applied to elements with hanging indents
2. **Fix TOC HTML structure** to support hanging indents and prevent truncation
3. **Fix class application** so hanging indent styles are properly targeted
4. **Add proper text wrapping attributes** to prevent truncation

### Phase 4: Integration and Testing

**Files to modify:**
- `lib/css/css-generator.js` - Ensure all CSS modules work together
- `lib/html/html-generator.js` - Validate HTML generation

**Tasks:**
1. **Integration testing** to ensure all components work together
2. **CSS specificity validation** to prevent style conflicts
3. **TOC rendering validation** to ensure proper alignment and no truncation
4. **Cross-browser compatibility** testing for hanging indent rendering

## Technical Implementation Details

### Key Issues to Address:

1. **Missing Hanging Indent Detection**: The current code may not be properly detecting hanging indents in all paragraph styles, especially for numbered headings and TOC entries.

2. **CSS Generation Gap**: The hanging indent CSS generation might not be applying to all elements that need it, particularly headings with numbering and TOC entries.

3. **TOC Truncation**: TOC entries are being truncated instead of wrapping with proper hanging indents.

4. **HTML Structure Mismatch**: The HTML structure might not be properly set up to support the CSS hanging indent implementation.

5. **Text Overflow Issues**: Long text is truncating instead of wrapping with hanging margin alignment.

### Expected Outcome:

After implementing this plan:

1. **Main Content**: Numbered headings like "c. Redelegation of the .TG domain..." will display with proper hanging indents where:
   - The number/letter ("c.") appears at the left margin
   - The text content is indented and wraps with proper alignment
   - All subsequent lines align with the first line of text (not the number)

2. **Table of Contents**: TOC entries will display with proper hanging indents where:
   - Long TOC entries wrap properly instead of truncating
   - All lines of wrapped text align consistently
   - Proper spacing and alignment throughout the TOC

3. **Text Wrapping**: All text will wrap properly with hanging margin alignment, preventing truncation and maintaining readability.

This will match Microsoft Word's hanging indent behavior exactly, as the app extracts the precise measurements from the DOCX XML structure.

## Implementation Priority

1. **High Priority**: TOC hanging indents and truncation fixes (most visible issue)
2. **High Priority**: Main content hanging indents for numbered headings
3. **Medium Priority**: General paragraph hanging indent improvements
4. **Low Priority**: Edge case handling and cross-browser compatibility

## Testing Strategy

1. **Unit Tests**: Test each CSS generator module independently
2. **Integration Tests**: Test complete DOCX to HTML conversion with hanging indents
3. **Visual Tests**: Compare output with original DOCX rendering
4. **Edge Case Tests**: Test with various hanging indent values and complex layouts
