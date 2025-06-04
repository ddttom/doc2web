# Bullet Point Margin Fix Implementation Plan

## Problem Summary
Bullet points in HTML output ignore document margins and position themselves at the viewport edge instead of aligning with regular paragraph content.

## Root Cause
- Document has `margin: 72pt 90pt 72pt 90pt` (from DOCX page margins)
- Bullet lists use `margin: 0 0 0 20pt` (positioned from viewport edge)
- Regular paragraphs respect document margins, bullet points don't

## Solution Overview
Modify bullet list CSS generation to include the document's base left margin in bullet positioning calculations.

## Key Implementation Steps

### 1. Modify `lib/css/generators/numbering-styles.js`
- **Function**: `generateBulletListStyles()`
- **Change**: Accept `defaultParagraphMargin` parameter
- **Update**: `ul.docx-bullet-list` margin calculation to include base margin

### 2. Update CSS Generation Logic
```javascript
// Current (problematic):
margin: 0 0 0 20pt !important;

// New (fixed):
margin: 0 0 0 calc(${defaultParagraphMargin} + 20pt) !important;
```

### 3. Update Function Calls
- Ensure `generateBulletListStyles()` receives the base margin parameter
- Verify margin is properly parsed and applied

### 4. Test Multi-level Lists
- Level 0: Base margin + 0pt
- Level 1: Base margin + 18pt
- Level 2: Base margin + 36pt
- Continue pattern for deeper levels

## Expected Outcome
- Bullet points align with regular paragraph left edge
- Nested bullets maintain proper relative indentation
- Solution works dynamically with any document margin settings
- Visual alignment matches original DOCX layout

## Files to Modify
1. `lib/css/generators/numbering-styles.js` (primary)
2. `lib/css/generators/utility-styles.js` (ensure margin calculation)
3. `lib/css/css-generator.js` (function call updates)

## Testing
- Verify bullet alignment with regular paragraphs
- Test with different document margin settings
- Check nested bullet list hierarchy
- Ensure no regression in other list types
