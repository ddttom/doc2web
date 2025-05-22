# doc2web Troubleshooting Guide

## Quick Fix Summary

The application has been fixed with the following improvements:

### 1. Enhanced Error Handling and Logging
- Better error reporting throughout the processing pipeline
- Detailed logging to help identify issues
- Graceful fallbacks when components fail

### 2. Fixed DOM Manipulation Issues
- Resolved content loss during HTML processing
- Improved JSDOM serialization handling
- Better preservation of document structure

### 3. Enhanced Validation
- Input file validation before processing
- Content validation during each processing step
- Output validation before saving files

### 4. Improved Accessibility Processing
- Fixed ARIA landmarks implementation
- Better error handling in accessibility enhancements
- Safer DOM manipulation in accessibility module

## How to Test the Fix

### 1. Run the Debug Script
First, test with the new debug script to identify any issues:

```bash
node debug-test.js path/to/your/document.docx
```

This will:
- Test each component individually
- Validate the full processing pipeline
- Generate detailed diagnostic information
- Save test output to `debug-output/` directory

### 2. Test with Main Application
```bash
node doc2web.js path/to/your/document.docx
```

### 3. Check the Output
Look for the following in the output directory:
- `document.html` - Should contain properly structured HTML
- `document.css` - Should contain generated styles
- `document.md` - Should contain converted Markdown (if not using --html-only)
- `images/` - Should contain extracted images

## Common Issues and Solutions

### Issue: "HTML output is too short or empty"

**Symptoms:**
- Generated HTML file is very small (< 100 characters)
- HTML contains only basic structure without content

**Solutions:**
1. Run the debug script to identify where content is lost
2. Check if the DOCX file is corrupted:
   ```bash
   # Try opening the file in Microsoft Word
   # Check file size: should be > 1KB for most documents
   ```
3. Verify file permissions:
   ```bash
   ls -la path/to/your/document.docx
   ```

### Issue: "Error parsing DOCX styles"

**Symptoms:**
- Errors about XML parsing or missing files
- Fallback styles being used

**Solutions:**
1. Verify the file is a valid DOCX document (not DOC or other format)
2. Check if the document was created with a supported version of Word
3. Try saving the document in Word again before processing

### Issue: "CSS output is very short"

**Symptoms:**
- Generated CSS file is minimal
- HTML doesn't have proper styling

**Solutions:**
1. Check if the document has custom styles defined
2. Verify that styles.xml exists in the DOCX file
3. The application generates CSS based on actual document styles - minimal CSS may be correct for simple documents

### Issue: "Images not extracted"

**Symptoms:**
- Images directory is empty
- Image references in HTML are broken

**Solutions:**
1. Ensure the images directory has write permissions
2. Check if the document actually contains images
3. Verify that images in the document are properly embedded (not linked)

### Issue: "Numbering not preserved"

**Symptoms:**
- Lists appear without proper numbering
- Heading numbers are missing

**Solutions:**
1. Ensure the document uses Word's built-in numbering features
2. Check if numbering.xml exists in the DOCX file
3. Verify that the document doesn't use manual numbering (typed numbers)

## Debug Output Analysis

When running the debug script, check these key metrics:

### Good Results:
- HTML length: > 1000 characters for typical documents
- CSS length: > 500 characters
- Numbering context entries: > 0 if document has numbered items
- No error messages in the output

### Warning Signs:
- HTML length < 100 characters
- Multiple error messages
- "Mammoth conversion result is very short"
- "Body is empty before serialization"

## Performance Tips

### For Large Documents:
```bash
# Increase Node.js memory limit
NODE_OPTIONS=--max-old-space-size=4096 node doc2web.js large-document.docx
```

### For Batch Processing:
```bash
# Process one file at a time for debugging
node doc2web.js directory/ --html-only
```

### For Complex Documents:
- Documents with many styles may take longer to process
- Track changes add processing overhead
- Large images increase processing time

## Getting Help

### 1. Check Debug Output
Always run the debug script first:
```bash
node debug-test.js problematic-document.docx
```

### 2. Check Log Messages
Look for specific error patterns:
- "XML parsing error" - Document structure issue
- "DOM manipulation error" - HTML processing issue  
- "Serialization error" - Output generation issue

### 3. Verify Environment
```bash
node --version  # Should be 14.x or higher
npm list        # Check installed dependencies
```

### 4. Test with Simple Document
Create a minimal test document:
1. Create a new Word document
2. Add a few paragraphs with basic formatting
3. Save as .docx
4. Test with doc2web

## Recent Fixes Applied

1. **Fixed HTML Generator** (`lib/html/html-generator.js`)
   - Enhanced error handling and logging
   - Better content preservation during DOM manipulation
   - Improved validation at each processing step

2. **Fixed Main Application** (`doc2web.js`)
   - Better input validation
   - Enhanced error reporting
   - Graceful error recovery

3. **Fixed Accessibility Processor** (`lib/accessibility/wcag-processor.js`)
   - Safer DOM manipulation
   - Better error handling
   - Fixed ARIA landmarks implementation

4. **Added Debug Script** (`debug-test.js`)
   - Comprehensive testing of all components
   - Detailed diagnostic output
   - Test file generation for inspection

## File Structure After Processing

```
output/
└── path/
    └── to/
        └── original/
            ├── document.html    # Main HTML output
            ├── document.css     # Generated styles
            ├── document.md      # Markdown version
            └── images/          # Extracted images
                ├── image1.png
                └── image2.jpg
```

The output preserves the original directory structure and creates all necessary files with proper naming conventions.