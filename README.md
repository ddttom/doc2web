# doc2web

doc2web is a powerful tool for converting Microsoft Word (.docx) documents to web-friendly formats, including Markdown and HTML with preserved styling.

## Overview

This tool extracts content from DOCX files while maintaining:

- Document structure and formatting
- Images and tables
- Styles and layout
- Unicode and multilingual content
- Table of Contents (TOC) with clickable navigation links
- Hierarchical list numbering and structure
- WCAG 2.1 Level AA accessibility compliance
- Document metadata preservation
- Track changes visualization and handling
- Exact DOCX numbering preservation through XML introspection
- Section IDs for direct navigation to numbered headings and paragraphs
- Return to top button for improved navigation and user experience

The app treats generated CSS and HTML as ephemeral - they are regenerated for each document based on its unique structure. All styling is extracted from the DOCX XML structure, not hardcoded patterns.

## Project Structure

The codebase follows a modular architecture with organized library modules:

```bash
doc2web/
├── doc2web.js             # Main entry point and orchestrator
├── markdownify.js         # HTML to Markdown converter
├── lib/                   # Modular library code
│   ├── index.js           # Main entry point that re-exports public API
│   ├── xml/               # XML parsing utilities
│   ├── parsers/           # DOCX parsing modules
│   ├── html/              # HTML processing modules
│   │   ├── generators/    # HTML generation utilities
│   │   └── processors/    # Content processing modules
│   ├── css/               # CSS generation modules
│   │   └── generators/    # CSS generation utilities
│   ├── accessibility/     # Accessibility enhancement modules
│   └── utils/             # Utility functions
```

For detailed technical architecture information, see [`docs/architecture.md`](docs/architecture.md).

## Getting Started

### Initial Setup

1. Clone or download all project files to a directory:

   ```bash
   git clone https://github.com/ddttom/doc2web.git
   cd doc2web
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

3. Make the helper scripts executable (Unix/Linux/macOS):

   ```bash
   chmod +x *.sh
   ```

## Usage

### Command Line

Process a single DOCX file:

```bash
node doc2web.js path/to/document.docx
```

Process all DOCX files in a directory:

```bash
node doc2web.js path/to/directory/
```

Process a list of files:

```bash
node doc2web.js file-list.txt --list
```

### Interactive Mode

For a user-friendly interface:

```bash
node doc2web-run.js
```

### Processing Find Command Output

```bash
find . -name "*.docx" > docx-files.txt
./process-find.sh docx-files.txt
```

## Output

All processed files are stored in the `./output` directory, preserving the original directory structure:

```bash
output/
└── original/
    └── path/
        ├── filename.md      # Markdown version
        ├── filename.html    # HTML version with styles
        ├── filename.css     # Extracted CSS styles
        └── images/          # Extracted images
```

## Options

- `--html-only`: Generate only HTML output, skip markdown
- `--list`: Treat the input file as a list of files to process

## Key Features

### Section IDs and Navigation

doc2web automatically generates section IDs for numbered headings based on their hierarchical position:

- `1. Introduction` becomes `id="section-1"`
- `1.2 Overview` becomes `id="section-1-2"`
- `1.2.a Details` becomes `id="section-1-2-a"`

This enables direct linking to sections, smooth scrolling navigation, and accessibility improvements.

### Table of Contents Navigation

doc2web automatically converts Table of Contents entries into clickable navigation links:

- **Automatic Detection**: Identifies TOC sections in DOCX documents and removes page numbers for web use
- **Smart Linking**: Intelligently matches TOC entries to corresponding document sections using hierarchical patterns
- **Accessibility Compliant**: WCAG 2.1 Level AA compliant navigation with proper ARIA attributes and keyboard support
- **Visual Enhancement**: Hover and focus states with smooth scrolling to target sections

### Document Header Processing

Automatically extracts and processes document headers from DOCX files:

- **Header Extraction**: Extracts content from Word header files and document analysis
- **Formatting Preservation**: Maintains original fonts, sizes, colors, alignment, and styling
- **Graphics Inclusion**: Preserves images and graphics within headers with proper positioning
- **Responsive Design**: Headers adapt to different screen sizes and print media
- **Accessibility**: Proper semantic HTML with ARIA roles for screen readers

### Return to Top Button

Provides convenient navigation back to the document start:

- **Always Visible**: Button remains visible in the bottom left corner for easy access
- **Theme-Aware Styling**: Automatically matches document colors and fonts for visual consistency
- **Smooth Scrolling**: Animated scrolling to the top of the document for better user experience
- **Accessibility Compliant**: Full keyboard navigation support with proper ARIA labels
- **Responsive Design**: Adapts to different screen sizes and devices
- **Visual Feedback**: Hover and focus states provide clear interaction cues

### Enhanced Table Formatting

Comprehensive table formatting improvements:

- **Professional Styling**: Clean borders, consistent padding, and alternating row colors
- **Semantic HTML Structure**: Proper `<thead>` and `<tbody>` sections with `<th>` elements
- **Accessibility Enhancements**: Header cells with `scope="col"` attributes and table captions
- **Responsive Design**: Tables adapt to different screen sizes with horizontal scrolling
- **Automatic Header Detection**: Intelligently converts first row to proper table headers

### Accessibility Compliance (WCAG 2.1 Level AA)

Generated HTML meets WCAG 2.1 Level AA accessibility standards:

- Proper semantic structure with HTML5 sectioning elements and ARIA landmarks
- Accessible tables with captions, header cells, and proper scope attributes
- Images with appropriate alt text and figure/figcaption elements
- Proper heading hierarchy with no skipped levels
- Skip navigation links for keyboard users
- High contrast mode support and screen reader compatibility

### Metadata Preservation

Extracts and preserves document metadata in the generated HTML:

- Title, description, and keywords
- Creation and modification dates
- Document statistics (pages, words, characters)
- Dublin Core metadata
- Open Graph and Twitter Card metadata for social sharing
- JSON-LD structured data for search engines

### Track Changes Support

Handles tracked changes in documents:

- Visual representation of insertions, deletions, moves, and formatting changes
- Multiple view modes (show changes, hide changes, accept all, reject all)
- Author and date information for each change
- Track changes legend with toggle functionality
- Keyboard shortcut (Alt+T) to toggle track changes visibility

## API Usage

```javascript
const { extractAndApplyStyles } = require('./lib');

async function convertDocument(docxPath) {
  const result = await extractAndApplyStyles(docxPath);
  console.log('Conversion completed:', result.html ? 'Success' : 'Failed');
}

convertDocument('document.docx').catch(console.error);
```

For detailed API options and configuration, see [`docs/architecture.md`](docs/architecture.md).

## Troubleshooting

### Using the Debug Script

If you encounter issues with document conversion, use the debug script:

```bash
node debug-test.js path/to/your/document.docx
```

This creates a debug-output/ directory with test files and diagnostic information.

### Common Issues

1. **Conversion Failures**
   - Ensure DOCX file is not corrupted or password-protected
   - Check that all required dependencies are installed
   - For large files, process individually rather than in batch mode

2. **Missing Content**
   - Update to latest version which includes content preservation fixes
   - Use the debug script to identify where content is being lost
   - Check error logs for specific issues

3. **Formatting Issues**
   - Ensure you're using the latest version for formatting fixes
   - Check that HTML files are > 1000 characters for typical documents
   - Verify CSS files contain generated styles appropriate for the document

4. **Accessibility Issues**
   - Update to latest version which includes accessibility processor fixes
   - Check that the enhanceAccessibility option is enabled
   - Use the debug script to test accessibility features

### Performance Tips

- For very large documents, increase Node.js memory limit:

  ```bash
  NODE_OPTIONS=--max-old-space-size=4096 node doc2web.js large-document.docx
  ```

- Process complex documents individually for better error isolation
- Use the debug script for troubleshooting specific conversion issues

## Recent Updates

**v1.3.3 (2025-06-02)**

- HTML formatting optimization with compact output
- Performance enhancement through streamlined processing

**v1.3.2 (2025-06-02)**

- Hanging indentation fixes with enhanced CSS specificity
- TOC page number removal for web-appropriate navigation

**v1.3.1 (2025-05-31)**

- Hanging margins implementation for TOC and numbered content
- Header image extraction and positioning functionality
- Table formatting enhancement with professional styling

**v1.3.0 (2025-05-26)**

- Modular refactoring with focused, maintainable modules
- Enhanced architecture with clear separation of concerns
- API preservation with improved internal structure

For complete version history and detailed technical implementation information, see [`docs/architecture.md`](docs/architecture.md).

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

## License

This project is licensed under the MIT License.
