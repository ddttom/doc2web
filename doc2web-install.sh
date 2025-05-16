#!/bin/bash
# doc2web-install.sh - Installation script for doc2web

echo "Installing doc2web..."

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "Node.js is not installed. Please install Node.js first."
    exit 1
fi

# Create package.json if it doesn't exist
if [ ! -f "package.json" ]; then
    echo "Creating package.json..."
    cat > package.json << EOF
{
  "name": "doc2web",
  "version": "1.0.0",
  "description": "Convert DOCX documents to markdown and HTML with style preservation",
  "main": "doc2web.js",
  "scripts": {
    "start": "node doc2web.js",
    "html": "node doc2web.js --html-only"
  },
  "keywords": [
    "docx",
    "markdown",
    "converter",
    "word",
    "document",
    "styles"
  ],
  "author": "",
  "license": "MIT",
  "dependencies": {
    "mammoth": "^1.6.0",
    "jsdom": "^22.1.0",
    "jszip": "^3.10.1",
    "xmldom": "^0.6.0",
    "xpath": "^0.0.33"
  }
}
EOF
fi

# Install dependencies
echo "Installing dependencies..."
npm install

# Create output directory
mkdir -p ./output

# Make scripts executable
chmod +x doc2web-run.js
chmod +x process-find.sh

echo "Installation complete!"
echo ""
echo "Usage examples:"
echo "  node doc2web.js document.docx                 # Process a single file"
echo "  node doc2web.js folder/                       # Process a directory"
echo "  node doc2web.js file-list.txt --list          # Process a list of files"
echo "  ./process-find.sh find-output.txt             # Process find command output"
echo "  node doc2web-run.js                           # Run interactive interface"
echo ""
echo "Happy converting!"
