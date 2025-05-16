#!/bin/bash
# init-doc2web.sh - Initialize doc2web project for version control

echo "=== Initializing doc2web project ==="
echo ""

# Check if git is installed
if ! command -v git &> /dev/null; then
    echo "Error: Git is not installed. Please install Git first."
    exit 1
fi

# Check if we're already in a git repository
if [ -d ".git" ]; then
    echo "Git repository already exists in this directory."
    echo "Skipping repository initialization."
else
    echo "Initializing Git repository..."
    git init
    
    # Create .gitignore file
    echo "Creating .gitignore file..."
    cat > .gitignore << EOF
# Dependencies
node_modules/

# Output directory (generated files)
output/

# Logs
logs
*.log
npm-debug.log*

# Runtime data
pids
*.pid
*.seed
*.pid.lock

# Coverage directory
coverage/

# Temporary files
.temp
.tmp
.cache
tmp/

# Editor directories and files
.idea/
.vscode/
*.swp
*.swo
*~

# OS specific files
.DS_Store
Thumbs.db
EOF
fi

# Check for required files
required_files=("doc2web.js" "markdownify.js" "style-extractor.js" "docx-style-parser.js" "process-find.sh" "doc2web-run.js" "doc2web-install.sh")
missing_files=()

echo "Checking for required files..."
for file in "${required_files[@]}"; do
    if [ ! -f "$file" ]; then
        missing_files+=("$file")
    fi
done

if [ ${#missing_files[@]} -gt 0 ]; then
    echo "Error: The following required files are missing:"
    for file in "${missing_files[@]}"; do
        echo "  - $file"
    done
    echo "Please ensure all required files are in the current directory."
    exit 1
fi

# Make scripts executable
echo "Making scripts executable..."
chmod +x doc2web-run.js
chmod +x process-find.sh
chmod +x doc2web-install.sh

# Run installation script
echo "Running installation script..."
./doc2web-install.sh

# Create README if it doesn't exist
if [ ! -f "README.md" ]; then
    echo "Creating README.md..."
    cat > README.md << EOF
# doc2web

Convert Microsoft Word (.docx) documents to markdown and HTML with style preservation.

## Installation

```bash
./doc2web-install.sh
```

## Usage

```bash
# Process a single file
node doc2web.js document.docx

# Process a directory
node doc2web.js ./my-documents/

# Process a list of files
node doc2web.js file-list.txt --list

# Interactive mode
node doc2web-run.js
```

## Features

- Converts DOCX to markdown and HTML formats
- Preserves document styling and structure
- Full Unicode and multilingual support
- Processes directories and file lists
- Extracts images and document styles
EOF
fi

# Add files to git and make initial commit
if [ -d ".git" ]; then
    echo "Adding files to Git repository..."
    git add .

    echo "Making initial commit..."
    git commit -m "Initial commit of doc2web project"
    
    # Display GitHub setup instructions
    echo ""
    echo "=== Project initialized successfully! ==="
    echo ""
    echo "To connect this repository to GitHub:"
    echo ""
    echo "1. Create a new repository on GitHub (don't initialize with README or .gitignore)"
    echo "2. Connect your local repository with:"
    echo ""
    echo "   git remote add origin https://github.com/username/doc2web.git"
    echo "   git branch -M main"
    echo "   git push -u origin main"
    echo ""
    echo "Replace 'username' with your GitHub username and 'doc2web' with your repository name."
fi

echo "Project initialization complete!"
