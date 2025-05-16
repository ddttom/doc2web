#!/bin/bash
# rename-doc2web-files.sh - Script to rename files to follow doc2web naming convention

echo "=== Renaming files to follow doc2web naming convention ==="
echo ""

# Create a backup directory
BACKUP_DIR="doc2web_original_files_$(date +%Y%m%d_%H%M%S)"
mkdir -p "$BACKUP_DIR"
echo "Created backup directory: $BACKUP_DIR"

# Define file mappings
declare -A file_mappings=(
    ["enhanced-script.js"]="doc2web.js"
    ["init-script.sh"]="init-doc2web.sh"
    ["install-script.sh"]="doc2web-install.sh"
    ["process-script.sh"]="process-find.sh"
    ["run-converter.js"]="doc2web-run.js"
    ["full-readme.md"]="README.md"
    ["enhanced-readme.md"]="_old_enhanced-readme.md"  # Not directly used, but renamed to avoid confusion
)

# Function to backup and rename a file
rename_file() {
    local source_file="$1"
    local target_file="$2"
    
    if [ -f "$source_file" ]; then
        # Backup the file
        cp "$source_file" "$BACKUP_DIR/"
        echo "✓ Backed up: $source_file → $BACKUP_DIR/$source_file"
        
        # Rename the file
        mv "$source_file" "$target_file"
        echo "✓ Renamed: $source_file → $target_file"
    else
        echo "⚠ Warning: $source_file not found, skipping"
    fi
}

# Process each file
echo "Renaming files..."
for source_file in "${!file_mappings[@]}"; do
    target_file="${file_mappings[$source_file]}"
    rename_file "$source_file" "$target_file"
done

echo ""
echo "=== File renaming complete ==="
echo "Original files were backed up to: $BACKUP_DIR"
echo ""

# Check for missing core files
missing_files=()
core_files=("doc2web.js" "markdownify.js" "style-extractor.js" "docx-style-parser.js")

echo "Checking for required core files..."
for file in "${core_files[@]}"; do
    if [ ! -f "$file" ]; then
        missing_files+=("$file")
    fi
done

if [ ${#missing_files[@]} -gt 0 ]; then
    echo "⚠ Warning: The following core files are still missing:"
    for file in "${missing_files[@]}"; do
        echo "  - $file"
    done
    echo "Please ensure all core files are present before running doc2web."
else
    echo "✓ All core files are present."
fi

echo ""
echo "Next steps:"
echo "1. Make scripts executable: chmod +x *.sh doc2web-run.js"
echo "2. Initialize the project: ./init-doc2web.sh"
echo "3. Test the converter: node doc2web.js [your-docx-file]"
