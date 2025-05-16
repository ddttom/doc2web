#!/bin/bash
# processFindOutput.sh - Helper script to process the output of find command

# Check if parameter is provided
if [ $# -eq 0 ]; then
    echo "Usage: $0 <find-output-file.txt>"
    exit 1
fi

# Input file with find command output
INPUT_FILE=$1

# Create a temporary file for processing
TEMP_FILE=$(mktemp)

# Clean up input file by removing any empty lines and normalizing paths
cat "$INPUT_FILE" | grep -v "^$" | sed 's/\r//g' > "$TEMP_FILE"

# Run the doc2web script
node doc2web.js "$TEMP_FILE" --list

# Remove temporary file
rm "$TEMP_FILE"

echo "Batch processing complete!"
