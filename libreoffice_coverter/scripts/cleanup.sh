#!/bin/bash

# Find and delete .DS_Store files and ._ files
find .. \( -name '*.DS_Store' -o -name '._*' \) -type f -delete

echo "Cleanup complete."