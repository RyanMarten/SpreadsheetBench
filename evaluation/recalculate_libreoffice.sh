#!/bin/bash
# Recalculate formulas in all xlsx files in a directory using LibreOffice headless.
# This replaces open_spreadsheet.py (win32com/Excel) for cross-platform evaluation.
#
# Usage:
#   bash recalculate_libreoffice.sh /path/to/spreadsheets
#   bash recalculate_libreoffice.sh /path/to/spreadsheets --recursive
#
# Requires: libreoffice (apt install libreoffice-calc or brew install --cask libreoffice)

set -euo pipefail

if [ $# -lt 1 ]; then
    echo "Usage: $0 <directory> [--recursive]"
    exit 1
fi

DIR_PATH="$1"
RECURSIVE="${2:-}"

if [ ! -d "$DIR_PATH" ]; then
    echo "ERROR: Not a valid directory: $DIR_PATH"
    exit 1
fi

# Find xlsx files
if [ "$RECURSIVE" = "--recursive" ]; then
    FILES=$(find "$DIR_PATH" -name "*.xlsx" -type f | sort)
else
    FILES=$(find "$DIR_PATH" -maxdepth 1 -name "*.xlsx" -type f | sort)
fi

COUNT=$(echo "$FILES" | grep -c . || true)
if [ "$COUNT" -eq 0 ]; then
    echo "No .xlsx files found in $DIR_PATH"
    exit 0
fi

# Detect LibreOffice binary
if command -v libreoffice &>/dev/null; then
    SOFFICE="libreoffice"
elif [ -x "/Applications/LibreOffice.app/Contents/MacOS/soffice" ]; then
    SOFFICE="/Applications/LibreOffice.app/Contents/MacOS/soffice"
else
    echo "ERROR: LibreOffice not found. Install with: brew install --cask libreoffice (macOS) or apt install libreoffice-calc (Linux)"
    exit 1
fi

echo "Recalculating formulas in $COUNT file(s) using LibreOffice ($SOFFICE)..."

SUCCESS=0
FAILED=0

for xlsx in $FILES; do
    basename=$(basename "$xlsx")
    tmpdir=$(mktemp -d)

    if "$SOFFICE" --headless --calc --convert-to "xlsx:Calc MS Excel 2007 XML" --outdir "$tmpdir" "$xlsx" 2>/dev/null; then
        name="${basename%.xlsx}"
        if [ -f "$tmpdir/$name.xlsx" ]; then
            mv "$tmpdir/$name.xlsx" "$xlsx"
            SUCCESS=$((SUCCESS + 1))
        else
            echo "WARNING: Converted file not found for $basename"
            FAILED=$((FAILED + 1))
        fi
    else
        echo "ERROR: LibreOffice failed on $basename"
        FAILED=$((FAILED + 1))
    fi

    rm -rf "$tmpdir"
done

echo ""
echo "Done. $SUCCESS succeeded, $FAILED failed out of $COUNT files."
