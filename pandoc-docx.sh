#!/bin/bash

# pandoc-tool.sh
# Converting between DOCX and reference directory for Pandoc template
# Usage: ./pandoc-tool.sh {zip|unzip}

set -e

# Color output
RED='\033[0;31m'
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m'

# Paths
SOURCE_DIR="$(pwd)/reference"
DOCX_PATH="$(pwd)/reference.docx"
ZIP_PATH="$(pwd)/reference.zip"

show_usage() {
    echo "Usage: $0 {zip|unzip}"
    echo "  zip    - Create DOCX from reference directory"
    echo "  unzip  - Extract DOCX to reference directory (with ID cleanup)"
    exit 1
}

# ZIP: Create DOCX from reference directory
do_zip() {
    echo -e "${BLUE}Creating DOCX from reference directory...${NC}"

    if [ ! -d "$SOURCE_DIR" ]; then
        echo -e "${RED}Error: reference directory not found!${NC}"
        exit 1
    fi

    # Cleanup existing files
    [ -f "$ZIP_PATH" ] && rm -f "$ZIP_PATH"
    [ -f "$DOCX_PATH" ] && rm -f "$DOCX_PATH"

    # Create ZIP archive
    cd "$SOURCE_DIR"
    zip -r -q -X -9 "$ZIP_PATH" .

    # Handle media files with no compression if they exist
    if [ -d "word/media" ] && ls word/media/* 1>/dev/null 2>&1; then
        zip -d "$ZIP_PATH" "word/media/*" 2>/dev/null || true
        cd word/media
        for file in *; do
            [ -f "$file" ] && zip -0 -q "$ZIP_PATH" "word/media/$file"
        done
        cd ../..
    fi

    cd - >/dev/null

    # Convert to DOCX
    mv "$ZIP_PATH" "$DOCX_PATH"

    if [ -f "$DOCX_PATH" ]; then
        file_size=$(ls -lh "$DOCX_PATH" | awk '{print $5}')
        echo -e "${GREEN}Successfully created reference.docx (${file_size})${NC}"
        # open "$DOCX_PATH"
    else
        echo -e "${RED}Error: Failed to create DOCX file${NC}"
        exit 1
    fi
}

# UNZIP: Extract DOCX to reference directory
do_unzip() {
    echo -e "${BLUE}Extracting DOCX to reference directory...${NC}"

    if [ ! -f "$DOCX_PATH" ]; then
        echo -e "${RED}Error: reference.docx not found!${NC}"
        exit 1
    fi

    # Setup temp paths
    TEMP_ZIP="/tmp/reference-temp-$$.zip"
    TEMP_EXTRACT_DIR="/tmp/reference-temp-$$"

    # Cleanup function
    cleanup() {
        [ -f "$TEMP_ZIP" ] && rm -f "$TEMP_ZIP"
        [ -d "$TEMP_EXTRACT_DIR" ] && rm -rf "$TEMP_EXTRACT_DIR"
    }
    trap cleanup EXIT

    # Extract DOCX
    cp "$DOCX_PATH" "$TEMP_ZIP"
    mkdir -p "$TEMP_EXTRACT_DIR"
    cd "$TEMP_EXTRACT_DIR"
    unzip -q "$TEMP_ZIP"
    cd - >/dev/null

    # Get file lists for comparison
    get_relative_paths() {
        local base_dir="$1"
        [ ! -d "$base_dir" ] && return
        local original_dir=$(pwd)
        cd "$base_dir"
        find . -type f -print | sed 's|^\./||'
        cd "$original_dir"
    }

    SOURCE_FILES_TMP="/tmp/source_files_$$"
    TARGET_FILES_TMP="/tmp/target_files_$$"

    get_relative_paths "$TEMP_EXTRACT_DIR" | sort >"$SOURCE_FILES_TMP"
    get_relative_paths "$SOURCE_DIR" | sort >"$TARGET_FILES_TMP"

    # Remove obsolete files
    while IFS= read -r rel_path; do
        if [ -n "$rel_path" ] && ! grep -Fxq "$rel_path" "$SOURCE_FILES_TMP"; then
            full_path="$SOURCE_DIR/$rel_path"
            [ -f "$full_path" ] && rm -f "$full_path"
        fi
    done <"$TARGET_FILES_TMP"

    # Protected files (skip if already exist)
    PROTECTED_FILES=("docProps/app.xml" "docProps/core.xml" "word/settings.xml" "word/glossary/settings.xml")

    is_protected_file() {
        local file="$1"
        for protected in "${PROTECTED_FILES[@]}"; do
            [[ "$file" == "$protected" ]] && return 0
        done
        return 1
    }

    # Copy new and updated files
    files_copied=0
    while IFS= read -r rel_path; do
        [ -z "$rel_path" ] && continue

        source_path="$TEMP_EXTRACT_DIR/$rel_path"
        target_path="$SOURCE_DIR/$rel_path"

        [ ! -f "$source_path" ] && continue

        # Skip protected files if they already exist
        if is_protected_file "$rel_path" && [ -f "$target_path" ]; then
            continue
        fi

        # Ensure target directory exists
        target_subdir=$(dirname "$target_path")
        [ ! -d "$target_subdir" ] && mkdir -p "$target_subdir"

        # Copy file
        if cp "$source_path" "$target_path"; then
            ((files_copied++))
        fi

    done <"$SOURCE_FILES_TMP"

    # Remove all Word-generated IDs for clean version control
    clean_word_ids() {
        local file="$1"
        local temp_file="${file}.cleanup"

        # Unified approach: remove all Word-generated IDs regardless of file type
        sed -E \
            -e 's/\s*<w:rsid w:val="[0-9A-F]{8}"\s*\/>//g' \
            -e 's/ w:rsidR="[0-9A-F]{8}"//g' \
            -e 's/ w:rsidRDefault="[0-9A-F]{8}"//g' \
            -e 's/ w:rsidRPr="[0-9A-F]{8}"//g' \
            -e 's/ w:rsidP="[0-9A-F]{8}"//g' \
            -e 's/ w:rsidTr="[0-9A-F]{8}"//g' \
            -e 's/ w:rsidSect="[0-9A-F]{8}"//g' \
            -e 's/ w14:textId="[0-9A-F]{8}"//g' \
            -e 's/ w14:paraId="[0-9A-F]{8}"//g' \
            "$file" >"$temp_file"

        if [ -s "$temp_file" ]; then
            mv "$temp_file" "$file"
            return 0
        else
            rm -f "$temp_file"
            return 1
        fi
    }

    # Format XML files for version control
    format_xml_file() {
        local file="$1"
        local temp_file="${file}.tmp"

        if xmllint --format "$file" >"$temp_file" 2>/dev/null; then
            if [ -s "$temp_file" ] && xmllint --noout "$temp_file" 2>/dev/null; then
                mv "$temp_file" "$file"
                return 0
            else
                rm -f "$temp_file"
                return 1
            fi
        else
            rm -f "$temp_file"
            return 1
        fi
    }

    # Clean IDs and format all XML files
    xml_files_formatted=0
    xml_files_cleaned=0
    while IFS= read -r -d '' xml_file; do
        if [ -f "$xml_file" ] && [ -s "$xml_file" ]; then
            # First clean Word-generated IDs
            if [[ "$xml_file" == */word/*.xml ]]; then
                clean_word_ids "$xml_file" && ((xml_files_cleaned++))
            fi
            # Then format
            format_xml_file "$xml_file" && ((xml_files_formatted++))
        fi
    done < <(find "$SOURCE_DIR" -name "*.xml" -type f -print0)

    # Cleanup temp files
    rm -f "$SOURCE_FILES_TMP" "$TARGET_FILES_TMP"

    # Summary
    total_files=$(find "$SOURCE_DIR" -type f | wc -l)
    echo -e "${GREEN}Successfully extracted to reference/ - $total_files files, $xml_files_cleaned XML files cleaned, $xml_files_formatted XML files formatted${NC}"
}

# Main logic
case "${1:-}" in
"zip")
    do_zip
    ;;
"unzip")
    do_unzip
    ;;
*)
    show_usage
    ;;
esac
