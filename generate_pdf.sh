#!/bin/bash

# Script to generate Wagner 950 Reference Sheets document

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo "Wagner 950 Reference Document Generator"
echo "========================================"

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo -e "${RED}Error: Node.js is not installed${NC}"
    echo "Please install Node.js from https://nodejs.org/"
    exit 1
fi

echo -e "${GREEN}✓ Node.js found: $(node --version)${NC}"

# Check if npm is installed
if ! command -v npm &> /dev/null; then
    echo -e "${RED}Error: npm is not installed${NC}"
    exit 1
fi

echo -e "${GREEN}✓ npm found: $(npm --version)${NC}"

# Check if node_modules exists, if not install dependencies
if [ ! -d "node_modules" ]; then
    echo -e "${YELLOW}Installing dependencies...${NC}"
    npm install docx
    if [ $? -ne 0 ]; then
        echo -e "${RED}Error: Failed to install dependencies${NC}"
        exit 1
    fi
    echo -e "${GREEN}✓ Dependencies installed${NC}"
else
    echo -e "${GREEN}✓ Dependencies already installed${NC}"
fi

# Run the JavaScript file
echo ""
echo "Generating document..."
node wagner_950_complete.js

# Check if the script ran successfully
if [ $? -eq 0 ]; then
    echo ""
    echo -e "${GREEN}✓ Document generated successfully!${NC}"
    echo -e "Output: ${YELLOW}$(pwd)/Wagner_950_Reference_Sheets.docx${NC}"
else
    echo -e "${RED}Error: Document generation failed${NC}"
    exit 1
fi
