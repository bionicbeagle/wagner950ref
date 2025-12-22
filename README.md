# Wagner Orion 950 Species Settings Reference Generator

A Node.js utility that generates comprehensive reference materials for the Wagner Orion 950 Moisture Meter, containing specific gravity settings for over 220 wood species and engineered materials.

## Overview

This tool generates reference materials in three formats:
- **DOCX** - Professional Word document with tables and formatting
- **PDF** - Portable document for printing and archiving
- **HTML** - Interactive web page with live search functionality

All formats include the same comprehensive data:
- 219 wood species with common names, scientific names, and specific gravity settings
- 5 engineered wood products (Plywood, OSB, MDF, HDF, Advantech™)
- Usage instructions and calibration notes
- Links to Wagner's online database

## Installation

```bash
npm install
```

This will install the required dependencies:
- `docx` - For Word document generation
- `pdfkit` - For PDF generation

## Usage

### Generate All Formats (Default)

```bash
node wagner_950_complete.js
# or
node wagner_950_complete.js all
```

This creates:
- `Wagner_950_Reference_Sheets.docx`
- `Wagner_950_Reference_Sheets.pdf`
- `Wagner_950_Reference_Sheets.html`

### Generate Specific Format

```bash
# Generate only Word document
node wagner_950_complete.js docx

# Generate only PDF
node wagner_950_complete.js pdf

# Generate only HTML page
node wagner_950_complete.js html
```

### Display Help

```bash
node wagner_950_complete.js help
```

## Output Formats

### DOCX (Word Document)
- Multi-page professional layout
- Title page with usage instructions
- Species tables (30 species per page)
- Engineered materials section
- Formatted with Wagner brand colors
- Perfect for printing or editing

### PDF
- Letter-size pages optimized for printing
- Identical layout to DOCX format
- Portable and universally compatible
- Professional appearance with branded styling

### HTML (Interactive Web Page)
- **Live search functionality** - filter by common name, scientific name, or setting value
- Real-time results counter
- Responsive design (works on desktop and mobile)
- No external dependencies - fully self-contained
- Modern, clean interface with gradient styling
- Sticky table header for easy scrolling
- Keyboard shortcuts:
  - Search auto-focuses on page load
  - Press `Escape` to clear search
- Engineered materials highlighted with badges

## Features

### Comprehensive Species Database
- 219 wood species organized alphabetically
- Scientific names for precise identification
- Specific gravity settings ranging from 0.14 (Balsa) to 1.27 (African Blackwood)
- 5 engineered wood products

### Multiple Output Formats
Each format is optimized for its use case:
- **DOCX**: Best for printing and offline reference
- **PDF**: Best for sharing and archiving
- **HTML**: Best for quick lookups and digital use

### Professional Formatting
- Wagner brand colors (#2E5C8A)
- Clean, readable typography
- Organized tables with alternating row colors
- Clear section headers

## File Structure

```
wagner950ref/
├── wagner_950_complete.js          # Main generation script
├── package.json                     # Project dependencies
├── package-lock.json                # Dependency lock file
├── README.md                        # This file
└── Output files:
    ├── Wagner_950_Reference_Sheets.docx
    ├── Wagner_950_Reference_Sheets.pdf
    └── Wagner_950_Reference_Sheets.html
```

## Species Data

The species database includes:
- Common name (e.g., "American Red Oak")
- Scientific name (e.g., "Quercus rubra")
- Specific gravity setting (e.g., "0.63")

All settings are based on specific gravity at 12% moisture content, with a natural variation of ±10% considered normal within species.

## HTML Page Features

The interactive HTML page includes:

### Search Functionality
- **Real-time filtering** - results update as you type
- **Multi-field search** - searches across all columns
- **Visual feedback** - counter shows matching results
- **No results message** - friendly message when no matches found

### User Interface
- **Gradient background** - modern purple/blue gradient
- **Card-based layout** - clean, contained design
- **Sticky header** - table header stays visible while scrolling
- **Hover effects** - rows highlight on mouse over
- **Responsive design** - adapts to screen size
- **Mobile-friendly** - works on phones and tablets

### Additional Information
- Usage instructions
- Important notes about settings and calibration
- Links to Wagner's website and support
- Contact information

## Technical Details

### Dependencies
- **docx (v9.5.1)** - Professional Word document generation with full formatting support
- **pdfkit (latest)** - High-quality PDF generation with drawing capabilities

### Browser Compatibility (HTML)
The generated HTML page works in all modern browsers:
- Chrome/Edge (latest)
- Firefox (latest)
- Safari (latest)
- Mobile browsers (iOS Safari, Chrome Mobile)

No external dependencies or internet connection required to use the HTML page.

## About Wagner Orion 950

The Wagner Orion 950 is a pinless moisture meter that uses electromagnetic wave technology to measure moisture content in wood. Each wood species has a different density (specific gravity), which must be set in the meter for accurate readings.

### How to Use These Settings

1. Locate your wood species in the reference materials
2. Note the specific gravity (SG) setting value
3. Enter this value into your Wagner 950 meter using the up/down arrows
4. Take your moisture reading

### Important Notes

- Settings are based on specific gravity at 12% moisture content
- Values represent average density for each species
- Natural variation of ±10% is normal within species
- Scientific names help ensure precise wood identification
- For species not listed, visit: https://www.wagnermeters.com/specific-gravity

## Support

For additional species or technical support:
- **Website**: https://www.wagnermeters.com
- **Online Database**: https://www.wagnermeters.com/specific-gravity (7,500+ species)
- **Phone**: (800) 634-9961
- **Email**: support@wagnermeters.com

## License

This reference generator is provided as-is for use with Wagner Orion 950 Moisture Meters. Species data is compiled from Wagner Meters' public resources.

---

**Generated with**: Node.js, docx library, pdfkit, and vanilla JavaScript
