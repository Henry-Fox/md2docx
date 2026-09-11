# Changelog

All notable changes to md2docx will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.3.0] - 2026-09-11

### Added

- **PDF Export Feature**: Template-driven PDF generation using pdf-lib
  - New `SimpleMd2Pdf` class for converting Markdown to PDF
  - PDF export button in the UI alongside Word export
  - Support for all major Markdown elements in PDF:
    - Headings (H1-H6)
    - Paragraphs with inline formatting
    - Ordered and unordered lists
    - Tables with borders
    - Code blocks with background
    - Block quotes
    - Horizontal rules
    - Links (as styled text)
  - Template system integration for consistent styling between DOCX and PDF
  - Multi-language support for PDF export UI strings (Chinese, English, French, Spanish, Russian, Arabic)

- **ExportManager**: Unified export interface for both DOCX and PDF formats
  - Centralized export logic
  - Template integration
  - Error handling and validation

- **Testing Infrastructure**:
  - Vitest testing framework
  - Unit tests for `TemplateManager`
  - Unit tests for `ExportManager`
  - Test coverage reporting
  - npm test scripts now functional

### Changed

- Updated `templateManager.js`:
  - Added `toPdfStyles()` method to convert templates to PDF-compatible styles
  - Added Chinese font mapping to PDF standard fonts
  - Enhanced template system to support both DOCX and PDF output

- Updated `App.js`:
  - Integrated `ExportManager` for unified export handling
  - Added `directConvertToPdf()` method
  - Refactored `directConvertToDocx()` to use ExportManager

- Updated UI:
  - Split export button into separate "Export Word" and "Export PDF" buttons
  - Added CSS styling for export button group
  - Improved button layout and responsiveness

- Updated package metadata:
  - Version bump to 1.3.0
  - Updated description to mention PDF support
  - Added "pdf" and "markdown-to-pdf" keywords

### Removed

- Removed duplicate files from root directory:
  - `simpleMd2Docx.js` (kept version in `js/` directory)
  - Test files: `test-*.js`, `test-*.md`, `test.json`

### Technical Details

**PDF Font Strategy**:
- Uses standard PDF fonts (Helvetica, Times-Roman, Courier) for broad compatibility
- Chinese fonts are mapped to standard fonts with basic support
- Future enhancement: custom font embedding for full Chinese character support

**Architecture**:
```
App.js → ExportManager → SimpleMd2Docx / SimpleMd2Pdf
                       ↓
                  TemplateManager (toDocxStyles / toPdfStyles)
```

### Known Limitations

- PDF Chinese font rendering uses standard fonts (limited glyph coverage)
- PDF does not support embedded images yet (planned for future release)
- PDF layout is simpler than DOCX (no page headers/footers yet)

### Migration Notes

For developers upgrading from 1.1.x:
- The export interface now goes through `ExportManager` instead of direct converter calls
- Templates now have both `toDocxStyles()` and `toPdfStyles()` methods
- Test suite requires Node.js environment with Vitest installed

---

## [1.1.2] - Prior Release

Previous stable release with DOCX-only export functionality.

### Features

- Markdown to DOCX conversion
- Template system with 5 built-in templates
- Real-time preview
- Word format import
- Multi-language support (6 languages)
- Browser-only operation (no server required)
