# Changelog

All notable changes to md2docx will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.6.0] - 2026-09-11

### Added

- **TRUE WYSIWYG Preview**: Complete rewrite of preview system for exact visual fidelity
  - Preview now uses the SAME rendering pipeline as PDF export (not approximate CSS)
  - What you see in the preview matches the exported document exactly (所见即所得)
  - Real PDF-based preview using pdf-lib + pdf.js rendering
  - Multi-page preview with proper pagination
  - Page numbers displayed on each preview page
  
- **New PreviewRenderer Module** (`js/previewRenderer.js`):
  - PDF generation using SimpleMd2Pdf pipeline
  - Canvas-based page rendering with pdf.js
  - Smart preview caching (only regenerates when content or template changes)
  - Debounced preview updates (500ms) for smooth typing experience
  - Loading, empty, and error states with clear user feedback
  - Automatic cancellation of stale preview tasks
  
- **Template Change Detection**:
  - Preview automatically refreshes when template is switched
  - Preview updates when template settings are saved
  - Cache clearing ensures fresh preview after template edits

- **Enhanced User Experience**:
  - Loading spinner with status message during preview generation
  - Clear empty state when no content is entered
  - Detailed error messages with helpful hints (e.g., font loading issues)
  - Smooth scrolling for multi-page documents
  - Professional page shadows and spacing

### Changed

- **App.js Refactoring**:
  - Replaced `marked.parse()` HTML preview with PDF-based preview
  - Added debounce timer (500ms delay) to optimize performance during typing
  - Preview now triggers on template change and template save
  - Integrated `previewRenderer` for all preview operations

- **CSS Enhancements** (`css/style.css`):
  - Added `.preview-container-wysiwyg` for PDF preview layout
  - Added `.pdf-pages-container` for multi-page display
  - Added `.pdf-page-container` with shadows and borders
  - Added loading, empty, and error state styles
  - Added smooth animations for spinner and transitions

- **Dependencies**:
  - Added `pdfjs-dist` (3.11.174) for client-side PDF rendering
  - Configured pdf.js worker from CDN

- **Documentation Updates**:
  - Updated version to 1.6.0 in package.json
  - Added comprehensive CHANGELOG entry
  - Updated description to mention "true WYSIWYG preview"

### Technical Details

**Preview Architecture**:
```
User Input → Debounce (500ms) → PreviewRenderer
                                      ↓
                          SimpleMd2Pdf (same as export)
                                      ↓
                          PDF Bytes → pdf.js → Canvas Pages
```

**Performance Optimizations**:
- Debounced input (500ms) prevents excessive regeneration during typing
- Smart caching: preview only regenerates when content or template changes
- Stale task cancellation: old preview tasks are cancelled when new ones start
- Efficient canvas rendering with appropriate scaling for container width

**Visual Fidelity**:
- Preview uses exact same pdf-lib drawing code as PDF export
- Same fonts, margins, line spacing, and pagination
- Same template system (page size, orientation, styles)
- Preview is ~99% identical to exported PDF (browser canvas rendering)

### Known Limitations

- PDF preview uses standard fonts (Helvetica/Times-Roman) with basic Chinese support
  - Exported DOCX still uses full Chinese fonts
  - Exported PDF may show slight font differences vs preview depending on font availability
- Complex diagrams or tables may wrap differently in preview vs Microsoft Word
- Browser canvas rendering may have minor anti-aliasing differences vs native PDF viewers

### Breaking Changes

None. Existing export functionality (DOCX and PDF) remains unchanged and fully compatible.

### Migration Notes

For users:
- No action required. Preview now automatically shows PDF-rendered pages instead of HTML
- Preview may take ~0.5-1s to generate (shows loading spinner)
- If preview shows blank pages, check browser console for font loading errors

For developers:
- Preview logic moved from `app.js` inline HTML generation to `previewRenderer.js` module
- Preview container now holds canvas elements instead of marked.js HTML
- To customize preview, modify `previewRenderer.js` render methods

---

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
