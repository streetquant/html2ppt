# HTML to PowerPoint Converter - Improvements

## Implemented Improvements

### 1. Aspect Ratio - 16:9 (Google Slides Standard)
- Changed slide dimensions from 1200x675 to 1280x720 pixels
- Maintains proper 16:9 aspect ratio matching Google Slides
- Updated conversion factors accordingly

### 2. Support for More HTML Elements
Extended selectors for:
- Images: Added support for `picture`, `canvas`, `video`, `figure`, `figcaption`, `.image`, `.photo`, `.graphic`, `.icon`, `.logo`, `.avatar`, `.thumbnail`, `.poster`, `.banner`, `.header-image`, `.footer-image`
- Text: Added support for `a`, `strong`, `b`, `em`, `i`, `u`, `s`, `sub`, `sup`, `td`, `th`, `caption`, `legend`, `label`, `button`, `input[type="button"]`, `.text`, `.content`, `.title`, `.subtitle`, `.headline`, `.subheadline`, `.paragraph`, `.description`, `.note`, `.caption`, `.quote`, `.blockquote`, `.highlight`, `.emphasis`, `.important`, `.warning`, `.alert`, `.info`, `.header`, `.footer`, `.sidebar`, `.nav`, `.menu`, `.breadcrumb`, `.tag`, `.badge`, `.chip`, `.tooltip`, `.popover`, `.modal`, `.card`, `.panel`, `.box`, `.section`, `.container`, `.wrapper`

### 3. Font Mapping for Better Typography Preservation
- Added comprehensive font mapping dictionary with 21+ common web fonts
- Maps web fonts to their PowerPoint equivalents (e.g., 'helvetica' → 'Helvetica', 'arial' → 'Arial', 'times new roman' → 'Times New Roman')
- Includes fallback mechanism for unknown fonts defaulting to 'Calibri'

### 4. Improved Text Extraction with Nested Structure Preservation
- Enhanced text extraction function that preserves nested HTML structure
- Added support for extracting text from nested elements while maintaining context
- Improved handling of elements with mixed content (text and child elements)

### 5. Image Optimization
- Added image optimization function that resizes images to maximum 1920x1080 resolution
- Uses Lanczos resampling for high-quality downscaling
- Compresses images with 85% quality to reduce file size while maintaining quality

### 6. Better Text Wrapping and Overflow Handling
- Added minimum dimension constraints for text boxes (0.5" width, 0.25" height)
- Implemented text frame margins for better appearance
- Added text fitting to shape functionality
- Disabled auto-size for better control over text placement

### 7. Hyperlink Support
- Added detection and preservation of hyperlinks in HTML elements
- Extracts href attributes from anchor tags and nested elements
- Preserves hyperlinks in the editable PowerPoint version
- Handles both direct anchor tags and elements nested within anchor tags

### 8. Additional Improvements
- Better font size conversion (px to points with proper scaling)
- Text decoration support (underline, strikethrough)
- Enhanced error handling and resource cleanup
- More robust CSS property extraction
- Improved element visibility checks