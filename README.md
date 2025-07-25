# DocShift

A pure client-side HTML ↔ DOCX conversion library for JavaScript that preserves styling by mapping between HTML and Word document formats.

## Features

- **Pure Client-Side**: No server dependency - runs entirely in the browser
- **Style Preservation**: mapping between HTML and DOCX styling
- **Rich Text Editor Compatible**: Works with TinyMCE, WordPress editor, and other popular rich text editors
- **Bidirectional**: Convert HTML to DOCX and DOCX to HTML
- **Compact, Self-Contained**: 240KB minified + gzipped, no other deps needed.
- Available as ESM npm package and vanilla HTML via CDN

## Installation

### NPM Package
```bash
npm install docshift
```

### CDN (Vanilla HTML)
```html
<script src="https://cdn.jsdelivr.net/npm/docshift@latest/dist/docshift.min.js"></script>
```

## Quick Start

### ESM Import
```javascript
import { toDocx, toHtml } from 'docshift';

// Convert DOCX to HTML
const docxFile = document.getElementById('fileInput').files[0];
const html = await toHtml(docxFile);
console.log(html);

// Convert HTML to DOCX
const htmlContent = '<p>Hello <strong>World</strong>!</p>';
const docxBlob = await toDocx(htmlContent);
```

### CDN Usage

When loaded via CDN, the library is available as a global variable `window.docshift`

```javascript
// Convert DOCX to HTML
const docxFile = document.getElementById('fileInput').files[0];
const html = await docshift.toHtml(docxFile);

// Convert HTML to DOCX
const htmlContent = '<p>Hello <strong>World</strong>!</p>';
const docxBlob = await docshift.toDocx(htmlContent);
```

## API Reference

### `toHtml(docxFile)`
Converts a DOCX file to HTML string.

**Parameters:**
- `docxFile` (File|Blob): DOCX file object (e.g., from file input)

**Returns:**
- `Promise<string>`: HTML string representation of the document

**Example:**
```javascript
const fileInput = document.getElementById('docx-input');
const docxFile = fileInput.files[0];
const html = await toHtml(docxFile);
document.getElementById('output').innerHTML = html;
```

### `toDocx(htmlContent)`
Converts HTML content to DOCX format.

**Parameters:**
- `htmlContent` (string|HTMLElement): HTML string or DOM element to convert

**Returns:**
- `Promise<Blob>`: DOCX file as Blob object

**Example:**
```javascript
// From HTML string
const html = '<p>This is a <em>sample</em> document with <strong>formatting</strong>.</p>';
const docxBlob = await toDocx(html);

// From DOM element
const contentDiv = document.getElementById('editor-content');
const docxBlob = await toDocx(contentDiv);

// Trigger download
const url = URL.createObjectURL(docxBlob);
const a = document.createElement('a');
a.href = url;
a.download = 'document.docx';
a.click();
```


### Example with TinyMCE
```javascript
// Export TinyMCE content to DOCX
async function exportToDocx() {
    const content = tinymce.get('editor').getContent();
    const docxBlob = await toDocx(content);
    
    // Download the file
    const url = URL.createObjectURL(docxBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'tinymce-export.docx';
    a.click();
}

// Import DOCX to TinyMCE
async function importFromDocx(file) {
    const html = await toHtml(file);
    tinymce.get('editor').setContent(html);
}
```

## Best Practices

### HTML Structure
OpenXML Word documents are built around paragraphs. For best results:

Organize content in paragraph tags
```html
<p>This is a paragraph with <strong>bold text</strong>.</p>
<p>This is another paragraph with <em>italic text</em>.</p>
```

Otherwise, DocShift will try to automatically group orphaned inline elements
```html
Some text with <strong>bold</strong> formatting.
<br>
More text on a new line.
<!-- DocShift converts this to proper paragraphs -->
<p>
  Some text with <strong>bold</strong> formatting.
  <br>
</p>
<p>More text on a new line.</p>
```

### Image Handling
Since DocShift runs client-side, CORS restrictions apply to images:

```javascript
// ❌ This may fail due to CORS
const html = '<p><img src="https://external-domain.com/image.jpg" /></p>';

// ✅ Better: Use blob URLs or same-origin images
const html = '<p><img src="data:image/jpeg;base64,/9j/4AAQ..." /></p>';
// or
const html = '<p><img src="/local-image.jpg" /></p>';
```

**Workaround for external images:**
1. Pre-process images through your server/proxy
2. Convert to blob URLs or data URLs
3. Replace src attributes before conversion


### Acknowledgments
DocShift is built on top of [mammoth.js](https://github.com/mwilliamson/mammoth.js) and [docx](https://github.com/dolanmiu/docx)

Thanks to these projects for doing the heavy lifting!

A few years back I built [win32.run](https://github.com/ducbao414/win32.run), a Windows XP recreation in the browser that included a basic imitation of MS Word with the ability to open and edit DOCX files directly in the browser.

Finally got around to extracting the document conversion bits into this standalone library.
