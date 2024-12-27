# ⚙️ PPTX In HTML Out

Convert PowerPoint presentations to HTML with high fidelity.

## Installation

```bash
npm install pptx-in-html-out
```

## Usage

```javascript
import fs from 'fs/promises';
import { PPTXInHTMLOut } from 'pptx-in-html-out';

// Read your PPTX file into a buffer
const pptxBuffer = await fs.readFile('presentation.pptx');

// Create converter instance with buffer
const converter = new PPTXInHTMLOut(pptxBuffer);

// Convert to HTML
const html = await converter.toHTML();
console.log(html);

// Or write to a file
await fs.writeFile('output.html', html);
```

## Features

- High-fidelity conversion of PowerPoint presentations to HTML
- OCR support for text extraction from images
- Preserves images, shapes, and text formatting
- Responsive output that works across devices
- Modern ESM package

## API

### `PPTXInHTMLOut`

Main class for converting PPTX files to HTML.

#### Constructor

```javascript
const converter = new PPTXInHTMLOut(pptxBuffer);
```

- `pptxBuffer`: Buffer containing the PPTX file data

#### Methods

##### `toHTML(options)`

Converts the presentation to HTML.

Parameters:
- `options` (optional): Configuration object
  - `includeStyles` (boolean, default: true): Whether to include default styles in the output HTML

```javascript
// With default styles
const html = await converter.toHTML();

// Without default styles (for custom styling)
const html = await converter.toHTML({ includeStyles: false });
```

Returns: `Promise<string>` - The generated HTML content

---

## Example

See the `example.js` file in the root of this project for an additional example that you can run.

## License

MIT
