import fs from 'fs/promises';
import path from 'path';
import { PPTXInHTMLOut } from './src/index.js';

// Read the PPTX file
const pptxBuffer = await fs.readFile('./example.pptx');
console.log('Read PPTX file of size:', pptxBuffer.length);

// Create converter instance with buffer
const converter = new PPTXInHTMLOut(pptxBuffer);
    
// Convert to HTML
console.log('Starting conversion...');
const html = await converter.toHTML();
    
// Output the HTML
console.log('HTML output:', html);
