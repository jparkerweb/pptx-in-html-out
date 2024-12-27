import JSZip from 'jszip';
import { parseStringPromise } from 'xml2js';
import sharp from 'sharp';
import { createWorker } from 'tesseract.js';

export class PPTXInHTMLOut {
  constructor(pptxBuffer) {
    this.pptxBuffer = pptxBuffer;
    this.zip = null;
    this.debug = false;
    this.ocrWorker = null;
    this.slideLayouts = new Map();
    this.slideMasters = new Map();
    this.slides = [];
    this.images = new Map();
    this.relationships = new Map();
  }

  setDebug(enabled) {
    this.debug = enabled;
    return this;
  }

  log(...args) {
    if (this.debug) {
      console.log(...args);
    }
  }

  async initialize() {
    try {
      this.zip = await JSZip.loadAsync(this.pptxBuffer);
      const files = Object.keys(this.zip.files);
      // console.log('Files in PPTX:', files);
    } catch (error) {
      console.error('Error initializing:', error);
      throw error;
    }
  }

  async load() {
    if (!Buffer.isBuffer(this.pptxBuffer)) {
      throw new Error('Input must be a Buffer');
    }

    try {
      this.log('Loading PPTX buffer of size:', this.pptxBuffer.length);
      await this.initialize();
      await this.validatePPTX();
    } catch (error) {
      this.log('Error during load:', error);
      throw error;
    }
  }

  async validatePPTX() {
    const files = Object.keys(this.zip.files);
    this.log('Files in PPTX:', files);

    const requiredFiles = [
      'ppt/presentation.xml',
      '_rels/.rels'
    ];

    for (const file of requiredFiles) {
      if (!files.includes(file)) {
        throw new Error(`Invalid PPTX file: missing ${file}`);
      }
    }

    const slideFiles = files.filter(f => f.match(/ppt\/slides\/slide[0-9]+\.xml/));
    if (slideFiles.length === 0) {
      throw new Error('Invalid PPTX file: no slides found');
    }

    this.slideFiles = slideFiles;
  }

  async parse() {
    try {
      await this.parseRelationships();
      await this.parseSlideLayouts();
      await this.parseSlideMasters();
      await this.parseSlides();
      await this.extractImages();
    } catch (error) {
      this.log('Error during parse:', error);
      throw error;
    }
  }

  async parseXml(content) {
    const parserOptions = {
      explicitArray: false,
      explicitRoot: true,
      normalizeTags: true,
      tagNameProcessors: [(name) => name.replace(/^[a-z]+:/, '')],
      attrNameProcessors: [(name) => name.replace(/^[a-z]+:/, '')],
      attrValueProcessors: [(value) => value],
      xmlns: true
    };
    return parseStringPromise(content, parserOptions);
  }

  async parseRelationships() {
    const relsFiles = Object.keys(this.zip.files).filter(name => name.endsWith('.rels'));
    for (const relsFile of relsFiles) {
      try {
        const content = await this.zip.file(relsFile).async('text');
        const result = await this.parseXml(content);
        
        if (result?.Relationships?.Relationship) {
          const relationships = Array.isArray(result.Relationships.Relationship) 
            ? result.Relationships.Relationship 
            : [result.Relationships.Relationship];
            
          const basePath = relsFile.split('/_rels/')[0];
          this.relationships.set(basePath, relationships.reduce((acc, rel) => {
            acc[rel.Id] = {
              type: rel.Type,
              target: rel.Target
            };
            return acc;
          }, {}));
        }
      } catch (error) {
        this.log(`Warning: Could not parse relationships in ${relsFile}:`, error);
      }
    }
  }

  async parseSlides() {
    const slideFiles = Object.keys(this.zip.files)
      .filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'))
      .sort();

    this.slides = [];
    
    for (const slideFile of slideFiles) {
      try {
        const slideContent = await this.zip.file(slideFile).async('string');
        const slideXml = await parseStringPromise(slideContent);
        
        if (this.debug) {
          console.log('Parsing slide:', slideFile);
          console.log('Full result:', JSON.stringify(slideXml));
        }
        
        this.slides.push({
          file: slideFile,
          content: slideXml
        });
      } catch (error) {
        console.error(`Error parsing slide ${slideFile}:`, error);
      }
    }
    
    if (this.debug) {
      console.log('Successfully parsed slides:', this.slides.length);
      if (this.slides.length > 0) {
        console.log('First slide content:', JSON.stringify(this.slides[0].content));
      }
    }
    
    return this.slides;
  }

  async convertSlideToHTML(slide) {
    if (!slide?.content) {
      console.error('Invalid slide content');
      return '';
    }

    const sld = slide.content['p:sld'];
    if (!sld) {
      console.error('No p:sld found in slide content');
      return '';
    }

    // Get the spTree from the correct path
    const spTree = sld['p:cSld']?.[0]?.['p:spTree']?.[0];
    
    if (!spTree) {
      console.error('No spTree found in slide, full content:', JSON.stringify(sld));
      return '';
    }

    let html = '<div class="slide"><div class="slide-content">';

    // Process shapes (text content)
    if (spTree['p:sp']) {
      for (const shape of spTree['p:sp']) {
        const txBody = shape['p:txBody']?.[0];
        if (txBody) {
          html += '<div class="shape">';
          
          // Process paragraphs
          const paragraphs = txBody['a:p'] || [];
          for (const paragraph of paragraphs) {
            html += '<p>';
            
            // Process text runs
            if (paragraph['a:r']) {
              for (const run of paragraph['a:r']) {
                const text = run['a:t']?.[0] || '';
                html += `<span>${text}</span>`;
              }
            }
            
            html += '</p>';
          }
          
          html += '</div>';
        }
      }
    }

    // Process pictures
    if (spTree['p:pic']) {
      for (const pic of spTree['p:pic']) {
        const blipFill = pic['p:blipFill']?.[0];
        const blip = blipFill?.['a:blip']?.[0];
        const rId = blip?.$?.['r:embed'];
        
        if (rId) {
          const imageData = await this.getImageData(slide.file, rId);
          if (imageData) {
            const processedImage = await this.processImage(imageData);
            if (processedImage?.text) {
              html += `<div class="picture">
                <p class="ocr-text">${processedImage.text}</p>
              </div>`;
            }
          }
        }
      }
    }

    html += '</div></div>';
    return html;
  }

  async convertShapeToHTML(shape, options) {
    const nvSpPr = shape?.nvsppr || shape?.['nvSpPr'];
    const txBody = shape?.txbody || shape?.['txBody'];
    const spPr = shape?.sppr || shape?.['spPr'];

    if (!txBody) {
      return '';
    }

    let shapeHtml = '<div class="shape">';

    if (txBody.p) {
      const paragraphs = Array.isArray(txBody.p) ? txBody.p : [txBody.p];
      for (const p of paragraphs) {
        shapeHtml += '<p>';
        if (p.r) {
          const runs = Array.isArray(p.r) ? p.r : [p.r];
          for (const r of runs) {
            const text = r?.t?._ || '';
            shapeHtml += `<span>${text}</span>`;
          }
        }
        shapeHtml += '</p>';
      }
    }

    shapeHtml += '</div>';
    return shapeHtml;
  }

  async convertPictureToHTML(pic, options) {
    const nvPicPr = pic?.nvpicpr || pic?.['nvPicPr'];
    const blipFill = pic?.blipfill || pic?.['blipFill'];
    const spPr = pic?.sppr || pic?.['spPr'];

    if (!blipFill?.blip?.$.embed) {
      return '';
    }

    const rId = blipFill.blip.$.embed;
    const imageData = await this.getImageData(pic.file, rId);
    if (!imageData) {
      return '';
    }

    return `<div class="picture">
        <img src="${imageData}" alt="Slide Image"/>
    </div>`;
  }

  async getSlideRels(slideFile) {
    try {
      // Get the relationships file for this slide
      const slideNumber = slideFile.match(/slide(\d+)\.xml/)[1];
      const relsFile = this.zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`);
      
      if (!relsFile) {
        console.error('No relationships file found for slide:', slideFile);
        return null;
      }
      
      const relsContent = await relsFile.async('string');
      const relsXml = await parseStringPromise(relsContent);
      
      // Parse relationships
      const rels = {};
      const relationships = relsXml?.['Relationships']?.['Relationship'] || [];
      
      for (const rel of relationships) {
        const id = rel.$?.Id;
        const target = rel.$?.Target;
        if (id && target) {
          rels[id] = {
            Id: id,
            Target: target
          };
        }
      }
      
      return rels;
    } catch (error) {
      console.error('Error getting slide relationships:', error);
      return null;
    }
  }

  async getImageData(slideFile, rId) {
    try {
      const rels = await this.getSlideRels(slideFile);
      if (!rels || !rels[rId]) {
        console.error('No relationship found for rId:', rId);
        return null;
      }

      const imagePath = rels[rId].Target;
      if (!imagePath) {
        console.error('No target path found for relationship:', rId);
        return null;
      }

      // Get full path to image file
      const fullPath = `ppt/media/${imagePath.split('/').pop()}`;
      const imageFile = this.zip.file(fullPath);
      
      if (!imageFile) {
        console.error('Image file not found:', fullPath);
        return null;
      }

      // Get image data as buffer
      const imageBuffer = await imageFile.async('nodebuffer');
      return imageBuffer;
    } catch (error) {
      console.error('Error getting image data:', error);
      return null;
    }
  }

  async processImage(imageBuffer) {
    let worker = null;
    try {
      worker = await createWorker();
      const { data: { text } } = await worker.recognize(imageBuffer);
      return { text };
    } catch (error) {
      console.error('Error processing image with OCR:', error);
      return null;
    } finally {
      if (worker) {
        await worker.terminate();
      }
    }
  }

  async convertShapeOrPicture(element) {
    if (element.pic) {
      // It's a picture, handle embedded image
      const rId = element.pic.blipFill?.blip?.$?.['r:embed'];
      if (rId) {
        const slideRels = await this.getSlideRels(element.file);
        const imagePath = slideRels?.[rId]?.Target;
        
        if (imagePath) {
          const fullImagePath = `ppt/media/${imagePath.split('/').pop()}`;
          const imageData = await this.zip.file(fullImagePath)?.async('nodebuffer');
          
          if (imageData) {
            const processedImage = await this.processImage(imageData);
            if (processedImage) {
              return `<div class="picture">
                  <p class="ocr-text">${processedImage.text}</p>
              </div>`;
            }
          }
        }
      }
    }
  }

  async parseSlideLayouts() {
    const layoutFiles = Object.keys(this.zip.files).filter(name => 
      name.includes('ppt/slideLayouts/slideLayout'));
    
    for (const layoutFile of layoutFiles) {
      const content = await this.zip.file(layoutFile).async('text');
      const result = await this.parseXml(content);
      this.slideLayouts.set(layoutFile, result);
    }
  }

  async parseSlideMasters() {
    const masterFiles = Object.keys(this.zip.files).filter(name => 
      name.includes('ppt/slideMasters/slideMaster'));
    
    for (const masterFile of masterFiles) {
      const content = await this.zip.file(masterFile).async('text');
      const result = await this.parseXml(content);
      this.slideMasters.set(masterFile, result);
    }
  }

  async extractImages() {
    const mediaFiles = Object.keys(this.zip.files).filter(name => 
      name.startsWith('ppt/media/'));
    
    for (const mediaFile of mediaFiles) {
      const data = await this.zip.file(mediaFile).async('nodebuffer');
      const image = await sharp(data);
      const metadata = await image.metadata();
      const base64 = data.toString('base64');
      
      this.images.set(mediaFile, {
        data: base64,
        metadata,
        type: metadata.format
      });
    }
  }

  async toHTML() {
    try {
      await this.initialize();
      const slides = await this.parseSlides();
      const html = await this.generateHTML(slides);
      return html;
    } catch (error) {
      console.error('Error converting to HTML:', error);
      throw error;
    }
  }

  async generateHTML(slides) {
    let slidesHTML = '';
    for (const slide of slides) {
      const slideHTML = await this.convertSlideToHTML(slide);
      slidesHTML += slideHTML;
    }

    return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PowerPoint Presentation</title>
  ${this.generateStyles()}
</head>
<body>${slidesHTML}</body>
</html>`;
  }

  generateStyles() {
    return `
      <style>
        .slide {
          position: relative;
          width: 100%;
          height: 0;
          padding-bottom: 56.25%;
          margin-bottom: 20px;
          background: white;
        }
        .slide-content {
          position: absolute;
          top: 0;
          left: 0;
          width: 100%;
          height: 100%;
        }
        .shape {
          position: absolute;
          box-sizing: border-box;
        }
        .text {
          word-wrap: break-word;
          overflow-wrap: break-word;
        }
        .image {
          max-width: 100%;
          height: auto;
        }
      </style>
    `;
  }
}
