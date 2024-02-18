const { createCanvas } = require('canvas');

function canvasConfig(imageSize, padding, textSpace, numPictures, img) {
  this.imageSize = imageSize;
  this.padding = padding;
  this.textSpace = textSpace;
  this.numPictures = numPictures;
    
  this.width = imageSize * numPictures + padding * (numPictures + 1);
  this.height = imageSize + textSpace;
    
  this.scale = imageSize / Math.max(img.naturalWidth, img.naturalHeight);
  this.imgWidth = this.scale * img.naturalWidth;
  this.imgHeight = this.scale * img.naturalHeight;
  this.dx = (this.imageSize - this.imgWidth) / 2;
  this.dy = (this.imageSize - this.imgHeight) / 2;
}

export function createImageBase64(img, words) {
    const numPictures = words.length;
    const imageSize = 245
    const padding = 20;
    const textSpace = 80; 

    const conf = new canvasConfig(imageSize, padding, textSpace, numPictures, img);

    const canvas = createCanvas(conf.width, conf.height);
    const ctx = canvas.getContext('2d');
  
    ctx.fillStyle = '#DDDDDD';
  
    for (let i = 1; i <= numPictures; i++) {
      ctx.fillRect(conf.padding * i + conf.imageSize * (i - 1), conf.padding, conf.imageSize, conf.imageSize);
      ctx.drawImage(img, conf.padding * i + conf.imageSize * (i - 1) + conf.dx, conf.padding + conf.dy, conf.imgWidth, conf.imgHeight);
    }
  
    // Set text style
    ctx.fillStyle = '#000000';
    ctx.font = '48px Arial';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
  
    for (let i = 1; i <= numPictures; i++) {
      ctx.fillText(words[i - 1], conf.padding * i + conf.imageSize * (i - 0.5), 0.9 * canvas.height);
    }
  
    // Convert canvas to Base64 string (without the data:image/png;base64, prefix)
    return canvas.toDataURL().split(',')[1];
  }
  