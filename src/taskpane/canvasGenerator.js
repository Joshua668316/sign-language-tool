const { createCanvas } = require('canvas');

export function createImageBase64(img, words) {
    const numPictures = words.length;
    const imageSize = 245
    const padding = 20;
    const textSpace = 80; 
    const width = imageSize * numPictures + padding * (numPictures + 1);
    const height = imageSize + textSpace
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext('2d');
  
    const scale = imageSize / Math.max(img.naturalWidth, img.naturalHeight);
    const imgWidth = scale * img.naturalWidth;
    const imgHeight = scale * img.naturalHeight;
    const dx = (imageSize - imgWidth) / 2;
    const dy = (imageSize - imgHeight) / 2;
  
    ctx.fillStyle = '#DDDDDD';
  
    for (let i = 1; i <= numPictures; i++) {
      ctx.fillRect(padding * i + imageSize * (i - 1), padding, imageSize, imageSize);
      ctx.drawImage(img, padding * i + imageSize * (i - 1) + dx, padding + dy, imgWidth, imgHeight);
    }
  
    // Set text style
    ctx.fillStyle = '#000000';
    ctx.font = '48px Arial';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
  
    for (let i = 1; i <= numPictures; i++) {
      ctx.fillText(words[i - 1], padding * i + imageSize * (i - 0.5), 0.9 * canvas.height);
    }
  
    // Convert canvas to Base64 string (without the data:image/png;base64, prefix)
    return canvas.toDataURL().split(',')[1];
  }
  