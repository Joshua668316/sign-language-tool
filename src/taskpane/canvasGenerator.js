import { getWords, getWordsLowerCase } from './wordMatching';
const { createCanvas } = require('canvas');


function canvasConfig(imageSize, padding, textSpace, numPictures) {
  this.imageSize = imageSize;
  this.padding = padding;
  this.textSpace = textSpace;
  this.numPictures = numPictures;
    
  this.width = imageSize * numPictures + padding * (numPictures + 1);
  this.height = imageSize + textSpace;
}

function imageTransformation(conf, naturalWidth, naturalHeight, i) {
  const scale = conf.imageSize / Math.max(naturalWidth, naturalHeight);
  const imgWidth = scale * naturalWidth;
  const imgHeight = scale * naturalHeight;
  const dx = (conf.imageSize - imgWidth) / 2;
  const dy = (conf.imageSize - imgHeight) / 2;
  const x = conf.padding * i + conf.imageSize * (i - 1) + dx;
  const y = conf.padding + dy;
  return {x, y, imgWidth, imgHeight};
}

function textTransformation(conf, i) {
  const x = conf.padding * i + conf.imageSize * (i - 0.5);
  const y = 0.9 * conf.height;
  return {x, y}
}

function drawImages(ctx, conf, images, words) {
  for (let i = 1; i <= conf.numPictures; i++) {
    if (images.has(words[i - 1])) {
      const img = images.get(words[i - 1]);
      let {x, y, imgWidth, imgHeight} = imageTransformation(conf, img.naturalWidth, img.naturalHeight, i);
      ctx.drawImage(img, x, y, imgWidth, imgHeight);
    }
  }
}

function drawText(ctx, conf, words) {
  // Set text style
  ctx.fillStyle = '#000000';
  ctx.font = '48px Arial';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';

  for (let i = 1; i <= conf.numPictures; i++) {
    let {x, y} = textTransformation(conf, i);
    ctx.fillText(words[i - 1], x, y);
  }
}

export function createCanvasBase64(images, text) {
    const words = getWords(text);
    const wordsLowerCase = getWordsLowerCase(text);

    const numPictures = words.length;
    const imageSize = 245
    const padding = 20;
    const textSpace = 80; 

    const conf = new canvasConfig(imageSize, padding, textSpace, numPictures);

    const canvas = createCanvas(conf.width, conf.height);
    const ctx = canvas.getContext('2d');
  
    drawImages(ctx, conf, images, wordsLowerCase);
    drawText(ctx, conf, words, images.size);
    // Convert canvas to Base64 string (without the data:image/png;base64, prefix)
    return canvas.toDataURL().split(',')[1];
  }
  