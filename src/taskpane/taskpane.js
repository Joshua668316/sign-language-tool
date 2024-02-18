/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { base64Image } from "./base64Image";
const { createCanvas } = require('canvas');

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("insert-image").onclick = () => clearMessage(submitTextAndImages);
  }
});

function readImages(handleResult) {
  const file = document.getElementById("fileElem").files[0];
  const reader = new FileReader();
  reader.onload = (e) => {
    handleResult(e.target.result);
  }
  reader.readAsDataURL(file);
}

function submitTextAndImages() {
  const file = document.getElementById("fileElem").files[0];
  readImages((result) => {
    const img = new Image();
    img.src = result;
    const base64Image = createImageBase64(img.src)
    insertImage(base64Image);
  });
}

function insertImage(image) {
  Office.context.document.setSelectedDataAsync(
    image,
    {
      coercionType: Office.CoercionType.Image
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
      }
    }
  );
}


function createImageBase64(base64Image) {
  var words = document.getElementById("text-input").value.split(/\s+/);
  const numPictures = words.length;
  const imageSize = 245
  const padding = 20;
  const textSpace = 80; 
  const width = imageSize * numPictures + padding * (numPictures + 1);
  const height = imageSize + textSpace
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext('2d');

  var img = new Image();
  img.src = base64Image;

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

async function clearMessage(callback) {
  document.getElementById("message").innerText = "";
  await callback();
}

function setMessage(message) {
  document.getElementById("message").innerText = message;
}
