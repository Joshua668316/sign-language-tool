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
    document.getElementById("insert-image").onclick = () => clearMessage(insertImage);
    document.getElementById("insert-text").onclick = () => clearMessage(insertText);
    // TODO6: Assign event handler for get-slide-metadata button.
    // TODO8: Assign event handlers for add-slides and the four navigation buttons.
  }
});


function insertImage() {
  Office.context.document.setSelectedDataAsync(
    createImageBase64(),
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


function createImageBase64() {
  var words = document.getElementById("text-input").value.split(/\s+/);
  const numPictures = words.length;
  const width = 1225;
  const height = 1500 / numPictures
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext('2d');

  const padding = 20;
  const imageSize = (canvas.width - (numPictures + 1) * padding) / numPictures
  ctx.fillStyle = '#DDDDDD';

  var img = new Image();
    
  img.src = "data:image/png;base64," + base64Image;

  for (let i = 1; i <= numPictures; i++) {
    ctx.fillRect(padding * i + imageSize * (i - 1), padding, imageSize, imageSize);
    ctx.drawImage(img, padding * i + imageSize * (i - 1), padding, imageSize, imageSize * 0.7);
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

function insertText() {
  Office.context.document.setSelectedDataAsync(
    document.getElementById("text-input").value,
    {
      coercionType: Office.CoercionType.Text
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
      }
    }
  )
}

// TODO7: Define the getSlideMetadata function.

// TODO9: Define the addSlides and navigation functions.

async function clearMessage(callback) {
  document.getElementById("message").innerText = "";
  await callback();
}

function setMessage(message) {
  document.getElementById("message").innerText = message;
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    document.getElementById("message").innerText = "";
    await callback();
  } catch (error) {
    setMessage("Error: " + error.toString());
  }
}