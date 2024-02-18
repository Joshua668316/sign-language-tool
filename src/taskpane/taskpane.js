/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { createImageBase64 } from "./canvasGenerator";

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
  var words = document.getElementById("text-input").value.split(/\s+/);
  readImages((result) => {
    const img = new Image();
    img.src = result;
    const base64Image = createImageBase64(img, words)
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

async function clearMessage(callback) {
  document.getElementById("message").innerText = "";
  await callback();
}

function setMessage(message) {
  document.getElementById("message").innerText = message;
}
