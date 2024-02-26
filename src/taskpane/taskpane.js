import { createCanvasBase64 } from "./canvasGenerator";
import { matchFiles } from "./wordMatching";
import { readImages } from "./io";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("insert-image").onclick = () => clearMessage(submitTextAndImages);
    document.getElementById("fileElem").onchange = () => clearMessage((e) => handleFiles(e));
  }
});

function getTextInput() {
  return document.getElementById("text-input").value;
}

function getFileInput() {
  return document.getElementById("fileElem").files;
}

function handleFiles(e) {
  var p = document.getElementById("file_names");
  var files = document.getElementById("fileElem").files;
  p.textContent = "";
  for (let i = 0; i < files.length; i++) {
    p.textContent += files[i].name;
    if (i !== files.length - 1) {
      p.textContent += ", ";
    }
  }
}

async function loadImages() {
  const files = matchFiles(getFileInput(), getTextInput());
  return readImages(files);
}

async function submitTextAndImages() {
  const images = await loadImages();
  const base64Image = createCanvasBase64(images, getTextInput());
  insertImage(base64Image);
}

function insertImage(image) {
  Office.context.document.setSelectedDataAsync(
    image,
    {
      coercionType: Office.CoercionType.Image,
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
