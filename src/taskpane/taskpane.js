import { createCanvasBase64 } from "./canvasGenerator";
import { matchFiles, getWords } from "./wordMatching";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("insert-image").onclick = () => clearMessage(submitTextAndImages);
    document.getElementById("fileElem").onchange = () => clearMessage(e => handleFiles(e));
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
      p.textContent += ", "
    }
  }
}

async function readImages() {
  const text = getTextInput();
  const files = getFileInput();

  let imagePromises = matchFiles(files, text).map(file => {
    const reader = new FileReader();
    return new Promise((resolve, reject) => {
      reader.onload = e => {
        const img = new Image();
        img.onload = () => resolve({ name: file.name.split(".")[0].toLowerCase(), image: img }); 
        img.onerror = reject; 
        img.src = e.target.result;
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  });
  let images = await Promise.all(imagePromises); 
  return new Map(images.map(obj => [obj.name, obj.image])); 
}

async function submitTextAndImages() {
  const images = await readImages();
  const base64Image = createCanvasBase64(images, getTextInput());
  insertImage(base64Image);
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
