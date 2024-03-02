export async function readImages(files) {
  let imagePromises = files.map((file) => {
    const reader = new FileReader();
    return new Promise((resolve, reject) => {
      reader.onload = (e) => {
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
  return new Map(images.map((obj) => [obj.name, obj.image]));
}

export async function readWordCSV() {
  const wordsPath = "assets/Tabelle_form_of.csv";
  const csv = await fetch(wordsPath);
  const text = await csv.text();
  const lines = text.trim().split(/\r?\n/);
  const map = new Map();
  lines.forEach(line => {
    const pair = line.split(";");
    if (pair.length === 2) {
      map.set(pair[0], pair[1]);
    }
  })
  return map;
}
