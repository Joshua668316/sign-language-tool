import { readWordCSV } from "./io";

const matchExpression = /([\wäöüßÄÖÜ]+('[\wäöüßÄÖÜ]+)*)/g;
let wordlist = new Map();

export async function initWordList() {
  wordlist = await readWordCSV();
}

export function getBasicWord(word) {
  return wordlist.get(word)
}

export function matchWordToImage(word, images) {
  word = word.toLowerCase();
  if (images.has(word)) {
    return images.get(word);
  } 
  return advancedWord2Image(word, images);
}

function advancedWord2Image(word, images) {
  if (!wordlist.has(word)) {
    return null;
  }
  const basicWord = wordlist.get(word);
  if (images.has(basicWord)) {
    return images.get(basicWord);
  }
  return null;
}

export function getWordsLowerCase(text) {
  return text.toLowerCase().match(matchExpression);
}

export function getWords(text) {
  return text.match(matchExpression);
}

export function matchFiles(files, text) {
  const words = getWordsLowerCase(text);
  words.forEach(word => {
    if (wordlist.has(word)) {
      words.push(wordlist.get(word));
    }
  })
  return Array.from(files).filter((file) => words.includes(file.name.split(".")[0].toLowerCase()));
}
