const matchExpression = /([\wäöüßÄÖÜ]+('[\wäöüßÄÖÜ]+)*)/g;

export function getWordsLowerCase(text) {
  return text.toLowerCase().match(matchExpression);
}

export function getWords(text) {
  return text.match(matchExpression);
}

export function matchFiles(files, text) {
  const words = getWordsLowerCase(text);
  return Array.from(files).filter((file) => words.includes(file.name.split(".")[0].toLowerCase()));
}
