// Stop words that match exactly
const STOP_EXACT = [
  // Preposition, articles and subjects
  "il", "lo", "la", "li", "i", "gli", "le", "un", "uno", "una",
  "di", "a", "da", "in", "con", "su", "per", "tra", "fra",
  "due", "tre", "dei", "del", "all", "ai", "al", "ad",
  "dal", "dall", "nel", "sul", "sugli", "degli", "agli", "negli",
  "verso",
  // Pronouns
  "mi", "ci", "ti", "si", "me", "cui",
  "io", "tu", "lui", "lei", "noi", "voi", "essi", "essa", "esso", "loro",
  // Common adjectives
  "bene", "male", "vari", "varie", "qualche",
  // Common names
  "sorta", "cosa", "qualcosa", "posto", "posti", "zona", "parte", "volta", "volte",
  // Adverbs and conjunctions
  "solamente", "sopra", "sotto", "giù", "già", "senza", "ora", "adesso",
  "avanti", "dietro", "comunque", "dove", "quando", "poi", "prima", "dopo", "mentre",
  "quindi", "ancora", "là", "lì", "più", "meno", "po’", "po",
  "e", "ed", "ma", "se", "perché", "perchè", "come", "che", "anche", "ne", "non",
  // Irregular verbs
  "avere", "ho", "hai", "ha", "abbiamo", "avete", "hanno",
  "essere", "sono", "sei", "è", "c’è", "siamo", "siete", "sono", "sia",
  "andare", "vado", "vai", "va", "andiamo", "andate", "vanno",
  "venire", "vengo", "vieni", "viene", "veniamo", "venite", "vengono",
  "dare", "do", "dai", "dà", "diamo", "date", "danno", "darmi",
  "dire", "dico", "dici", "dice", "diciamo", "dite", "dicono",
  "fare", "faccio", "fai", "fa", "facciamo", "fate", "fanno", "farmi", "farlo", "farli",
  "potere", "posso", "puoi", "può", "possiamo", "potete", "possono",
  "volere", "voglio", "vuoi", "vuole", "vogliamo", "volete", "vogliono", "volermi",
  "sapere", "so", "sai", "sa", "sappiamo", "sapete", "sanno",
  "stare", "sto", "stai", "sta", "stiamo", "state", "stanno", "starmi",
  // Regular verbs
  "riuscire", "riusciamo", "riuscite", "riescono", "risucirmi",
  "dovere", "dobbiamo", "dovete", "devono", "dovermi",
  "parlare", "parliamo", "parlate", "parlano", "parlarmi",
  "chiedere", "chiediamo", "chiedete", "chiedono", "chiedermi",
  "prendere", "prendiamo", "prendete", "prendono", "prendermi",
  "decidere", "decidiamo", "decidete", "decidono",
  "sembrare", "sembriamo", "sembrate", "sembrano", "sembrarmi",
  "rendere", "rendiamo", "rendete", "rendono", "rendermi",
  "provare", "proviamo", "provate", "provano", "provarmi",
  "arrivare", "arriviamo", "arrivate", "arrivano",
  "pare",
  // Sense regular verbs
  "vedere", "vediamo", "vedete", "vedono", "vedermi", "visto", "visti", "viste",
  "guardare", "guardiamo", "guardate", "guardano", "guardarmi",
  "notare", "notiamo", "notate", "notano"
];

// Wildcard stop patterns
const STOP_WILDCARD = [
  // Common nouns
  "oggett*", "lat*", "qualcun*",
  // Common prepositions
  "quest*", "quell*", "dell*", "del*",
  "all*", "sull*", "dall*", "nell*",
  // Common adjectives
  "molt*", "poc*", "poch*", "tropp*", "altr*", "vari*",
  "tutt*", "alt*", "bass*", "lung*", "alcun*", "grand*",
  // Pronouns
  "mi*", "tu*", "su*",
  // Irregular Verbs
  "stat*", "avut*",
  "fatt*", "arrivat*",
  // Regular verbs
  "riesc*", "dev*",
  "parl*", "chied*",
  "prend*", "decid*", "sembr*",
  "rend*", "prov*", "arriv*",
  // Sense regular verbs
  "ved*", "guard*", "not*"
];

// Stop phrases to delete entirely
const STOP_PHRASES = [
  "rendere conto", "rendo conto", "rendi conto", "rende conto",
  "rendiamo conto", "rendete conto", "rendono conto", "rendermi conto",
  "andare via", "vado via", "andiamo via", "mettere via",
  "metterlo via", "metterla via"
];

// Synonyms list for single words
const EQUIVALENT_WORDS = {
  "avanti": "davanti",
  "retro": "dietro",
  "sù": "sopra",
  "giù": "sotto",

  // Temporal
  "precedentemente": "prima",
  "poi": "dopo",
}


function normalizeWordEquivalents(word) {
  return EQUIVALENT_WORDS[word] || word;
}


// Returns true if words matches any exact stop word or any wildcard pattern
function isStopWord(w) {
  if (STOP_EXACT.includes(w)) return true;

  return STOP_WILDCARD.some(pat => {
    if (!pat.endsWith('*')) return false;
    const base = pat.slice(0, -1);
    // match only if word starts with base AND has exactly one more character
    return w.startsWith(base) && w.length === base.length + 1;
  });
}



// Returns the singular or plural form(s) of a word that appear in the text, or [] if none
function getBaseForms(word, rawFreq) {
  const forms = [];
  const len = word.length;

  const IRREGULAR_MAP = {
    // Singular -> Plural
    "uomo": "uomini", "dio": "dei", "tempio": "templi", "bue": "buoi",
    "moglie": "mogli", "ciliegia": "ciliegie", "arma": "armi", "ala": "ali",
    "osso": "ossa", "uovo": "uova",
    // Plural -> Singular
    "uomini": "uomo", "dei": "dio", "templi": "tempio", "buoi": "bue",
    "mogli": "moglie", "ciliegie": "ciliegia", "armi": "arma", "ali": "ala",
    "ossa": "osso", "uova": "uovo",
  };

  // Check irregular forms and return those present in text
  if (IRREGULAR_MAP[word]) {
    forms.push(IRREGULAR_MAP[word]);
    return forms.filter(f => rawFreq[f]);
  }

  // Masculine nouns
  if (word.endsWith("i") && len > 3) {
    forms.push(word.slice(0, -1) + "o");    // dadi -> dado
    forms.push(word.slice(0, -1) + "e");    // fiori -> fiore
    forms.push(word.slice(0, -1) + "io");   // tizi -> tizio
  }
  if (word.endsWith("o") && len > 3) {
    forms.push(word.slice(0, -1) + "i");    // dado -> dadi
    forms.push(word.slice(0, -1));          // tizio -> tizi
  }
  if (word.endsWith("io") && len > 3) {
    forms.push(word.slice(0, -2) + "i");    // tizio -> tizi
  }
  if (word.endsWith("e") && len > 3) {
    forms.push(word.slice(0, -1) + "i");    // fiore -> fiori
  }

  // "H" added for masculine plural forms
  if (word.endsWith("co")) {
    forms.push(word.slice(0, -2) + "chi");  // parco -> parchi
  }
  if (word.endsWith("go")) {
    forms.push(word.slice(0, -2) + "ghi");  // fungo -> funghi
  }
  if (word.endsWith("chi")) forms.push(word.slice(0, -3) + "co");   // parchi -> parco
  if (word.endsWith("ghi")) forms.push(word.slice(0, -3) + "go");   // funghi -> fungo


  // Feminine nouns
  if (word.endsWith("e") && len > 3) {
    forms.push(word.slice(0, -1) + "a");    // case -> casa
  }
  if (word.endsWith("a") && len > 3) {
    forms.push(word.slice(0, -1) + "e");    // casa -> case
  }

  // "H" added for feminine plural forms
  if (word.endsWith("ca")) {
    forms.push(word.slice(0, -2) + "che");  // esca -> esche
  }
  if (word.endsWith("ga")) {
    forms.push(word.slice(0, -2) + "ghe");  // alga -> alghe
  }
  if (word.endsWith("che")) {
    forms.push(word.slice(0, -3) + "ca");   // esche -> esca
  }
  if (word.endsWith("ghe")) {
    forms.push(word.slice(0, -3) + "ga");   // alghe -> alga
  }

  // Other feminine rules
  if (word.length > 4) {
    const preChar = word[word.length - 4].toLowerCase();

    if (word.endsWith("cia") && isVowel(preChar)) {
      forms.push(word.slice(0, -3) + "ce");             // arancia -> arance
    } else if (word.endsWith("ce") && isVowel(preChar)) {
      forms.push(word.slice(0, -2) + "cia");            // arance -> arancia
    }

    if (word.endsWith("gia") && isVowel(preChar)) {
      forms.push(word.slice(0, -3) + "ge");             // ciliegia -> ciliege
    } else if (word.endsWith("ge") && isVowel(preChar)) {
      forms.push(word.slice(0, -2) + "gia");            // ciliege -> ciliegia
    }
  }


  // Other forms
  if (word.endsWith("ista")) forms.push(word.slice(0, -4) + "isti");  // turista -> turisti
  if (word.endsWith("isti")) forms.push(word.slice(0, -4) + "ista");  // turisti -> turista

  // Keep -iste separate for simplicity's sake and to avoid forming triplets
  //if (word.endsWith("ista")) forms.push(word.slice(0, -4) + "iste");  // turista -> turiste
  //if (word.endsWith("iste")) forms.push(word.slice(0, -4) + "ista");  // turiste -> turista


  // Words ending with accented vowels are invariant; return no forms
  // caffè -> caffè
  if (["à", "è", "ì", "ò", "ù"].some(accent => word.endsWith(accent))) {
    return [];
  }

  // Return only forms that exist in the text
  const filtered = forms.filter(f => rawFreq[f]);
  return filtered;
}

function isVowel(char) {
  return 'aeiou'.includes(char.toLowerCase());
}


// Orders singular/plural pairs that aren't automatically sorted correctly (alphabetically)
function orderSingularPlural(w1, w2) {
  const endsO1 = w1.endsWith('o');
  const endsI1 = w1.endsWith('i');
  const endsO2 = w2.endsWith('o');
  const endsI2 = w2.endsWith('i');

  if (endsO1 && endsI2) return [w1, w2].join("/");
  if (endsI1 && endsO2) return [w2, w1].join("/");

  // Fallback alphabetical
  return [w1, w2].sort().join("/");
}


// Tokenize and filter by:
// - changing to lowercase
// - removing STOP phrases and STOP words, 
// - removing punctuation
// - removing non-dream text inside curly bracers {comment} 
function tokenizeAndFilter(rawText) {
  let txt = rawText.toLowerCase();

  // Remove specific STOP phrases
  STOP_PHRASES.forEach(phrase => {
    const re = new RegExp(`\\b${phrase}\\b`, "gi");
    txt = txt.replace(re, " ");
  });

  // Strip comments and non-dream text inside {}
  txt = txt.replace(/\{.*?\}/g, "");

  // Replace all kinds of punctuation with space
  txt = txt.replace(/[.,;:!?"“”«»()\[\]{}—–\-…'’]/g, " ");

  // Normalize whitespace
  txt = txt.replace(/\s+/g, " ").trim();

  // Tokenize, filter short words and stopwords
  return txt
    .split(" ")
    .filter(w => w.length >= 2 && !isStopWord(w));
}

// Returns the longest common prefix of two strings 
function longestCommonPrefix(a, b) {
  const n = Math.min(a.length, b.length);
  let i = 0;
  while (i < n && a[i] === b[i]) {
    i++;
  }
  return a.slice(0, i);
}

// Shortens singluar-plural couple for visualizing in Words Cloud
function shortenSingPlu(display) {
  if (typeof display !== 'string' || !display.includes('/')) return display;

  const [sing, plu] = display.split('/');
  const prefix = longestCommonPrefix(sing, plu);

  // Exclude short words 
  if (prefix.length < 2) return display;

  const tail = plu.slice(prefix.length);
  return tail ? `${sing}/${tail}` : display;
}

// Helper function to calculate percentile
function percentile(arr, p) {
  if (arr.length === 0) return 0;
  const sorted = arr.slice().sort((a, b) => a - b);
  const index = (p / 100) * (sorted.length - 1);
  const lower = Math.floor(index);
  const upper = lower + 1;
  const weight = index - lower;
  if (upper >= sorted.length) return sorted[lower];
  return sorted[lower] * (1 - weight) + sorted[upper] * weight;
}



