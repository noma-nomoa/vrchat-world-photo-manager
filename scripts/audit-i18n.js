const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const REPO_ROOT = path.resolve(__dirname, '..');
const SOURCE_FILES = [
  'src/index.html',
  'src/renderer.js',
];
if (process.argv.includes('--include-main')) {
  SOURCE_FILES.push('src/index.js');
}
const LANGUAGES = ['en', 'ko'];
const HAS_JAPANESE_RE = /[\u3040-\u30ff\u3400-\u9fff]/;
const ALLOWED_UNTRANSLATED_TEXTS = new Set(['日本語']);
const STRING_LITERAL_RE = /(["'`])(?:\\[\s\S]|(?!\1)[^\\])*\1/g;
const HTML_TEXT_RE = />\s*([^<>]*[\u3040-\u30ff\u3400-\u9fff][^<>]*)\s*</g;
const HTML_ATTR_RE =
  /\b(?:aria-label|title|placeholder|label|alt)\s*=\s*(["'])([\s\S]*?)\1/g;

function readSource(relativePath) {
  return fs.readFileSync(path.join(REPO_ROOT, relativePath), 'utf8');
}

function createTranslator(language) {
  const source = readSource('src/i18n.js');
  const listeners = new Map();
  const context = {
    console,
    Intl,
    Map,
    Set,
    WeakMap,
    String,
    Number,
    Date,
    RegExp,
    Event: class Event {
      constructor(type) {
        this.type = type;
      }
    },
    CustomEvent: class CustomEvent {
      constructor(type, options = {}) {
        this.type = type;
        this.detail = options.detail;
      }
    },
    Element: class Element {},
    Node: {
      TEXT_NODE: 3,
      ELEMENT_NODE: 1,
    },
    NodeFilter: {
      SHOW_ELEMENT: 1,
      SHOW_TEXT: 4,
    },
    MutationObserver: class MutationObserver {
      observe() {}
    },
  };

  context.window = {
    localStorage: {
      getItem: () => language,
      setItem: () => {},
    },
    dispatchEvent: () => {},
    addEventListener: (eventName, listener) => {
      listeners.set(eventName, listener);
    },
  };
  context.document = {
    body: null,
    documentElement: {
      lang: '',
    },
    addEventListener: (eventName, listener) => {
      listeners.set(eventName, listener);
    },
  };
  context.window.window = context.window;
  context.window.document = context.document;

  vm.runInNewContext(source, context, {
    filename: path.join(REPO_ROOT, 'src/i18n.js'),
  });

  return (text) => context.window.WorldShotI18n.t(text);
}

function decodeJsLiteral(literal) {
  if (literal.startsWith('`') && literal.includes('${')) {
    return null;
  }

  try {
    return vm.runInNewContext(literal);
  } catch {
    return null;
  }
}

function normalizeCandidate(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function shouldKeepCandidate(value) {
  const text = normalizeCandidate(value);
  if (!text || !HAS_JAPANESE_RE.test(text)) {
    return false;
  }

  if (/<\/?[a-z][\s\S]*>/i.test(text)) {
    return false;
  }

  if (text.length > 180) {
    return false;
  }

  if (/^[\d\s/年月日:.:-]+$/.test(text)) {
    return false;
  }

  return true;
}

function collectCandidates() {
  const candidates = new Map();

  function addCandidate(text, source) {
    const normalized = normalizeCandidate(text);
    if (!shouldKeepCandidate(normalized)) {
      return;
    }

    if (!candidates.has(normalized)) {
      candidates.set(normalized, {
        text: normalized,
        sources: new Set(),
      });
    }
    candidates.get(normalized).sources.add(source);
  }

  for (const relativePath of SOURCE_FILES) {
    const source = readSource(relativePath);

    if (relativePath.endsWith('.html')) {
      for (const match of source.matchAll(HTML_TEXT_RE)) {
        addCandidate(match[1], relativePath);
      }
      for (const match of source.matchAll(HTML_ATTR_RE)) {
        addCandidate(match[2], relativePath);
      }
      continue;
    }

    for (const match of source.matchAll(STRING_LITERAL_RE)) {
      const decoded = decodeJsLiteral(match[0]);
      if (typeof decoded === 'string') {
        addCandidate(decoded, relativePath);
      }
    }
  }

  return Array.from(candidates.values()).map((entry) => ({
    text: entry.text,
    sources: Array.from(entry.sources).sort(),
  }));
}

function isMissingTranslation(original, translated) {
  const normalizedOriginal = normalizeCandidate(original);
  const normalizedTranslated = normalizeCandidate(translated);
  if (ALLOWED_UNTRANSLATED_TEXTS.has(normalizedOriginal)) {
    return false;
  }

  return (
    normalizedTranslated === normalizedOriginal ||
    HAS_JAPANESE_RE.test(normalizedTranslated)
  );
}

function runAudit() {
  const translators = Object.fromEntries(
    LANGUAGES.map((language) => [language, createTranslator(language)])
  );
  const candidates = collectCandidates();
  const missing = [];

  for (const candidate of candidates) {
    for (const language of LANGUAGES) {
      const translated = translators[language](candidate.text);
      if (isMissingTranslation(candidate.text, translated)) {
        missing.push({
          language,
          text: candidate.text,
          translated,
          sources: candidate.sources,
        });
      }
    }
  }

  console.log(`i18n audit candidates: ${candidates.length}`);
  console.log(`i18n audit missing: ${missing.length}`);

  if (missing.length > 0) {
    for (const item of missing.slice(0, 80)) {
      console.log(
        `[${item.language}] ${item.text} -> ${item.translated} (${item.sources.join(', ')})`
      );
    }
  }

  if (process.argv.includes('--json')) {
    console.log(
      JSON.stringify(
        {
          candidateCount: candidates.length,
          missingCount: missing.length,
          missing,
        },
        null,
        2
      )
    );
  }

  if (process.argv.includes('--strict') && missing.length > 0) {
    process.exitCode = 1;
  }
}

runAudit();
