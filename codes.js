// Code generation. Every code is some prefix followed by three letters, and the
// letters come from hashing the person with linear probing on collision.
//
//   school code       07KBO     district + letters
//   school student    0710KEL   district + class + letters
//   individual        IN10KEL   class + letters, no district at all
//
// Each prefix is its own independent slot space, so a school student and an
// individual can never be issued the same code.

export const DISTRICTS = {
  '01': 'Thiruvananthapuram',
  '02': 'Kollam',
  '03': 'Pathanamthitta',
  '04': 'Alappuzha',
  '05': 'Kottayam',
  '06': 'Idukki',
  '07': 'Ernakulam',
  '08': 'Thrissur',
  '09': 'Palakkad',
  '10': 'Malappuram',
  '11': 'Kozhikode',
  '12': 'Wayanad',
  '13': 'Kannur',
  '14': 'Kasaragod',
};

export const CLASSES = ['8', '9', '10'];

const A = 26;
const LETTERS = 3;
export const slots = A ** LETTERS; // 17,576 per prefix

/**
 * Accept 8 / 9 / 10 however it arrives — "10", 10, "Class 10", "X", "8th" —
 * and return it zero-padded for use inside a code. null if it is not 8-10.
 */
export function parseClass(value) {
  // An array stringifies to "1,0" and became class 10; an object to
  // "[object Object]". Only a string or a number is a class.
  if (typeof value !== 'string' && typeof value !== 'number') return null;
  const raw = String(value).trim();
  if (!raw) return null;

  let n = null;
  const digits = raw.replace(/[^0-9]/g, '');

  if (digits) {
    // Digits always win. Checking roman numerals first made "Exam 9" parse as X
    // (the stray X) and land a class-9 student in class 10.
    // Number() also drops the padding Excel leaves on text cells: "08" -> 8.
    n = String(Number(digits));
  } else {
    // Roman only when the whole token is one, so "Sixth" does not become IX.
    // A trailing section letter is allowed: "VIII-A" is class 8.
    const WORD = /CLASS|STD|STANDARD|GRADE/;
    const alpha = raw.toUpperCase().replace(/[^A-Z]/g, '')
      .replace(new RegExp('^(' + WORD.source + ')'), '')
      .replace(new RegExp('(' + WORD.source + ')$'), '')
      // ordinal suffix before the section letter, or "Xth" loses its h to [A-H]
      .replace(/(TH|ST|ND|RD)$/, '')
      .replace(/[A-H]$/, '');
    n = { X: '10', IX: '9', VIII: '8' }[alpha] ?? null;
  }

  if (!CLASSES.includes(n)) return null;
  return n.padStart(2, '0'); // 8 -> "08", 10 -> "10"
}

/**
 * FNV-1a. Deterministic, dependency-free, and spreads near-identical strings
 * (ameya@x.com / ameyb@x.com) far apart, which is what keeps probe runs short.
 */
export function hash(str) {
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h >>> 0;
}

/** slot 0 -> "AAA", 1 -> "AAB", 17575 -> "ZZZ" */
export function toLetters(slot) {
  let out = '';
  for (let i = 0; i < LETTERS; i++) {
    out = String.fromCharCode(65 + (slot % A)) + out;
    slot = Math.floor(slot / A);
  }
  return out;
}

/**
 * Key used to decide whether two school names are "the same school".
 *
 * "St. Thomas H.S.S" and "st thomas hss" must collapse together, or a school
 * registering twice gets two codes. But letters of EVERY script have to survive:
 * stripping to [a-z0-9] turned every Malayalam-script name into the empty
 * string, so only one such school per district could ever register.
 *
 * Accents are folded rather than deleted, so "Kochí HSS" and "Kochi HSS" are one
 * school instead of two.
 */
export function normalize(name) {
  return String(name ?? '')
    .normalize('NFD')
    // Latin diacritics only (U+0300-U+036F). Stripping every combining mark
    // would delete Malayalam vowel signs and virama, which collapses genuinely
    // different names: കേരളം and കരളം are not the same school.
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    // keep marks, so Indic scripts survive intact
    .replace(/[^\p{L}\p{N}\p{M}]/gu, '');
}

/**
 * Walk the probe sequence for a key, yielding each candidate code in order. The
 * caller tries to claim each one and stops at the first that sticks, so the
 * database's unique index is what settles races — not a read-then-write check
 * that two simultaneous signups could both pass.
 */
export function* probe(prefix, key) {
  const start = hash(normalize(key)) % slots;
  for (let i = 0; i < slots; i++) {
    yield prefix + toLetters((start + i) % slots);
  }
}

/** 07KBO — a school, keyed on its name so re-registering returns the same code */
export const schoolProbe = (district, schoolName) => probe(district, schoolName);

/** 0710KEL — a student uploaded by a school in that district, in that class */
export const schoolStudentProbe = (district, classPadded, email) => probe(district + classPadded, email);

/** IN10KEL — someone who registered on their own. No district involved. */
export const individualProbe = (classPadded, email) => probe('IN' + classPadded, email);

/**
 * Claim the first free code in the probe sequence.
 *
 * The insert itself is the claim. A unique-violation on the code column means
 * somebody else took that slot, so move to the next one; any other error is
 * real and must not be swallowed.
 */
export async function claim(probeSeq, insert, codeConstraint) {
  let attempts = 0;
  for (const code of probeSeq) {
    attempts++;
    try {
      await insert(code);
      return { code, attempts };
    } catch (e) {
      // Only a violation of THIS code's own constraint means the slot is taken.
      // A /pkey/ pattern match would also swallow a 23505 from any other table,
      // turning an unrelated bug into 17,576 pointless retries.
      const tookThatSlot = e?.code === '23505' && !!codeConstraint && e.constraint === codeConstraint;
      if (tookThatSlot) continue;
      throw e;
    }
  }
  throw new Error('no codes left for this prefix');
}
