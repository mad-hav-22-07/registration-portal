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
  const raw = String(value ?? '').trim();
  if (!raw) return null;

  const roman = { X: '10', IX: '9', VIII: '8' }[raw.toUpperCase().replace(/[^IVX]/g, '')];
  const digits = roman ?? raw.replace(/[^\d]/g, '');
  if (!CLASSES.includes(digits)) return null;
  return digits.padStart(2, '0'); // 8 -> "08", 10 -> "10"
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

export function normalize(name) {
  // "St. Thomas H.S.S" and "st thomas hss" must hash to the same slot, or one
  // school registering twice would be given two different codes.
  return name.toLowerCase().replace(/[^a-z0-9]/g, '');
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
      const tookThatSlot =
        e.code === '23505' && (e.constraint === codeConstraint || /pkey/.test(e.constraint || ''));
      if (tookThatSlot) continue;
      throw e;
    }
  }
  throw new Error('no codes left for this prefix');
}
