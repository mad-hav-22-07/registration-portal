// Code generation: district number + letters, assigned by hashing with linear
// probing on collision. Kept in its own file so it can be tested without a DB.

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

const A = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.length; // 26

// 3 letters each = 17,576 slots per district. Two letters would only give 676,
// and some Kerala districts have more than 900 schools — a district could
// actually run out of codes.
export const ROLL_LETTERS = 3;
export const SCHOOL_LETTERS = 3;
export const rollSlots = A ** ROLL_LETTERS;
export const schoolSlots = A ** SCHOOL_LETTERS;

/**
 * FNV-1a. Any deterministic hash works; this one is short, has no dependencies
 * and spreads similar strings (ameya@x.com / ameyb@x.com) far apart, which is
 * what keeps probe runs short.
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
export function toLetters(slot, width) {
  let out = '';
  for (let i = 0; i < width; i++) {
    out = String.fromCharCode(65 + (slot % A)) + out;
    slot = Math.floor(slot / A);
  }
  return out;
}

export function normalize(name) {
  // "St. Thomas H.S.S" and "st thomas hss" must hash to the same slot, or the
  // same school registering twice gets two different codes.
  return name.toLowerCase().replace(/[^a-z0-9]/g, '');
}

/**
 * Walk the probe sequence for a key, yielding each candidate code in order.
 * The caller tries to claim each one and stops at the first that sticks — that
 * way the database's unique index is what resolves races, not a read-then-write
 * check that two simultaneous signups could both pass.
 */
export function* probe(district, key, { letters, slots }) {
  const start = hash(normalize(key)) % slots;
  for (let i = 0; i < slots; i++) {
    yield district + toLetters((start + i) % slots, letters);
  }
}

export const rollProbe = (district, email) =>
  probe(district, email, { letters: ROLL_LETTERS, slots: rollSlots });

export const schoolProbe = (district, schoolName) =>
  probe(district, schoolName, { letters: SCHOOL_LETTERS, slots: schoolSlots });

/**
 * Claim the first free code in the probe sequence.
 *
 * The insert itself is the claim — never "check if free, then take it", because
 * two people registering in the same millisecond would both pass that check and
 * be handed the same code. A unique-violation on the code column means somebody
 * else got that slot, so move to the next one.
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
      throw e; // a different unique violation (duplicate email) is a real error
    }
  }
  throw new Error('no codes left for this district');
}
