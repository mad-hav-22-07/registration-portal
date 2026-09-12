import dotenv from 'dotenv';
import express from 'express';
import pg from 'pg';
import ExcelJS from 'exceljs';
import multer from 'multer';
import {
  DISTRICTS, CLASSES, parseClass, normalize, claim,
  schoolProbe, schoolStudentProbe, individualProbe,
} from './codes.js';
import { parseSheet, buildTemplate } from './sheet.js';

// `vercel env pull` writes .env.local, so read that first and fall back to .env.
// Earlier entries win, which keeps hand-written values in .env from being lost.
dotenv.config({ path: ['.env.local', '.env'], quiet: true });

const { DATABASE_URL, ADMIN_KEY, PORT = 3002 } = process.env;

const pool = new pg.Pool({
  connectionString: DATABASE_URL,
  // Verify the certificate properly - this connection carries names, emails and
  // phone numbers. A local postgres normally has no TLS, hence the exception.
  ssl: /localhost|127\.0\.0\.1/.test(DATABASE_URL || '') ? false : true,
  // One instance per serverless invocation, so keep the pool small; Neon caps
  // total connections across all warm instances.
  max: process.env.VERCEL ? 2 : 5,
  // Without these an overloaded pool makes requests queue forever instead of
  // failing fast.
  connectionTimeoutMillis: 10_000,
  statement_timeout: 15_000,
});

async function setup() {
  await pool.query(`
    -- People who registered on their own. Completely independent of schools;
    -- this is the set that will feed the payment gateway later.
    CREATE TABLE IF NOT EXISTS individuals (
      roll        TEXT PRIMARY KEY,
      class       TEXT NOT NULL,
      name        TEXT NOT NULL,
      email       TEXT NOT NULL,
      mobile      TEXT NOT NULL,
      created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
      CONSTRAINT individuals_email_unique UNIQUE (email)
    );

    CREATE TABLE IF NOT EXISTS schools (
      code        TEXT PRIMARY KEY,
      district    TEXT NOT NULL,
      name        TEXT NOT NULL,
      name_key    TEXT NOT NULL,
      poc_name    TEXT NOT NULL,
      poc_mobile  TEXT NOT NULL,
      email       TEXT NOT NULL,
      created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
      -- the same school registering twice gets its code back, not a second row
      CONSTRAINT schools_unique_per_district UNIQUE (district, name_key)
    );

    -- Students a school uploaded. Separate from individuals on purpose.
    CREATE TABLE IF NOT EXISTS school_students (
      roll        TEXT PRIMARY KEY,
      school_code TEXT NOT NULL REFERENCES schools(code),
      class       TEXT NOT NULL,
      name        TEXT NOT NULL,
      email       TEXT NOT NULL,
      mobile      TEXT NOT NULL,
      created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
      CONSTRAINT school_students_email_unique UNIQUE (email)
    );

    -- Postgres does not index a referencing column automatically, and both the
    -- school lookup and the export count students per school.
    CREATE INDEX IF NOT EXISTS school_students_by_school ON school_students (school_code);
  `);
  console.log('database ready');
}

/**
 * Trim and cap a user string. Control characters are stripped because Postgres
 * rejects 0x00 in a text parameter outright ("invalid byte sequence for encoding
 * UTF8"), and a single %00 in a URL was enough to take the whole server down.
 */
const clean = (v, max = 200) => {
  // Only strings; otherwise {"a":1} stringifies to "[object Object]" and
  // registers as a school name.
  const str = typeof v === 'string' ? v : typeof v === 'number' ? String(v) : '';
  return str.replace(/[\u0000-\u001f\u007f]/g, '').trim().slice(0, max);
};

/**
 * Same as clean(), but refuses anything over the limit instead of truncating.
 * Silent truncation merged two different long school names into one key and told
 * the second school it was already registered.
 */
/** Marks an error as the caller's fault, so the handler knows to answer 400. */
class BadInput extends Error {
  constructor(message) {
    super(message);
    this.name = 'BadInput';
    this.badInput = true;
  }
}
const bad = (m) => new BadInput(m);

const exact = (v, max, label) => {
  const str = typeof v === 'string' ? v : typeof v === 'number' ? String(v) : '';
  const out = str.replace(/[\u0000-\u001f\u007f]/g, '').trim();
  if ([...out].length > max) throw bad(`${label} is too long (max ${max} characters)`);
  return out;
};
/**
 * Deliberately stricter than the old /[^\s@]+@[^\s@]+\.[^\s@]{2,}/, which let
 * through paste artefacts from mail clients ("<ravi@x.in>", "ravi@x.in,",
 * "mailto:ravi@x.in") and addresses that cannot exist ("a@x..in", ".a@x.in",
 * "a@x.in."). Each of those gets a roll number and then never receives anything.
 */
const isEmail = (v) => {
  if (typeof v !== 'string') return false;
  if (!/^[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(\.[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*@[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?)+$/.test(v)) return false;
  return /\.[A-Za-z]{2,}$/.test(v); // a real TLD, so "a@b.1" is out
};
/**
 * Indian mobile, normalized to bare 10 digits.
 *
 * People type "09876543210", "+91 98765 43210", "0091-9876543210",
 * "91 9876543210" and "(+91) 9876543210". Strip leading zeros first, then a
 * country code, because "0091..." needs both.
 */
const mobile10 = (v) => {
  if (typeof v !== 'string' && typeof v !== 'number') return null;
  const raw = String(v).trim();
  // Refuse anything that is not recognisably a phone number. Stripping every
  // non-digit first meant "98765abc43299" was quietly stored as 9876543299 and
  // "my number is 9876512347" was accepted — a typo became a real number
  // belonging to someone else, and SMS went to the wrong person.
  if (!/^[+\d][\d\s().+-]*$/.test(raw)) return null;

  let d = raw.replace(/\D/g, '').replace(/^0+/, '');
  if (d.length === 12 && d.startsWith('91')) d = d.slice(2);
  return /^[6-9]\d{9}$/.test(d) ? d : null;
};

const app = express();
app.use(express.json());
app.use(express.static('public'));

/**
 * Serverless only: make sure the schema exists before the first request is
 * served, once per instance. Doing this at module load blocked cold starts for
 * up to 20s — longer than the function budget — and process.exit() on failure
 * would kill the whole instance. Must sit ahead of the routes to apply to them.
 */
let schemaReady = null;
if (process.env.VERCEL) {
  app.use((req, res, next) => {
    schemaReady ??= setupWithRetry(2);
    schemaReady.then(() => next()).catch((e) => {
      schemaReady = null; // let the next request retry rather than wedge forever
      next(e);
    });
  });
}

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024, files: 1 } });

/**
 * Express 4 does not catch rejections from async handlers: they become
 * unhandled rejections, and Node 24 exits the process on those. So every async
 * route goes through this, or one malformed request kills the site for everyone.
 */
const route = (fn) => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

/**
 * Small fixed-window limiter, per IP.
 *
 * The connection pool holds 5; without this, a few dozen concurrent lookups
 * starved it and a real registration went from 0.3s to 19s. It also slows
 * brute-forcing of the 5-character school codes to a crawl. In-memory on
 * purpose — per-instance is enough to protect the pool, and a shared store
 * would be a whole dependency for a one-week registration window.
 */
function rateLimit({ max, windowMs, message }) {
  const hits = new Map();
  return (req, res, next) => {
    const now = Date.now();
    const ip = req.ip || req.socket.remoteAddress || 'unknown';
    const slot = hits.get(ip);

    if (!slot || now > slot.resetAt) {
      hits.set(ip, { n: 1, resetAt: now + windowMs });
    } else if (++slot.n > max) {
      res.setHeader('Retry-After', Math.ceil((slot.resetAt - now) / 1000));
      return res.status(429).json({ error: message });
    }

    // Cheap sweep so the map cannot grow without bound.
    if (hits.size > 5000) {
      for (const [k, v] of hits) if (now > v.resetAt) hits.delete(k);
    }
    next();
  };
}

// Attached per route, not via app.use: a path prefix of '/api/schools' also
// matches '/api/schools/:code/upload', which would have locked a school out of
// re-uploading a corrected sheet after 20 attempts.
const readLimit = rateLimit({
  max: 60, windowMs: 60_000,
  message: 'too many requests — wait a minute and try again',
});
const writeLimit = rateLimit({
  max: 20, windowMs: 60_000,
  message: 'too many registration attempts — wait a minute and try again',
});
const uploadLimit = rateLimit({
  max: 12, windowMs: 60_000,
  message: 'too many uploads — wait a minute and try again',
});

// Districts go out as a sorted array, not an object: JS orders integer-like keys
// first, so a plain object would put 10-14 above 01-09 in the dropdown.
app.get('/api/options', (req, res) =>
  res.json({
    districts: Object.entries(DISTRICTS).sort(([a], [b]) => a.localeCompare(b)).map(([code, name]) => ({ code, name })),
    classes: CLASSES,
  }));

/* -- individual registration ---------------------------------------------- */

app.post('/api/individuals', writeLimit, route(async (req, res) => {
  try {
    const cls = parseClass(req.body.class);
    if (!cls) throw bad('pick your class (8, 9 or 10)');

    const name = exact(req.body.name, 100, 'name');
    if (name.length < 2) throw bad('enter your name');

    const email = exact(req.body.email, 254, 'email').toLowerCase();
    if (!isEmail(email)) throw bad('enter a valid email address');

    const mob = mobile10(req.body.mobile);
    if (!mob) throw bad('enter a valid 10-digit contact number');

    const { code: roll, attempts } = await claim(
      individualProbe(cls, email),
      (c) => pool.query(
        'INSERT INTO individuals (roll, class, name, email, mobile) VALUES ($1,$2,$3,$4,$5)',
        [c, cls, name, email, mob],
      ),
      'individuals_pkey',
    );

    console.log(`individual ${roll} - ${name}, class ${Number(cls)} (${attempts} probe${attempts > 1 ? 's' : ''})`);
    res.json({ roll, name, class: Number(cls) });
  } catch (e) {
    if (e.constraint === 'individuals_email_unique') {
      // Looking up their existing roll is a courtesy; if that query also fails we
      // still owe them the 409 rather than a hung request.
      const existing = await pool
        .query('SELECT roll FROM individuals WHERE email=$1', [clean(req.body.email, 254).toLowerCase()])
        .then((r) => r.rows[0]?.roll)
        .catch(() => null);
      return res.status(409).json({ error: `that email is already registered${existing ? ` - your roll number is ${existing}` : ''}` });
    }
    if (!e.badInput) throw e; // a DB failure is not the student's fault
    res.status(400).json({ error: e.message });
  }
}));

/* -- school registration --------------------------------------------------- */

app.post('/api/schools', writeLimit, route(async (req, res) => {
  try {
    // Compare the WHOLE value. Slicing to 2 characters turned "1234567" into
    // "12" and silently registered the school in Wayanad.
    const district = typeof req.body.district === 'string' ? req.body.district.trim() : '';
    if (!DISTRICTS[district]) throw bad('pick a district');

    const name = exact(req.body.name, 150, 'school name');
    if (name.length < 3) throw bad('enter the school name');
    // A name of only punctuation or emoji normalizes to an empty key, which
    // would collide with every other such name in the district.
    if (!normalize(name)) throw bad('the school name needs letters or digits in it');

    const pocName = exact(req.body.pocName, 100, 'contact person name');
    if (pocName.length < 2) throw bad('enter the contact person name');

    const poc = mobile10(req.body.pocMobile);
    if (!poc) throw bad('enter a valid 10-digit contact number');

    const email = exact(req.body.email, 254, 'email').toLowerCase();
    if (!isEmail(email)) throw bad('enter a valid email address');

    const nameKey = normalize(name);

    // Already registered? Hand back the existing code - the contact person has
    // almost certainly just lost it. Take the chance to update their contact
    // details, since re-registering is exactly how someone fixes a wrong number;
    // previously the new details were accepted and silently discarded.
    const found = await pool.query(
      `SELECT code, name, email, poc_mobile,
              (SELECT COUNT(*)::int FROM school_students s WHERE s.school_code = schools.code) AS students
         FROM schools WHERE district=$1 AND name_key=$2`,
      [district, nameKey],
    );

    if (found.rows.length) {
      const row = found.rows[0];
      // The school code is the only thing guarding student uploads, so do not
      // hand it to anyone who merely knows the school's name. The contact email
      // or number has to match what was registered.
      if (row.email !== email && row.poc_mobile !== poc) {
        return res.status(409).json({
          error: 'a school with this name is already registered in that district. '
               + 'Use the email or contact number it was registered with, or ask the organisers for the code.',
        });
      }
      // Matched, so this really is the school: refresh their contact details,
      // since re-registering is how someone fixes a wrong number.
      await pool.query('UPDATE schools SET poc_name=$2, poc_mobile=$3, email=$4 WHERE code=$1', [row.code, pocName, poc, email]);
      return res.json({
        code: row.code, name: row.name, district: DISTRICTS[district],
        students: row.students, already: true,
      });
    }

    const { code } = await claim(
      schoolProbe(district, name),
      (c) => pool.query(
        `INSERT INTO schools (code, district, name, name_key, poc_name, poc_mobile, email)
         VALUES ($1,$2,$3,$4,$5,$6,$7)`,
        [c, district, name, nameKey, pocName, poc, email],
      ),
      'schools_pkey',
    );

    console.log(`school ${code} - ${name}`);
    res.json({ code, name, district: DISTRICTS[district], students: 0, already: false });
  } catch (e) {
    if (e.constraint === 'schools_unique_per_district') {
      // Two simultaneous registrations: the unique index closed the window our
      // read-then-insert left open. Hand back the existing code rather than an
      // error with nothing in it — this is exactly the case the "already"
      // path exists for, and the school has no other way to get its code.
      const row = await pool
        .query(`SELECT code, name, district,
                       (SELECT COUNT(*)::int FROM school_students s WHERE s.school_code = schools.code) AS students
                  FROM schools WHERE district=$1 AND name_key=$2`,
               [typeof req.body.district === 'string' ? req.body.district.trim() : '', normalize(clean(req.body.name, 150))])
        .then((q) => q.rows[0])
        .catch(() => null);
      if (row) {
        return res.json({ code: row.code, name: row.name, district: DISTRICTS[row.district], students: row.students, already: true });
      }
      return res.status(409).json({ error: 'this school is already registered in that district' });
    }
    if (!e.badInput) throw e;
    res.status(400).json({ error: e.message });
  }
}));

/** Look up a school by code, so a returning school can confirm it is the right one. */
app.get('/api/schools/:code', readLimit, route(async (req, res) => {
  const code = clean(req.params.code, 10).toUpperCase();
  const { rows } = await pool.query(`
    SELECT code, name, district,
           (SELECT COUNT(*)::int FROM school_students s WHERE s.school_code = schools.code) AS students
      FROM schools WHERE code=$1`, [code]);
  if (!rows.length) return res.status(404).json({ error: 'no school with that code' });
  // Same shape as the POST responses: `district` is always the readable name.
  const r = rows[0];
  res.json({ code: r.code, name: r.name, district: DISTRICTS[r.district], districtCode: r.district, students: r.students });
}));

/* -- school uploads its student sheet -------------------------------------- */

/** Blank sheet for schools to fill in, so the columns come back as expected. */
app.get('/api/template.xlsx', route(async (req, res) => {
  const wb = await buildTemplate();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="student-list-template.xlsx"');
  await wb.xlsx.write(res);
  res.end();
}));

app.post('/api/schools/:code/upload', uploadLimit, upload.single('file'), route(async (req, res) => {
  try {
    const schoolCode = clean(req.params.code, 10).toUpperCase();
    const school = await pool.query('SELECT code, district, name, email, poc_mobile FROM schools WHERE code=$1', [schoolCode]);
    if (!school.rows.length) throw bad('that school code does not exist');

    // A 5-character code is guessable inside a 17,576-code district space, and
    // roll numbers are permanent once issued — so prove you are the school by
    // also giving the email or contact number it was registered with.
    const who = clean(req.body.email, 150).toLowerCase();
    const whoMobile = mobile10(req.body.email);
    const row = school.rows[0];
    if (who !== row.email && whoMobile !== row.poc_mobile) {
      throw bad('the email or contact number does not match the one this school registered with');
    }

    if (!req.file) throw bad('attach the filled-in sheet');

    const { rows, columns, blankRows, sheetName } = await parseSheet(req.file.buffer, req.file.originalname);
    if (!rows.length) throw bad('the sheet has column headings but no rows below them');

    const district = school.rows[0].district;
    const added = [];
    const skipped = [];
    const seen = new Set(); // the same student listed twice in one sheet

    for (const r of rows) {
      const name = clean(r.name, 300);
      const email = clean(r.email, 300).toLowerCase();
      const mob = mobile10(r.mobile);
      const cls = parseClass(r.class);

      if (name.length < 2) { skipped.push({ ...r, reason: 'no name' }); continue; }
      // Report over-length rather than trimming: a sliced email is a different,
      // non-existent address and that student would never be contacted.
      if (name.length > 100) { skipped.push({ ...r, reason: 'name is longer than 100 characters' }); continue; }
      if (email.length > 254) { skipped.push({ ...r, reason: 'email is too long' }); continue; }
      if (!isEmail(email)) { skipped.push({ ...r, reason: email ? `"${email}" is not a valid email` : 'no email' }); continue; }
      if (!mob) { skipped.push({ ...r, reason: r.mobile ? `"${r.mobile}" is not a valid 10-digit number` : 'no contact number' }); continue; }
      if (!cls) { skipped.push({ ...r, reason: r.class ? `class "${r.class}" is not 8, 9 or 10` : 'no class' }); continue; }
      if (seen.has(email)) { skipped.push({ ...r, reason: 'listed twice in this sheet' }); continue; }
      seen.add(email);

      try {
        const { code: roll } = await claim(
          schoolStudentProbe(district, cls, email),
          (c) => pool.query(
            'INSERT INTO school_students (roll, school_code, class, name, email, mobile) VALUES ($1,$2,$3,$4,$5,$6)',
            [c, schoolCode, cls, name, email, mob],
          ),
          'school_students_pkey',
        );
        added.push({ row: r.row, name, email, mobile: mob, class: Number(cls), roll });
      } catch (e) {
        if (e.constraint === 'school_students_email_unique') {
          const existing = await pool
            .query('SELECT roll FROM school_students WHERE email=$1', [email])
            .then((q) => q.rows[0]?.roll)
            .catch(() => null);
          skipped.push({ ...r, reason: `already has a roll number${existing ? ` - ${existing}` : ''}` });
        } else {
          skipped.push({ ...r, reason: e.message });
        }
      }
    }

    console.log(`upload for ${schoolCode}: ${added.length} added, ${skipped.length} skipped of ${rows.length} rows`);
    // dataRows lets the school check nothing was lost: added + skipped must equal it
    res.json({
      schoolCode, schoolName: school.rows[0].name, sheetName,
      columnsFound: columns, dataRows: rows.length, blankRows,
      added, skipped,
    });
  } catch (e) {
    // parseSheet's messages are all user-facing guidance about their file
    if (!e.badInput && !/sheet|file|column|heading|xlsx|csv/i.test(e.message || '')) throw e;
    console.error('upload rejected:', e.message);
    res.status(400).json({ error: e.message });
  }
}));

/* -- admin: counts + Excel export ------------------------------------------ */

function checkAdmin(req, res) {
  if (!ADMIN_KEY) {
    res.status(500).json({ error: 'ADMIN_KEY is not set on the server' });
    return false;
  }
  if ((req.query.key ?? req.header('x-admin-key')) !== ADMIN_KEY) {
    res.status(401).json({ error: 'wrong admin key' });
    return false;
  }
  return true;
}

app.get('/api/stats', route(async (req, res) => {
  if (!checkAdmin(req, res)) return;
  const q = async (sql) => (await pool.query(sql)).rows[0].n;
  res.json({
    individuals: await q('SELECT COUNT(*)::int n FROM individuals'),
    schools: await q('SELECT COUNT(*)::int n FROM schools'),
    schoolStudents: await q('SELECT COUNT(*)::int n FROM school_students'),
  });
}));

app.get('/api/export.xlsx', route(async (req, res) => {
  if (!checkAdmin(req, res)) return;

  const individuals = await pool.query('SELECT roll, name, email, mobile, class, created_at FROM individuals ORDER BY class, roll');
  const students = await pool.query(`
    SELECT s.roll, s.name, s.email, s.mobile, s.class, s.school_code, sc.name AS school_name, sc.district, s.created_at
      FROM school_students s JOIN schools sc ON sc.code = s.school_code
     ORDER BY sc.district, s.school_code, s.class, s.roll`);
  const schools = await pool.query(`
    SELECT code, name, poc_name, poc_mobile, email, district,
           (SELECT COUNT(*)::int FROM school_students s WHERE s.school_code = schools.code) AS students
      FROM schools ORDER BY district, code`);

  const wb = new ExcelJS.Workbook();
  const sheet = (name, columns, rows) => {
    const ws = wb.addWorksheet(name);
    ws.columns = columns;
    ws.addRows(rows);
    ws.getRow(1).font = { bold: true };
    ws.views = [{ state: 'frozen', ySplit: 1 }]; // header stays put while scrolling
    ws.autoFilter = { from: 'A1', to: { row: 1, column: columns.length } };
    // Keep phone numbers as text, or Excel shows them as 9.87654E+09.
    const mobCol = columns.findIndex((c) => c.key === 'mobile' || c.key === 'poc_mobile');
    if (mobCol >= 0) ws.getColumn(mobCol + 1).numFmt = '@';
  };

  const when = (r) => ({ ...r, registered_at: new Date(r.created_at).toISOString().slice(0, 16).replace('T', ' ') });

  sheet('Individuals', [
    { header: 'Roll Number', key: 'roll', width: 14 },
    { header: 'Name', key: 'name', width: 26 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Contact Number', key: 'mobile', width: 16 },
    { header: 'Class', key: 'class_n', width: 8 },
    { header: 'Registered At', key: 'registered_at', width: 18 },
  ], individuals.rows.map((r) => ({ ...when(r), class_n: Number(r.class) })));

  sheet('School Students', [
    { header: 'Roll Number', key: 'roll', width: 14 },
    { header: 'Name', key: 'name', width: 26 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Contact Number', key: 'mobile', width: 16 },
    { header: 'Class', key: 'class_n', width: 8 },
    { header: 'School Code', key: 'school_code', width: 13 },
    { header: 'School Name', key: 'school_name', width: 32 },
    { header: 'District', key: 'district_name', width: 20 },
    { header: 'Registered At', key: 'registered_at', width: 18 },
  ], students.rows.map((r) => ({ ...when(r), class_n: Number(r.class), district_name: DISTRICTS[r.district] })));

  sheet('Schools', [
    { header: 'School Code', key: 'code', width: 13 },
    { header: 'School Name', key: 'name', width: 32 },
    { header: 'Contact Person', key: 'poc_name', width: 22 },
    { header: 'Contact Number', key: 'poc_mobile', width: 16 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'District', key: 'district_name', width: 20 },
    { header: 'Students Uploaded', key: 'students', width: 18 },
  ], schools.rows.map((r) => ({ ...r, district_name: DISTRICTS[r.district] })));

  const stamp = new Date().toISOString().slice(0, 10);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="registrations-${stamp}.xlsx"`);
  await wb.xlsx.write(res);
  res.end();
  console.log(`exported ${individuals.rowCount} individuals, ${students.rowCount} school students, ${schools.rowCount} schools`);
}));

// An unmatched /api path otherwise falls through to Express's HTML 404 page, and
// the browser's `r.json()` then throws "Unexpected token '<'", which the user
// sees as nothing happening at all.
app.use('/api', (req, res) => res.status(404).json({ error: 'no such endpoint' }));

/**
 * Last line of defence. Anything a route throws lands here, so the client gets
 * a clean message and the process stays up. Internal detail (SQL text, the
 * connection string, stack traces) is logged but never sent to the browser.
 */
app.use((err, req, res, next) => {
  if (err?.code === 'LIMIT_FILE_SIZE') {
    return res.status(413).json({ error: 'that file is larger than 10MB — split the sheet or remove images from it' });
  }
  if (err?.code === 'LIMIT_FILE_COUNT' || err?.code === 'LIMIT_UNEXPECTED_FILE') {
    return res.status(400).json({ error: 'attach a single sheet' });
  }
  // express.json() rejects bad input with a SyntaxError; that is the caller's
  // fault, not ours, so it must not read as a server error
  if (err?.type === 'entity.parse.failed') {
    return res.status(400).json({ error: 'the request body was not valid JSON' });
  }
  if (err?.type === 'entity.too.large') {
    return res.status(413).json({ error: 'that request was too large' });
  }
  console.error(`${req.method} ${req.path} failed:`, err?.message || err);
  if (res.headersSent) return next(err);
  res.status(500).json({ error: 'something went wrong on the server — try again' });
});

/**
 * Neon's free tier suspends the database when idle, so the first connection
 * after a quiet spell can be refused - sometimes with an empty error message.
 * Retry rather than crashing on a sleeping database.
 */
async function setupWithRetry(tries = 5) {
  for (let i = 1; i <= tries; i++) {
    try {
      await setup();
      return;
    } catch (e) {
      const why = e.message || e.code || 'connection refused';
      if (i === tries) throw new Error(why);
      console.log(`database not ready (${why}) - retrying in ${2 * i}s [${i}/${tries - 1}]`);
      await new Promise((r) => setTimeout(r, 2000 * i));
    }
  }
}

if (!process.env.VERCEL) {
  setupWithRetry()
    .then(() => app.listen(PORT, () => console.log(`open http://localhost:${PORT}`)))
    .catch((e) => {
      console.error('could not start:', e.message);
      if (!DATABASE_URL) console.error('DATABASE_URL is not set - see README');
      process.exit(1);
    });
}

export default app;
