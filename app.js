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
  max: 5,
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
  `);
  console.log('database ready');
}

const clean = (v, max = 200) => String(v ?? '').trim().slice(0, max);
const isEmail = (v) => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v);
// Indian mobile: 10 digits starting 6-9, with or without +91
const mobile10 = (v) => {
  const d = String(v ?? '').replace(/\D/g, '').replace(/^91(?=\d{10}$)/, '');
  return /^[6-9]\d{9}$/.test(d) ? d : null;
};

const app = express();
app.use(express.json());
app.use(express.static('public'));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024, files: 1 } });

// Districts go out as a sorted array, not an object: JS orders integer-like keys
// first, so a plain object would put 10-14 above 01-09 in the dropdown.
app.get('/api/options', (req, res) =>
  res.json({
    districts: Object.entries(DISTRICTS).sort(([a], [b]) => a.localeCompare(b)).map(([code, name]) => ({ code, name })),
    classes: CLASSES,
  }));

/* -- individual registration ---------------------------------------------- */

app.post('/api/individuals', async (req, res) => {
  try {
    const cls = parseClass(req.body.class);
    if (!cls) throw new Error('pick your class (8, 9 or 10)');

    const name = clean(req.body.name, 100);
    if (name.length < 2) throw new Error('enter your name');

    const email = clean(req.body.email, 150).toLowerCase();
    if (!isEmail(email)) throw new Error('enter a valid email address');

    const mob = mobile10(req.body.mobile);
    if (!mob) throw new Error('enter a valid 10-digit contact number');

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
      const { rows } = await pool.query('SELECT roll FROM individuals WHERE email=$1', [clean(req.body.email, 150).toLowerCase()]);
      return res.status(409).json({ error: `that email is already registered${rows.length ? ` - your roll number is ${rows[0].roll}` : ''}` });
    }
    console.error('individual failed:', e.message);
    res.status(400).json({ error: e.message });
  }
});

/* -- school registration --------------------------------------------------- */

app.post('/api/schools', async (req, res) => {
  try {
    const district = clean(req.body.district, 2);
    if (!DISTRICTS[district]) throw new Error('pick a district');

    const name = clean(req.body.name);
    if (name.length < 3) throw new Error('enter the school name');

    const pocName = clean(req.body.pocName, 100);
    if (pocName.length < 2) throw new Error('enter the contact person name');

    const poc = mobile10(req.body.pocMobile);
    if (!poc) throw new Error('enter a valid 10-digit contact number');

    const email = clean(req.body.email, 150).toLowerCase();
    if (!isEmail(email)) throw new Error('enter a valid email address');

    const nameKey = normalize(name);

    // Already registered? Hand back the existing code - the contact person has
    // almost certainly just lost it.
    const existing = await pool.query('SELECT code, name FROM schools WHERE district=$1 AND name_key=$2', [district, nameKey]);
    if (existing.rows.length) {
      return res.json({ code: existing.rows[0].code, name: existing.rows[0].name, already: true });
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
    res.json({ code, name, district: DISTRICTS[district] });
  } catch (e) {
    if (e.constraint === 'schools_unique_per_district') {
      return res.status(409).json({ error: 'this school is already registered in that district' });
    }
    console.error('school failed:', e.message);
    res.status(400).json({ error: e.message });
  }
});

/** Look up a school by code, so a returning school can confirm it is the right one. */
app.get('/api/schools/:code', async (req, res) => {
  const code = clean(req.params.code, 10).toUpperCase();
  const { rows } = await pool.query(`
    SELECT code, name, district,
           (SELECT COUNT(*)::int FROM school_students s WHERE s.school_code = schools.code) AS students
      FROM schools WHERE code=$1`, [code]);
  if (!rows.length) return res.status(404).json({ error: 'no school with that code' });
  res.json({ ...rows[0], districtName: DISTRICTS[rows[0].district] });
});

/* -- school uploads its student sheet -------------------------------------- */

/** Blank sheet for schools to fill in, so the columns come back as expected. */
app.get('/api/template.xlsx', async (req, res) => {
  const wb = await buildTemplate();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="student-list-template.xlsx"');
  await wb.xlsx.write(res);
  res.end();
});

app.post('/api/schools/:code/upload', upload.single('file'), async (req, res) => {
  try {
    const schoolCode = clean(req.params.code, 10).toUpperCase();
    const school = await pool.query('SELECT code, district, name FROM schools WHERE code=$1', [schoolCode]);
    if (!school.rows.length) throw new Error('that school code does not exist');
    if (!req.file) throw new Error('attach the filled-in sheet');

    const { rows, columns } = await parseSheet(req.file.buffer, req.file.originalname);
    if (!rows.length) throw new Error('the sheet has column headings but no rows below them');

    const district = school.rows[0].district;
    const added = [];
    const skipped = [];
    const seen = new Set(); // the same student listed twice in one sheet

    for (const r of rows) {
      const name = clean(r.name, 100);
      const email = clean(r.email, 150).toLowerCase();
      const mob = mobile10(r.mobile);
      const cls = parseClass(r.class);

      if (name.length < 2) { skipped.push({ ...r, reason: 'no name' }); continue; }
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
          const { rows: ex } = await pool.query('SELECT roll FROM school_students WHERE email=$1', [email]);
          skipped.push({ ...r, reason: `already has a roll number${ex.length ? ` - ${ex[0].roll}` : ''}` });
        } else {
          skipped.push({ ...r, reason: e.message });
        }
      }
    }

    console.log(`upload for ${schoolCode}: ${added.length} added, ${skipped.length} skipped`);
    res.json({ schoolCode, schoolName: school.rows[0].name, columnsFound: columns, added, skipped });
  } catch (e) {
    console.error('upload failed:', e.message);
    res.status(400).json({ error: e.message });
  }
});

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

app.get('/api/stats', async (req, res) => {
  if (!checkAdmin(req, res)) return;
  const q = async (sql) => (await pool.query(sql)).rows[0].n;
  res.json({
    individuals: await q('SELECT COUNT(*)::int n FROM individuals'),
    schools: await q('SELECT COUNT(*)::int n FROM schools'),
    schoolStudents: await q('SELECT COUNT(*)::int n FROM school_students'),
  });
});

app.get('/api/export.xlsx', async (req, res) => {
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

setupWithRetry()
  .then(() => app.listen(PORT, () => console.log(`open http://localhost:${PORT}`)))
  .catch((e) => {
    console.error('could not start:', e.message);
    if (!DATABASE_URL) console.error('DATABASE_URL is not set - see README');
    process.exit(1);
  });

export default app;
