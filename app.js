import dotenv from 'dotenv';
import express from 'express';
import pg from 'pg';
import ExcelJS from 'exceljs';
import { DISTRICTS, rollProbe, schoolProbe, normalize, claim } from './codes.js';

// `vercel env pull` writes .env.local, so read that first and fall back to .env.
// Earlier entries win, which keeps hand-written values in .env from being lost.
dotenv.config({ path: ['.env.local', '.env'], quiet: true });

const { DATABASE_URL, ADMIN_KEY, PORT = 3002 } = process.env;

const pool = new pg.Pool({
  connectionString: DATABASE_URL,
  // Verify the server certificate properly — this connection carries names,
  // emails and phone numbers. Neon and RDS both present publicly trusted certs.
  // A local postgres normally has no TLS at all, hence the exception.
  ssl: /localhost|127\.0\.0\.1/.test(DATABASE_URL || '') ? false : true,
  max: 5,
});

async function setup() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS schools (
      code        TEXT PRIMARY KEY,
      district    TEXT NOT NULL,
      name        TEXT NOT NULL,
      name_key    TEXT NOT NULL,
      poc_name    TEXT NOT NULL,
      poc_mobile  TEXT NOT NULL,
      email       TEXT NOT NULL,
      created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
      -- the same school registering twice must get the same code back, not a second row
      CONSTRAINT schools_unique_per_district UNIQUE (district, name_key)
    );

    CREATE TABLE IF NOT EXISTS participants (
      roll        TEXT PRIMARY KEY,
      school_code TEXT NOT NULL REFERENCES schools(code),
      role        TEXT NOT NULL,
      name        TEXT NOT NULL,
      email       TEXT NOT NULL,
      mobile      TEXT NOT NULL,
      created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
      -- one registration per person
      CONSTRAINT participants_email_unique UNIQUE (email)
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

// Sent as a sorted array, not the object: JS orders integer-like keys first, so
// a plain object would put 10-14 above 01-09 in the dropdown.
app.get('/api/districts', (req, res) =>
  res.json(Object.entries(DISTRICTS).sort(([a], [b]) => a.localeCompare(b)).map(([code, name]) => ({ code, name }))));

/* ── school registration ─────────────────────────────────────────────────── */

app.post('/api/schools', async (req, res) => {
  try {
    const district = clean(req.body.district, 2);
    if (!DISTRICTS[district]) throw new Error('pick a district');

    const name = clean(req.body.name);
    if (name.length < 3) throw new Error('enter the school name');

    const pocName = clean(req.body.pocName, 100);
    if (pocName.length < 2) throw new Error('enter the contact person name');

    const poc = mobile10(req.body.pocMobile);
    if (!poc) throw new Error('enter a valid 10-digit mobile number');

    const email = clean(req.body.email, 150).toLowerCase();
    if (!isEmail(email)) throw new Error('enter a valid email address');

    const nameKey = normalize(name);

    // Already registered? Hand back the existing code instead of erroring — the
    // POC has almost certainly just lost the code.
    const existing = await pool.query('SELECT code, name FROM schools WHERE district=$1 AND name_key=$2', [district, nameKey]);
    if (existing.rows.length) {
      return res.json({ code: existing.rows[0].code, name: existing.rows[0].name, already: true });
    }

    const { code, attempts } = await claim(
      schoolProbe(district, name),
      (c) => pool.query(
        `INSERT INTO schools (code, district, name, name_key, poc_name, poc_mobile, email)
         VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING code`,
        [c, district, name, nameKey, pocName, poc, email],
      ),
      'schools_pkey',
    );

    console.log(`school ${code} — ${name} (${attempts} probe${attempts > 1 ? 's' : ''})`);
    res.json({ code, name, district: DISTRICTS[district] });
  } catch (e) {
    if (e.constraint === 'schools_unique_per_district') {
      return res.status(409).json({ error: 'this school is already registered in that district' });
    }
    console.error('school failed:', e.message);
    res.status(400).json({ error: e.message });
  }
});

/** Look up a school by code so the participant form can confirm the name. */
app.get('/api/schools/:code', async (req, res) => {
  const { rows } = await pool.query('SELECT code, name, district FROM schools WHERE code=$1', [clean(req.params.code, 10).toUpperCase()]);
  if (!rows.length) return res.status(404).json({ error: 'no school with that code' });
  res.json({ ...rows[0], districtName: DISTRICTS[rows[0].district] });
});

/* ── participant registration ────────────────────────────────────────────── */

app.post('/api/participants', async (req, res) => {
  try {
    const schoolCode = clean(req.body.schoolCode, 10).toUpperCase();
    const school = await pool.query('SELECT code, district, name FROM schools WHERE code=$1', [schoolCode]);
    if (!school.rows.length) throw new Error('that school code does not exist — ask your school for it');

    const role = clean(req.body.role, 10).toLowerCase();
    if (role !== 'student' && role !== 'teacher') throw new Error('pick student or teacher');

    const name = clean(req.body.name, 100);
    if (name.length < 2) throw new Error('enter your name');

    const email = clean(req.body.email, 150).toLowerCase();
    if (!isEmail(email)) throw new Error('enter a valid email address');

    const mob = mobile10(req.body.mobile);
    if (!mob) throw new Error('enter a valid 10-digit mobile number');

    const district = school.rows[0].district;
    const { code: roll, attempts } = await claim(
      rollProbe(district, email),
      (c) => pool.query(
        `INSERT INTO participants (roll, school_code, role, name, email, mobile)
         VALUES ($1,$2,$3,$4,$5,$6) RETURNING roll`,
        [c, schoolCode, role, name, email, mob],
      ),
      'participants_pkey',
    );

    console.log(`${role} ${roll} — ${name} (${attempts} probe${attempts > 1 ? 's' : ''})`);
    res.json({ roll, name, school: school.rows[0].name, district: DISTRICTS[district] });
  } catch (e) {
    if (e.constraint === 'participants_email_unique') {
      const { rows } = await pool.query('SELECT roll FROM participants WHERE email=$1', [clean(req.body.email, 150).toLowerCase()]);
      return res.status(409).json({ error: `that email is already registered${rows.length ? ` — roll number ${rows[0].roll}` : ''}` });
    }
    console.error('participant failed:', e.message);
    res.status(400).json({ error: e.message });
  }
});

/* ── admin: counts + Excel download ──────────────────────────────────────── */

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
  const [s, p] = await Promise.all([
    pool.query('SELECT COUNT(*)::int n FROM schools'),
    pool.query(`SELECT role, COUNT(*)::int n FROM participants GROUP BY role`),
  ]);
  const byRole = Object.fromEntries(p.rows.map((r) => [r.role, r.n]));
  res.json({ schools: s.rows[0].n, students: byRole.student ?? 0, teachers: byRole.teacher ?? 0 });
});

app.get('/api/export.xlsx', async (req, res) => {
  if (!checkAdmin(req, res)) return;

  const participants = await pool.query(`
    SELECT p.roll, p.name, p.email, p.mobile, p.role, p.school_code, s.name AS school_name, s.district
      FROM participants p JOIN schools s ON s.code = p.school_code
     ORDER BY s.district, p.school_code, p.roll`);

  const schools = await pool.query(`
    SELECT code, name, poc_name, poc_mobile, email, district,
           (SELECT COUNT(*)::int FROM participants p WHERE p.school_code = schools.code) AS registered
      FROM schools ORDER BY district, code`);

  const wb = new ExcelJS.Workbook();

  const sheet = (name, columns, rows) => {
    const ws = wb.addWorksheet(name);
    ws.columns = columns;
    ws.addRows(rows);
    ws.getRow(1).font = { bold: true };
    ws.views = [{ state: 'frozen', ySplit: 1 }]; // header stays visible while scrolling
    ws.autoFilter = { from: 'A1', to: { row: 1, column: columns.length } };
    return ws;
  };

  sheet('Participants', [
    { header: 'Roll Number', key: 'roll', width: 14 },
    { header: 'Name', key: 'name', width: 26 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Mobile Number', key: 'mobile', width: 16 },
    { header: 'Role', key: 'role', width: 10 },
    { header: 'School Code', key: 'school_code', width: 14 },
    { header: 'School Name', key: 'school_name', width: 32 },
    { header: 'District', key: 'district_name', width: 20 },
  ], participants.rows.map((r) => ({ ...r, district_name: DISTRICTS[r.district] })));

  sheet('Schools', [
    { header: 'School Code', key: 'code', width: 14 },
    { header: 'School Name', key: 'name', width: 32 },
    { header: 'POC Name', key: 'poc_name', width: 22 },
    { header: 'POC Number', key: 'poc_mobile', width: 16 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'District', key: 'district_name', width: 20 },
    { header: 'Registered', key: 'registered', width: 12 },
  ], schools.rows.map((r) => ({ ...r, district_name: DISTRICTS[r.district] })));

  const stamp = new Date().toISOString().slice(0, 10);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="registrations-${stamp}.xlsx"`);
  await wb.xlsx.write(res);
  res.end();
  console.log(`exported ${participants.rowCount} participants, ${schools.rowCount} schools`);
});

/**
 * Neon's free tier suspends the database when idle, so the first connection
 * after a quiet spell can be refused outright — sometimes with an empty error
 * message. Retry a few times rather than crashing on a sleeping database.
 */
async function setupWithRetry(tries = 5) {
  for (let i = 1; i <= tries; i++) {
    try {
      await setup();
      return;
    } catch (e) {
      const why = e.message || e.code || 'connection refused';
      if (i === tries) throw new Error(why);
      const wait = 2000 * i;
      console.log(`database not ready (${why}) — retrying in ${wait / 1000}s [${i}/${tries - 1}]`);
      await new Promise((r) => setTimeout(r, wait));
    }
  }
}

setupWithRetry()
  .then(() => app.listen(PORT, () => console.log(`open http://localhost:${PORT}`)))
  .catch((e) => {
    console.error('could not start:', e.message);
    if (!DATABASE_URL) console.error('DATABASE_URL is not set — see README');
    process.exit(1);
  });

export default app;
