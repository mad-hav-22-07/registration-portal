import ExcelJS from 'exceljs';

/**
 * Reading the student sheet a school sends you.
 *
 * Real sheets are messy: columns in any order, headings worded differently, a
 * title block above them, the students on the second tab, blank rows in the
 * middle, phone numbers Excel has turned into numbers, and cells a teacher
 * bolded halfway through. So match columns by meaning, search every tab, and
 * report problems per row rather than rejecting the file.
 *
 * The invariant that matters most: every data row must come back either added
 * or skipped. A student silently vanishing is far worse than a rejected file.
 */

const ALIASES = {
  name: [
    'name', 'fullname', 'studentname', 'participantname', 'nameofstudent',
    'nameofthestudent', 'nameoftheparticipant', 'studentsname', 'candidatename',
  ],
  email: ['email', 'emailid', 'emailaddress', 'mail', 'mailid', 'gmail', 'emailidofstudent'],
  mobile: [
    'mobile', 'mobilenumber', 'mobileno', 'mobno', 'mob', 'phone', 'phonenumber',
    'phoneno', 'phno', 'phnumber', 'ph', 'contact', 'contactnumber', 'contactno',
    'contactnumberofstudent', 'whatsapp', 'whatsappnumber', 'whatsappno', 'cell', 'cellnumber',
  ],
  class: ['class', 'std', 'standard', 'grade', 'classstd', 'studyingin', 'classgrade', 'classstandard'],
};

/** How many rows from the top to search for the heading row. */
const HEADER_SEARCH_ROWS = 30;

const squash = (v) => String(v ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');

/**
 * Read a cell as the text a human sees.
 *
 * ExcelJS's `.text` getter already resolves rich text, formula results,
 * hyperlinks and dates. Reading `.value` by hand missed `{richText:[...]}` —
 * which is what any partially bolded or pasted-from-Word cell becomes — so the
 * cell read as empty, the row was dropped as blank, and the student disappeared
 * without even appearing in the skipped list.
 */
function cellText(cell) {
  if (!cell) return '';
  try {
    const t = cell.text;
    if (t !== undefined && t !== null && t !== '') return String(t);
  } catch {
    /* fall through to .value */
  }
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (typeof v !== 'object') return String(v);
  if (v.result !== undefined && v.result !== null) return String(v.result);
  if (Array.isArray(v.richText)) return v.richText.map((p) => p.text ?? '').join('');
  if (v.text !== undefined) return String(v.text);
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return '';
}

/* -- working out what a column holds by looking at it ---------------------- */

const looksEmail = (v) => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v);

const looksMobile = (v) => {
  if (!/^[+\d][\d\s().+-]*$/.test(v)) return false;
  let d = v.replace(/\D/g, '').replace(/^0+/, '');
  if (d.length === 12 && d.startsWith('91')) d = d.slice(2);
  return /^[6-9]\d{9}$/.test(d);
};

const looksClass = (v) => {
  const digits = v.replace(/[^0-9]/g, '');
  if (digits) return ['8', '9', '10'].includes(String(Number(digits)));
  return /^(CLASS|STD|STANDARD|GRADE)?\s*(VIII|IX|X)(TH|ST|ND|RD)?[-\s]?[A-H]?$/i.test(v.trim());
};

// A name is whatever is left: has a letter, and is not one of the above.
const looksName = (v) => /\p{L}/u.test(v) && !looksEmail(v) && !looksMobile(v) && !looksClass(v);

/**
 * Work out which column is which from the VALUES, ignoring the headings.
 *
 * This is what makes a wrong template still work. A school that renames the
 * columns, reorders them, writes them in Malayalam, or sends a sheet with no
 * heading row at all still gets parsed, because an email looks like an email and
 * a 10-digit number looks like a phone number whatever the column is called.
 */
export function inferColumns(ws, firstDataRow) {
  const lastRow = Math.min(ws.rowCount, firstDataRow + 40);
  const lastCol = Math.min(ws.columnCount || 20, 30);
  const tests = { email: looksEmail, mobile: looksMobile, class: looksClass, name: looksName };
  const scores = {};

  for (let col = 1; col <= lastCol; col++) {
    const values = [];
    for (let r = firstDataRow; r <= lastRow; r++) {
      const t = cellText(ws.getRow(r).getCell(col)).trim();
      if (t) values.push(t);
    }
    if (!values.length) continue;
    scores[col] = {};
    for (const [field, test] of Object.entries(tests)) {
      scores[col][field] = values.filter(test).length / values.length;
    }
  }

  // Assign most-distinctive first, one column per field. A phone number also
  // "looks like" a class if it is short, so order matters.
  const cols = {};
  const confidence = {};
  const taken = new Set();

  for (const field of ['email', 'mobile', 'class', 'name']) {
    let bestCol = null;
    let bestScore = 0;
    for (const [col, sc] of Object.entries(scores)) {
      if (taken.has(col)) continue;
      if (sc[field] > bestScore) { bestScore = sc[field]; bestCol = col; }
    }
    // Over half the cells have to agree, or we are guessing.
    if (bestCol && bestScore > 0.5) {
      cols[field] = Number(bestCol);
      confidence[field] = Math.round(bestScore * 100);
      taken.add(bestCol);
    }
  }
  return { cols, confidence };
}

/** Spreadsheet column letter, for telling the user what we picked. */
export function colLetter(n) {
  let out = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** Non-empty cells below a column, to tell a real column from an empty duplicate. */
function filledBelow(ws, col, fromRow) {
  let n = 0;
  const last = Math.min(ws.rowCount, fromRow + 50);
  for (let r = fromRow; r <= last; r++) {
    if (cellText(ws.getRow(r).getCell(col)).trim()) n++;
  }
  return n;
}

/**
 * Find the heading row in one worksheet.
 *
 * Scores every candidate row by how many distinct fields it matches and keeps
 * the best, rather than taking the first row that happens to mention a name and
 * an email — a school info line like "Name | GHSS ALUVA | Email | principal@..."
 * was being picked over the real headings three rows below it.
 */
function findHeader(ws) {
  const limit = Math.min(ws.rowCount, HEADER_SEARCH_ROWS);
  let best = null;

  for (let r = 1; r <= limit; r++) {
    const row = ws.getRow(r);
    const candidates = {};

    row.eachCell({ includeEmpty: false }, (cell, col) => {
      const key = squash(cellText(cell));
      if (!key) return;
      for (const [field, names] of Object.entries(ALIASES)) {
        if (names.includes(key)) (candidates[field] ??= []).push(col);
      }
    });

    const fields = Object.keys(candidates);
    if (!fields.includes('name')) continue;

    // When a heading appears twice (someone pasted a second table alongside the
    // first), take the column that actually has data under it.
    const cols = {};
    for (const [field, list] of Object.entries(candidates)) {
      cols[field] = list.length === 1
        ? list[0]
        : list.reduce((a, b) => (filledBelow(ws, b, r + 1) > filledBelow(ws, a, r + 1) ? b : a));
    }

    if (!best || fields.length > best.score) best = { headerRow: r, cols, score: fields.length, fields };
  }

  return best;
}

/**
 * @param override optional {name,email,mobile,class} of 1-based column numbers,
 *        set by the teacher when the automatic guess got it wrong.
 */
export async function parseSheet(buffer, filename = '', override = null) {
  const wb = new ExcelJS.Workbook();

  if (/\.csv$/i.test(filename)) {
    const { Readable } = await import('node:stream');
    const head = buffer.subarray(0, 4096).toString('utf8').split(/\r?\n/)[0] ?? '';
    // Excel writes semicolon-separated CSV in many locales, and tab-separated
    // when saving as "Unicode text".
    const tally = [[',', (head.match(/,/g) || []).length], [';', (head.match(/;/g) || []).length], ['\t', (head.match(/\t/g) || []).length]]
      .sort((a, b) => b[1] - a[1]);
    await wb.csv.read(Readable.from(buffer), { parserOptions: { delimiter: tally[0][1] > 0 ? tally[0][0] : ',' } });
  } else {
    // .xlsx is a zip. A legacy .xls, a renamed file or a corrupt upload makes
    // exceljs throw a JSZip message with a documentation link, which means
    // nothing to a teacher - say what to actually do instead.
    const isZip = buffer.length > 1 && buffer[0] === 0x50 && buffer[1] === 0x4b;
    if (!isZip) {
      const ole = buffer.length > 7 && buffer[0] === 0xd0 && buffer[1] === 0xcf;
      throw new Error(ole
        ? 'That looks like an old .xls file, or a password-protected workbook. Open it in Excel, remove any password, then File > Save As > Excel Workbook (.xlsx) and upload again.'
        : 'That file is not a readable .xlsx or .csv. Download the template and fill that in.');
    }
    await wb.xlsx.load(buffer);
  }

  if (!wb.worksheets.length) throw new Error('the file has no sheets');

  // Search every tab, not just the first: an "Instructions" or "Summary" tab in
  // front of the student list used to make the whole file unreadable.
  let picked = null;
  for (const ws of wb.worksheets) {
    if (ws.rowCount === 0) continue;
    const found = findHeader(ws);
    if (found && (!picked || found.score > picked.found.score)) picked = { ws, found };
  }

  // No recognisable headings anywhere? Fall back to reading the values. A sheet
  // with no heading row at all, or headings we have never seen, is still usable.
  let ws;
  let headerRow;
  let cols;
  let how;

  if (picked) {
    ws = picked.ws;
    headerRow = picked.found.headerRow;
    cols = { ...picked.found.cols };
    how = 'headings';
  } else {
    ws = wb.worksheets.reduce((a, b) => (b.rowCount > (a?.rowCount ?? 0) ? b : a), null);
    if (!ws || ws.rowCount === 0) throw new Error('the file has no rows in it');
    // There may still be a heading row, just one worded in a way we do not
    // recognise (or in Malayalam). A heading row contains no email and no phone
    // number, so if row 1 has neither and a later row does, skip row 1 — it
    // would otherwise be reported as a student called "Full Name Of Pupil".
    headerRow = !rowHasRealValues(ws, 1) && rowHasRealValues(ws, 2) ? 1 : 0;
    cols = {};
    how = 'contents';
  }

  // Check every mapping against the actual values, and fill in or correct it.
  // A column headed "Email" that holds phone numbers is trusted less than the
  // column that actually holds the emails.
  const inferred = inferColumns(ws, headerRow + 1);
  const corrected = [];

  for (const field of ['name', 'email', 'mobile', 'class']) {
    const guess = inferred.cols[field];
    if (cols[field] === undefined) {
      if (guess) { cols[field] = guess; corrected.push(`${field} read from column ${colLetter(guess)}`); }
      continue;
    }
    // The heading claims this field — but does the column actually hold it?
    if (guess && guess !== cols[field] && columnScore(ws, cols[field], headerRow + 1, field) < 0.5) {
      cols[field] = guess;
      corrected.push(`${field} taken from column ${colLetter(guess)}, which is where the ${field} values actually are`);
    }
  }

  // An explicit choice from the teacher always wins.
  if (override) {
    for (const field of ['name', 'email', 'mobile', 'class']) {
      const n = Number(override[field]);
      if (Number.isInteger(n) && n > 0) cols[field] = n;
    }
    how = 'your column choices';
  }

  const LABEL = { name: 'Name', email: 'Email', mobile: 'Contact Number', class: 'Class' };
  const missing = ['name', 'email', 'mobile', 'class'].filter((f) => cols[f] === undefined);
  if (missing.length) {
    throw new Error(
      `could not work out which column holds the ${missing.map((m) => LABEL[m]).join(' and ')}. `
      + 'Check the sheet has one column each for Name, Email, Contact Number and Class, '
      + 'or set the columns yourself below.',
    );
  }

  const rows = [];
  let blank = 0;

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const get = (field) => (cols[field] ? cellText(row.getCell(cols[field])).trim() : '');

    const name = get('name');
    const email = get('email');
    const mobile = get('mobile');
    const cls = get('class');

    if (!name && !email && !mobile && !cls) {
      blank++;
      continue; // genuinely empty spacer row
    }

    rows.push({
      row: r, // the real spreadsheet row number, so a teacher can find it
      name,
      email: email.toLowerCase(),
      mobile,
      class: cls,
    });
  }

  return {
    rows,
    headerRow,
    columns: Object.keys(cols),
    sheetName: ws.name,
    blankRows: blank,
    // what was used, so the teacher can see it and correct it if wrong
    mapping: Object.fromEntries(Object.entries(cols).map(([f, c]) => [f, { column: c, letter: colLetter(c) }])),
    detectedBy: how,
    corrections: corrected,
    confidence: inferred.confidence,
    sheetNames: wb.worksheets.map((w) => w.name),
  };
}

/** Does this row contain an actual email or phone number, i.e. is it data? */
function rowHasRealValues(ws, r) {
  if (r > ws.rowCount) return false;
  let found = false;
  ws.getRow(r).eachCell({ includeEmpty: false }, (cell) => {
    const t = cellText(cell).trim();
    if (looksEmail(t) || looksMobile(t)) found = true;
  });
  return found;
}

/** How well one column's values match one field. */
function columnScore(ws, col, firstDataRow, field) {
  const tests = { email: looksEmail, mobile: looksMobile, class: looksClass, name: looksName };
  const last = Math.min(ws.rowCount, firstDataRow + 40);
  const values = [];
  for (let r = firstDataRow; r <= last; r++) {
    const t = cellText(ws.getRow(r).getCell(col)).trim();
    if (t) values.push(t);
  }
  if (!values.length) return 0;
  return values.filter(tests[field]).length / values.length;
}

/** The blank sheet schools should fill in. */
export async function buildTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Students');

  ws.columns = [
    { header: 'Name', key: 'name', width: 28 },
    { header: 'Email', key: 'email', width: 32 },
    { header: 'Contact Number', key: 'mobile', width: 18 },
    { header: 'Class', key: 'class', width: 10 },
  ];
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  ws.addRows([
    { name: 'Arjun Nair', email: 'arjun@example.com', mobile: '9876543210', class: 10 },
    { name: 'Deepa Raj', email: 'deepa@example.com', mobile: '9876543211', class: 9 },
    { name: 'Fathima S', email: 'fathima@example.com', mobile: '9876543212', class: 8 },
  ]);

  // Mobile numbers must stay text, or Excel eats a leading zero and turns long
  // numbers into 9.87654E+09.
  ws.getColumn('mobile').numFmt = '@';

  const note = ws.getCell('F2');
  note.value = 'Replace the example rows with your students. Every student needs a name, '
    + 'email, 10-digit contact number and a class of 8, 9 or 10. Do not rename the headings.';
  note.font = { italic: true, size: 10 };
  note.alignment = { wrapText: true };
  ws.getColumn('F').width = 58;

  return wb;
}
