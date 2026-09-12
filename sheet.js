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

export async function parseSheet(buffer, filename = '') {
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

  if (!picked) {
    throw new Error(
      'could not find the column headings. The sheet needs a Name column plus Email, '
      + `Contact Number and Class - within the first ${HEADER_SEARCH_ROWS} rows. `
      + 'Download the template and fill that in.',
    );
  }

  const { ws, found } = picked;
  const { headerRow, cols, fields } = found;

  // Say this once, up front, instead of repeating "no contact number" against
  // every single row of an otherwise perfect sheet.
  const LABEL = { email: 'Email', mobile: 'Contact Number', class: 'Class' };
  const missing = ['email', 'mobile', 'class'].filter((f) => !fields.includes(f));
  if (missing.length) {
    throw new Error(`the sheet has no ${missing.map((m) => LABEL[m]).join(' or ')} column, and every student needs one. Download the template and fill that in.`);
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
  };
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
