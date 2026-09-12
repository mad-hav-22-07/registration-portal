import ExcelJS from 'exceljs';

/**
 * Reading the student sheet a school sends you.
 *
 * Real sheets are messy: columns in any order, headers worded differently,
 * a title row above the headers, blank rows in the middle, phone numbers that
 * Excel has helpfully turned into numbers and stripped the leading zero from.
 * So match columns by meaning rather than by position, and report problems per
 * row instead of rejecting the whole file.
 */

const ALIASES = {
  name: ['name', 'fullname', 'studentname', 'participantname', 'teachername', 'nameofstudent', 'nameoftheparticipant'],
  email: ['email', 'emailid', 'emailaddress', 'mail', 'mailid', 'gmail'],
  mobile: ['mobile', 'mobilenumber', 'mobileno', 'phone', 'phonenumber', 'phoneno', 'contact', 'contactnumber', 'contactno', 'whatsapp', 'whatsappnumber'],
  class: ['class', 'std', 'standard', 'grade', 'classstd', 'studyingin', 'classgrade'],
};

const squash = (v) => String(v ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');

/** A cell can hold a string, a number, a date, or a {text, hyperlink} object. */
function cellText(cell) {
  const v = cell?.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v.text !== undefined) return String(v.text); // hyperlink / rich text
    if (v.result !== undefined) return String(v.result); // formula
    if (v instanceof Date) return v.toISOString();
    return '';
  }
  return String(v);
}

/** Find the row that holds the headers, and which column each field is in. */
function findHeader(ws) {
  const limit = Math.min(ws.rowCount, 20); // a title block above the headers is common
  for (let r = 1; r <= limit; r++) {
    const row = ws.getRow(r);
    const map = {};
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      const key = squash(cellText(cell));
      if (!key) return;
      for (const [field, names] of Object.entries(ALIASES)) {
        if (map[field] === undefined && names.includes(key)) map[field] = col;
      }
    });
    // name + one way to contact them is the minimum worth reading; a missing
    // class column is reported per row rather than rejecting the whole file
    if (map.name !== undefined && (map.email !== undefined || map.mobile !== undefined)) {
      return { headerRow: r, cols: map };
    }
  }
  return null;
}

export async function parseSheet(buffer, filename = '') {
  const wb = new ExcelJS.Workbook();

  if (/\.csv$/i.test(filename)) {
    const { Readable } = await import('node:stream');
    await wb.csv.read(Readable.from(buffer));
  } else {
    await wb.xlsx.load(buffer);
  }

  const ws = wb.worksheets[0];
  if (!ws || ws.rowCount === 0) throw new Error('the file has no sheets or is empty');

  const found = findHeader(ws);
  if (!found) {
    throw new Error('could not find the column headings. The sheet needs Name, Email, Contact Number and Class columns. Download the template and fill that in.');
  }

  const { headerRow, cols } = found;
  const rows = [];

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const get = (field) => (cols[field] ? cellText(row.getCell(cols[field])).trim() : '');

    const name = get('name');
    const email = get('email');
    const mobile = get('mobile');
    const cls = get('class');

    if (!name && !email && !mobile) continue; // blank spacer row

    rows.push({
      row: r, // the real spreadsheet row number, so errors are findable
      name,
      email: email.toLowerCase(),
      mobile,
      class: cls,
    });
  }

  return { rows, headerRow, columns: Object.keys(cols), sheetName: ws.name };
}

/** The blank sheet schools should fill in. */
export async function buildTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Participants');

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
  note.value = 'Replace the example rows with your students. Class must be 8, 9 or 10 and is required for every student. Do not rename the headings.';
  note.font = { italic: true, size: 10 };
  note.alignment = { wrapText: true };
  ws.getColumn('F').width = 60;

  return wb;
}
