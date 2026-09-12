# Registration Portal

Two completely separate registration paths.

**Individual** — `/` — a person enters name, email, contact number and class, and
gets a roll number. No school involved. This is the set that will feed the
payment gateway.

**School** — `/school.html` — a teacher enters district, school name, contact
person, contact number and email, gets a **school code**, then uploads an Excel
sheet of students. Every valid row gets a roll number.

**Admin** — `/admin.html` — counts, and an Excel export with three sheets.

## Run it

```bash
npm install
cp .env.example .env      # put your ADMIN_KEY in
npm start                 # http://localhost:3002
```

Both `.env` and `.env.local` are read, `.env.local` winning, so
`vercel env pull` can supply the database credentials without disturbing
anything typed by hand.

## Roll numbers

Every code is a prefix plus three letters. The letters come from hashing the
person, with **linear probing** when a slot is taken:

```
hash(email) % 17576  ->  slot 9839  ->  "NGL"
                         taken?     ->  9840 -> "NGM"   (probe +1)
                         taken?     ->  9841 -> "NGN"   (probe +2)
```

| | Prefix | Example | Reads as |
|---|---|---|---|
| School code | district | `07KBO` | Ernakulam |
| School student | district + class | `0710KEL` | Ernakulam, class 10 |
| Individual | `IN` + class | `IN10KEL` | class 10, no district |

Each prefix is its own slot space (17,576 codes), so a school student and an
individual can never be issued the same code. Individuals start with a letter,
so the two are distinguishable at a glance.

Two properties worth knowing:

- **The insert is the claim.** A code is never "checked then taken" — the
  database's unique index decides and a rejection means probe on. Two people
  registering in the same millisecond cannot be given the same code.
- **Re-registering is safe.** A school that registers twice gets its original
  code back. A person who registers twice is told their existing roll number.
- **Identity is the person, not the mailbox.** Siblings on a parent's address
  both register; only the same name on the same address is a duplicate.
- **Codes are allocated without failed inserts.** `ON CONFLICT (roll) DO NOTHING`
  in batches, because node-postgres destroys a client that threw and the next
  query then pays a ~2.6s reconnect. Re-uploading a 500-student sheet went from
  about 46 minutes to 1.5 seconds.

## The sheet a school uploads

Template is on the page (`/api/template.xlsx`): **Name, Email, Contact Number,
Class**. Class must be 8, 9 or 10 and is required for every student, because the
roll number is built from it.

**A wrong template still works.** The headings are a hint, not a requirement —
what a column holds decides what it is:

- **No heading row at all?** Fine. An email looks like an email and a 10-digit
  number looks like a phone number, so the columns are worked out from the values.
- **Headings we have never seen** (`Pupil Identifier`, `Reach At`, `Which Std`,
  or headings in Malayalam)? Read from the values instead.
- **Headings that lie?** A column headed `Email` that holds phone numbers is
  overruled by the column that actually holds the emails, and the swap is
  reported back.
- **Still wrong?** The page shows which column it read as what, plus the first
  six rows, *before anything is saved*. If the guess is wrong there are
  dropdowns to set the columns by hand and re-read.

Nothing is written until the teacher looks at that preview and agrees.

On top of that:

- **Columns matched by meaning, in any order** — `Student Name` / `Full Name`;
  `Phone Number` / `Contact No` / `WhatsApp`; `E-Mail ID` / `Mail`;
  `Class` / `Std` / `Standard` / `Grade`.
- **Class in any form** — `10`, `Class 10`, `10th`, `X`, `VIII` all work.
- **A title block above the headings is fine** — the header row is searched for
  in the first 20 rows.
- **Phone numbers survive Excel** — cells that became numbers, `98765 00102`
  and `+91 9876500103` all normalize to 10 digits.
- **Blank spacer rows skipped**, emails lowercased.
- **One bad row does not reject the file.** Valid rows are registered; the rest
  come back with their real spreadsheet row number and the reason.
- **Re-uploading a corrected sheet is safe** — anyone who already has a roll
  number is reported, not given a second one.

CSV works too.

## The Excel export

**Individuals** — Roll Number · Name · Email · Contact Number · Class · Registered At

**School Students** — Roll Number · Name · Email · Contact Number · Class ·
School Code · School Name · District · Registered At

**Schools** — School Code · School Name · Contact Person · Contact Number ·
Email · District · Students Uploaded

Phone columns are formatted as text so Excel does not turn them into
`9.87654E+09`.

## Deploying to Vercel

```bash
npx vercel link
npx vercel integration add neon   # provisions Postgres, writes .env.local
npx vercel env add ADMIN_KEY
npx vercel deploy --prod
```

Tables are created automatically on first boot. Deployment Protection must be
**off** in project settings, or visitors hit a Vercel login page.

## Moving to AWS later

Plain Express and plain Postgres. Delete `vercel.json`, point `DATABASE_URL` at
RDS, run it as a container. No code changes.

## The files

| | |
|---|---|
| `app.js` | server — routes, validation, Excel export |
| `codes.js` | the code generator and the probing logic |
| `sheet.js` | reading the uploaded sheet, building the template |
| `public/index.html` | individual registration |
| `public/school.html` | school registration + upload |
| `public/admin.html` | counts + export |
