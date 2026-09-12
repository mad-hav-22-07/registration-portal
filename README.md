# Registration Portal

Two ways in:

- **A school registers** → gets a **School Code** → uploads an Excel sheet of
  all its students and teachers → every row gets a **Roll Number**
- **An individual registers** → enters their school's code → gets a Roll Number

You download everything as an Excel file with two sheets.

## Run it

```bash
npm install
cp .env.example .env      # put your ADMIN_KEY in
npm start
```

Both `.env` and `.env.local` are read, with `.env.local` winning — so
`vercel env pull` can drop the database credentials in without disturbing
anything you typed by hand.

- `http://localhost:3002` — registration (two tabs: Student/Teacher, School)
- `http://localhost:3002/admin.html` — counts and the Excel download

## The code generator

District number (`01`–`14`) followed by three letters. The letters come from
hashing, with **linear probing** when a slot is already taken:

```
hash(email) % 17576  ->  slot 9839  ->  "NGL"   ->  roll = 07NGL
                         slot taken?  ->  9840  ->  "NGM"   (probe +1)
                         still taken? ->  9841  ->  "NGN"   (probe +2)
```

Three letters gives **17,576 codes per district** for participants and the same
for schools. Two letters would have given only 676, and some Kerala districts
have more than 900 schools — a district could have run out.

District numbers are the usual south-to-north order: `01` Thiruvananthapuram,
`02` Kollam, `03` Pathanamthitta, `04` Alappuzha, `05` Kottayam, `06` Idukki,
`07` Ernakulam, `08` Thrissur, `09` Palakkad, `10` Malappuram, `11` Kozhikode,
`12` Wayanad, `13` Kannur, `14` Kasaragod.

Two properties worth knowing:

- **The insert is the claim.** The code is never "checked then taken" — the
  database's unique index decides, and a rejection means probe on. Two people
  registering in the same millisecond therefore cannot be given the same code.
- **Re-registering is safe.** A school that registers twice gets its original
  code back, not a second row, because the school name is normalized first
  (`St. Thomas H.S.S` and `st thomas hss` hash to the same slot). A participant
  who registers twice is told their existing roll number.

## The sheet a school uploads

They download the template from the portal (`/api/template.xlsx`): **Name,
Email, Mobile Number, Role**. Role is `student` or `teacher`, blank means
student.

Real sheets never look like the template, so the parser is deliberately loose:

- **Columns are matched by meaning, in any order.** `Student Name`, `Full Name`,
  `Name of Participant` all map to name; `Phone Number`, `Contact No`,
  `WhatsApp` all map to mobile; `E-Mail ID`, `Mail` to email; `Category`,
  `Type`, `Designation` to role.
- **A title block above the headers is fine** — the header row is searched for
  in the first 20 rows.
- **Phone numbers survive Excel's meddling** — cells that became numbers,
  `98765 00002`, and `+91 9876500003` all normalize to 10 digits.
- **Blank spacer rows are skipped**, emails are lowercased.
- **One bad row does not reject the file.** Valid rows are registered; the rest
  come back with the real spreadsheet row number and the reason, so they can be
  fixed and the sheet uploaded again.
- **Re-uploading is safe.** Anyone who already has a roll number is reported as
  already registered rather than issued a second one.

CSV works too.

## The Excel file

**Participants** — Roll Number · Name · Email · Mobile Number · Role ·
School Code · School Name · District

**Schools** — School Code · School Name · POC Name · POC Number · Email ·
District · Registered

School Code and District are on the participants sheet too, so the two sheets
can be joined and filtered without a lookup.

## Deploying to Vercel

```bash
npm i -g vercel
vercel login
vercel link
vercel integration add neon      # provisions Postgres, writes .env.local
vercel env add ADMIN_KEY         # paste your admin key
vercel deploy --prod
```

Tables are created automatically on first boot.

## Moving to AWS later

Nothing here is Vercel-specific except `vercel.json`. It is plain Express and
plain Postgres, so on AWS it is a container (ECS/App Runner/EC2) with
`DATABASE_URL` pointed at RDS. No code changes.

## The files

| | |
|---|---|
| `app.js` | server — registration, validation, Excel export |
| `codes.js` | the code generator and the probing logic |
| `public/index.html` | registration page |
| `public/admin.html` | counts + download |
