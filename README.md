# Registration Portal

Schools register and get a **School Code**. Students and teachers register with
that code and get a **Roll Number**. You download everything as an Excel file
with two sheets.

## Run it

```bash
npm install
cp .env.example .env      # put your DATABASE_URL and ADMIN_KEY in
npm start
```

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
vercel integration add neon      # provisions Postgres, sets DATABASE_URL
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
