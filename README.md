# Personal Statement Splitter

A Cloudflare Pages + D1 app for uploading credit card statements, parsing transactions, assigning/splitting charges by person, and tracking person-level payments or Splitwise transfers.

## What it supports

- American Express Excel uploads
  - Ignores rows 1-6
  - Uses row 7 as headers
  - Pulls Date, Description, Card Member, Amount
- Chase copy-pastable PDF uploads
  - Ignores pages 1-2
  - Starts on page 3
  - Pulls transaction date, description, amount from the PURCHASE section
  - Leaves Person blank because Chase does not give cardholder details
- Capital One copy-pastable PDF uploads
  - Ignores pages 1-2
  - Starts on page 3
  - Pulls transactions by cardholder section
  - Pulls credits/refunds as negative amounts
  - Skips CAPITAL ONE ONLINE PYMT
- Merchant cleanup rules, for example `MTA*NYCT PAYGO` -> `Subway`
- Editable review table before saving
- Even and custom splits, with total/difference checker
- Split transactions are saved as multiple line items
- Person dashboard with assigned totals, payments, Splitwise transfers, adjustments, and open balances
- Separate user logins with data isolated by `user_id`
- CSV export from dashboard summary

## Cloudflare setup

### 1. Create the D1 database

```bash
npx wrangler d1 create personal_statement_splitter
```

Copy the returned `database_id` into `wrangler.toml`.

### 2. Run the migration

```bash
npx wrangler d1 migrations apply personal_statement_splitter --remote
```

For local development:

```bash
npx wrangler d1 migrations apply personal_statement_splitter --local
```

### 3. Add a session secret

In Cloudflare Pages settings, add an environment variable:

```text
SESSION_SECRET = put-a-long-random-string-here
```

### 4. Cloudflare Pages build settings

Framework preset: `Vite`

Build command:

```bash
npm run build
```

Output directory:

```bash
dist
```

### 5. Bind D1 to Pages

In Cloudflare Pages → Settings → Functions → D1 database bindings:

```text
Variable name: DB
Database: personal_statement_splitter
```

### 6. First login

When you open the deployed site for the first time, it will ask you to create the first admin user.

After that, admins can create separate logins from the Users page. Each user only sees their own statements, people, rules, payments, and dashboard.

## Notes on PDF accuracy

This uses text extraction through PDF.js, not OCR. It works best when the PDF text is selectable/copy-pastable. Since Chase and Capital One PDFs can vary a bit, review the parsed rows before saving. The app is built around that review step so you can edit anything before it hits the database.
