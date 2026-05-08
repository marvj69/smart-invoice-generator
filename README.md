Generates Invoices Easily

## Backend

This app uses Vercel Functions for account auth and user data sync. The API expects a Postgres-compatible `DATABASE_URL`; Neon through Vercel Marketplace is the intended setup.

Required Vercel environment variable:

```env
DATABASE_URL="postgres://USER:PASSWORD@HOST/DB?sslmode=require"
```

Once `DATABASE_URL` is present, users can create accounts, sign in, and sync their current invoice draft, saved templates, and default company details across devices.

Useful commands:

```bash
npm run check
npm run dev
npm run deploy
```
