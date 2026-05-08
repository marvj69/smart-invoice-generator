const { neon } = require('@neondatabase/serverless');

let sqlClient = null;
let schemaReady = null;

function getDatabaseUrl() {
  const url = process.env.DATABASE_URL || process.env.POSTGRES_URL || process.env.NEON_DATABASE_URL;
  if (!url) {
    const error = new Error('Cloud database is not configured.');
    error.statusCode = 503;
    error.code = 'DATABASE_NOT_CONFIGURED';
    throw error;
  }
  return url;
}

function getSql() {
  if (!sqlClient) {
    sqlClient = neon(getDatabaseUrl());
  }
  return sqlClient;
}

async function ensureSchema() {
  if (schemaReady) {
    return schemaReady;
  }

  schemaReady = (async () => {
    const sql = getSql();

    await sql`
      CREATE TABLE IF NOT EXISTS app_users (
        id text PRIMARY KEY,
        email text NOT NULL UNIQUE,
        password_hash text NOT NULL,
        created_at timestamptz NOT NULL DEFAULT now(),
        updated_at timestamptz NOT NULL DEFAULT now()
      )
    `;

    await sql`
      CREATE TABLE IF NOT EXISTS app_sessions (
        id text PRIMARY KEY,
        user_id text NOT NULL REFERENCES app_users(id) ON DELETE CASCADE,
        token_hash text NOT NULL UNIQUE,
        expires_at timestamptz NOT NULL,
        created_at timestamptz NOT NULL DEFAULT now()
      )
    `;

    await sql`
      CREATE INDEX IF NOT EXISTS app_sessions_user_id_idx
      ON app_sessions (user_id)
    `;

    await sql`
      CREATE INDEX IF NOT EXISTS app_sessions_expires_at_idx
      ON app_sessions (expires_at)
    `;

    await sql`
      CREATE TABLE IF NOT EXISTS user_invoice_data (
        user_id text PRIMARY KEY REFERENCES app_users(id) ON DELETE CASCADE,
        current_invoice jsonb NOT NULL DEFAULT '{}'::jsonb,
        templates jsonb NOT NULL DEFAULT '[]'::jsonb,
        default_company jsonb NOT NULL DEFAULT '{}'::jsonb,
        updated_at timestamptz NOT NULL DEFAULT now()
      )
    `;
  })().catch((error) => {
    schemaReady = null;
    throw error;
  });

  return schemaReady;
}

async function query(strings, ...values) {
  await ensureSchema();
  return getSql()(strings, ...values);
}

module.exports = {
  ensureSchema,
  query
};
