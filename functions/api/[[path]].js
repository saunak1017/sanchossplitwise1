const encoder = new TextEncoder();

function json(data, status = 200, headers = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8', ...headers }
  });
}

async function readJson(request) {
  try { return await request.json(); } catch { return {}; }
}

function getCookie(request, name) {
  const cookie = request.headers.get('cookie') || '';
  return cookie.split(';').map(x => x.trim()).find(x => x.startsWith(name + '='))?.slice(name.length + 1) || '';
}

function b64url(bytes) {
  const str = typeof bytes === 'string' ? bytes : String.fromCharCode(...new Uint8Array(bytes));
  return btoa(str).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function fromB64url(str) {
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  while (str.length % 4) str += '=';
  return Uint8Array.from(atob(str), c => c.charCodeAt(0));
}

async function hmac(secret, message) {
  const key = await crypto.subtle.importKey('raw', encoder.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  return b64url(await crypto.subtle.sign('HMAC', key, encoder.encode(message)));
}

async function makeToken(env, user) {
  const payload = b64url(JSON.stringify({ id: user.id, email: user.email, role: user.role, exp: Date.now() + 1000 * 60 * 60 * 24 * 14 }));
  const sig = await hmac(env.SESSION_SECRET || 'dev-secret-change-me', payload);
  return `${payload}.${sig}`;
}

async function verifyToken(env, token) {
  if (!token || !token.includes('.')) return null;
  const [payload, sig] = token.split('.');
  const expected = await hmac(env.SESSION_SECRET || 'dev-secret-change-me', payload);
  if (sig !== expected) return null;
  try {
    const data = JSON.parse(new TextDecoder().decode(fromB64url(payload)));
    if (!data.exp || Date.now() > data.exp) return null;
    return data;
  } catch { return null; }
}

async function hashPassword(password, salt) {
  const key = await crypto.subtle.importKey('raw', encoder.encode(password), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', salt: encoder.encode(salt), iterations: 120000, hash: 'SHA-256' }, key, 256);
  return b64url(bits);
}

function newSalt() {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return b64url(bytes);
}

async function requireUser(env, request) {
  const token = await verifyToken(env, getCookie(request, 'sst_session'));
  if (!token) return null;
  const user = await env.DB.prepare('SELECT id, name, email, role FROM users WHERE id = ?').bind(token.id).first();
  return user || null;
}

function cookieHeader(token) {
  return `sst_session=${token}; Path=/; HttpOnly; SameSite=Lax; Secure; Max-Age=${60 * 60 * 24 * 14}`;
}

async function getOrCreatePerson(env, userId, name) {
  const cleaned = String(name || '').trim();
  if (!cleaned) return null;
  let p = await env.DB.prepare('SELECT id, name FROM people WHERE user_id = ? AND lower(name) = lower(?)').bind(userId, cleaned).first();
  if (p) return p;
  await env.DB.prepare('INSERT INTO people (user_id, name) VALUES (?, ?)').bind(userId, cleaned).run();
  return await env.DB.prepare('SELECT id, name FROM people WHERE user_id = ? AND lower(name) = lower(?)').bind(userId, cleaned).first();
}

async function route(request, env) {
  const url = new URL(request.url);
  const path = url.pathname.replace(/^\/api\/?/, '');
  const parts = path.split('/').filter(Boolean);
  const method = request.method;

  if (method === 'GET' && parts[0] === 'health') return json({ ok: true });

  if (method === 'GET' && parts[0] === 'setup-status') {
    const row = await env.DB.prepare('SELECT COUNT(*) AS count FROM users').first();
    return json({ needsSetup: !row || row.count === 0 });
  }

  if (method === 'POST' && parts[0] === 'setup') {
    const count = await env.DB.prepare('SELECT COUNT(*) AS count FROM users').first();
    if (count?.count > 0) return json({ error: 'Setup has already been completed.' }, 400);
    const body = await readJson(request);
    if (!body.name || !body.email || !body.password) return json({ error: 'Name, email, and password are required.' }, 400);
    const salt = newSalt();
    const pwHash = await hashPassword(body.password, salt);
    await env.DB.prepare('INSERT INTO users (name, email, password_salt, password_hash, role) VALUES (?, ?, ?, ?, ?)')
      .bind(body.name.trim(), body.email.trim().toLowerCase(), salt, pwHash, 'admin').run();
    const user = await env.DB.prepare('SELECT id, name, email, role FROM users WHERE email = ?').bind(body.email.trim().toLowerCase()).first();
    const token = await makeToken(env, user);
    return json({ user }, 200, { 'set-cookie': cookieHeader(token) });
  }

  if (method === 'POST' && parts[0] === 'login') {
    const body = await readJson(request);
    const user = await env.DB.prepare('SELECT * FROM users WHERE email = ?').bind(String(body.email || '').trim().toLowerCase()).first();
    if (!user) return json({ error: 'Invalid login.' }, 401);
    const pwHash = await hashPassword(String(body.password || ''), user.password_salt);
    if (pwHash !== user.password_hash) return json({ error: 'Invalid login.' }, 401);
    const safe = { id: user.id, name: user.name, email: user.email, role: user.role };
    const token = await makeToken(env, safe);
    return json({ user: safe }, 200, { 'set-cookie': cookieHeader(token) });
  }

  if (method === 'POST' && parts[0] === 'logout') {
    return json({ ok: true }, 200, { 'set-cookie': 'sst_session=; Path=/; HttpOnly; SameSite=Lax; Secure; Max-Age=0' });
  }

  const user = await requireUser(env, request);
  if (!user) return json({ error: 'Unauthorized' }, 401);

  if (method === 'GET' && parts[0] === 'me') return json({ user });

  if (parts[0] === 'users') {
    if (user.role !== 'admin') return json({ error: 'Admin only.' }, 403);
    if (method === 'GET') {
      const rows = await env.DB.prepare('SELECT id, name, email, role, created_at FROM users ORDER BY created_at DESC').all();
      return json({ users: rows.results || [] });
    }
    if (method === 'POST') {
      const body = await readJson(request);
      if (!body.name || !body.email || !body.password) return json({ error: 'Name, email, and password are required.' }, 400);
      const salt = newSalt();
      const pwHash = await hashPassword(body.password, salt);
      await env.DB.prepare('INSERT INTO users (name, email, password_salt, password_hash, role) VALUES (?, ?, ?, ?, ?)')
        .bind(body.name.trim(), body.email.trim().toLowerCase(), salt, pwHash, body.role === 'admin' ? 'admin' : 'user').run();
      return json({ ok: true });
    }
  }

  if (parts[0] === 'people') {
    if (method === 'GET') {
      const rows = await env.DB.prepare('SELECT id, name FROM people WHERE user_id = ? ORDER BY name').bind(user.id).all();
      return json({ people: rows.results || [] });
    }
    if (method === 'POST') {
      const body = await readJson(request);
      const p = await getOrCreatePerson(env, user.id, body.name);
      return json({ person: p });
    }
  }

  if (parts[0] === 'merchant-rules') {
    if (method === 'GET') {
      const rows = await env.DB.prepare('SELECT id, match_text, clean_name FROM merchant_rules WHERE user_id = ? ORDER BY match_text').bind(user.id).all();
      return json({ rules: rows.results || [] });
    }
    if (method === 'POST') {
      const body = await readJson(request);
      if (!body.match_text || !body.clean_name) return json({ error: 'Both fields are required.' }, 400);
      await env.DB.prepare('INSERT INTO merchant_rules (user_id, match_text, clean_name) VALUES (?, ?, ?)')
        .bind(user.id, body.match_text.trim(), body.clean_name.trim()).run();
      return json({ ok: true });
    }
    if (method === 'DELETE' && parts[1]) {
      await env.DB.prepare('DELETE FROM merchant_rules WHERE user_id = ? AND id = ?').bind(user.id, parts[1]).run();
      return json({ ok: true });
    }
  }

  if (parts[0] === 'statements') {
    if (method === 'GET' && !parts[1]) {
      const rows = await env.DB.prepare(`
        SELECT s.*, COALESCE(SUM(li.amount), 0) AS total
        FROM statements s
        LEFT JOIN line_items li ON li.statement_id = s.id AND li.user_id = s.user_id
        WHERE s.user_id = ?
        GROUP BY s.id
        ORDER BY s.created_at DESC
      `).bind(user.id).all();
      return json({ statements: rows.results || [] });
    }
    if (method === 'POST') {
      const body = await readJson(request);
      if (!body.title || !body.issuer || !Array.isArray(body.rows)) return json({ error: 'Missing statement data.' }, 400);
      const statement = await env.DB.prepare('INSERT INTO statements (user_id, issuer, title, period_start, period_end, notes) VALUES (?, ?, ?, ?, ?, ?) RETURNING id')
        .bind(user.id, body.issuer, body.title.trim(), body.period_start || null, body.period_end || null, body.notes || null).first();
      const statementId = statement.id;
      for (let i = 0; i < body.rows.length; i++) {
        const row = body.rows[i];
        const tx = await env.DB.prepare('INSERT INTO transactions (user_id, statement_id, transaction_date, merchant, original_description, amount, sort_order) VALUES (?, ?, ?, ?, ?, ?, ?) RETURNING id')
          .bind(user.id, statementId, row.date || null, row.merchant || 'Unknown', row.original || row.merchant || '', Number(row.originalAmount ?? row.amount ?? 0), i).first();
        const items = Array.isArray(row.lineItems) && row.lineItems.length ? row.lineItems : [{ person: row.person, amount: row.amount }];
        for (const item of items) {
          if (!item.person) continue;
          const p = await getOrCreatePerson(env, user.id, item.person);
          if (!p) continue;
          await env.DB.prepare('INSERT INTO line_items (user_id, statement_id, transaction_id, person_id, amount, notes) VALUES (?, ?, ?, ?, ?, ?)')
            .bind(user.id, statementId, tx.id, p.id, Number(item.amount || 0), item.notes || null).run();
        }
      }
      return json({ id: statementId });
    }
    if (method === 'GET' && parts[1]) {
      const statement = await env.DB.prepare('SELECT * FROM statements WHERE user_id = ? AND id = ?').bind(user.id, parts[1]).first();
      if (!statement) return json({ error: 'Statement not found.' }, 404);
      const txs = await env.DB.prepare(`
        SELECT t.* FROM transactions t WHERE t.user_id = ? AND t.statement_id = ? ORDER BY t.sort_order, t.id
      `).bind(user.id, parts[1]).all();
      const items = await env.DB.prepare(`
        SELECT li.*, p.name AS person, t.merchant, t.original_description, t.transaction_date, t.amount AS original_amount
        FROM line_items li
        JOIN people p ON p.id = li.person_id
        JOIN transactions t ON t.id = li.transaction_id
        WHERE li.user_id = ? AND li.statement_id = ?
        ORDER BY t.sort_order, li.id
      `).bind(user.id, parts[1]).all();
      return json({ statement, transactions: txs.results || [], lineItems: items.results || [] });
    }
    if (method === 'DELETE' && parts[1]) {
      await env.DB.prepare('DELETE FROM statements WHERE user_id = ? AND id = ?').bind(user.id, parts[1]).run();
      return json({ ok: true });
    }
  }

  if (parts[0] === 'payments') {
    if (method === 'GET') {
      const rows = await env.DB.prepare(`
        SELECT pay.*, p.name AS person FROM payments pay JOIN people p ON p.id = pay.person_id
        WHERE pay.user_id = ? ORDER BY pay.payment_date DESC, pay.created_at DESC
      `).bind(user.id).all();
      return json({ payments: rows.results || [] });
    }
    if (method === 'POST') {
      const body = await readJson(request);
      const p = await getOrCreatePerson(env, user.id, body.person);
      if (!p) return json({ error: 'Person is required.' }, 400);
      await env.DB.prepare('INSERT INTO payments (user_id, person_id, payment_date, type, method, amount, notes) VALUES (?, ?, ?, ?, ?, ?, ?)')
        .bind(user.id, p.id, body.payment_date || new Date().toISOString().slice(0,10), body.type || 'Paid', body.method || 'Venmo', Number(body.amount || 0), body.notes || null).run();
      return json({ ok: true });
    }
    if (method === 'DELETE' && parts[1]) {
      await env.DB.prepare('DELETE FROM payments WHERE user_id = ? AND id = ?').bind(user.id, parts[1]).run();
      return json({ ok: true });
    }
  }

  if (parts[0] === 'dashboard' && method === 'GET') {
    const people = await env.DB.prepare('SELECT id, name FROM people WHERE user_id = ? ORDER BY name').bind(user.id).all();
    const assigned = await env.DB.prepare(`
      SELECT p.id AS person_id, COALESCE(SUM(li.amount),0) AS assigned
      FROM people p LEFT JOIN line_items li ON li.person_id = p.id AND li.user_id = p.user_id
      WHERE p.user_id = ? GROUP BY p.id
    `).bind(user.id).all();
    const payments = await env.DB.prepare(`
      SELECT p.id AS person_id,
        SUM(CASE WHEN pay.type = 'Paid' THEN pay.amount ELSE 0 END) AS paid,
        SUM(CASE WHEN pay.type = 'Moved to Splitwise' THEN pay.amount ELSE 0 END) AS splitwise,
        SUM(CASE WHEN pay.type = 'Adjustment' THEN pay.amount ELSE 0 END) AS adjustment
      FROM people p LEFT JOIN payments pay ON pay.person_id = p.id AND pay.user_id = p.user_id
      WHERE p.user_id = ? GROUP BY p.id
    `).bind(user.id).all();
    const detail = await env.DB.prepare(`
      SELECT p.id AS person_id, p.name AS person, s.id AS statement_id, s.title AS statement_title, s.issuer,
             t.transaction_date, t.merchant, t.original_description, t.amount AS original_amount, li.amount AS line_amount
      FROM line_items li
      JOIN people p ON p.id = li.person_id
      JOIN statements s ON s.id = li.statement_id
      JOIN transactions t ON t.id = li.transaction_id
      WHERE li.user_id = ?
      ORDER BY p.name, s.created_at DESC, t.sort_order, li.id
    `).bind(user.id).all();
    const payDetail = await env.DB.prepare(`
      SELECT pay.*, p.name AS person FROM payments pay JOIN people p ON p.id = pay.person_id
      WHERE pay.user_id = ? ORDER BY pay.payment_date DESC, pay.created_at DESC
    `).bind(user.id).all();
    const a = Object.fromEntries((assigned.results || []).map(r => [r.person_id, Number(r.assigned || 0)]));
    const pm = Object.fromEntries((payments.results || []).map(r => [r.person_id, { paid: Number(r.paid || 0), splitwise: Number(r.splitwise || 0), adjustment: Number(r.adjustment || 0) }]));
    const summary = (people.results || []).map(p => {
      const pay = pm[p.id] || { paid: 0, splitwise: 0, adjustment: 0 };
      const assignedTotal = a[p.id] || 0;
      return { ...p, assigned: assignedTotal, ...pay, open: assignedTotal - pay.paid - pay.splitwise - pay.adjustment };
    });
    return json({ summary, detail: detail.results || [], payments: payDetail.results || [] });
  }

  return json({ error: 'Not found' }, 404);
}

export const onRequest = ({ request, env }) => route(request, env);
