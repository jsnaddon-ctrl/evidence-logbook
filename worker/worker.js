/* Evidence Logbook — AI helper (Cloudflare Worker)
 * Keeps the Anthropic API key private and passes requests from the app to Claude.
 *
 * Settings (Worker > Settings > Variables and Secrets):
 *   ANTHROPIC_API_KEY  (Secret)  your key from console.anthropic.com
 *   ALLOWED_ORIGINS    (Text)    e.g. https://yourname.github.io
 *   ACCESS_CODES       (Text)    optional, comma-separated codes you give apprentices, e.g. symonite2026,diesel-ben
 *   MODEL              (Text)    optional, defaults to claude-sonnet-5-5
 */
export default {
  async fetch(req, env) {
    const origin = req.headers.get('Origin') || '';
    const allowed = (env.ALLOWED_ORIGINS || '').split(',').map(s => s.trim().replace(/\/$/, '')).filter(Boolean);
    const okOrigin = allowed.includes(origin);
    const cors = {
      'Access-Control-Allow-Origin': okOrigin ? origin : 'null',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, X-Access-Code',
      'Access-Control-Max-Age': '86400',
      'Vary': 'Origin'
    };
    const json = (obj, status = 200) => new Response(JSON.stringify(obj), { status, headers: { ...cors, 'Content-Type': 'application/json' } });

    if (req.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors });
    if (!okOrigin) return json({ error: 'origin_not_allowed' }, 403);
    if (req.method !== 'POST') return json({ error: 'method_not_allowed' }, 405);
    if (!env.ANTHROPIC_API_KEY) return json({ error: 'not_configured' }, 500);

    const codes = (env.ACCESS_CODES || '').split(',').map(s => s.trim()).filter(Boolean);
    if (codes.length && !codes.includes((req.headers.get('X-Access-Code') || '').trim())) return json({ error: 'bad_code' }, 401);

    let body;
    try { body = await req.json(); } catch { return json({ error: 'bad_request' }, 400); }
    const prompt = String(body.prompt || '');
    if (!prompt) return json({ error: 'bad_request' }, 400);
    if (prompt.length > 400000) return json({ error: 'too_large' }, 413);
    const images = (Array.isArray(body.images) ? body.images : []).slice(0, 20)
      .filter(im => im && typeof im.data === 'string')
      .map(im => ({ type: 'image', source: { type: 'base64', media_type: /^image\/(jpeg|png|webp|gif)$/.test(im.media_type) ? im.media_type : 'image/jpeg', data: im.data } }));

    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'x-api-key': env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      body: JSON.stringify({
        model: env.MODEL || 'claude-sonnet-5-5',
        max_tokens: Math.min(Math.max(parseInt(body.max_tokens) || 2000, 200), 8000),
        system: 'You help apprentices organise their workplace evidence. Reply with only valid JSON, no other text.',
        messages: [{ role: 'user', content: [...images, { type: 'text', text: prompt }] }]
      })
    });
    const data = await r.json().catch(() => ({}));
    if (!r.ok) {
      const tooBig = r.status === 413 || /too long|too large/i.test((data.error && data.error.message) || '');
      return json({ error: tooBig ? 'too_large' : 'upstream', status: r.status }, 502);
    }
    const text = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
    return json({ text });
  }
};
