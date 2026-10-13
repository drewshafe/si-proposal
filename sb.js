// sb.js — ShipInsure calculators shared Supabase client (Tier 2: saved links / folders)
// The publishable key is PUBLIC by design (client-side), fenced by row-level security.
// This is NOT a service/secret key. See supabase-schema.sql for the security model.
(function () {
  const SB_URL = 'https://oiljklutlmtztascnkpm.supabase.co';
  const SB_KEY = 'sb_publishable_DTbJR5CxMe6wFYKKhogEtQ_XJyp4nMs';
  const CDN = 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/dist/umd/supabase.js';

  let _client = null, _loading = null;
  async function client() {
    if (_client) return _client;
    if (!_loading) {
      _loading = new Promise((resolve, reject) => {
        if (window.supabase && window.supabase.createClient) return resolve();
        const s = document.createElement('script');
        s.src = CDN; s.onload = resolve;
        s.onerror = () => reject(new Error('supabase-js failed to load'));
        document.head.appendChild(s);
      }).then(() => {
        _client = window.supabase.createClient(SB_URL, SB_KEY, {
          auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true }
        });
      });
    }
    await _loading;
    return _client;
  }

  // ── Auth (rep magic-link) ──────────────────────────────────────────────
  async function currentUser() { const c = await client(); const { data } = await c.auth.getUser(); return data ? data.user : null; }
  async function signInMagic(email) {
    const c = await client();
    // Public signups are disabled — only already-invited reps can sign in, so don't
    // attempt to create a user (that would be rejected). Existing users get a link.
    return c.auth.signInWithOtp({ email: email, options: { shouldCreateUser: false, emailRedirectTo: location.href.split('#')[0] } });
  }
  async function signOut() { const c = await client(); return c.auth.signOut(); }

  // ── Storage (logo images) ──────────────────────────────────────────────
  // Upload a data: URL to the logos bucket; return its public URL.
  async function uploadLogo(dataUrl) {
    if (!dataUrl || dataUrl.indexOf('data:') !== 0) return dataUrl || null;
    const c = await client();
    const blob = await (await fetch(dataUrl)).blob();
    const ext = ((blob.type.split('/')[1]) || 'png').replace('jpeg', 'jpg').replace('svg+xml', 'svg');
    const path = 'l/' + Date.now() + '-' + Math.random().toString(36).slice(2, 8) + '.' + ext;
    const { error } = await c.storage.from('logos').upload(path, blob, { contentType: blob.type, upsert: false });
    if (error) throw error;
    return c.storage.from('logos').getPublicUrl(path).data.publicUrl;
  }

  // ── Configs (saved calculator links) ───────────────────────────────────
  // save({id?, merchant, kind, title, data, logoUrl?}) → { id, edit_token }
  async function saveConfig(cfg) {
    const c = await client();
    const row = {
      merchant: cfg.merchant || '', kind: cfg.kind || 'proposal',
      title: cfg.title || '', data: cfg.data || {},
      updated_by: 'rep', updated_at: new Date().toISOString()
    };
    if (cfg.logoUrl != null) row.logo_path = cfg.logoUrl;
    if (cfg.merchant_ref != null) row.merchant_ref = cfg.merchant_ref;
    const q = c.from('configs');
    const res = cfg.id
      ? await q.update(row).eq('id', cfg.id).select('id, edit_token').single()
      : await q.insert(row).select('id, edit_token').single();
    if (res.error) throw res.error;
    return res.data;
  }
  async function loadConfig(id) {
    const c = await client();
    const { data, error } = await c.from('configs')
      .select('id, merchant, merchant_ref, kind, title, data, logo_path, updated_by, updated_at').eq('id', id).single();
    if (error) throw error;
    return data;
  }
  async function listConfigs() {
    const c = await client();
    const { data, error } = await c.from('configs')
      .select('id, merchant, merchant_ref, kind, title, logo_path, updated_by, updated_at')
      .order('merchant', { ascending: true }).order('updated_at', { ascending: false });
    if (error) throw error;
    return data || [];
  }

  // ── Merchants (folder-level branding; single source for all scenarios) ──
  // bld = cfg() output (for rendering), cb = cbSnapshot() (for editing), logo in storage.
  async function saveMerchant(m) {
    const c = await client();
    const row = { name: m.name || '', bld: m.bld || {}, cb: m.cb || {}, updated_at: new Date().toISOString() };
    if (m.logoUrl != null) row.logo_path = m.logoUrl;
    let id = m.id;
    if (!id && m.name) { const ex = await getMerchantByName(m.name); if (ex) id = ex.id; }
    const q = c.from('merchants');
    const res = id ? await q.update(row).eq('id', id).select('id').single()
                   : await q.insert(row).select('id').single();
    if (res.error) throw res.error;
    return res.data;
  }
  async function getMerchant(id) {
    const c = await client();
    const { data, error } = await c.from('merchants').select('id, name, bld, cb, logo_path, updated_at').eq('id', id).single();
    if (error) throw error;
    return data;
  }
  async function getMerchantByName(name) {
    const c = await client();
    const { data, error } = await c.from('merchants').select('id, name, bld, cb, logo_path').ilike('name', name).limit(1);
    if (error) throw error;
    return (data && data[0]) || null;
  }
  async function listMerchants() {
    const c = await client();
    const { data, error } = await c.from('merchants').select('id, name, logo_path, updated_at').order('name', { ascending: true });
    if (error) throw error;
    return data || [];
  }
  // Merchant save-back — token-gated, no login required.
  async function updateByToken(id, token, data) {
    const c = await client();
    const { error } = await c.rpc('update_config_by_token', { p_id: id, p_token: token, p_data: data });
    if (error) throw error;
    return true;
  }
  async function deleteConfig(id) { const c = await client(); const { error } = await c.from('configs').delete().eq('id', id); if (error) throw error; return true; }

  window.SIsb = {
    URL: SB_URL, client, currentUser, signInMagic, signOut,
    uploadLogo, saveConfig, loadConfig, listConfigs, updateByToken, deleteConfig,
    saveMerchant, getMerchant, getMerchantByName, listMerchants
  };
})();
