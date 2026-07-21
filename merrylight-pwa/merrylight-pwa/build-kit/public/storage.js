window.__mlBackend = { configured: false, status: "untested", detail: "" };
(function () {
  const cfg = () => window.MLERP_BACKEND || {};
  const on = () => !!(cfg().url && cfg().anonKey);
  setTimeout(function () {
    window.__mlBackend.configured = on();
    if (!on()) { window.__mlBackend.status = "off"; return; }
    fetch(cfg().url.replace(/\/+$/, "").replace(/\/rest\/v1$/, "") + "/rest/v1/kv?select=key&limit=1", {
      headers: { apikey: cfg().anonKey, Authorization: "Bearer " + cfg().anonKey },
    }).then(function (res) {
      if (res.ok) { window.__mlBackend.status = "ok"; }
      else res.text().then(function (t) {
        window.__mlBackend.status = "error";
        window.__mlBackend.detail = "HTTP " + res.status + ": " + (t || "").slice(0, 160);
      });
    }).catch(function (e) {
      window.__mlBackend.status = "error";
      window.__mlBackend.detail = String(e).slice(0, 160);
    });
  }, 0);
  const H = () => ({
    apikey: cfg().anonKey,
    Authorization: "Bearer " + cfg().anonKey,
    "Content-Type": "application/json",
  });
  const norm = () => cfg().url.replace(/\/+$/, "").replace(/\/rest\/v1$/, "");
  const base = () => norm() + "/rest/v1/kv";
  const lsGet = (k) => localStorage.getItem("mlerp:" + k);
  const lsSet = (k, v) => { try { localStorage.setItem("mlerp:" + k, v); } catch {} };
  const lsDel = (k) => localStorage.removeItem("mlerp:" + k);

  window.storage = {
    async get(key, shared) {
      if (on()) {
        try {
          const res = await fetch(base() + "?select=value&key=eq." + encodeURIComponent(key), { headers: H() });
          if (res.ok) {
            const rows = await res.json();
            if (rows.length) { lsSet(key, rows[0].value); return { key, value: rows[0].value, shared: !!shared }; }
            const v = lsGet(key); // offline-written value not yet on server
            if (v != null) return { key, value: v, shared: !!shared };
            throw new Error("key not found: " + key);
          }
        } catch (e) {
          const v = lsGet(key); // network down: serve cache
          if (v != null) return { key, value: v, shared: !!shared };
          throw e;
        }
      }
      const v = lsGet(key);
      if (v == null) throw new Error("key not found: " + key);
      return { key, value: v, shared: !!shared };
    },
    async set(key, value, shared) {
      lsSet(key, value);
      if (on()) {
        try {
          const res = await fetch(base(), {
            method: "POST",
            headers: { ...H(), Prefer: "resolution=merge-duplicates" },
            body: JSON.stringify({ key, value, updated_at: new Date().toISOString() }),
          });
          if (!res.ok) {
            window.__mlBackend.status = "error";
            const t = await res.text().catch(() => "");
            window.__mlBackend.detail = "write HTTP " + res.status + ": " + (t || "").slice(0, 160);
          }
        } catch (e) {
          window.__mlBackend.status = "error";
          window.__mlBackend.detail = "write: " + String(e).slice(0, 160);
        }
      }
      return { key, value, shared: !!shared };
    },
    async delete(key, shared) {
      lsDel(key);
      if (on()) {
        try { await fetch(base() + "?key=eq." + encodeURIComponent(key), { method: "DELETE", headers: H() }); } catch {}
      }
      return { key, deleted: true, shared: !!shared };
    },
    async list(prefix, shared) {
      if (on()) {
        try {
          const q = prefix ? "&key=like." + encodeURIComponent(prefix) + "*" : "";
          const res = await fetch(base() + "?select=key" + q, { headers: H() });
          if (res.ok) {
            const rows = await res.json();
            return { keys: rows.map((r) => r.key), prefix, shared: !!shared };
          }
        } catch {}
      }
      const p = "mlerp:" + (prefix || "");
      const keys = Object.keys(localStorage).filter((k) => k.startsWith(p)).map((k) => k.slice(6));
      return { keys, prefix, shared: !!shared };
    },
  };
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("sw.js").catch(() => {});
    });
  }
})();
