import { createClient } from "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm";

const cfg = window.PORTAL_CONFIG;
const el = (id) => document.getElementById(id);

function showError(msg) {
  const box = el("errorBox");
  if (!box) return;
  box.hidden = !msg;
  box.textContent = msg || "";
}

function showView(name) {
  el("viewLogin").hidden = name !== "login";
  el("viewApps").hidden = name !== "apps";
}

let supabase;

function initClient() {
  if (!cfg?.supabaseUrl || !cfg?.supabaseKey) {
    showError("Falta configurar config.js (URL e chave anon/publishable).");
    return null;
  }
  const key = String(cfg.supabaseKey || "").trim();
  if (!key || key.includes("COLOCA_AQUI")) {
    showError(
      "Edita config.js: cola a chave completa em supabaseKey (Publishable sb_publishable_… ou legacy anon eyJ…). Sem espaços nem texto a mais.",
    );
    return null;
  }
  const url = String(cfg.supabaseUrl || "").trim().replace(/\/+$/, "");
  return createClient(url, key, {
    auth: { persistSession: true, autoRefreshToken: true },
  });
}

async function loadApps() {
  const list = el("appList");
  list.innerHTML = '<p class="muted">A carregar…</p>';

  const { data, error } = await supabase
    .from("applications")
    .select("id, slug, name, description, external_url, sort_order")
    .eq("is_active", true)
    .order("sort_order", { ascending: true });

  if (error) {
    list.innerHTML = "";
    showError(error.message || "Erro ao ler aplicações.");
    return;
  }

  showError("");
  if (!data || data.length === 0) {
    list.innerHTML =
      '<p class="muted">Não tens acesso a nenhuma aplicação (ou ainda não há permissões). Um administrador pode dar-te acesso em <code>user_application_access</code>.</p>';
    return;
  }

  list.innerHTML = "";
  for (const app of data) {
    const card = document.createElement("article");
    card.className = "app-card";
    const href = app.external_url && app.external_url.trim() ? app.external_url.trim() : "#";
    const canOpen = href !== "#";
    card.innerHTML = `
      <h3 class="app-card__title">${escapeHtml(app.name)}</h3>
      <p class="app-card__desc">${escapeHtml(app.description || "")}</p>
      <p class="app-card__slug"><code>${escapeHtml(app.slug)}</code></p>
      ${
        canOpen
          ? `<a class="btn btn--primary" href="${escapeAttr(href)}" ${appLinkAttrs(href)}>Abrir</a>`
          : '<p class="app-card__hint muted">Sem link público ainda. No Supabase → Table Editor → <code>applications</code> → coluna <code>external_url</code> (ex.: <code>…/agenda/</code> no deploy unificado ou URL próprio da agenda).</p><button type="button" class="btn btn--ghost" disabled title="Defina external_url no Supabase">Abrir (indisponível)</button>'
      }
    `;
    list.appendChild(card);
  }
}

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s;
  return d.innerHTML;
}

function escapeAttr(s) {
  return String(s).replace(/"/g, "&quot;");
}

/** No mesmo domínio que o portal: abre no mesmo separador (um só “sítio”). */
function appLinkAttrs(href) {
  try {
    const u = new URL(href, window.location.href);
    if (u.origin === window.location.origin) {
      return 'target="_self"';
    }
  } catch (_) {
    /* ignore */
  }
  return 'target="_blank" rel="noopener noreferrer"';
}

async function refreshSessionUI() {
  const {
    data: { session },
  } = await supabase.auth.getSession();
  if (!session) {
    showView("login");
    el("userEmail").textContent = "";
    return;
  }
  showView("apps");
  el("userEmail").textContent = session.user.email || session.user.id;
  await loadApps();
}

el("formLogin")?.addEventListener("submit", async (e) => {
  e.preventDefault();
  showError("");
  const email = el("email").value.trim();
  const password = el("password").value;
  const { error } = await supabase.auth.signInWithPassword({ email, password });
  if (error) {
    showError(error.message);
    return;
  }
  await refreshSessionUI();
});

el("btnLogout")?.addEventListener("click", async () => {
  await supabase.auth.signOut();
  showView("login");
  el("userEmail").textContent = "";
  showError("");
});

supabase = initClient();
if (supabase) {
  supabase.auth.onAuthStateChange(() => {
    void refreshSessionUI();
  });
  void refreshSessionUI();
}
