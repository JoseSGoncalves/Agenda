/**
 * Copia para config.js e preenche com o teu projecto Supabase.
 * cp config.example.js config.js
 */
const SUPABASE_SITE_CONFIG = {
  supabaseUrl: "https://SEU_PROJECTO.supabase.co",
  supabaseKey: "sb_publishable_...",
};

window.PORTAL_CONFIG = SUPABASE_SITE_CONFIG;
window.AGENDA_CONFIG = {
  ...SUPABASE_SITE_CONFIG,
  localOnly: false,
  persistSession: true,
};
