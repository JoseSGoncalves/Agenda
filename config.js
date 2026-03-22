/**
 * Um ficheiro para portal (raiz) + agenda (/agenda/).
 * Com URL + chave válidos, a agenda usa sempre Supabase (localOnly no config é ignorado).
 */
const SUPABASE_SITE_CONFIG = {
  supabaseUrl: "https://zzmcnygjzcdsvohovdnq.supabase.co",
  supabaseKey: "sb_publishable_WBSczkEDUn_8J6DDVnNisA_ar289e4l",
};

window.PORTAL_CONFIG = SUPABASE_SITE_CONFIG;
window.AGENDA_CONFIG = {
  ...SUPABASE_SITE_CONFIG,
  localOnly: false,
  persistSession: true,
};
