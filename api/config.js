// Serves the Supabase settings to the app from Vercel Environment Variables,
// so the keys live in Vercel (and in your local .env), never in the GitHub repo.
module.exports = (req, res) => {
  const cfg = {
    supabaseUrl: process.env.SUPABASE_URL || "",
    supabaseAnonKey: process.env.SUPABASE_ANON_KEY || process.env.SUPABASE_PUBLISHABLE_KEY || ""
  };
  res.setHeader("Content-Type", "application/javascript; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  res.end("window.MB_CONFIG = " + JSON.stringify(cfg) + ";");
};
