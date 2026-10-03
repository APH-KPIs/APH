// Shared database connection (Supabase). Both values are public by design: the key only allows
// reading the dashboard data and calling the login-checked functions in supabase/setup.sql.
// Leave supabaseUrl empty to run in local mode (data kept in this browser only).
window.APP_CONFIG = {
    supabaseUrl: "https://oolyviictebdfsrydrhm.supabase.co",
    supabaseKey: "sb_publishable_ZVptrcILdJqD-8ndvz5Blw_JkgAXDQm"
};
