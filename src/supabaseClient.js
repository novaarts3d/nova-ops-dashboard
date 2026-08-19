import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // Fails loudly at build/run time rather than silently breaking every
  // save/load — much easier to diagnose than mysterious empty data.
  console.error(
    "Missing VITE_SUPABASE_URL or VITE_SUPABASE_ANON_KEY. Add them to a .env file locally, and to Vercel → Project Settings → Environment Variables for the deployed site."
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
