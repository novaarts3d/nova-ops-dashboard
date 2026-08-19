import React, { useState } from "react";
import { supabase } from "../supabaseClient.js";

export default function Login() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    const { error: signInError } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (signInError) setError(signInError.message);
  };

  return (
    <div className="min-h-screen bg-neutral-50 flex items-center justify-center p-4">
      <div className="w-full max-w-sm bg-white border border-neutral-200 rounded-2xl shadow-sm p-6">
        <div className="text-red-600 font-bold text-lg mb-1">NOVA</div>
        <h1 className="text-xl font-bold text-neutral-900 mb-1">Sign in</h1>
        <p className="text-sm text-neutral-500 mb-5">Operations Dashboard</p>
        <form onSubmit={handleSubmit} className="space-y-3">
          <label className="flex flex-col gap-1 text-sm">
            <span className="text-neutral-600 font-medium">Email</span>
            <input
              type="email"
              required
              autoFocus
              className="border border-neutral-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-red-500 focus:border-red-500"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
            />
          </label>
          <label className="flex flex-col gap-1 text-sm">
            <span className="text-neutral-600 font-medium">Password</span>
            <input
              type="password"
              required
              className="border border-neutral-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-red-500 focus:border-red-500"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
          </label>
          {error && <p className="text-xs text-red-600">{error}</p>}
          <button
            type="submit"
            disabled={loading}
            className="w-full bg-red-600 hover:bg-red-700 disabled:opacity-50 text-white text-sm font-semibold py-2 rounded-lg transition"
          >
            {loading ? "Signing in…" : "Sign in"}
          </button>
        </form>
        <p className="text-[11px] text-neutral-400 mt-4">
          Don't have an account yet? Ask your admin to create one for you in Supabase (Authentication → Users) and add you in the Access Control tab.
        </p>
      </div>
    </div>
  );
}
