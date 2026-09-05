import React, { useState } from "react";
import { supabase } from "../supabaseClient.js";

const Icon = ({ path, className = "" }) => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"
       strokeLinecap="round" strokeLinejoin="round" className={className}>
    {path}
  </svg>
);

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
      <div className="w-full max-w-3xl bg-white rounded-2xl shadow-sm overflow-hidden grid grid-cols-1 md:grid-cols-2">
        {/* Left brand panel */}
        <div className="bg-red-700 p-10 flex flex-col justify-between">
          <div>
            <div className="w-10 h-10 rounded-xl bg-white/15 flex items-center justify-center mb-6">
              <Icon className="w-5 h-5 text-white" path={<path d="M13 2 3 14h7l-1 8 10-12h-7l1-8z" />} />
            </div>
            <h1 className="text-white text-xl font-bold mb-2">Nova</h1>
            <p className="text-white/75 text-sm leading-relaxed">
              Operations dashboard for inventory, payroll, production and finance — all in one place.
            </p>
          </div>
          <div className="flex gap-2 flex-wrap">
            <div className="bg-white/15 rounded-lg p-2 flex items-center justify-center">
              <Icon className="w-[18px] h-[18px] text-white"
                path={<><path d="M21 8 12 3 3 8v8l9 5 9-5V8z" /><path d="M3 8l9 5 9-5" /><path d="M12 13v8" /></>} />
            </div>
            <div className="bg-white/15 rounded-lg p-2 flex items-center justify-center">
              <Icon className="w-[18px] h-[18px] text-white"
                path={<><circle cx="9" cy="8" r="3" /><path d="M2 20c0-3 3-5 7-5s7 2 7 5" /><circle cx="17" cy="8" r="2.5" /><path d="M16 15c2.5 0 5 1.5 5 5" /></>} />
            </div>
            <div className="bg-white/15 rounded-lg p-2 flex items-center justify-center">
              <Icon className="w-[18px] h-[18px] text-white"
                path={<path d="M4 20V10M10 20V4M16 20v-7M22 20H2" />} />
            </div>
            <div className="bg-white/15 rounded-lg p-2 flex items-center justify-center">
              <Icon className="w-[18px] h-[18px] text-white"
                path={<><rect x="1" y="7" width="13" height="9" rx="1" /><path d="M14 10h4l3 3v3h-7z" /><circle cx="5.5" cy="18" r="1.5" /><circle cx="17.5" cy="18" r="1.5" /></>} />
            </div>
          </div>
        </div>

        {/* Right form panel */}
        <div className="p-10 flex flex-col justify-center">
          <h2 className="text-xl font-bold text-neutral-900 mb-1">Sign in</h2>
          <p className="text-sm text-neutral-500 mb-6">Welcome back, enter your details.</p>

          <form onSubmit={handleSubmit} className="space-y-4">
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

            <div className="text-right -mt-2">
              <span className="text-xs text-red-700 cursor-pointer hover:underline">Forgot password?</span>
            </div>

            {error && <p className="text-xs text-red-600">{error}</p>}

            <button
              type="submit"
              disabled={loading}
              className="w-full bg-red-700 hover:bg-red-800 disabled:opacity-50 text-white text-sm font-semibold py-2.5 rounded-lg transition"
            >
              {loading ? "Signing in…" : "Sign in"}
            </button>
          </form>

          <p className="text-[11px] text-neutral-400 mt-4 leading-relaxed">
            Don't have an account? Ask your admin to add you in Access Control.
          </p>
        </div>
      </div>
    </div>
  );
}
