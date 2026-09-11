import React, { useState, useEffect } from "react";
import { supabase } from "../supabaseClient.js";

// Route this component at /reset-password (must match redirectTo in Login.jsx
// AND be added to Supabase → Authentication → URL Configuration → Redirect URLs)
export default function ResetPassword() {
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [success, setSuccess] = useState(false);
  const [sessionReady, setSessionReady] = useState(false);

  // When the user lands here from the emailed link, Supabase's client
  // automatically detects the recovery token in the URL and creates a
  // temporary session. We just wait for that event before showing the form.
  useEffect(() => {
    const { data: listener } = supabase.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY" || event === "SIGNED_IN") {
        setSessionReady(true);
      }
    });

    // Fallback in case the event already fired before this component mounted
    supabase.auth.getSession().then(({ data }) => {
      if (data?.session) setSessionReady(true);
    });

    return () => listener?.subscription?.unsubscribe();
  }, []);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");

    if (password.length < 8) {
      setError("Password must be at least 8 characters.");
      return;
    }
    if (password !== confirmPassword) {
      setError("Passwords do not match.");
      return;
    }

    setLoading(true);
    const { error: updateError } = await supabase.auth.updateUser({ password });
    setLoading(false);

    if (updateError) {
      setError(updateError.message);
    } else {
      setSuccess(true);
    }
  };

  return (
    <div className="min-h-screen bg-neutral-50 flex items-center justify-center p-4">
      <div className="w-full max-w-md bg-white rounded-2xl shadow-sm overflow-hidden p-10">
        <h2 className="text-xl font-bold text-neutral-900 mb-1">Set a new password</h2>

        {!sessionReady && !success && (
          <p className="text-sm text-neutral-500 mt-4">
            Verifying your reset link…
          </p>
        )}

        {sessionReady && !success && (
          <>
            <p className="text-sm text-neutral-500 mb-6">
              Choose a new password for your account.
            </p>
            <form onSubmit={handleSubmit} className="space-y-4">
              <label className="flex flex-col gap-1 text-sm">
                <span className="text-neutral-600 font-medium">New password</span>
                <input
                  type="password"
                  required
                  autoFocus
                  className="border border-neutral-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-red-500 focus:border-red-500"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                />
              </label>

              <label className="flex flex-col gap-1 text-sm">
                <span className="text-neutral-600 font-medium">Confirm password</span>
                <input
                  type="password"
                  required
                  className="border border-neutral-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-red-500 focus:border-red-500"
                  value={confirmPassword}
                  onChange={(e) => setConfirmPassword(e.target.value)}
                />
              </label>

              {error && <p className="text-xs text-red-600">{error}</p>}

              <button
                type="submit"
                disabled={loading}
                className="w-full bg-red-700 hover:bg-red-800 disabled:opacity-50 text-white text-sm font-semibold py-2.5 rounded-lg transition"
              >
                {loading ? "Updating…" : "Update password"}
              </button>
            </form>
          </>
        )}

        {success && (
          <div className="mt-4">
            <p className="text-sm text-green-700 mb-4">
              Your password has been updated. You can now sign in.
            </p>
            <a
              href="/login"
              className="block text-center w-full bg-red-700 hover:bg-red-800 text-white text-sm font-semibold py-2.5 rounded-lg transition"
            >
              Go to sign in
            </a>
          </div>
        )}
      </div>
    </div>
  );
}
