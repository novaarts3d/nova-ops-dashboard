import React from "react";
import { useAuth } from "./AuthContext.jsx";
import Login from "./Login.jsx";

export default function AppGate({ children }) {
  const { session, permissions, permLoading, permError, signOut } = useAuth();

  if (session === undefined) {
    return <div className="min-h-screen flex items-center justify-center text-neutral-400 text-sm">Loading…</div>;
  }
  if (!session) return <Login />;

  if (permLoading) {
    return <div className="min-h-screen flex items-center justify-center text-neutral-400 text-sm">Checking your access…</div>;
  }

  if (permError) {
    return (
      <div className="min-h-screen flex items-center justify-center p-4">
        <div className="max-w-sm text-center">
          <p className="text-sm text-neutral-700 mb-4">{permError}</p>
          <button
            onClick={signOut}
            className="text-sm font-semibold text-red-600 hover:text-red-700"
          >
            Sign out
          </button>
        </div>
      </div>
    );
  }

  if (!permissions) return null;

  return children;
}
