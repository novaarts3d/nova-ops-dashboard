import React, { createContext, useContext, useEffect, useState } from "react";
import { supabase } from "../supabaseClient.js";

const AuthContext = createContext(null);

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error("useAuth must be used inside <AuthProvider>");
  return ctx;
}

export function AuthProvider({ children }) {
  const [session, setSession] = useState(undefined); // undefined = still checking
  const [permissions, setPermissions] = useState(null);
  const [permLoading, setPermLoading] = useState(false);
  const [permError, setPermError] = useState("");

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) {
      setPermissions(null);
      return;
    }
    setPermLoading(true);
    setPermError("");
    supabase
      .from("user_permissions")
      .select("*")
      .eq("email", session.user.email)
      .maybeSingle()
      .then(({ data, error }) => {
        if (error) {
          setPermError(error.message);
        } else if (!data) {
          // Logged in via Supabase Auth but nobody's added them to
          // user_permissions yet — block access with a clear message rather
          // than silently showing (or hiding) the whole app.
          setPermError(
            "Your account isn't set up with any dashboard access yet. Ask your admin to add you in the Access Control tab."
          );
        } else {
          setPermissions(data);
        }
        setPermLoading(false);
      });
  }, [session]);

  const signOut = () => supabase.auth.signOut();

  return (
    <AuthContext.Provider value={{ session, permissions, permLoading, permError, signOut }}>
      {children}
    </AuthContext.Provider>
  );
}
