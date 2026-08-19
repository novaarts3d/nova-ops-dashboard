import React from "react";
import ReactDOM from "react-dom/client";
import NovaOps from "./App.jsx";
import { AuthProvider } from "./auth/AuthContext.jsx";
import AppGate from "./auth/AppGate.jsx";
import "./index.css";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <AuthProvider>
      <AppGate>
        <NovaOps />
      </AppGate>
    </AuthProvider>
  </React.StrictMode>
);
