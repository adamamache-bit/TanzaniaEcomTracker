import React, { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "./index.css";
import App from "./App.jsx";

class AppErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, message: "" };
  }

  static getDerivedStateFromError(error) {
    return {
      hasError: true,
      message: error instanceof Error ? error.message : "Unknown application error",
    };
  }

  componentDidCatch(error, errorInfo) {
    try {
      console.error("TanzaniaEcomTracker runtime error", error, errorInfo);
      window.localStorage.setItem(
        "tanzaniaecomtracker:last-runtime-error",
        JSON.stringify({
          at: new Date().toISOString(),
          message: error instanceof Error ? error.message : "Unknown application error",
          stack: error instanceof Error ? error.stack || "" : "",
          componentStack: errorInfo?.componentStack || "",
        })
      );
    } catch {
      // ignore logging issues
    }
  }

  handleReload = () => {
    window.location.reload();
  };

  handleClearLocalCache = () => {
    try {
      const keysToRemove = [
        "tanzania-ecom-tracker-v16",
        "tanzania-ecom-tracker-auto-backup-v1",
        "tanzania-ecom-tracker-auto-backup-meta-v1",
        "tanzania-ecom-tracker-import-meta-v1",
        "tanzaniaecomtracker:last-runtime-error",
      ];
      keysToRemove.forEach((key) => window.localStorage.removeItem(key));
    } catch {
      // ignore
    }
    window.location.reload();
  };

  render() {
    if (this.state.hasError) {
      return (
        <div
          style={{
            minHeight: "100vh",
            display: "grid",
            placeItems: "center",
            padding: 24,
            background: "#f5f8ff",
            color: "#1f2a44",
            fontFamily: "Inter, Segoe UI, sans-serif",
          }}
        >
          <div
            style={{
              width: "100%",
              maxWidth: 720,
              background: "#ffffff",
              border: "1px solid #d7e3ff",
              borderRadius: 24,
              boxShadow: "0 20px 60px rgba(28, 63, 141, 0.10)",
              padding: 28,
            }}
          >
            <div style={{ fontSize: 12, fontWeight: 800, letterSpacing: "0.08em", textTransform: "uppercase", color: "#2f61e3" }}>
              Runtime recovery
            </div>
            <div style={{ fontSize: 32, fontWeight: 900, marginTop: 12 }}>The app hit an unexpected error</div>
            <div style={{ marginTop: 12, lineHeight: 1.7, color: "#5e6c8d" }}>
              The page did not disappear silently anymore. You can reload the app, or clear only the local browser cache if the last imported
              data corrupted the current session.
            </div>
            <div
              style={{
                marginTop: 18,
                padding: 16,
                borderRadius: 18,
                background: "#f8fbff",
                border: "1px solid #dfe9ff",
                color: "#22304d",
                fontSize: 14,
                wordBreak: "break-word",
              }}
            >
              {this.state.message || "Unknown application error"}
            </div>
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginTop: 22 }}>
              <button
                type="button"
                onClick={this.handleReload}
                style={{
                  border: "none",
                  borderRadius: 14,
                  background: "#264fb8",
                  color: "#fff",
                  fontWeight: 800,
                  padding: "14px 18px",
                  cursor: "pointer",
                }}
              >
                Reload app
              </button>
              <button
                type="button"
                onClick={this.handleClearLocalCache}
                style={{
                  borderRadius: 14,
                  border: "1px solid #d7e3ff",
                  background: "#fff",
                  color: "#1f2a44",
                  fontWeight: 800,
                  padding: "14px 18px",
                  cursor: "pointer",
                }}
              >
                Clear local cache and reload
              </button>
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <AppErrorBoundary>
      <App />
    </AppErrorBoundary>
  </StrictMode>
);
