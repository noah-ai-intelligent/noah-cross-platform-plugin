import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";

class ErrorBoundary extends React.Component<{children: React.ReactNode}, {hasError: boolean, error: any}> {
  constructor(props: any) { super(props); this.state = { hasError: false, error: null }; }
  static getDerivedStateFromError(error: any) { return { hasError: true, error }; }
  render() {
    if (this.state.hasError) {
      return <div style={{padding: 20, color: 'red'}}>
        <h3>App Crashed</h3>
        <pre>{String(this.state.error?.stack || this.state.error)}</pre>
      </div>;
    }
    return this.props.children;
  }
}

const renderApp = () => {
  createRoot(document.getElementById("root")!).render(
    <React.StrictMode>
      <ErrorBoundary>
        <App />
      </ErrorBoundary>
    </React.StrictMode>
  );
};

let appRendered = false;
const safeRenderApp = () => {
  if (appRendered) return;
  appRendered = true;
  renderApp();
};

if (window.Office) {
  Office.onReady(() => safeRenderApp());
  setTimeout(safeRenderApp, 1000);
} else {
  safeRenderApp();
}
