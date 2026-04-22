import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./styles.css";

function render(): void {
  const root = document.getElementById("root");
  if (!root) {
    return;
  }
  createRoot(root).render(<App />);
}

const office = (globalThis as any).Office;
if (office?.onReady) {
  office.onReady(() => render());
} else {
  render();
}
