import { createRoot } from "react-dom/client";
import { App } from "./App.js";
import "./style.css";

function render(): void {
  createRoot(document.getElementById("root") ?? document.body).render(<App />);
}

if (typeof Office === "undefined") render();
else Office.onReady(render);
