import { bootstrapApp } from "./app.js";

document.addEventListener("DOMContentLoaded", () => {
  const root = document.querySelector("#app");
  if (!root) {
    throw new Error("Elemento raiz #app não encontrado");
  }

  bootstrapApp(root);
});
