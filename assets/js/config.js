import { state } from "./state.js";

export async function loadConfig() {
  const response = await fetch("/.netlify/functions/config");
  if (!response.ok) {
    throw new Error(await response.text());
  }
  state.config = await response.json();
  return state.config;
}
