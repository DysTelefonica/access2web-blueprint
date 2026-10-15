// allow-hosts.js — fragmento portable (DysTelefonica/team-skills#528).
// Helper Node equivalente a la fixture `allow_network` del fragmento
// `pytest-socket`. Por defecto la suite corre con la red cerrada
// (`nock.disableNetConnect()`); este módulo expone `allowHosts` para
// los tests que necesiten abrir hosts concretos durante su ejecución.
// Plantilla copy-paste: la auditoría HR-28 sigue siendo el gate
// `check_env_isolation.py`; este módulo es la materialización en el
// lado del test runner.

"use strict";

function loadNock() {
  try {
    return require("nock");
  } catch (err) {
    return null;
  }
}

function parseEnvHosts() {
  const raw = (process.env.NOCK_ALLOW_HOSTS || "").trim();
  if (!raw) return [];
  return raw.split(",").map((h) => h.trim()).filter(Boolean);
}

function setupGlobalAllowlist() {
  const nock = loadNock();
  if (!nock) return;
  nock.disableNetConnect();
  for (const host of parseEnvHosts()) {
    nock.enableNetConnect(host);
  }
}

function allowHosts(...hosts) {
  const nock = loadNock();
  if (!nock) {
    return function restore() { /* noop: nock not installed */ };
  }
  const flat = hosts
    .flatMap((h) => (Array.isArray(h) ? h : [h]))
    .map((h) => String(h).trim())
    .filter(Boolean);
  nock.enableNetConnect(flat.join(","));
  return function restore() { nock.disableNetConnect(); };
}

module.exports = { allowHosts, setupGlobalAllowlist, parseEnvHosts };