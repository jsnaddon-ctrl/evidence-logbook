// ===== Evidence Logbook settings =====
// Fill these in during setup (see SETUP.md). They are not secret.
window.APP_CONFIG = {
  // Microsoft app registration "Application (client) ID"
  clientId: "PASTE-CLIENT-ID-HERE",

  // Your Cloudflare Worker address, e.g. "https://evidence-ai.yourname.workers.dev"
  workerUrl: "PASTE-WORKER-URL-HERE",

  // Folder created in each apprentice's OneDrive
  rootFolder: "Apprenticeship Evidence",

  // Leave blank unless sign-in complains about the redirect address.
  // Must match the Redirect URI in the Microsoft app registration exactly.
  redirectUri: ""
};
