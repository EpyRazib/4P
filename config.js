/* ============================================================
   EWO dashboard settings

   This file is YOURS. Dashboard updates replace index.html but never
   this file, so anything you set here survives every update.

   Only the values you actually want to change need to be here.
   Anything you leave out falls back to the dashboard's defaults.
   ============================================================ */
window.EWO_CONFIG = {

  /* The Cloudflare Worker that starts the GitHub Action for you.
     With this set, "Rebuild from OneDrive" runs the whole rebuild in
     place and shows its progress in the status text. See
     rebuild-worker.js for how to create it. Leave "" to fall back to
     opening the Action page instead. */
  REBUILD_API: "https://4p.razib-hossain.workers.dev/",

  /* The Action page, used only when REBUILD_API is empty. */
  REBUILD_URL: "https://github.com/EpyRazib/4P/actions/workflows/refresh-data.yml",

  /* How often, in seconds, an open page checks for newly published data. */
  POLL_SECONDS: 120

  /* Rarely needed:
     DATA_URL:    "data.json",
     FLOW_URL:    "",     a Power Automate endpoint, if you ever use one
     FLOW_METHOD: "POST"
  */
};
