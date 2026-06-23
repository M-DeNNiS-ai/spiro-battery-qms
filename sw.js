<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover" />

  <!-- ── PWA Meta Tags ── -->
  <meta name="application-name" content="Spiro Battery QMS" />
  <meta name="description" content="Spiro Uganda Battery Quality Management System" />
  <meta name="theme-color" content="#0D1B3E" />
  <meta name="background-color" content="#0D1B3E" />
  <meta name="mobile-web-app-capable" content="yes" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent" />
  <meta name="apple-mobile-web-app-title" content="SpiroQMS" />

  <!-- ── iOS Splash / Icons ── -->
  <link rel="apple-touch-icon" href="/icons/icon-192x192.png" />
  <link rel="apple-touch-icon" sizes="152x152" href="/icons/icon-152x152.png" />
  <link rel="apple-touch-icon" sizes="144x144" href="/icons/icon-144x144.png" />
  <link rel="apple-touch-icon" sizes="128x128" href="/icons/icon-128x128.png" />

  <!-- ── PWA Manifest ── -->
  <link rel="manifest" href="/manifest.json" />
  <link rel="icon" type="image/png" href="/favicon.png" />

  <title>Spiro Battery QMS</title>

  <!-- ── Fonts (cached by SW after first load) ── -->
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet" />

  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    html, body {
      height: 100%;
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      background: #0D1B3E;
      color: #1A1F36;
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
      overscroll-behavior: none;
    }

    #root {
      height: 100%;
      min-height: 100vh;
      min-height: 100dvh; /* dynamic viewport height — handles mobile browser bars */
    }

    /* Safe area insets for notched phones (iPhone X+) */
    #root > div {
      padding-left: env(safe-area-inset-left);
      padding-right: env(safe-area-inset-right);
    }

    /* PWA Install Banner */
    #install-banner {
      display: none;
      position: fixed;
      bottom: 80px;
      left: 12px;
      right: 12px;
      background: #1565C0;
      color: white;
      border-radius: 12px;
      padding: 12px 14px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.3);
      z-index: 9999;
      font-size: 13px;
      align-items: center;
      gap: 10px;
    }
    #install-banner.visible { display: flex; }
    #install-banner .msg { flex: 1; line-height: 1.4; }
    #install-banner .msg strong { display: block; font-size: 14px; }
    #install-banner button {
      padding: 7px 14px;
      border-radius: 8px;
      border: none;
      font-family: inherit;
      font-weight: 700;
      font-size: 13px;
      cursor: pointer;
      white-space: nowrap;
    }
    #install-btn { background: white; color: #1565C0; }
    #dismiss-btn { background: rgba(255,255,255,0.2); color: white; }

    /* Offline indicator */
    #offline-bar {
      display: none;
      position: fixed;
      top: 0; left: 0; right: 0;
      background: #E65100;
      color: white;
      text-align: center;
      font-size: 12px;
      font-weight: 600;
      padding: 4px 8px;
      z-index: 10000;
    }
    #offline-bar.visible { display: block; }

    /* Splash loader shown before React mounts */
    #splash {
      position: fixed;
      inset: 0;
      background: #0D1B3E;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      z-index: 9998;
      transition: opacity 0.4s ease;
    }
    #splash.hidden { opacity: 0; pointer-events: none; }
    #splash .logo { font-size: 56px; margin-bottom: 14px; }
    #splash h1 { color: white; font-size: 22px; font-weight: 800; letter-spacing: 0.5px; }
    #splash p { color: rgba(255,255,255,0.5); font-size: 13px; margin-top: 6px; }
    #splash .bar {
      width: 120px; height: 3px;
      background: rgba(255,255,255,0.15);
      border-radius: 2px;
      margin-top: 24px;
      overflow: hidden;
    }
    #splash .bar-fill {
      height: 100%;
      background: #42A5F5;
      border-radius: 2px;
      animation: load 1.2s ease forwards;
    }
    @keyframes load { from { width: 0; } to { width: 100%; } }

    /* Prevent text selection on UI elements */
    button, select, label { user-select: none; -webkit-user-select: none; }

    /* Smooth scrolling */
    * { scroll-behavior: smooth; }

    /* Hide scrollbars on mobile */
    ::-webkit-scrollbar { width: 0; height: 0; }
  </style>
</head>
<body>

  <!-- Offline indicator -->
  <div id="offline-bar">⚠️ No internet connection — data saves locally and will sync when reconnected</div>

  <!-- Splash screen -->
  <div id="splash">
    <div class="logo">⚡</div>
    <h1>Spiro Battery QMS</h1>
    <p>Battery Quality Management System</p>
    <div class="bar"><div class="bar-fill"></div></div>
  </div>

  <!-- PWA Install Banner (Android) -->
  <div id="install-banner">
    <div class="msg">
      <strong>Install Spiro QMS</strong>
      Add to your home screen for offline access
    </div>
    <button id="install-btn">Install</button>
    <button id="dismiss-btn">✕</button>
  </div>

  <!-- React App Mount Point -->
  <div id="root"></div>

  <!-- React + ReactDOM from CDN (cached offline by SW) -->
  <script crossorigin src="https://unpkg.com/react@18/umd/react.production.min.js"></script>
  <script crossorigin src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>

  <!-- Babel standalone (transpiles JSX in browser) -->
  <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>

  <!-- SheetJS for Excel export -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

  <!-- Load the app -->
  <script type="text/babel" src="/app.jsx" data-type="module"></script>

  <script>
    // ── Hide splash when app renders ──────────────────────────────────────────
    const splash = document.getElementById("splash");
    const observer = new MutationObserver(() => {
      if (document.getElementById("root").children.length > 0) {
        splash.classList.add("hidden");
        setTimeout(() => splash.remove(), 500);
        observer.disconnect();
      }
    });
    observer.observe(document.getElementById("root"), { childList: true, subtree: true });
    // Fallback: hide splash after 4s regardless
    setTimeout(() => { splash?.classList.add("hidden"); setTimeout(() => splash?.remove(), 500); }, 4000);

    // ── Register Service Worker ───────────────────────────────────────────────
    if ("serviceWorker" in navigator) {
      window.addEventListener("load", async () => {
        try {
          const reg = await navigator.serviceWorker.register("/sw.js", { scope: "/" });
          console.log("[App] Service Worker registered:", reg.scope);

          // Listen for SW updates
          reg.addEventListener("updatefound", () => {
            const newSW = reg.installing;
            newSW.addEventListener("statechange", () => {
              if (newSW.state === "installed" && navigator.serviceWorker.controller) {
                // New version available — show update prompt
                if (confirm("⚡ Spiro QMS has been updated! Reload to get the latest version?")) {
                  newSW.postMessage({ type: "SKIP_WAITING" });
                  window.location.reload();
                }
              }
            });
          });

          // Background sync registration (for failed OneDrive uploads)
          if ("sync" in reg) {
            window.registerBackgroundSync = async () => {
              try { await reg.sync.register("spiro-onedrive-sync"); }
              catch (e) { console.log("[App] Background sync not available:", e); }
            };
          }

          // Listen for SW messages (retry sync)
          navigator.serviceWorker.addEventListener("message", (event) => {
            if (event.data?.type === "RETRY_SYNC") {
              window.dispatchEvent(new CustomEvent("spiro-retry-sync"));
            }
          });

        } catch (e) {
          console.log("[App] Service Worker registration failed:", e);
        }
      });
    }

    // ── PWA Install Prompt (Android Chrome) ──────────────────────────────────
    let deferredInstallPrompt = null;
    const installBanner = document.getElementById("install-banner");
    const installBtn    = document.getElementById("install-btn");
    const dismissBtn    = document.getElementById("dismiss-btn");

    window.addEventListener("beforeinstallprompt", (e) => {
      e.preventDefault();
      deferredInstallPrompt = e;
      // Show banner after 3 seconds (give user time to see the app first)
      setTimeout(() => {
        if (!localStorage.getItem("spiro_pwa_dismissed")) {
          installBanner.classList.add("visible");
        }
      }, 3000);
    });

    installBtn.addEventListener("click", async () => {
      installBanner.classList.remove("visible");
      if (deferredInstallPrompt) {
        deferredInstallPrompt.prompt();
        const { outcome } = await deferredInstallPrompt.userChoice;
        console.log("[App] PWA install outcome:", outcome);
        deferredInstallPrompt = null;
      }
    });

    dismissBtn.addEventListener("click", () => {
      installBanner.classList.remove("visible");
      localStorage.setItem("spiro_pwa_dismissed", "1");
    });

    window.addEventListener("appinstalled", () => {
      console.log("[App] PWA installed successfully ✓");
      installBanner.classList.remove("visible");
    });

    // ── Offline / Online detection ────────────────────────────────────────────
    const offlineBar = document.getElementById("offline-bar");
    function updateOnlineStatus() {
      if (!navigator.onLine) {
        offlineBar.classList.add("visible");
      } else {
        offlineBar.classList.remove("visible");
        // Trigger sync retry when back online
        window.dispatchEvent(new CustomEvent("spiro-retry-sync"));
      }
    }
    window.addEventListener("online",  updateOnlineStatus);
    window.addEventListener("offline", updateOnlineStatus);
    updateOnlineStatus();

    // ── iOS "Add to Home Screen" hint ────────────────────────────────────────
    const isIOS = /iphone|ipad|ipod/i.test(navigator.userAgent);
    const isInStandalone = window.navigator.standalone;
    if (isIOS && !isInStandalone && !localStorage.getItem("spiro_ios_hint")) {
      setTimeout(() => {
        const hint = document.createElement("div");
        hint.style.cssText = "position:fixed;bottom:80px;left:12px;right:12px;background:#0D1B3E;color:#fff;border-radius:12px;padding:14px;z-index:9999;font-size:13px;box-shadow:0 4px 20px rgba(0,0,0,0.4);line-height:1.5;";
        hint.innerHTML = `<strong style="display:block;margin-bottom:4px">📱 Install on iPhone</strong>Tap <strong>Share ↑</strong> then <strong>"Add to Home Screen"</strong> to install Spiro QMS as an app.<button onclick="this.parentElement.remove();localStorage.setItem('spiro_ios_hint','1')" style="float:right;background:rgba(255,255,255,0.2);border:none;color:#fff;padding:4px 10px;border-radius:6px;font-family:inherit;cursor:pointer;font-size:12px;">Got it</button>`;
        document.body.appendChild(hint);
      }, 4000);
    }
  </script>

</body>
</html>
