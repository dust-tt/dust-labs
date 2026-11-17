/* global Office */

// Full-screen notice inviting users of this legacy add-in to install the new
// Dust add-in from Microsoft Marketplace. It is shown when the task pane opens,
// unless the user chose to keep using this add-in during the last 24 hours.
(function () {
  const NEW_ADDIN_URL =
    "https://marketplace.microsoft.com/en-us/product/WA200013110";
  const NEW_ADDIN_MANIFEST_URL = "https://addins.dust.tt/office/manifest.xml";
  const SNOOZE_DURATION_MS = 24 * 60 * 60 * 1000;

  function storageKey(prefix) {
    return `${prefix}_migrationNoticeDismissedAt`;
  }

  function wasDismissedRecently(prefix) {
    try {
      const dismissedAt = Number(localStorage.getItem(storageKey(prefix)));
      return (
        Number.isFinite(dismissedAt) &&
        dismissedAt > 0 &&
        Date.now() - dismissedAt < SNOOZE_DURATION_MS
      );
    } catch (e) {
      return false;
    }
  }

  function rememberDismissal(prefix) {
    try {
      localStorage.setItem(storageKey(prefix), String(Date.now()));
    } catch (e) {
      // Storage unavailable: the notice will simply show again next time.
    }
  }

  // openBrowserWindow opens the system browser from desktop Office, where a
  // plain link would otherwise navigate inside the task pane.
  function openInBrowser(event) {
    if (
      typeof Office !== "undefined" &&
      Office.context &&
      Office.context.ui &&
      Office.context.ui.openBrowserWindow
    ) {
      event.preventDefault();
      Office.context.ui.openBrowserWindow(event.currentTarget.href);
    }
  }

  const DUST_LOGO = `
    <svg width="96" height="24" viewBox="0 0 96 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-label="Dust">
      <path d="M84 0H72V24H84V0Z" fill="#FE9C1A" />
      <path d="M12 24C18.6274 24 24 18.6274 24 12C24 5.37258 18.6274 0 12 0C5.37258 0 0 5.37258 0 12C0 18.6274 5.37258 24 12 24Z" fill="#E2F78C" />
      <path d="M36 24C42.6274 24 48 18.6274 48 12C48 5.37258 42.6274 0 36 0C29.3726 0 24 5.37258 24 12C24 18.6274 29.3726 24 36 24Z" fill="#FFC3DF" />
      <path d="M12 0H0V24H12V0Z" fill="#418B5C" />
      <path d="M48 0H24V12H48V0Z" fill="#E14322" />
      <path fill-rule="evenodd" clip-rule="evenodd" d="M60 12C56.6863 12 54 9.31371 54 6C54 2.68629 56.6863 0 60 0H96V12H60Z" fill="#3B82F6" />
      <path fill-rule="evenodd" clip-rule="evenodd" d="M48 24V12H60C63.3137 12 66 14.6863 66 18C66 21.3137 63.3137 24 60 24H48Z" fill="#9FDBFF" />
    </svg>`;

  const CHECK_ICON = `
    <svg class="mn-check" width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M20 6L9 17L4 12" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" />
    </svg>`;

  function render(appName, onContinue, prefix) {
    const notice = document.createElement("div");
    notice.className = "mn-overlay";
    notice.setAttribute("role", "dialog");
    notice.setAttribute("aria-modal", "true");
    notice.setAttribute("aria-labelledby", "mn-title");
    notice.innerHTML = `
      <div class="mn-content">
        <div class="mn-logo">${DUST_LOGO}</div>
        <span class="mn-chip">New</span>
        <h1 id="mn-title" class="mn-title">Dust in ${appName}, reimagined</h1>
        <p class="mn-subtitle">
          A new Dust add-in is here. Faster to set up and more powerful,
          it brings your agents right next to your work.
        </p>
        <ul class="mn-benefits">
          <li>${CHECK_ICON}<span>Sign in with your Dust account, no API keys to copy</span></li>
          <li>${CHECK_ICON}<span>Chat with your agents without leaving ${appName}</span></li>
          <li>${CHECK_ICON}<span>Always up to date via Microsoft Marketplace</span></li>
        </ul>
        <div class="mn-actions">
          <a class="mn-btn mn-btn-primary" href="${NEW_ADDIN_URL}" target="_blank" rel="noopener noreferrer">
            Get the new add-in
          </a>
          <button type="button" class="mn-btn mn-btn-outline">
            Continue with this version
          </button>
        </div>
      </div>
      <p class="mn-footnote">
        Can't access Microsoft Marketplace?
        <a class="mn-link" href="${NEW_ADDIN_MANIFEST_URL}" target="_blank" rel="noopener noreferrer">Download the manifest</a>
      </p>`;

    notice.querySelectorAll("a").forEach(function (link) {
      link.addEventListener("click", openInBrowser);
    });
    notice
      .querySelector(".mn-btn-outline")
      .addEventListener("click", function () {
        rememberDismissal(prefix);
        notice.remove();
        document.documentElement.classList.remove("mn-open");
        onContinue();
      });

    // The add-in's page underneath is taller than the pane; lock its scroll so
    // only the notice can scroll.
    document.documentElement.classList.add("mn-open");
    document.body.appendChild(notice);
  }

  // Shows the notice if needed, then calls onContinue once the user chooses to
  // keep this add-in (or immediately if they did so in the last 24 hours).
  window.showMigrationNoticeOrContinue = function ({
    appName,
    storagePrefix,
    onContinue,
  }) {
    if (wasDismissedRecently(storagePrefix)) {
      onContinue();
      return;
    }
    render(appName, onContinue, storagePrefix);
  };
})();
