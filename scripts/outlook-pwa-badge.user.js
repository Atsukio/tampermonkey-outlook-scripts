// ==UserScript==
// @name         Outlook PWA: unread badge + reload after sleep
// @namespace    local.outlook-pwa
// @version      1.5
// @description  Shows the Inbox unread count on the PWA taskbar icon and reloads Outlook after the laptop wakes from sleep.
// @icon         data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA2NCA2NCI+PHJlY3QgeD0iMyIgeT0iMTAiIHdpZHRoPSI1MCIgaGVpZ2h0PSIzOCIgcng9IjciIGZpbGw9IiMzQjZGRTAiLz48cGF0aCBkPSJNOCAxN2wyMCAxNSAyMC0xNSIgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjZmZmIiBzdHJva2Utd2lkdGg9IjQuNSIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIi8+PGNpcmNsZSBjeD0iNDkiIGN5PSI0NiIgcj0iMTMiIGZpbGw9IiNFNTQ4NEQiIHN0cm9rZT0iI2ZmZiIgc3Ryb2tlLXdpZHRoPSIzLjUiLz48dGV4dCB4PSI0OSIgeT0iNTIiIGZvbnQtc2l6ZT0iMTciIGZvbnQtZmFtaWx5PSJTZWdvZSBVSSxBcmlhbCxzYW5zLXNlcmlmIiBmb250LXdlaWdodD0iNzAwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmaWxsPSIjZmZmIj4zPC90ZXh0Pjwvc3ZnPg==
// @match        https://outlook.office.com/*
// @match        https://outlook.office365.com/*
// @match        https://outlook.cloud.microsoft/*
// @match        https://outlook.live.com/*
// @homepageURL  https://github.com/Atsukio/tampermonkey-outlook-scripts
// @supportURL   https://github.com/Atsukio/tampermonkey-outlook-scripts/issues
// @downloadURL  https://raw.githubusercontent.com/Atsukio/tampermonkey-outlook-scripts/main/scripts/outlook-pwa-badge.user.js
// @updateURL    https://raw.githubusercontent.com/Atsukio/tampermonkey-outlook-scripts/main/scripts/outlook-pwa-badge.user.js
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(() => {
  'use strict';

  const FOLDER = 'Inbox';               // change if your Outlook UI isn't in English
  const POLL_MS = 10 * 1000;            // how often to refresh the badge
  const SLEEP_GAP_MS = 5 * 60 * 1000;   // a timer gap this long means the machine slept
  const DEBUG = true;                   // set to false once the badge works

  if (DEBUG) console.log('[Outlook badge] v1.5 started on', location.href);

  const canBadge = 'setAppBadge' in navigator;
  if (!canBadge) console.warn('[Outlook badge] Badging API not available in this window.');

  // Keep our own handles to the badge API, then stop Outlook from clearing the badge
  // itself (it appears to clear it when the window regains focus).
  const setBadge = canBadge ? navigator.setAppBadge.bind(navigator) : null;
  const clearBadge = canBadge ? navigator.clearAppBadge.bind(navigator) : null;
  if (canBadge) {
    try {
      navigator.setAppBadge = () => Promise.resolve();
      navigator.clearAppBadge = () => Promise.resolve();
    } catch (err) {
      console.warn('[Outlook badge] could not block Outlook badge calls:', err);
    }
  }

  // Reads the unread count from the Inbox entry in the folder pane.
  // Returns null when the folder pane isn't visible (e.g. on Calendar), so the badge is left as is.
  function getUnread() {
    for (const el of document.querySelectorAll('[role="treeitem"]')) {
      const aria = (el.getAttribute('aria-label') || '').trim();
      const title = (el.getAttribute('title') || '').trim();
      const text = (el.textContent || '').replace(/\s+/g, ' ').trim();
      const name = aria || title || text;
      if (!name.toLowerCase().startsWith(FOLDER.toLowerCase())) continue;

      const all = `${aria} | ${title} | ${text}`;

      // Prefer an explicit "N unread" anywhere in the labels
      let m = all.match(/(\d+)\s+unread/i);
      if (m) return parseInt(m[1], 10);

      // Otherwise take the first number after the folder name in the visible text
      m = text.slice(FOLDER.length).match(/\d+/);
      return m ? parseInt(m[0], 10) : 0;
    }
    return null;
  }

  // Re-applies the badge on every check, so it comes back even if something clears it.
  let lastCount = -1;
  function updateBadge() {
    if (!canBadge) return;
    const n = getUnread();
    if (n === null) return;
    if (DEBUG && n !== lastCount) console.log('[Outlook badge] unread count is now', n);
    lastCount = n;
    (n > 0 ? setBadge(n) : clearBadge())
      .catch(err => console.warn('[Outlook badge]', err));
  }

  // Restore the badge shortly after the window regains focus, after Outlook's own handling runs
  const reapplySoon = () => setTimeout(updateBadge, 1000);
  window.addEventListener('focus', reapplySoon);
  document.addEventListener('visibilitychange', reapplySoon);

  // Update as soon as Outlook changes a folder's count (its title attribute), instead of
  // waiting for the next poll. Debounced so a burst of changes triggers one update.
  let pending = null;
  new MutationObserver(() => {
    if (pending) return;
    pending = setTimeout(() => { pending = null; updateBadge(); }, 250);
  }).observe(document.body, { subtree: true, attributes: true, attributeFilter: ['title'] });

  let lastTick = Date.now();
  setInterval(() => {
    const now = Date.now();
    if (now - lastTick > SLEEP_GAP_MS) {
      location.reload(); // woke from sleep: reconnect to pick up missed mail
      return;
    }
    lastTick = now;
    updateBadge();
  }, POLL_MS);

  updateBadge();
})();
