// Theme handling for Clear Enough To Lead.
// Loaded synchronously in <head> so the theme is applied before first paint.
(function () {
  "use strict";

  var STORAGE_KEY = "cetl-theme";

  function storedTheme() {
    try {
      var value = localStorage.getItem(STORAGE_KEY);
      return value === "light" || value === "dark" ? value : null;
    } catch (e) {
      return null;
    }
  }

  function systemTheme() {
    return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches
      ? "dark"
      : "light";
  }

  function applyTheme(theme) {
    document.documentElement.setAttribute("data-theme", theme);
  }

  applyTheme(storedTheme() || systemTheme());

  document.addEventListener("DOMContentLoaded", function () {
    var toggle = document.querySelector(".theme-toggle");
    if (!toggle) return;

    function refreshLabel() {
      var current = document.documentElement.getAttribute("data-theme");
      var next = current === "dark" ? "light" : "dark";
      toggle.textContent = current === "dark" ? "☀ Light" : "☾ Dark";
      toggle.setAttribute("aria-label", "Switch to " + next + " theme");
    }

    toggle.addEventListener("click", function () {
      var next = document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark";
      applyTheme(next);
      try {
        localStorage.setItem(STORAGE_KEY, next);
      } catch (e) {
        /* private browsing — theme still applies for this page view */
      }
      refreshLabel();
    });

    refreshLabel();
  });
})();
