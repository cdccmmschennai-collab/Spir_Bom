/* Shared appearance (light/dark) handling for every SPIR-BOM Tool page.
   Loaded as the very first thing in <head>, before any CSS, so the
   `data-theme` attribute is already set on <html> by the time styles are
   applied -- no flash of the wrong theme. The preference lives in
   localStorage (this browser, this origin), so it persists across page
   navigation and across sessions without needing a server round-trip. */
(function () {
  var KEY = 'bomtool-theme';

  function apply(theme) {
    if (theme === 'dark') document.documentElement.setAttribute('data-theme', 'dark');
    else document.documentElement.removeAttribute('data-theme');
  }

  function get() {
    try { return localStorage.getItem(KEY) || 'light'; } catch (e) { return 'light'; }
  }

  function set(theme) {
    try { localStorage.setItem(KEY, theme); } catch (e) { /* private mode etc -- still apply for this load */ }
    apply(theme);
  }

  apply(get());
  window.bomToolTheme = { get: get, set: set };
})();
