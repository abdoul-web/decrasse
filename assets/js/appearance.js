/* Préférence d'apparence uniquement ; aucun fichier utilisateur n'est conservé. */
(function () {
  'use strict';
  const root = document.documentElement;
  const button = document.getElementById('themeToggle');
  const system = window.matchMedia('(prefers-color-scheme: dark)');
  let preference = null;
  try { preference = localStorage.getItem('decrasse-theme'); } catch { /* Navigation privée ou stockage indisponible. */ }
  if (!['light', 'dark'].includes(preference)) preference = null;
  const apply = (theme) => {
    root.dataset.theme = theme;
    if (button) {
      button.setAttribute('aria-pressed', String(theme === 'dark'));
      button.setAttribute('aria-label', theme === 'dark' ? 'Passer au thème clair' : 'Passer au thème sombre');
    }
    document.querySelector('meta[name="theme-color"]')?.setAttribute('content', theme === 'dark' ? '#141A12' : '#F6F5EE');
  };
  apply(preference || (system.matches ? 'dark' : 'light'));
  button?.addEventListener('click', () => {
    preference = root.dataset.theme === 'dark' ? 'light' : 'dark';
    apply(preference);
    try { localStorage.setItem('decrasse-theme', preference); } catch { /* Le thème fonctionne aussi sans stockage. */ }
  });
  system.addEventListener('change', (e) => { if (!preference) apply(e.matches ? 'dark' : 'light'); });
})();
