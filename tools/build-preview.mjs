// Construit une version monofichier de Décrasse (CSS et JS intégrés) dans dist/decrasse-preview.html.
// Utile pour une démo hors ligne ou un aperçu partagé. Le site réel n'a besoin d'aucune étape de build.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';

const root = new URL('../', import.meta.url);
const read = (p) => readFileSync(new URL(p, root), 'utf8');
let html = read('index.html');

html = html.replace(/<link rel="stylesheet" href="assets\/css\/app.css">/, () => `<style>\n${read('assets/css/app.css')}\n</style>`);
html = html.replace(/<script src="(assets\/js\/[\w-]+\.js)"><\/script>/g, (_, p) => `<script>\n${read(p).replace(/<\/script/gi, '<\\/script')}\n</script>`);
html = html.replace(/<link rel="(icon|apple-touch-icon|manifest|canonical)"[^>]*>\n?/g, '');
html = html.replace(/href="confidentialite\.html"/g, 'href="#faq"');

const fragment = process.argv.includes('--fragment');
if (fragment) {
  // Format « fragment » : sans doctype, html, head ni body (pour un hébergeur qui fournit déjà le squelette)
  const head = html.match(/<head>([\s\S]*?)<\/head>/)[1].replace(/<meta charset="utf-8">\n?|<meta name="viewport"[^>]*>\n?/g, '');
  const body = html.match(/<body>([\s\S]*?)<\/body>/)[1];
  html = head + body;
}
mkdirSync(new URL('dist/', root), { recursive: true });
const out = new URL(fragment ? 'dist/decrasse-fragment.html' : 'dist/decrasse-preview.html', root);
writeFileSync(out, html);
console.log('Écrit :', out.pathname, Math.round(html.length / 1024) + ' Ko');
