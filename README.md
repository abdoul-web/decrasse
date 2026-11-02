# Décrasse

Nettoyage automatique de fichiers Excel et CSV, entièrement dans le navigateur.
Dates, montants, doublons, orthographe, courriels, téléphones et codes postaux sont corrigés.
Plusieurs fichiers sont fusionnés et le résultat est rendu sous forme de classeur Excel avec la liste des corrections.

Aucun serveur, aucune base de données, aucune étape de build : le site est un ensemble de fichiers statiques.
L'hébergement est donc gratuit.

## Structure

```
index.html              Page d'accueil et outil
confidentialite.html    Politique de confidentialité
404.html                Page introuvable
assets/css/app.css      Styles (thème clair et sombre)
assets/js/zip.js        Lecture et écriture ZIP (conteneur des .xlsx)
assets/js/io.js         Lecture CSV / XLSX (maison) / XLS (SheetJS, chargé au besoin)
assets/js/xlsx-write.js Écriture de classeurs .xlsx typés (dates, montants, filtres)
assets/js/engine.js     Moteur de nettoyage (fonctions pures, testées)
assets/js/sample.js     Configuration et fichiers d'exemple
assets/js/app.js        Interface
tests/                  Tests du moteur et du format Excel (node --test)
tools/build-preview.mjs Version monofichier pour une démo hors ligne
_headers                En-têtes de sécurité (Cloudflare Pages, Netlify)
```

## Essayer en local

```bash
npm start        # serveur de production sans dépendance (http://localhost:5173, ou $PORT)
npm test         # 12 tests, aucune dépendance à installer
```

Ou, sans Node : `python3 -m http.server 5173`, puis ouvrir http://localhost:5173.

## Mettre en ligne sur Render (gratuit)

1. Pousser ce dossier dans un dépôt GitHub.
2. Sur https://dashboard.render.com : **New › Blueprint**, choisir le dépôt. Render lit `render.yaml`
   (site statique, aucun build, en-têtes de sécurité) et crée le service `decrasse`.
3. Le site est en ligne à `https://decrasse.onrender.com` (ou un nom proche). Chaque `git push` redéploie.

Les sites statiques de Render sont gratuits et ne se mettent pas en veille.

## Mettre en ligne gratuitement (Cloudflare Pages)

1. Créer un dépôt GitHub (par exemple `decrasse`) et y pousser ce dossier :
   ```bash
   git init && git add . && git commit -m "Décrasse v1"
   git branch -M main
   git remote add origin https://github.com/<votre-compte>/decrasse.git
   git push -u origin main
   ```
2. Sur https://dash.cloudflare.com : Workers & Pages › Créer › Pages › Connecter à Git, choisir le dépôt.
3. Paramètres de build : **aucun préréglage**, commande de build **vide**, répertoire de sortie **`/`**.
4. Déployer. Le site est en ligne à `https://decrasse.pages.dev` (ou un nom proche), avec HTTPS.
   Chaque `git push` redéploie automatiquement.

Autres options gratuites : Netlify (glisser-déposer le dossier sur app.netlify.com/drop) ou GitHub Pages
(Settings › Pages › branche `main`, dossier `/`). GitHub Pages ignore le fichier `_headers`.

## À faire avant le lancement public

- Remplacer `https://decrasse.ca` par l'adresse réelle dans `index.html` (balises canonical et og), `robots.txt` et `sitemap.xml`.
- Remplacer `contact@decrasse.ca` dans `confidentialite.html` par une adresse qui existe.
- Nom de domaine (facultatif, environ 15 $ par an pour un .ca) : l'ajouter dans Cloudflare Pages › Domaines personnalisés.
- Pour ne dépendre d'aucun CDN : télécharger `xlsx.full.min.js` (SheetJS 0.20.3, https://cdn.sheetjs.com) dans `assets/vendor/`.
  Il n'est utilisé que pour les vieux fichiers .xls. Les polices peuvent aussi être hébergées localement.
- Mesure d'audience sans témoins, si souhaitée : Cloudflare Web Analytics (gratuit), à ajouter dans la politique de confidentialité.

## Fonctionnement du moteur

1. **Structure** : repère la vraie ligne d'en-têtes (ignore les titres), retire lignes vides, totaux et en-têtes répétés.
2. **En-têtes** : synonymes français et anglais ramenés à un nom commun (`Qté`, `quantity` → `Quantité`), ce qui permet de fusionner des fichiers différents.
3. **Types** : chaque colonne est classée (date, montant, nombre, pourcentage, courriel, téléphone, code postal, oui/non, texte).
   Les codes avec zéros initiaux restent du texte. Le sens JJ/MM ou MM/JJ est déduit de la colonne entière.
4. **Cellules** : conversion selon le type ; les valeurs illisibles sont retirées et listées, les contacts douteux sont signalés sans être effacés.
5. **Orthographe** : les variantes d'une même valeur (casse, accents, espaces, traits d'union) prennent la forme la plus fréquente.
6. **Doublons** et **colonnes vides** retirés, puis journal complet des corrections.
