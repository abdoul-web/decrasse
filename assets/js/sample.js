/* Décrasse — configuration et fichiers d'exemple (données fictives). */
(function (D) {
  'use strict';
  D.config = {
    // Lecteur .xls de secours, essayé dans l'ordre. Déposez xlsx.full.min.js dans assets/vendor/ pour ne dépendre d'aucun CDN.
    sheetjsSources: [
      'assets/vendor/xlsx.full.min.js',
      'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
      'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    ],
    previewRows: 300,
    maxFileMB: 80,
  };

  D.sample = [
    {
      name: 'ventes_janvier.csv',
      text: `Rapport des ventes — janvier 2026;;;;;;;
Exporté le 01/02/2026;;;;;;;
Date;Nom client;Courriel;Tél.;Ville;Article;Qté;Montant ($)
2026-01-03;Boulangerie Lavoie;Info@BoulangerieLavoie.ca ;819.555.0142;Trois-Rivières;Pneus hiver;4;1 250,00 $
05/01/2026; Dépanneur Chez Ti-Jean ;tijean.depanneur@gmail.com;(819) 555-0177;trois-rivieres;Changement d'huile;1;89,99
2026-01-07;Toitures Bergeron;contact@toituresbergeron.com;418 555 0190;Québec;Freins avant;2;430,50 $
;;;;;;;
12 janv. 2026;Boulangerie Lavoie;info@boulangerielavoie.ca;8195550142;TROIS-RIVIÈRES ;Alignement;1;119,95
2026-01-14;Clinique Vet Mauricie;accueil@vetmauricie.ca;819-555-0133;Shawinigan;Batterie;1;214,00 $
2026-01-14;Clinique Vet Mauricie;accueil@vetmauricie.ca;819-555-0133;Shawinigan;Batterie;1;214,00 $
16/01/2026;Transport Gagnon;dispatch@transportgagnon;1-418-555-0111;quebec;Pneus hiver;8;2 480,00 $
2026-01-19;Fleuriste Rose & Co;bonjour@roseetco.ca;514 555 0108;Montréal;Changement d'huile;2;179,98
2026-01-21;Dépanneur Chez Ti-Jean;tijean.depanneur@gmail.com;819 555-0177;Trois-Rivieres;Freins avant;1;N/A
22 janvier 2026;Toitures Bergeron;contact@toituresbergeron.com;418-555-0190 poste 12;Québec ;Changement d'huile;3;269,97 $
2026-01-26;Transport Gagnon;dispatch@transportgagnon.ca;418.555.0111;Québec;Alignement;4;479,80
28/01/2026;Fleuriste Rose & Co;BONJOUR@ROSEETCO.CA;514-555-0108;montreal;Batterie;1;199,50 $
2026-01-30;Boulangerie Lavoie;info@boulangerielavoie.ca;819 555-0142;Trois-Rivières;freins avant;2;  415,00
TOTAL;;;;;;;6 849,04 $`,
    },
    {
      name: 'sales_feb_export.csv',
      text: `date,client,email,phone,city,product,quantity,total
2026-02-02,Transport Gagnon,dispatch@transportgagnon.ca,418-555-0111,Québec,Pneus hiver,4,"$1,240.00"
2026-02-04,Boulangerie  Lavoie,info@boulangerielavoie.ca,819-555-0142,Trois-Rivières,Changement d'huile,1,89.99
2026/02/06,Clinique Vet Mauricie,accueil@vetmauricie.ca,8195550133,shawinigan,Alignement,1,119.95
2026-02-09,Toitures Bergeron,contact@toituresbergeron.com,418-555-0190,Quebec,Batterie,2,428.00
2026-02-11,Fleuriste Rose & Co,bonjour@roseetco.ca,514-555-0108,Montréal,Freins avant,1,215.25
2026-02-11,Fleuriste Rose & Co,bonjour@roseetco.ca,514-555-0108,Montréal,Freins avant,1,215.25
13/02/2026,Garage Mécanique Plus,service@mecaniqueplus.ca,819 555 0164,Trois-Rivières,Pneus hiver,4,"1 310,00 $"
2026-02-16,Dépanneur Chez Ti-Jean,tijean.depanneur@gmail.com,819-555-0177,Trois-Rivières,Batterie,1,
2026-02-18,Transport Gagnon,dispatch@transportgagnon.ca,418-555-0111,QUÉBEC,Alignement,6,719.70
2026-02-20,Clinique Vet Mauricie,accueil@vetmauricie.ca,819-555-0133,Shawinigan,Changement d'huile,2,179.98
,,,,,,,
2026-02-23,Toitures Bergeron,contact@toituresbergeron.com,555-0190,Québec,Pneus hiver,4,"$1,295.00"
2026-02-25,Boulangerie Lavoie,info@boulangerielavoie.ca,819-555-0142,trois-rivières,Freins avant,1,207.50
2026-02-27,Garage Mécanique Plus,service@mecaniqueplus.ca,819-555-0164,Trois-Rivières,changement d'huile,3,269.97`,
    },
  ];
})(globalThis.Decrasse || (globalThis.Decrasse = {}));
