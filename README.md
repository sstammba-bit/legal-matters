# Générateur du modèle de mémoire juridique

Ce dossier contient le code source qui génère `..\Modele_Memoire_Juridique.dotx`.
Il n'a pas besoin d'être touché pour simplement utiliser le modèle — voir
"Utilisation du modèle" ci-dessous. Il sert uniquement si vous voulez
modifier la structure du modèle (styles, numérotation, en-tête, etc.).

## Utilisation du modèle (l'essentiel)

1. Ouvrez `Modele_Memoire_Juridique.dotx` depuis Word, ou passez par
   **Fichier → Nouveau → Personnel** (le fichier doit se trouver dans votre
   dossier de modèles Word pour apparaître ici — sinon, double-cliquez
   simplement sur le `.dotx`, Word crée un nouveau document basé dessus).
2. **Supprimez le chapitre d'exemple** (tout le contenu marqué "EXEMPLE —"),
   qui sert uniquement à démontrer chaque style.
3. Rédigez avec les styles définis dans la galerie Word : `Titre 1/2/3` pour
   les chapitres, `Allégation`, `Preuve`, `Pièce`, `Citation juridique`,
   `Observation`, `Conclusions`.
4. Mettez à jour la table des matières par un clic droit dessus → **Mettre à
   jour les champs**, ou en sélectionnant tout (Ctrl+A) puis **F9**.
5. Le nom du dossier dans l'en-tête (`[Nom du dossier]`) se modifie par un
   double-clic dans l'en-tête.

### Renvois (Insertion → Renvoi)

- Renvoi vers un **titre** (chapitre/sous-chapitre) : fonctionne nativement,
  aucune préparation nécessaire.
- Renvoi vers un **élément numéroté** (ex. "cf. allégué 3") : fonctionne
  nativement sur toutes les allégations, sans signet, car elles appartiennent
  à une liste numérotée dédiée.
- Renvoi vers le **texte ou la page** d'une allégation ou d'un passage
  précis : posez d'abord un signet dessus (**Insertion → Signet**), puis
  utilisez Renvoi → Signet ou Renvoi → Page du signet. Le script ne peut pas
  deviner à l'avance quels passages de vos futurs mémoires vous voudrez
  cibler — ce geste reste manuel dans Word, comme pour n'importe quel
  document Word.

## Régénérer ou modifier le modèle

Le modèle est produit par un script Node.js (librairie `docx`), pas édité
à la main dans Word. Structure :

```
src/
  units.js            conversions cm/pt -> twips
  numbering.js         3 numérotations multiniveaux indépendantes
  styles.js             tous les styles de paragraphe personnalisés
  header-footer.js      en-tête + pied de page "Page X sur Y"
  toc.js                  table des matières native
  example-chapter.js    chapitre de démonstration + bookmarks + conclusions
  pieces-table.js       chapitre "Liste des pièces"
  annexes.js             chapitre "Annexes"
  build-document.js     assemble le Document final
convert-to-dotx.js       post-traitement : transforme le .docx généré en .dotx
generate-template.js     point d'entrée
```

### ⚠️ Important : n'installez PAS les dépendances directement ici

Ce dépôt (`G:\My Drive\20-IT-Dev\legal-matters`) est synchronisé par Google
Drive. `npm install` y échoue (Google Drive verrouille les fichiers pendant
que npm écrit les milliers de petits fichiers de `node_modules`, ce qui
produit des erreurs `EBADF`/`EPERM`). Suivre le pattern des autres projets
(retirement-sim, swiss-tax) : le code est versionné ici sur Drive, mais
s'exécute depuis une copie locale hors Drive.

1. Copie/clone de travail : `C:\Dev\legal-matters` (hors Google Drive).
   ```
   cd C:\Dev\legal-matters
   npm install
   node generate-template.js "G:\My Drive\05. - Admin\09. Legal matters\Modele_Memoire_Juridique.dotx"
   ```
2. Le script réécrit directement le `.dotx` dans le dossier Legal matters
   (un seul fichier réécrit, ce qui ne pose pas de problème avec Google
   Drive — c'est uniquement `npm install`, avec ses milliers de petits
   fichiers, qui échoue).
3. Après avoir modifié le code dans `C:\Dev\legal-matters`, reporter les
   changements dans ce dépôt (`G:\My Drive\20-IT-Dev\legal-matters`) puis
   commit/push — c'est le dépôt Drive qui fait foi pour l'historique git.

### Pourquoi `.dotx` et pas nativement supporté par la librairie `docx`

La librairie `docx` ne génère que des `.docx`. `.docx` et `.dotx` ne
diffèrent, au niveau du format OOXML, que par un seul content-type déclaré
dans `[Content_Types].xml` (document vs. modèle) et l'extension de fichier.
`convert-to-dotx.js` génère un `.docx` normalement puis corrige cette seule
ligne avant d'écrire le fichier en `.dotx`.
