# Fichator 3000

Logiciel Windows qui génère les **fiches mise en place & banquets** de Chez le Père Gras :
on remplit le formulaire, la fiche Word sort avec la mise en page de la trame, déjà
renommée (`JJ MM AAAA MR/MME NOM Entreprise ou événement.docx`) et rangée dans le dossier réseau.

Une idée TL Studio.

---

## Installer sur un PC du resto (une seule fois par poste)

1. Ouvre la page **Releases** du dépôt : https://github.com/6bz2tt5244-crypto/fichator/releases/latest
2. Télécharge `Fichator-Setup-x.y.z.exe` et lance-le.
3. Windows affiche « Windows a protégé votre ordinateur » (le logiciel n'est pas signé) :
   clique **Informations complémentaires** → **Exécuter quand même**.
4. L'installation est automatique (pas besoin de droits admin). Une icône apparaît sur le bureau.
5. Au premier enregistrement, choisis le **dossier réseau des fiches**. C'est mémorisé.

Ensuite, plus rien à faire : les mises à jour arrivent toutes seules (voir plus bas).

## Utilisation

- **Enregistrer la fiche** (ou `Ctrl+S`) : écrit le Word directement dans le dossier des fiches.
- **Enregistrer sous…** : pour la ranger ailleurs (dossier d'un client par exemple).
- Après l'enregistrement : boutons **Ouvrir dans Word** et **Voir le dossier**.
- Dans le menu, une ligne entre étoiles passe en rouge gras : `*besoin de place pour une activité*`.
- Le formulaire se garde tout seul en brouillon. **Nouvelle fiche** le vide.

## Développer (sur Mac ou PC)

```bash
npm install
npm start          # lance le logiciel en mode dev
```

L'interface est dans `src/` (`index.html`, `app.js`). Elle marche aussi en ouvrant
simplement `src/index.html` dans Chrome (mode navigateur, sans enregistrement direct).

## Publier une mise à jour

1. Ajoute les nouveautés en haut de `src/changelog.js` (avec le prochain numéro de version).
2. Commit tes modifs, puis :

```bash
npm run release        # 1.0.0 → 1.0.1 (petite modif)
npm run release:minor  # 1.0.1 → 1.1.0 (nouvelle fonction)
```

GitHub compile l'installateur Windows (~5 min, onglet **Actions**) et publie la version.
Sur chaque PC, le Fichator la télécharge en fond, affiche un bandeau
« La version x.y.z est prête — Redémarrer et mettre à jour », puis montre les nouveautés.

## Structure

```
main.js                       fenêtre, écriture des fichiers, mises à jour auto
preload.js                    pont sécurisé interface ↔ Windows
src/index.html                interface
src/app.js                    formulaire, aperçu, génération du Word (.docx)
src/changelog.js              nouveautés affichées à l'équipe
build/icon.ico                icône
.github/workflows/release.yml compilation Windows + publication
```

Les fiches clients ne passent jamais par GitHub : elles restent sur le réseau du resto.
