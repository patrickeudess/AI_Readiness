# 🌍 AfriData Ready

**From Raw Data to FAIR & AI-Ready Data**

Application web 100 % locale qui accompagne l'utilisateur, comme un assistant intelligent, pour transformer des données brutes (tous domaines : recherche, santé, agriculture, éducation, développement) en un jeu de données propre, documenté, gouverné, conforme aux principes FAIR et prêt pour l'intelligence artificielle.

**➡️ Application en ligne : https://patrickeudess.github.io/AfriData-Ready/**

## Points clés

- **100 % local** : tout le traitement se fait dans le navigateur, aucune donnée n'est envoyée sur Internet, aucun serveur.
- **Assistant AfriData** : à chaque étape, l'assistant explique les problèmes et propose des corrections automatiques ou manuelles, avec le gain réel sur l'AQI.
- **AfriData Quality Index (AQI)** : indice de qualité sur 100, inspiré de standards internationaux (évaluation interne, non officielle).

## Parcours en 8 étapes

1. **Importation** — CSV, Excel (multi-feuilles) ou JSON, avec diagnostic initial et AQI initial
2. **Qualité des données** — diagnostic (manquants, doublons, valeurs aberrantes, colonnes vides) et corrections, dont la correction des valeurs aberrantes en un clic
3. **Documentation** — description de chaque variable, avec suggestions automatiques
4. **Métadonnées** — titre, auteur, licence, mots-clés, DOI, etc.
5. **FAIR** — chaque critère non atteint est rendu actionnable
6. **AI Readiness** — préparation des données pour l'apprentissage automatique
7. **Gouvernance** — sécurité, pseudonymisation (RGPD), provenance, data lineage, versionnement
8. **Tableau de bord AQI** — score global, radar, jauge, progression et visualisations

Puis **export du « Data Product »** : données corrigées et originales, dictionnaire, métadonnées (JSON, YAML, DataCite XML), rapports qualité/FAIR/AI/AQI, recommandations, journal des modifications, README, citation, rapport Word et certificat AQI imprimable.

## Deux modes

- **🚀 Démarrage rapide** : importer, traiter, exporter, sans sauvegarde imposée.
- **📂 Projet AfriData Ready** : sauvegarde automatique locale (IndexedDB), reprise exacte plus tard, export/import de projets (fichiers `.afridata`).

## Utilisation

Aucune installation : ouvrez simplement l'application dans un navigateur récent.

Pour l'exécuter localement, il suffit de servir le dossier en statique, par exemple :

```bash
python -m http.server 8000
# puis ouvrir http://localhost:8000
```

## Technologies

HTML / CSS / JavaScript, sans dépendance externe (bibliothèques embarquées localement : PapaParse, SheetJS, JSZip, FileSaver). Graphiques en SVG pur.

## Références

- Wilkinson et al. (2016) : FAIR Principles
- ISO 8000 / ISO 25012 : qualité et modèle de qualité des données
- ISO/IEC 11179, Dublin Core, DataCite, DCAT : métadonnées et interopérabilité
- Wang & Strong (1996) : Data Quality — Sambasivan et al. (2021) : Data Cascades in AI — Andrew Ng : Data-Centric AI

---
*Mémoire DU Données : UCAD/EBAD : Patrick-Eudess Zatty ALLA : 2026*
