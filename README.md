# 🍷 Projet : Modernisation et Optimisation du Pipeline d'Analyse des Ventes (BottleNeck)

## 📌 1. Contexte & Objectifs du Projet
Ce projet s'inscrit dans le cadre de la modernisation d'un pipeline d'analyse de données e-commerce (vin et spiritueux) existant (Projet 6). L'objectif est de refactorisé pour gagner en robustesse, en performances et en maintenabilité en combinant **Pandas**, **DuckDB** et des visualisations sous **Matplotlib/Seaborn**.

### Problématique Métier
Comment identifier efficacement les produits moteurs du chiffre d'affaires, repérer les anomalies (marges négatives, prix négatif, quantité négative, écart produits ERP...) et optimiser la gestion du fonds de roulement à travers l'analyse des stocks dormants ?

---

## 🛠️ 2. Choix Technologiques & Veille (Comparatif)
Pour répondre aux exigences de performance et de qualité, deux options majeures ont été testées :

| Critère | Option A : Pandas pur | Option B : Pandas + DuckDB (Retenue) |
| :--- | :--- | :--- |
| **Performance (Volumétrie)** | Ralentissements sur les gros volumes et jointures complexes | **Excellente** (moteur SQL vectorisé ultra-rapide) |
| **Lisibilité du code** | Multiplicité des méthodes chaînées verbeuses (`groupby`, `merge`) | **Clarté accrue** grâce à des requêtes SQL standardisées |
| **Maintenabilité** | Difficile à auditer pour des non-data scientists | **Forte** (le SQL est universel pour les équipes métiers) |
| **Décision** | Écarté pour les jointures lourdes | **Validé et adopté** pour l'ensemble du pipeline |

---

## 🧾 3. Livrables
- Notebook modifié
- Rapport santé financière (Export PDF, Excel)
- Dataframe global (export Excel)
- Dataframe marge positive (Export Excel)

---

## 📂 4. Structure du Répertoire
```text
projet-bottleneck/
│
├── data/                       # Données brutes sources
├── Graphiques/                 # Exports graphiques (PNG Haute Résolution)
│   ├── Quantites_Pareto_Comparatif.png
│   ├── rotation_stock.png
│   ├── marge_cat_produit.png
│   └── corr_stock_ventes_prix.png
├── Rapport/                    # Livrables finaux pour le management
│   ├── dataframe_final_bottleneck.xlsx
│   └── marge_positive_bottleneck.xlsx
├── analyse_vente_bottleneck_IA.ipynb  # Notebook principal modernisé
└── README.md                   # Documentation du projet

---
## 🚀 5. Guide d'Exécution & Reproductibilité
Prérequis techniques : Python 3.10+
Bibliothèques requises : pandas, duckdb, matplotlib, seaborn, numpy, openpyxl

---
## 👀 6. Instructions d'installation
1. Cloner le dépôt ou placer le notebook dans votre environnement de travail (JupyterLab).
2. Installer les dépendances si nécessaire :
# pip install pandas matplotlib seaborn numpy openpyxl --> 
# pip install pandera duckdb --> pour créer des schémas et des requetes sql
# pip install reportlab xlsxwriter pillow kaleido --> pour exporter le rapport en PDF et Excel
3. Ouvrir et exécuter séquentiellement les cellules du notebook analyse_vente_bottleneck_IA.ipynb



