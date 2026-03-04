## Automatisation et Modélisation actuarielle – Assurance Décès (Excel + VBA)

**Construction d'une table de mortalité, calcul de primes pures, estimation des provisions techniques, et évaluation SCR d’un contrat décès (Solvabilité II – Pilier 1).**  

Outil développé sous **Excel VBA** à partir des **données INSEE (TD 2025, Homme, France)**.

<div align="center">
  <img src="Graphs/bouton_prime_pure.PNG" alt="Bouton Prime Pure" width="40%" />
  <img src="Graphs/Provision_technique.PNG" alt="Provision et Solvabilité II " width="55%" />
</div>

---
### 🎯 Objectif du projet

L’assurance décès repose sur un risque intrinsèquement incertain (date du sinistre, coût).  
Ce projet vise à **automatiser la chaîne actuarielle complète** :

**Données INSEE → Table de mortalité → Prime pure → Best Estimate → Risk Margin → Provision S2 → SCR/MCR & ratios**

L’objectif est de proposer un modèle **structuré, paramétrable et auditable**, développée sous Excel avec automatisation VBA, réduisant les manipulations manuelles (risque d’erreur, lenteur, difficulté d’audit).

---
### **Contexte : Pourquoi ce projet ?**

En assurance décès, **l’assureur s’engage à verser un capital en cas de décès**, mais ne connaît ni la date du sinistre ni son coût réel. **La prime pure est au cœur du problème** : c’est le **coût technique du risque**, calculé comme la **valeur actuelle des engagements futurs** (capital × probabilité de décès à chaque âge), **sans marge ni frais**.

**Sans une tarification précise, deux risques majeurs apparaissent** :
- **Sous-tarification** → Les cotisations ne couvrent pas les sinistres (pertes financières).
- **Sur-tarification** → Les clients fuient vers la concurrence (perte de parts de marché).

Pour éviter cela, **les normes Solvabilité II et IFRS 17 imposent aux assureurs de calculer cette prime pure à partir d’une table de mortalité fiable**, tout en justifiant leurs hypothèses auprès des régulateurs.

**Problème** : Les outils actuels (Excel manuel, scripts non documentés) sont lents, sources d’erreurs et difficiles à auditer.

---



### Fonctionnalités clés

| Fonctionnalité | Description | Exemple d’utilisation (dans le projet) |
|---|---|---|
| **Import automatique INSEE (TD 2025)** | Import et consolidation des données **INSEE TD 2025** (CSV) dans la feuille `Données_Brutes` (âges 0–120), prêtes pour les calculs. | Charger la table **TD 2025 – Homme (France)** pour alimenter automatiquement `Données_Brutes`. |
| **Construction de la table de mortalité** | Génération de la feuille `Table_Mortalité` avec les colonnes actuarielles usuelles : \( q_x \), \( p_x \), \( l_x \), \( d_x \), \( L_x \), \( T_x \), \( e_x \). Remplissage via bouton/macro. | Cliquer sur **“Remplir les formules”** pour obtenir une table prête pour la tarification et le provisionnement. |
| **Contrôles qualité (QA) intégrés** | Tests simples et auditables : bornes \( 0 \le q_x \le 1 \), cohérence \( p_x = 1 - q_x \), \( d_x = l_x \cdot q_x \), chaînage \( l_{x+1} = l_x - d_x \), monotonicité de \( l_x \), etc. | Détecter rapidement un mauvais import (décimal/%), un décalage d’âge, des lignes manquantes ou une incohérence de formules. |
| **Tarification : calcul de prime pure** | Calcul de la **prime pure** d’un contrat décès via actualisation des flux probabilisés (hypothèse standard : décès en **milieu d’année**). Feuille `Prime_Pure` + tableau multi-âges + graphique. | Calculer la prime pour **30 ans**, **100 000 €**, **durée 30 ans**, **taux 2%**, puis visualiser l’évolution par âge (ex. 20–60 ans). |
| **Provisions techniques (vision Solvabilité II)** | Calcul du **Best Estimate (BE)**, de la **Risk Margin** (approche **Cost of Capital – CoC = 6%**), puis de la **Provision S2 = BE + RM** dans `Provisions_Techniques et Solvabilité`. | Sur le cas principal **(Homme 40 ans, 200 000 €, 20 ans, taux 2%)**, obtenir **BE / RM / Provision S2** à la souscription *(t = 0, Janvier 2026)*. |
| **Solvabilité II – Pilier 1 (SCR / MCR)** | Estimation du **SCR** (modules : mortalité, longévité, rachat + agrégation) et du **MCR**, puis calcul des **ratios de couverture** : \( \frac{FP}{SCR} \) et \( \frac{FP}{MCR} \). | Avec **FP = 25 000 €**, afficher **Ratio SCR** et **Ratio MCR** et conclure sur la conformité. |
| **Comparaison Décès / Vie / Mixte** | Paramètre **Type de contrat** (Décès / Vie / Mixte) : les cashflows projetés, le BE, la RM et le SCR s’adaptent au profil de garantie. | Comparer les résultats et montrer l’impact sur **provisions** et **solvabilité** selon le type de garantie. |
| **Visualisations automatiques** | Graphiques démographiques ( \( l_x \), \( q_x \), \( e_x \), \( d_x \) ) et solvabilité (décomposition SCR, ratios SCR/MCR) dans la feuille `Graphiques`. | Mettre à jour les graphiques après recalcul pour analyser rapidement la mortalité et les indicateurs de solvabilité. |
---
### **📂 Structure du projet**

```
MORTEX/
│
├── README.md                  # Documentation du projet (ce fichier)
├── MORTEX.xlsm                # Fichier Excel principal avec le code VBA
│
├── Documentation/             # Documents techniques et schémas
│   ├── Cahier_des_charges.pdf # Cahier des charges détaillé
│   └── Schema_Architecture.png # Schéma de l'architecture du projet
│
├── Data/                      # Données sources
│   └── Table_INSEE_source.csv # Données de mortalité INSEE
│
└── Captures/                  # Captures d'écran des résultats
    ├── Dashboard.png          # Exemple de dashboard interactif
    └── Exemple_Calcul.png     # Exemple de calcul de prime pure
```

---

## **🛠️ Guide d’utilisation**

### **1️⃣ Prérequis**
- **Excel 2016 ou supérieur** (avec macros activées).
- **Données INSEE** : Un fichier CSV contenant les taux de mortalité par âge (exemple fourni dans `Data/`).

### **2️⃣ Installation**
1. Téléchargez le dépôt GitHub (`git clone` ou téléchargement ZIP).
2. Ouvrez `MORTEX.xlsm` et activez les macros si demandé.

### **3️⃣ Utilisation pas à pas**
1. **Importer les données** :
   - Cliquez sur le bouton **"Importer données INSEE"** et sélectionnez `Table_INSEE_source.csv`.
   - *Exemple de format attendu* :
     ```
     Age;Taux_mortalite
     0;0.0005
     1;0.0003
     ...
     ```

2. **Générer la table de mortalité** :
   - Allez dans l’onglet **"Table de mortalité"**.
   - Sélectionnez les paramètres (ex : hypothèse d’amélioration de la longévité).
   - Cliquez sur **"Générer la table"**.

3. **Calculer la prime pure** :
   - Renseignez dans l’onglet **"Prime pure"** :
     - Âge d’entrée (ex : 30 ans).
     - Capital assuré (ex : 100 000 €).
     - Taux d’actualisation (ex : 2%).
   - Cliquez sur **"Calculer la prime"**.

4. **Visualiser les résultats** :
   - Consultez le **dashboard interactif** pour voir les courbes de mortalité et l’évolution des primes par âge.
   - Exportez les résultats en PDF ou CSV si besoin.

---

## **📸 Captures d’écran**

### **1️⃣ Dashboard interactif**
![Dashboard](Captures/Dashboard.png)
*Visualisation des probabilités de décès et des primes pures par âge.*

### **2️⃣ Exemple de calcul**
![Exemple de calcul](Captures/Exemple_Calcul.png)
*Calcul de la prime pure pour un capital de 100 000 € à 30 ans.*

---

## **🔍 Exemples concrets**

### **Cas 1 : Impact d’un changement de table de mortalité**
- **Scénario** : Comparer les primes pures calculées avec la table INSEE 2020 vs. TGH05-10.
- **Résultat** : La prime pour un assuré de 40 ans passe de **120 €/an** (INSEE) à **105 €/an** (TGH05-10), soit une **réduction de 12,5%** grâce à l’amélioration de la longévité.

### **Cas 2 : Sensibilité au taux d’actualisation**
- **Scénario** : Calculer la prime pour un capital de 50 000 € à 50 ans avec un taux de 1% vs. 3%.
- **Résultat** :
  - Taux à 1% → **Prime = 250 €/an**.
  - Taux à 3% → **Prime = 200 €/an** (baisse de 20% due à l’actualisation plus forte).

---

## **📌 Pourquoi utiliser ce projet ?**

| Public cible          | Bénéfices                                                                 |
|-----------------------|---------------------------------------------------------------------------|
| **Actuaires**         | Automatisation des calculs, réduction des erreurs, conformité réglementaire. |
| **Risk Managers**     | Analyse rapide des scénarios (choc de mortalité, taux d’actualisation).   |
| **Directions**        | Prise de décision éclairée sur la tarification et la compétitivité.       |
| **Étudiants/Enseignants** | Outil pédagogique pour comprendre les tables de mortalité et la tarification. |

---

## **🛡️ Gestion des risques**

| Risque                          | Mitigation                                                                 |
|---------------------------------|----------------------------------------------------------------------------|
| **Données erronées**            | Nettoyage automatique des données (vérification des formats, valeurs aberrantes). |
| **Erreurs de calcul**           | Tests unitaires pour valider les formules (ex : somme des probabilités = 1). |
| **Non-conformité réglementaire** | Documentation claire des hypothèses et méthodologies.                     |
| **Performance lente**           | Optimisation du code VBA (boucles, tableaux dynamiques).                  |

---

## **🤝 Contribuer**

Vous souhaitez améliorer ce projet ?
1. **Forkez** le dépôt (`git fork`).
2. **Créez une branche** (`git checkout -b feature/ma-fonctionnalite`).
3. **Commitez vos changements** (`git commit -m "Ajout de la fonction X"`).
4. **Pushez** (`git push origin feature/ma-fonctionnalite`).
5. **Ouvrez une Pull Request** pour revue.

**Idées d’améliorations** :
- Ajouter une interface utilisateur plus intuitive.
- Intégrer des données de mortalité internationales.
- Automatiser les tests avec un framework VBA.

---

## **📜 Licence**
Ce projet est sous licence **MIT** – libre d’utilisation, modification et distribution.

---

## **📧 Contact**
Pour toute question ou suggestion :
📩 [votre.email@example.com](jordanjatsa@gmail.com)
🔗 [LinkedIn](https://www.linkedin.com/in/jordan-jatsa-lekane/)

--
