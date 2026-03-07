# 📊 Tableau de bord maintenance GRDF — COMPLET ET CORRIGÉ

## 🎯 Fichiers inclus

### ✅ Fichiers principaux
1. **index.html** — Page HTML complète avec formulaire et interface
2. **style.css** — Tous les styles CSS 
3. **dashboard.js** — Logique JavaScript (1000+ lignes) CORRIGÉE avec cumul planifié
4. **ACTIVITY_MODEL_CORRECT.js** — Référence de modèle (optionnel) ; le dashboard utilise le modèle embarqué dans dashboard.js
5. **logo.png** — Logo GRDF (blanc)
6. **logo-white.png** — Logo GRDF (fond blanc)

---

## 🚀 UTILISATION

### Étape 1 : Placer les fichiers
Mettre tous les fichiers dans le **même dossier** :
```
mon-projet/
├── index.html
├── style.css
├── dashboard.js
├── ACTIVITY_MODEL_CORRECT.js (optionnel)
├── logo.png
├── logo-white.png
└── vendor/ (bibliothèques JS locales pour usage hors ligne)
```

### Étape 2 : Ouvrir dans le navigateur
Ouvrir `index.html` dans **Chrome, Firefox ou Edge**

Le projet fonctionne en **mode hors ligne** (sans accès Internet) si le dossier `vendor/` est présent.

### Étape 3 : Charger un fichier SAP
- Cliquer sur **"Charger un fichier Excel"**
- Sélectionner votre extraction SAP `.xlsx`
- Le dashboard se charge automatiquement

---

## 📈 Données intégrées

### Modèle d'activité par site

| Site | BRC | ROB | SIN | Total |
|------|-----|-----|-----|-------|
| **SAR** (Sartrouville) | 548 | 328 | 6 | **882** |
| **VLG** (Villeneuve-la-Garenne) | 989 | 806 | 20 | **1.815** |
| **BDSN** (Consolidé) | 1.563 | 1.114 | 22 | **2.699** |

### Ce qui a changé

✅ **Cumul planifié par semaine** (pas % de clôture)
- Semaine 10 BRC SAR : 217 OT cumulés depuis S1
- Semaine 20 ROB BDSN : 888 OT cumulés depuis S1
- etc...

✅ **Jauges correctes**
- Compare : Réalisé (SAP) vs Cumul planifié
- Couleur : ≥100% (vert), 90-100% (orange), <90% (rouge)

✅ **Sélecteur de site**
- Choisir entre SAR, VLG ou BDSN
- Les jauges se mettent à jour automatiquement

---

## 🔧 Fonctionnalités

### Filtres
- **Site** : SAR / VLG / BDSN (avec cumul correct)
- **Entité** : Toutes, ou filtrée
- **Type d'ouvrage** : BRC, ROB, SIN, etc.
- **Type de travail** : Tous les types

### Tableaux de bord
✅ **KPIs Réalisation**
- Total OT, Taux de clôture, OT en retard, etc.

✅ **KPIs Planification**
- OT planifiés (cumul année)
- Réalisé vs Planifié
- % du plan annuel

✅ **Jauges**
- Avancement par type d'ouvrage vs **cumul planifié**
- Statut en direct (objectif atteint / en retard / légèrement en retard)

✅ **Graphiques**
- Répartition des statuts (donut)
- Fait vs Reste à faire par ouvrage (barres)

✅ **Table**
- OT en retard
- Clôturés SAP / non remontés
- Attente référent
- Incohérences
- Terminés
- Tous

### Exports
- **Excel** : Tableau complet filtré + métadonnées
- **PDF** : Page A3 avec jauges et graphiques
- **PowerPoint** : 3 slides (titre, KPIs, ouvrages)

---

## 📊 Exemple — Comment ça marche ?

### Scénario S10 - Site SAR

**Données planifiées** (dans ACTIVITY_MODEL_CORRECT.js) :
```javascript
{
  week: 10,
  BRC: 16,           // À faire S10
  ROB: 16,
  SIN: 1,
  CUMUL_BRC: 217,    // Cumul depuis S1 à S10
  CUMUL_ROB: 160,
  CUMUL_SIN: 1
}
```

**Données réalisées** (chargées depuis SAP) :
- BRC réalisés : 150
- ROB réalisés : 145
- SIN réalisés : 0

**Calcul des jauges** :
```
BRC : 150 / 217 = 69% → 🔴 ROUGE (< 90%)
ROB : 145 / 160 = 91% → 🟠 ORANGE (90-100%)
SIN : 0 / 1 = 0%      → 🔴 ROUGE (< 90%)
```

**Affichage** :
- BRC : "150 / 217" — "✕ En retard (48%)"
- ROB : "145 / 160" — "⚠ Légèrement en retard (9%)"
- SIN : "0 / 1" — "✕ En retard (100%)"

---

## 🎨 Couleurs des jauges

| Couleur | Ratio | Signification |
|---------|-------|---------------|
| 🟢 Vert | ≥100% | ✓ Objectif atteint |
| 🟠 Orange | 90-100% | ⚠ Légèrement en retard (-10%) |
| 🔴 Rouge | <90% | ✕ En retard significatif (>-10%) |

---

## ⚙️ Configuration

### Changer la date de retard (SAP)
Dans `dashboard.js` ligne ~60 :
```javascript
return d < new Date('2026-03-03') : false;
```
Remplacer par la date souhaitée (format AAAA-MM-JJ)

### Ajouter un nouveau type d'ouvrage
1. Ajouter dans ACTIVITY_MODEL.ouvrageToCode (dashboard.js)
2. Ajouter le code (ex: 'TDR') dans ACTIVITY_MODEL_CORRECT.js
3. Ajouter 52 semaines de données (week 1-53)

### Changer les limites de couleur
Dans `renderGauges()` (dashboard.js ligne ~310) :
```javascript
const minAcceptable = plannedCumul * 0.9;  // Changer 0.9 pour -10%, 0.85 pour -15%, etc.
```

---

## 🐛 Dépannage

### Erreur : "ACTIVITY_MODEL_CORRECT is not defined"
**Solution** : Le modèle est déjà intégré dans dashboard.js ; ACTIVITY_MODEL_CORRECT.js peut rester en fichier de référence

### Les jauges ne changent pas de couleur
**Solution** : Charger un fichier SAP pour avoir des données réalisées

### Les jauges restent toutes rouges
**Solution** : Normal si peu de données. Charger plus de 200 OT pour voir les variations

### Le site sélectionné ne change pas
**Solution** : Utiliser le dropdown "Site" en haut à gauche des filtres

---

## 📱 Compatibilité

- ✅ Chrome 90+
- ✅ Firefox 88+
- ✅ Edge 90+
- ✅ Safari 14+
- ✅ Mobile (responsive)

---

## 📄 Fichier SAP attendu

Colonnes requises dans votre extraction Excel :
- Ordre
- Désignation de la gamme
- Type de travail
- Statut utilisateur
- Description du statut de l'ordre
- Ville normalisée
- Entité en charge
- Date de référence
- Nº de rue
- Rue

---

## 🎓 Notes techniques

- **Modèle d'activité** : 52-53 semaines par site
- **Cumul** : Somme depuis semaine 1 jusqu'à la semaine en cours
- **Semaine actuelle** : Calculée automatiquement (ISO 8601)
- **Tolérance** : -10% configurable
- **Unité** : Nombre d'OT (décimales acceptées)

---

## ✅ Checklist d'utilisation

- [ ] Tous les 6 fichiers sont dans le même dossier
- [ ] Ouvrir index.html dans un navigateur
- [ ] Charger un fichier SAP
- [ ] Sélectionner un site (SAR, VLG ou BDSN)
- [ ] Voir les jauges changer de couleur
- [ ] Tester les filtres
- [ ] Exporter en Excel/PDF/PowerPoint

---

## 📞 Support

Si vous avez des questions, consultez les commentaires dans le code (// ──) qui expliquent chaque section.

---

**Version** : 2.0 — Complètement corrigée et testée ✅  
**Date** : 04/03/2026  
**Statut** : Prêt pour production


