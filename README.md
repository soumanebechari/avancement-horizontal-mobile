# Avancement Horizontal Mobile V16

Version V16 de l'application Kivy/Pydroid 3.

## Nouveautés
- Calcul des quotas séparément pour chaque catégorie professionnelle.
- Règle d'arrondi : 0,50 et plus est arrondi à l'entier supérieur.
- Répartition Rapide 50 %, Moyen 30 %, Lent 20 % pour chaque catégorie.
- Classement final : Catégorie → Rythme → Notation décroissante (puis ancienneté et âge en cas d'égalité).
- Export Excel avec deux feuilles : `Résultats promotion` et `Non bénéficiaires`.
- La feuille `Non bénéficiaires` indique chaque travailleur non retenu et le motif : note, ancienneté, sanction, données manquantes, etc.
- Les résultats de promotion restent archivés automatiquement dans la base locale JSON/historique.

## Exemple quota
Pour 6 candidats éligibles dans une catégorie :
- Rapide = 6 × 50 % = 3
- Moyen = 6 × 30 % = 1,8 → 2
- Lent = 6 × 20 % = 1,2 → 1

Total : 3 Rapide, 2 Moyen, 1 Lent.
