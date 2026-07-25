# Système de Paie & RH RDC — Morning SARL

## Contexte

Produit Excel générique commercialisable (paie + RH conforme au droit congolais)
construit pour Morning SARL (filiale de Lobii Group SARL, Kinshasa). Client de
référence pendant la construction : STELLAIRE SARLU (entreprise de construction,
30+ agents, mais le fichier de travail actuel n'a qu'1 agent d'exemple, Jean
MUKENDI, MSOC-0001).

**Fichier principal** : `Morning_SARL_Systeme_Paie_RH_RDC.xlsx` — à placer dans
ce dossier de projet (télécharger depuis la conversation Claude.ai d'origine,
ce fichier ne se transfère pas automatiquement).

Construit entièrement en **formules Excel pures** (pas de VBA) via Python/
openpyxl, pour rester ouvrable sans macro dans n'importe quel Excel/LibreOffice.

## Architecture — 8 feuilles, dans cet ordre de dépendance

1. **PARAMETRES** — paramètres légaux RDC (CNSS, INPP, ONEM, IPR, SMIG, jours
   fériés), calendrier de référence 2026-2035, taux de change. Toutes les
   valeurs légales sont vérifiées contre des textes officiels (voir section
   "Sources légales vérifiées" plus bas) — sauf exceptions explicitement
   signalées comme non vérifiées dans le fichier lui-même (échéance INPP,
   échéance IPR : sources secondaires seulement).
2. **EMPLOYES** — référentiel RH, 60 lignes pré-formatées (agents 1-60 =
   lignes 15-74). Contient depuis peu : mode de négociation (Brut de base /
   Net salaire seul / Net tout compris) + moteur d'inversion net→brut par
   barème IPR (colonnes X, Y).
3. **VARIABLES** — pointage journalier complet (grille P/D/A/F, 31 jours ×
   12 mois × 60 agents), jours réels/prévus/absence calculés automatiquement
   (pas de saisie manuelle d'absence).
4. **BULLETIN** — moteur de calcul, un agent à la fois (sélecteur matricule).
5. **JOURNAL** — même cascade que BULLETIN mais pour les 60 agents en
   parallèle (calcul indépendant, sert de validation croisée).
6. **DECLARATIONS** — bordereaux CNSS/INPP/ONEM/IPR à partir de JOURNAL.
7. **DASHBOARD_RH** — effectifs, ancienneté, consolidation annuelle honnête
   (mois clos uniquement), sélecteur de mois indépendant pour consulter
   l'historique, listes filtrées (absences/avances).
8. **HISTORIQUE** — archive figée des mois clôturés. **Clôture manuelle**
   (Collage spécial > Valeurs depuis JOURNAL) — une macro VBA automatiserait
   ça mais n'a jamais pu être injectée proprement via le pipeline Python/
   openpyxl utilisé jusqu'ici. Si Claude Code peut manipuler le classeur
   directement en VBA (COM/win32com, ou simplement en étant un environnement
   qui sait éditer un .xlsm), **c'est le principal gain possible** :
   implémenter la bascule mois/année + archivage automatique à la fermeture,
   sur le modèle observé dans un fichier tiers étudié (`SIMULATION_10ans_
   30agents.xlsm`, non fourni ici mais dont la logique est décrite plus bas).

## Pièges déjà trouvés — à ne pas re-découvrir à ses frais

Ce sont des bugs réels, chacun confirmé par test isolé avant correction :

1. **`SUMPRODUCT` + cellule texte = `#VALUE!` silencieux sous `IFERROR`.**
   Si une colonne "vide par formule" renvoie `""` (texte) plutôt qu'un vrai
   blanc, et qu'elle est ensuite multipliée dans un `SUMPRODUCT` (via
   `IFERROR(SUMPRODUCT(...),0)`), le résultat retombe silencieusement à 0
   — masquant une vraie valeur. **Toujours renvoyer 0, pas `""`, pour les
   colonnes numériques consommées par SUMPRODUCT en aval.** Touché deux fois
   dans ce projet (jours fériés texte vs date ; jours d'absence "" vs 0).

2. **Mise en forme conditionnelle + référence inter-feuille = ne se déclenche
   jamais**, testé et confirmé avec ce moteur de rendu (LibreOffice via le
   script `recalc.py`). Si une règle CF doit dépendre d'une valeur dans une
   AUTRE feuille (ex. le régime d'un agent stocké dans EMPLOYES), il faut
   d'abord la **rapatrier dans une colonne miroir de la même feuille**, puis
   faire référence à cette colonne miroir dans la règle CF.

3. **Noms définis à portée classeur ne prennent PAS le préfixe `NomFeuille!`.**
   `PARAMETRES!nom_du_parametre` est invalide si `nom_du_parametre` est un nom
   global — ça se traduit en formule renvoyée en minuscules par LibreOffice
   (signe distinctif d'une formule qu'il n'a pas su analyser) puis `#NAME?`.

4. **`LET()` est une fonction Excel récente, à éviter** — compatibilité
   LibreOffice non garantie. Préférer la répétition explicite de la
   sous-expression plutôt qu'une variable nommée en formule.

5. **Insertion de lignes/colonnes au milieu d'une feuille existante = risque
   de casse silencieuse.** openpyxl ne réajuste pas automatiquement les
   formules des autres feuilles qui référencent des numéros de ligne fixes.
   Règle du projet : **toujours ajouter en fin de feuille (append-only)**,
   jamais insérer au milieu, sauf reconstruction complète assumée (et dans ce
   cas, mettre à jour toutes les tables de références en aval — voir
   `BULLETIN!$B$201:$D$212`, la table technique des bornes de blocs
   VARIABLES par mois, qui doit être resynchronisée à chaque reconstruction
   de VARIABLES).

6. **Zéro erreur au recalcul ne prouve pas que les formules sont justes** —
   seulement qu'elles s'évaluent. Toujours revérifier les valeurs calculées
   par un recalcul indépendant (Python) sur au moins un cas non trivial.

## Conventions du fichier

- **Couleurs** : bleu gras sur fond jaune pâle = cellule à saisir. Noir = 
  calculé, ne pas modifier. Rouge gras sur fond orange = avertissement/non
  vérifié.
- **Garde-fou ligne vide** : quasi toutes les formules par agent commencent
  par `IF($C{ligne}="","",...)` pour que les lignes d'agents non remplies
  n'affichent rien (et ne polluent pas les totaux).
- **Skill xlsx obligatoire** : lire `/mnt/skills/public/xlsx/SKILL.md` avant
  toute construction. Le script `recalc.py` (LibreOffice headless) est
  systématiquement utilisé après CHAQUE modification, avec vérification du
  compte d'erreurs ET d'un échantillon de valeurs.
- **Named ranges** : très nombreux, définis pour presque tous les paramètres
  légaux (`t_cnss_pension_salarie`, `ipr_taux1..4`, `smig_journalier`,
  `jours_ouvrables_6j_mois_courant`, etc.) — préférer leur usage direct
  (sans préfixe de feuille, voir piège n°3) plutôt que des références de
  cellule brutes.

## Sources légales vérifiées (recherches web menées dans la conversation d'origine)

- CNSS : Décret n°18/041 du 24/11/2018 — 18 % total (5 % salarié pension +
  13 % employeur : 5 % pension + 6,5 % familles + 1,5 % risques). Déclaration
  et paiement dans les 15 jours suivant le mois.
- SMIG : Décret n°25/22 du 30/05/2025 — 21 500 CDF/jour depuis janvier 2026.
- ONEM : 0,5 % depuis août/septembre 2025 (Arrêté du 05/08/2025, porté de
  0,2 %). Déclaration 10 jours, paiement 15 jours.
- INPP : Arrêté du 14/02/2006 — barème par tranche d'effectif (3 % / 2 % /
  1 % selon 1-50 / 51-300 / >300 salariés). **Versement trimestriel**
  (Ordonnance 84-186 du 15/10/1984 : 30 avril, 31 juillet, 31 octobre,
  31 janvier) — tension avec l'arrêté 2006 qui parle de base "mensuelle",
  signalée mais non résolue dans le fichier.
- IPR : Ordonnance-Loi 69/009 — barème progressif 4 tranches (3 %/15 %/30 %/
  40 %), plancher 2000 CDF/mois, plafond 30 % du revenu imposable. Déclaration
  10 jours (sources secondaires seulement, pas de texte primaire consulté).
- Heures supplémentaires : Arrêté 68/11 du 17/05/1968 + art. 119 Loi 16/010 —
  30 % (6 premières heures/semaine), 60 % (au-delà), 100 % (jour de repos,
  NON modélisé séparément — VARIABLES ne distingue pas ce cas).
- Jours fériés : Ordonnance n°23-042 du 30/03/2023 — 10 jours fériés légaux
  (liste complète dans PARAMETRES section 9), avec règle de report automatique
  dimanche→veille programmée en formule ; cas samedi laissé en ajustement
  manuel (pratique gouvernementale incohérente d'une année à l'autre).

## Chantier « devise contractuelle par agent » — terminé (déjà présent au transfert)

Contrairement à ce que ce fichier disait au moment du transfert, la version
du classeur reçue avait déjà cette fonctionnalité entièrement construite :
EMPLOYES colonnes Z (« Devise du contrat », CDF/USD, liste déroulante) et AA
(« Montant négocié converti (CDF) » = `IF(Z="USD", P*opt_taux_change, P)`),
avec **taux unique** (`opt_taux_change`, PARAMETRES!C22 — confirmé le choix
de conception avec l'utilisateur). BULLETIN et JOURNAL consomment bien le
montant converti (colonne Y, résolu à partir de AA), jamais le montant brut
négocié (P). Vérifié : recalcul indépendant en Python de l'inversion
net→brut sur l'agent 16 (500 USD, mode "Net salaire seul") — `0,95B −
IPR(B) = 1 450 000` → `B = 1 788 200 CDF`, exact match avec la valeur en
cache du fichier. Rien à construire ; aucune modification faite sur ce point.

## Chantier « clôture automatique HISTORIQUE » — macro VBA livrée en .xlsm séparé

`Morning_SARL_Systeme_Paie_RH_RDC.xlsm` (généré par `tools/build_xlsm.py` —
relancer ce script après toute modification du .xlsx pour regénérer le
.xlsm à jour) est une **copie séparée**, avec macro, du .xlsx canonique (qui reste volontairement sans macro, cf. « Construit
entièrement en formules Excel pures » ci-dessus). Il ajoute une macro
`ClotureDuMois` (Alt+F8 pour l'exécuter — pas de bouton graphique, voir
pourquoi ci-dessous) qui automatise exactement la procédure manuelle déjà
documentée dans HISTORIQUE!B5:B9 : copie figée (valeurs) de JOURNAL vers le
bloc du mois en cours dans HISTORIQUE, statut « Clos » + date, puis avance
PARAMETRES!C30/C29 au mois suivant. Refuse d'écraser un bloc déjà clos.

**Limite connue, non résolue** : HISTORIQUE n'a que 12 blocs nommés par mois
(pas par année) — la macro (comme la procédure manuelle) ne peut donc
archiver qu'une seule année à la fois. Fermer un mois d'une deuxième année
alors que son bloc est déjà « Clos » est bloqué explicitement (message
d'erreur clair) plutôt que d'écraser l'archive. Étendre HISTORIQUE pour
plusieurs années (nouveaux blocs en fin de feuille, append-only) reste à
faire — c'est justement la question laissée ouverte par le fichier tiers
`SIMULATION_10ans_30agents.xlsm` mentionné plus haut, jamais tranchée.

**Comment le vbaProject.bin a été construit sans Excel/Windows/COM** (cet
environnement Claude Code est Linux) : voir la docstring de
`tools/build_xlsm.py` pour le détail complet. En résumé — un vbaProject.bin
réel généré par Excel (tiré des exemples de la bibliothèque XlsxWriter,
BSD, prévu pour cet usage) sert de gabarit ; son module `Module1` est
remplacé par le code de `tools/vba/ClotureDuMois.bas`, recompressé avec une
implémentation MS-OVBA (2.4.1) écrite pour l'occasion (`tools/vba/
ovba_compress.py`, mode littéral seul, vérifiée par aller-retour via le
décompresseur de référence d'oletools) ; le conteneur OLE2/MS-CFB est
reconstruit de zéro (`tools/vba/cfb_writer.py`) car `olefile` ne sait
réécrire un flux qu'à taille identique.

**Vérifié** : extraction du code via `oletools.olevba`, sur le `.bin` seul
et sur le `.xlsm` final, identique octet pour octet à `ClotureDuMois.bas` ;
structure OLE2 valide (`olefile`) ; chargement complet du `.xlsm` par
`openpyxl` avec les mêmes 65 noms définis, les mêmes 16 072 formules et les
mêmes 186 validations de données que le `.xlsx` source (donc le passage par
openpyxl pour ajouter la note d'instruction n'a rien cassé).
**Non vérifié — à faire avant tout usage réel** : aucun test dans Excel
réel (impossible dans cet environnement, pas de Windows/COM disponible).
La macro n'a donc jamais tourné pour de vrai ; seule sa présence et son
intégrité structurelle sont prouvées. C'est pourquoi il n'y a pas de bouton
graphique (Form Control) — sa géométrie/rendu VML n'aurait pas pu être
vérifié non plus, et un XML mal formé y aurait été plus risqué qu'utile
pour peu de gain face à Alt+F8.

## Discipline de travail attendue

L'utilisateur (Luisance) travaille par petites étapes validées une par une,
avec vérification stricte (recalcul + valeurs, jamais juste "zéro erreur") à
chaque étape, et signalement explicite de toute approximation ou incertitude
plutôt que de la masquer. Ne jamais présenter une valeur non vérifiée comme
un fait établi.
