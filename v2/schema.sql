-- Base personnelle de l'app Finances v2 (SQLite). Une base par personne, dans son OneDrive.
-- Montants : en cents entiers (_c). Exception : historique importé, où l'Excel contenait des parts à plus de 2 décimales
-- (part_hist, en dollars, valeur exacte de l'Excel) : on la garde telle quelle pour que les totaux tombent au cent près.
-- Aucun total n'est enregistré : tout est recalculé (voir calculs.py / calculs.js).
PRAGMA foreign_keys = ON;

CREATE TABLE reglages (cle TEXT PRIMARY KEY, valeur TEXT);       -- valeurs en JSON

CREATE TABLE categories (
  id INTEGER PRIMARY KEY, nom TEXT NOT NULL UNIQUE, ordre INTEGER NOT NULL, couleur TEXT, active INTEGER NOT NULL DEFAULT 1);

CREATE TABLE mots_cles (                                         -- le dernier mot trouvé dans la description l'emporte
  id INTEGER PRIMARY KEY, mot TEXT NOT NULL, categorie_id INTEGER NOT NULL REFERENCES categories(id), ordre INTEGER NOT NULL);

CREATE TABLE reglements (                                        -- un « Solde réglé »
  id INTEGER PRIMARY KEY, date TEXT NOT NULL, ordre INTEGER NOT NULL,
  confirme INTEGER NOT NULL DEFAULT 1, source TEXT);              -- source : 'import' | 'app' | 'sync'

CREATE TABLE depenses (
  id INTEGER PRIMARY KEY,
  mois TEXT NOT NULL,                                            -- période (AAAA-MM) : le mois où la dépense compte
                                                                 -- (= mois de la date, sauf exception reprise de l'Excel)
  date TEXT,                                                     -- AAAA-MM-JJ
  description TEXT NOT NULL DEFAULT '',
  montant_c INTEGER,                                             -- montant payé (cents) ; NULL pour une ligne d'historique
  calcul TEXT,                                                   -- saisie d'origine, ex. « 33,34+15,53 »
  type TEXT NOT NULL,                                            -- P | 50 | SP | M (100 % partenaire) | R (ancien Remboursement)
  categorie_id INTEGER REFERENCES categories(id),                -- NULL = catégorie automatique (mots-clés)
  paye_par TEXT NOT NULL DEFAULT 'moi',                          -- 'moi' | 'partenaire' (part payée par la carte de l'autre)
  reglement_id INTEGER REFERENCES reglements(id),                -- règlement qui a soldé la part (NULL = pas encore réglée)
  id_partage TEXT,                                               -- identifiant dans le fichier commun du couple
  historique INTEGER NOT NULL DEFAULT 0,                         -- 1 = mois d'avant la v1 (lecture seule)
  part_hist REAL,                                                -- historique : « Montant » de l'Excel (ma part, exacte)
  ordre INTEGER NOT NULL,                                        -- ordre de saisie (départage les règlements)
  origine TEXT);                                                 -- traçabilité de l'import, ex. « Octobre 2026 !L12 »

CREATE TABLE autres_montants (                                   -- ajustements avec le partenaire (+ il/elle me doit, − je dois)
  id INTEGER PRIMARY KEY, date TEXT NOT NULL, libelle TEXT NOT NULL, montant_c INTEGER NOT NULL,
  reglement_id INTEGER REFERENCES reglements(id), ordre INTEGER NOT NULL, origine TEXT);

CREATE TABLE pourboires (
  id INTEGER PRIMARY KEY, mois TEXT NOT NULL, date TEXT, especes_c INTEGER, carte_c INTEGER, heures REAL,
  recu INTEGER,                                                  -- 1 reçu, 0 non reçu, NULL non indiqué
  origine TEXT);

CREATE TABLE revenus (
  id INTEGER PRIMARY KEY, mois TEXT NOT NULL, date TEXT, description TEXT NOT NULL DEFAULT '', montant_c INTEGER NOT NULL,
  calcul TEXT, cases_colonne TEXT,                               -- 'especes' | 'carte' quand le revenu additionne des jours de pourboires
  cases_ajout_c INTEGER NOT NULL DEFAULT 0, historique INTEGER NOT NULL DEFAULT 0, origine TEXT);
CREATE TABLE revenus_pourboires (revenu_id INTEGER NOT NULL REFERENCES revenus(id) ON DELETE CASCADE,
  pourboire_id INTEGER NOT NULL REFERENCES pourboires(id), PRIMARY KEY (revenu_id, pourboire_id));

CREATE TABLE repartition (                                       -- un mois (AAAA-MM)
  mois TEXT PRIMARY KEY, revenu_moi_c INTEGER, revenu_partenaire_c INTEGER,
  pct_partenaire REAL NOT NULL, pct_manuel INTEGER NOT NULL DEFAULT 0);

CREATE TABLE comptes_placement (
  id INTEGER PRIMARY KEY, nom TEXT NOT NULL UNIQUE, valeur_fixe INTEGER NOT NULL DEFAULT 0, ordre INTEGER NOT NULL);
CREATE TABLE placements_mois (
  compte_id INTEGER NOT NULL REFERENCES comptes_placement(id), mois TEXT NOT NULL,
  valeur_c INTEGER, investi_c INTEGER, investi_manuel INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (compte_id, mois));
CREATE TABLE versements_epargne (                                -- argent sorti du compte vers un placement
  id INTEGER PRIMARY KEY, mois TEXT NOT NULL, date TEXT NOT NULL, compte_id INTEGER REFERENCES comptes_placement(id), montant_c INTEGER NOT NULL, origine TEXT);

CREATE TABLE soldes_mois (                                       -- un mois : point de départ et ajustement repris de l'Excel
  mois TEXT PRIMARY KEY,
  debut_c INTEGER,                                               -- point de départ explicite (premier mois, ou solde réel de la banque
                                                                 -- inscrit à la fin du mois précédent) ; NULL = fin du mois précédent
  debut_raison TEXT,
  ajustement_c INTEGER NOT NULL DEFAULT 0,                       -- écart entre le solde de fin de l'Excel et le calcul standard
  ajustement_raison TEXT);
CREATE TABLE ajustements_solde (                                 -- « Ajustement de solde » inscrit par la personne (écart inexpliqué)
  id INTEGER PRIMARY KEY, mois TEXT NOT NULL, date TEXT NOT NULL, montant_c INTEGER NOT NULL, note TEXT);
CREATE TABLE financements (                                      -- prêts et autres financements (ex. prêt étudiant)
  id INTEGER PRIMARY KEY, nom TEXT NOT NULL, ordre INTEGER NOT NULL, cree TEXT);
CREATE TABLE financements_recus (                                -- argent reçu d'un financement : entre au compte, mais ce n'est PAS un revenu
  id INTEGER PRIMARY KEY, financement_id INTEGER NOT NULL REFERENCES financements(id), mois TEXT NOT NULL, date TEXT NOT NULL,
  montant_c INTEGER NOT NULL, note TEXT);
CREATE TABLE financements_soldes (                               -- solde dû inscrit à la main (ex. relevé du prêt)
  id INTEGER PRIMARY KEY, financement_id INTEGER NOT NULL REFERENCES financements(id), date TEXT NOT NULL, solde_du_c INTEGER NOT NULL);
CREATE TABLE releves_banque (id INTEGER PRIMARY KEY, date TEXT NOT NULL, solde_c INTEGER NOT NULL, origine TEXT);

CREATE TABLE budgets (categorie_id INTEGER NOT NULL REFERENCES categories(id), mois TEXT NOT NULL DEFAULT '',  -- '' = défaut
  montant_c INTEGER NOT NULL, PRIMARY KEY (categorie_id, mois));

CREATE TABLE exclusions_rappro (id INTEGER PRIMARY KEY, texte TEXT NOT NULL, compte TEXT);

CREATE TABLE journal (id INTEGER PRIMARY KEY, quand TEXT NOT NULL, table_nom TEXT NOT NULL, ligne_id INTEGER,
  action TEXT NOT NULL, avant TEXT, apres TEXT);

CREATE INDEX i_dep_mois ON depenses(mois);
CREATE INDEX i_rev_mois ON revenus(mois);
