"""
analyseDetensionnement.py

Analyse un (ou plusieurs) fichier(s) Excel d'enregistreur de température
(type SMARTDAC+, colonnes CHxxxx / CHCxxx) et génère, pour chacun, une
image regroupant :
  - un graphique gauche (pleine hauteur) montrant l'ensemble des voies
    retenues après filtrage (grisées),
  - 10 petits graphiques à droite (répartis sur 2 colonnes), un par
    sous-phase du cycle : montée 1, regroupement 1, maintien 1, montée 2,
    regroupement 2, maintien 2, sortie 2, descente 3, regroupement 3,
    descente 4 — chaque sous-phase est isolée dans son propre graphique
    pour ne pas surcharger un même graphique avec trop d'informations,
ainsi qu'un second fichier texte contenant les valeurs de synthèse du
cycle (température initiale, vitesses de montée/descente, durées de
regroupement/maintien/sortie, température de maintien 2), au format
imposé par un exemple fourni (voir hypothèse 8).

Les paramètres se trouvent dans un fichier YAML (par défaut
analyseDetensionnement.yaml), avec quelques surcharges possibles en
ligne de commande pour les seuils principaux.

Usage :
    python analyseDetensionnement.py fichier1.xlsx fichier2.xlsx
    (sans fichier -> traite tous les .xlsx du dossier courant, hors fichiers
    déjà suffixés)
    (paramètres absents de la ligne de commande -> lus dans
    analyseDetensionnement.yaml)

-------------------------------------------------------------------------
HYPOTHÈSES / CHOIX DE CONCEPTION (assumés faute de spécification exacte) :

1. Structure du fichier Excel : la ligne d'en-tête des voies est repérée
   automatiquement en cherchant la première ligne contenant des cellules du
   type "CH0001", "CH0002"... ou "CHC001", "CHC002"... (pas de numéro de
   ligne fixe). La ligne de début des données est repérée en cherchant,
   après cette ligne, la ligne dont la colonne A vaut "Date" et la colonne
   B vaut "Time" : les données commencent juste après. La lecture s'arrête
   à la première ligne dont la colonne "Date" est vide. Aucune correction
   de chronologie n'est appliquée.

2. Cellule `titre_graphique` : repérée en cherchant, en colonne A, la ligne
   dont le libellé vaut exactement "Batch No." (fixe), et en lisant la
   valeur en colonne C (fixe également — non exposées dans le YAML, à la
   demande explicite d'une itération précédente).

3. `CH_ignore_dT` : une voie est CONSERVÉE si (max - min) de ses valeurs
   sur l'ensemble du fichier est STRICTEMENT SUPÉRIEUR à `CH_ignore_dT` ;
   sinon elle est ignorée. Une voie contenant au moins une valeur de type
   "-OVER" est ignorée avant même ce calcul.

4. "Première courbe" / "dernière courbe" : à CHAQUE étape de détection, on
   considère TOUTES les voies retenues et on cherche, pour le seuil de
   cette étape, laquelle le franchit LA PREMIÈRE (l'instant le plus tôt
   parmi toutes les voies qui le franchissent) ou LA DERNIÈRE (l'instant
   le plus tardif). Ce n'est donc pas une voie fixe pour tout le fichier :
   la voie "première" ou "dernière" peut différer d'une étape à l'autre.
   Le nom de la voie retenue à chaque étape est conservé et affiché dans
   l'annotation du point correspondant.

5. Détection des instants (recherche séquentielle ; chaque étape reprend
   la recherche, sur TOUTES les voies, à partir de l'instant trouvé à
   l'étape précédente) :
     - t0 (début montée) = 1ère voie à franchir (montant)
       température initiale + `temp_delta_ambiant`
     - t1 (fin montée 1) = 1ère voie à franchir (montant)
       `temp_palier_1` - `temp_delta_1` -> vitesse de montée 1 (°C/h)
       calculée entre t0 et t1
     - t2 (regroupement 1) = DERNIÈRE voie à franchir (montant) le même
       seuil que t1 -> durée de regroupement 1 = t2 - t1
     - t3 (maintien 1) = 1ère voie à franchir (montant)
       `temp_palier_1` + `temp_delta_1` -> durée de maintien 1 = t3 - t2
     - t4 (fin montée 2) = 1ère voie à franchir (montant)
       `temp_palier_2` - `temp_delta_2` -> vitesse de montée 2 (°C/h)
       calculée entre t3 et t4
     - t5 (regroupement 2) = DERNIÈRE voie à franchir (montant) le même
       seuil que t4 -> durée de regroupement 2 = t5 - t4
       [ÉTAPE AJOUTÉE : le cahier des charges ne décrivait pas
       explicitement de palier "regroupement 2" symétrique du
       regroupement 1, mais la liste des 12 valeurs de sortie demandées
       (hypothèse 8) inclut une "Durée de regroupement 2" — cette étape a
       donc été ajoutée par symétrie avec le palier 1]
     - t6 (maintien 2) = 1ère voie à franchir (montant)
       `temp_palier_2` + `temp_delta_2` -> durée de maintien 2 = t6 - t5 ;
       la "température de maintien 2" est le maximum atteint, TOUTES
       voies retenues confondues, entre t5 et t6
     - t7 (sortie 2) = DERNIÈRE voie à franchir (descendant)
       `temp_palier_2` - `temp_delta_2` -> durée de sortie 2 = t7 - t6
     - t8 (fin descente 3) = 1ère voie à franchir (descendant)
       `temp_palier_3` + `temp_delta_3` -> vitesse de descente 3 (°C/h)
       calculée entre t7 et t8
     - t9 (regroupement 3) = DERNIÈRE voie à franchir (descendant)
       `temp_palier_3` - `temp_delta_3` -> durée de regroupement 3 = t9 - t8
     - t10 (fin descente 4) = 1ère voie à franchir (descendant)
       `temp_palier_4` -> vitesse de descente 4 (°C/h) calculée entre t9
       et t10
   Tous les franchissements sont interpolés linéairement entre les deux
   échantillons encadrants. Si une étape n'est pas trouvée, les étapes
   suivantes ne sont pas calculées (valeurs "N/A" dans les sorties).

6. Fenêtres de zoom : deux paramètres uniques `zoom_marge_temperature`
   (°C) et `zoom_marge_temps` (minutes), appliqués aux 10 graphiques de
   droite, ajoutés de part et d'autre de la plage utile de chaque
   sous-phase.

7. Fichier de sortie image : `<nom_fichier_excel>_<suffixe>.png`.
   Fichier de sortie résultats : `<nom_fichier_excel>_<suffixe_resultats>.txt`.

8. Formalisme STRICT du fichier de résultats (imposé, reproduit tel quel) :
   12 lignes de libellés puis 12 lignes de valeurs, dans cet ordre :
       température initiale                    (tabulation AVANT la valeur)
       vitesse de montée 1                      (avant)
       durée de regroupement 1                  (avant)
       durée du maintien 1                      (avant)
       Vitesse de montée 2                      (APRÈS la valeur)
       Durée de regroupement 2                  (après)
       Température de maintien 2                (après)
       Durée du maintien 2                      (après)
       Durée de la sortie 2                     (après)
       Vitesse de descente 3                    (après)
       Durée du regroupement 3                  (avant)
       vitesse de descente 4                    (avant)
   soit, pour les valeurs (exemple donné) :
       \t12°C
       \t20°C/h
       \t5min
       \t40min
       20°C/h\t
       5min\t
       40°C\t
       3h02min\t
       2min\t
       20°C/h\t
       \t2min
       \t10°C/h
   Le fichier final concatène le bloc des 12 libellés puis le bloc des 12
   valeurs, chacun respectant la position de tabulation indiquée. [Les
   fautes de frappe "maintient"/"vitese" d'une itération précédente ont
   été corrigées en "maintien"/"vitesse" à la demande de l'utilisateur.]

9. Température initiale : puisqu'il n'y a plus de voie de référence fixe
   (voir hypothèse 4), la température initiale est la MOYENNE des voies
   retenues au premier instant du fichier.

10. Toutes les valeurs de sortie sont arrondies à l'entier le plus proche,
    sans décimale. Durées formatées sans espace, en minutes en dessous de
    60 min ("5min"), sinon en heures+minutes ("3h02min"). Les vitesses de
    montée/descente sont exprimées en °C/h.

11. Couleurs : chaque point d'intérêt (t0 à t10) a sa propre couleur fixe,
    réutilisée à l'identique sur tous les graphiques. Les annotations sont
    posées sur fond blanc semi-opaque, avec le nom de la voie à l'origine
    du point, et leur position (au-dessus/en dessous, à gauche/à droite)
    est choisie automatiquement selon la position du point dans la
    fenêtre affichée. Chaque sous-phase (montée, regroupement, maintien,
    sortie, descente) est isolée dans son propre petit graphique (2 points
    au plus par graphique, 3 pour le maintien 2 qui inclut le marqueur de
    température maximale) afin d'éviter la superposition des étiquettes.
-------------------------------------------------------------------------
"""

import argparse
import os
import re
import sys
from datetime import datetime, timedelta

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import openpyxl
import yaml


# =========================================================================
# Constantes non paramétrables (retirées du YAML à la demande de l'utilisateur)
# =========================================================================

FORMAT_IMAGE = 'png'
FORMAT_RESULTATS = 'txt'
FEUILLE = None
TITRE_GRAPHIQUE_COL = 3
TITRE_GRAPHIQUE_LABEL = "Batch No."

DEFAULT_CONFIG_FILE = "analyseDetensionnement.yaml"


# =========================================================================
# Lecture du fichier Excel
# =========================================================================

_RE_CHANNEL = re.compile(r'^CH[A-Z]?\d{3,4}$', re.IGNORECASE)


def trouver_ligne_entete_channels(ws, max_rows_scan=200):
    """Cherche la première ligne contenant des libellés de voie CHxxxx /
    CHCxxx. Renvoie (numero_ligne, {colonne: nom_voie})."""
    for r in range(1, min(max_rows_scan, ws.max_row) + 1):
        canaux = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and _RE_CHANNEL.match(v.strip()):
                canaux[c] = v.strip()
        if canaux:
            return r, canaux
    raise ValueError("Impossible de trouver la ligne d'en-tête des voies (CHxxxx / CHCxxx).")


def trouver_ligne_debut_donnees(ws, ligne_apres, max_rows_scan=50):
    """Cherche, après la ligne d'en-tête des voies, la ligne 'Date'/'Time' et
    renvoie le numéro de la première ligne de données (juste après)."""
    borne = min(ligne_apres + max_rows_scan, ws.max_row)
    for r in range(ligne_apres, borne + 1):
        v1 = ws.cell(row=r, column=1).value
        v2 = ws.cell(row=r, column=2).value
        if isinstance(v1, str) and v1.strip().lower() == 'date' and \
                isinstance(v2, str) and v2.strip().lower() == 'time':
            return r + 1
    raise ValueError("Impossible de trouver la ligne d'en-tête 'Date' / 'Time'.")


def trouver_titre_graphique(ws, max_rows_scan=200):
    for r in range(1, min(max_rows_scan, ws.max_row) + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip() == TITRE_GRAPHIQUE_LABEL:
            val = ws.cell(row=r, column=TITRE_GRAPHIQUE_COL).value
            return "" if val is None else str(val).strip()
    return ""


def parser_date_heure(val_date, val_time):
    if isinstance(val_date, datetime):
        d = val_date.date()
    else:
        d = datetime.strptime(str(val_date).strip(), "%Y/%m/%d").date()

    if hasattr(val_time, 'hour'):
        t = val_time.time() if isinstance(val_time, datetime) else val_time
    else:
        s = str(val_time).strip()
        fmt = "%H:%M:%S" if s.count(':') == 2 else "%H:%M"
        t = datetime.strptime(s, fmt).time()

    return datetime.combine(d, t)


def lire_donnees_excel(fichier):
    wb = openpyxl.load_workbook(fichier, data_only=True)
    ws = wb[FEUILLE] if FEUILLE else wb[wb.sheetnames[0]]

    ligne_ch, canaux_col = trouver_ligne_entete_channels(ws)
    ligne_debut = trouver_ligne_debut_donnees(ws, ligne_ch)
    titre = trouver_titre_graphique(ws)

    dates = []
    valeurs_brutes = {nom: [] for nom in canaux_col.values()}

    r = ligne_debut
    while True:
        v_date = ws.cell(row=r, column=1).value
        if v_date is None or str(v_date).strip() == '':
            break
        v_time = ws.cell(row=r, column=2).value
        dates.append(parser_date_heure(v_date, v_time))
        for c, nom in canaux_col.items():
            valeurs_brutes[nom].append(ws.cell(row=r, column=c).value)
        r += 1

    if not dates:
        raise ValueError(f"Aucune donnée trouvée dans {fichier}.")

    return titre, dates, valeurs_brutes


# =========================================================================
# Filtrage des voies
# =========================================================================

def _est_valeur_over(v):
    return isinstance(v, str) and 'OVER' in v.upper()


def filtrer_voies(valeurs_brutes, ch_ignore_dT):
    """Retire les voies contenant une valeur de type -OVER, puis celles dont
    l'amplitude (max-min) est <= ch_ignore_dT. Renvoie {nom: [floats]},
    en conservant l'ordre des colonnes du fichier Excel d'origine."""
    voies = {}
    for nom, valeurs in valeurs_brutes.items():
        if any(_est_valeur_over(v) for v in valeurs):
            continue
        try:
            floats = [float(v) for v in valeurs]
        except (TypeError, ValueError):
            continue
        if (max(floats) - min(floats)) > ch_ignore_dT:
            voies[nom] = floats
    return voies


# =========================================================================
# Détection des instants de franchissement de seuil
# =========================================================================

def detecter_franchissement(dates, valeurs, seuil, sens, idx_debut=0):
    """sens = 'montant' (franchissement vers le haut) ou 'descente'
    (franchissement vers le bas). Recherche à partir de l'indice idx_debut.
    Renvoie (instant interpolé, indice de l'échantillon suivant le
    franchissement) ou (None, None) si le seuil n'est jamais franchi."""
    depart = max(1, idx_debut)
    for i in range(depart, len(valeurs)):
        v0, v1 = valeurs[i - 1], valeurs[i]
        if sens == 'montant':
            franchi = v0 < seuil <= v1
        else:
            franchi = v0 >= seuil > v1
        if not franchi:
            continue
        t0, t1 = dates[i - 1], dates[i]
        frac = 0.0 if v1 == v0 else (seuil - v0) / (v1 - v0)
        frac = max(0.0, min(1.0, frac))
        return t0 + (t1 - t0) * frac, i
    return None, None


def instant_extreme(dates, voies, seuil, sens, mode, idx_debut=0):
    """Cherche, parmi TOUTES les voies, le franchissement du seuil (sens
    'montant'/'descente') à partir de idx_debut, et renvoie celui de
    toutes les voies qui franchit LE PLUS TÔT (mode='first') ou LE PLUS
    TARD (mode='last'). Renvoie (instant, nom_voie, indice_global) ou
    (None, None, idx_debut) si aucune voie ne franchit le seuil."""
    candidats = []
    for nom, valeurs in voies.items():
        t, i = detecter_franchissement(dates, valeurs, seuil, sens, idx_debut)
        if t is not None:
            candidats.append((t, nom, i))
    if not candidats:
        return None, None, idx_debut
    if mode == 'first':
        return min(candidats, key=lambda c: c[0])
    return max(candidats, key=lambda c: c[0])


def calculer_cycle(dates, voies, config):
    """Calcule les instants t0..t10 et les valeurs de synthèse du cycle
    de détensionnement (voir hypothèse 5 en tête de fichier). Renvoie un
    dict."""
    noms = list(voies)
    temp_initiale = sum(voies[nom][0] for nom in noms) / len(noms) if noms else None

    cycle = {f't{i}': None for i in range(11)}
    cycle.update({f'seuil_t{i}': None for i in range(11)})
    cycle.update({f'voie_t{i}': None for i in range(11)})
    cycle.update({
        'temp_initiale': temp_initiale,
        'vitesse_montee_1': None,
        'duree_regroupement_1_min': None,
        'duree_maintien_1_min': None,
        'vitesse_montee_2': None,
        'duree_regroupement_2_min': None,
        'temp_maintien_2': None,
        'instant_maintien_2': None,
        'voie_maintien_2': None,
        'duree_maintien_2_min': None,
        'duree_sortie_2_min': None,
        'vitesse_descente_3': None,
        'duree_regroupement_3_min': None,
        'vitesse_descente_4': None,
    })

    if temp_initiale is None:
        return cycle

    seuil_t0 = temp_initiale + config['temp_delta_ambiant']
    cycle['seuil_t0'] = seuil_t0
    t0, nom0, i0 = instant_extreme(dates, voies, seuil_t0, 'montant', 'first', 0)
    cycle['t0'], cycle['voie_t0'] = t0, nom0
    if t0 is None:
        return cycle

    seuil_t1 = config['temp_palier_1'] - config['temp_delta_1']
    cycle['seuil_t1'] = seuil_t1
    t1, nom1, i1 = instant_extreme(dates, voies, seuil_t1, 'montant', 'first', i0)
    cycle['t1'], cycle['voie_t1'] = t1, nom1
    if t1 is None:
        return cycle
    heures = (t1 - t0).total_seconds() / 3600.0
    if heures > 0:
        cycle['vitesse_montee_1'] = (seuil_t1 - seuil_t0) / heures

    cycle['seuil_t2'] = seuil_t1
    t2, nom2, i2 = instant_extreme(dates, voies, seuil_t1, 'montant', 'last', i1)
    cycle['t2'], cycle['voie_t2'] = t2, nom2
    if t2 is None:
        return cycle
    cycle['duree_regroupement_1_min'] = (t2 - t1).total_seconds() / 60.0

    seuil_t3 = config['temp_palier_1'] + config['temp_delta_1']
    cycle['seuil_t3'] = seuil_t3
    t3, nom3, i3 = instant_extreme(dates, voies, seuil_t3, 'montant', 'first', i2)
    cycle['t3'], cycle['voie_t3'] = t3, nom3
    if t3 is None:
        return cycle
    cycle['duree_maintien_1_min'] = (t3 - t2).total_seconds() / 60.0

    seuil_t4 = config['temp_palier_2'] - config['temp_delta_2']
    cycle['seuil_t4'] = seuil_t4
    t4, nom4, i4 = instant_extreme(dates, voies, seuil_t4, 'montant', 'first', i3)
    cycle['t4'], cycle['voie_t4'] = t4, nom4
    if t4 is None:
        return cycle
    heures = (t4 - t3).total_seconds() / 3600.0
    if heures > 0:
        cycle['vitesse_montee_2'] = (seuil_t4 - seuil_t3) / heures

    cycle['seuil_t5'] = seuil_t4
    t5, nom5, i5 = instant_extreme(dates, voies, seuil_t4, 'montant', 'last', i4)
    cycle['t5'], cycle['voie_t5'] = t5, nom5
    if t5 is None:
        return cycle
    cycle['duree_regroupement_2_min'] = (t5 - t4).total_seconds() / 60.0

    seuil_t6 = config['temp_palier_2'] + config['temp_delta_2']
    cycle['seuil_t6'] = seuil_t6
    t6, nom6, i6 = instant_extreme(dates, voies, seuil_t6, 'montant', 'first', i5)
    cycle['t6'], cycle['voie_t6'] = t6, nom6
    if t6 is None:
        return cycle
    cycle['duree_maintien_2_min'] = (t6 - t5).total_seconds() / 60.0

    debut_idx, fin_idx = min(i5, i6), max(i5, i6)
    segment_dates = dates[debut_idx:fin_idx + 1]
    meilleur_valeur, meilleur_instant, meilleur_nom = None, None, None
    for nom, valeurs in voies.items():
        segment = valeurs[debut_idx:fin_idx + 1]
        for t, v in zip(segment_dates, segment):
            if meilleur_valeur is None or v > meilleur_valeur:
                meilleur_valeur, meilleur_instant, meilleur_nom = v, t, nom
    cycle['temp_maintien_2'] = meilleur_valeur
    cycle['instant_maintien_2'] = meilleur_instant
    cycle['voie_maintien_2'] = meilleur_nom

    seuil_t7 = config['temp_palier_2'] - config['temp_delta_2']
    cycle['seuil_t7'] = seuil_t7
    t7, nom7, i7 = instant_extreme(dates, voies, seuil_t7, 'descente', 'last', i6)
    cycle['t7'], cycle['voie_t7'] = t7, nom7
    if t7 is None:
        return cycle
    cycle['duree_sortie_2_min'] = (t7 - t6).total_seconds() / 60.0

    seuil_t8 = config['temp_palier_3'] + config['temp_delta_3']
    cycle['seuil_t8'] = seuil_t8
    t8, nom8, i8 = instant_extreme(dates, voies, seuil_t8, 'descente', 'first', i7)
    cycle['t8'], cycle['voie_t8'] = t8, nom8
    if t8 is None:
        return cycle
    heures = (t8 - t7).total_seconds() / 3600.0
    if heures > 0:
        cycle['vitesse_descente_3'] = (seuil_t7 - seuil_t8) / heures

    seuil_t9 = config['temp_palier_3'] - config['temp_delta_3']
    cycle['seuil_t9'] = seuil_t9
    t9, nom9, i9 = instant_extreme(dates, voies, seuil_t9, 'descente', 'last', i8)
    cycle['t9'], cycle['voie_t9'] = t9, nom9
    if t9 is None:
        return cycle
    cycle['duree_regroupement_3_min'] = (t9 - t8).total_seconds() / 60.0

    seuil_t10 = config['temp_palier_4']
    cycle['seuil_t10'] = seuil_t10
    t10, nom10, i10 = instant_extreme(dates, voies, seuil_t10, 'descente', 'first', i9)
    cycle['t10'], cycle['voie_t10'] = t10, nom10
    if t10 is None:
        return cycle
    heures = (t10 - t9).total_seconds() / 3600.0
    if heures > 0:
        cycle['vitesse_descente_4'] = (seuil_t9 - seuil_t10) / heures

    return cycle


# =========================================================================
# Formatage
# =========================================================================

def fmt_num(x, decimales=0):
    if x is None:
        return "N/A"
    s = f"{x:.{decimales}f}"
    if decimales > 0 and s.endswith('.' + '0' * decimales):
        s = s.split('.')[0]
    return s


def fmt_temperature(x):
    return "N/A" if x is None else f"{fmt_num(x)}°C"


def fmt_duree(minutes):
    """Durée sans espace : "5min" en dessous de 60 min, sinon "3h02min"."""
    if minutes is None:
        return "N/A"
    total = round(minutes)
    signe = '-' if total < 0 else ''
    total_abs = abs(total)
    if total_abs >= 60:
        h, m = divmod(total_abs, 60)
        return f"{signe}{h}h{m:02d}min"
    return f"{signe}{total_abs}min"


def fmt_vitesse_h(x):
    return "N/A" if x is None else f"{fmt_num(x)}°C/h"


def _texte_point(instant, valeur, nom_voie):
    if instant is None or valeur is None:
        return None
    base = f"{instant.strftime('%H:%M:%S')}\n{fmt_temperature(valeur)}"
    return f"{base}\n({nom_voie})" if nom_voie else base


# =========================================================================
# Graphique
# =========================================================================

_COULEUR_GRISE = '0.65'
_COULEUR_INITIALE = '#8c564b'
_COULEUR_MAX_MAINTIEN = 'teal'

_ORDRE_POINTS = [f't{i}' for i in range(11)]
_PALETTE = plt.get_cmap('tab20').colors
COULEURS_POINTS = {cle: _PALETTE[i] for i, cle in enumerate(_ORDRE_POINTS)}
LABELS_POINTS = {
    't0': 'Début montée',
    't1': 'Fin montée 1',
    't2': 'Regroupement 1',
    't3': 'Maintien 1',
    't4': 'Fin montée 2',
    't5': 'Regroupement 2',
    't6': 'Maintien 2',
    't7': 'Sortie 2',
    't8': 'Fin descente 3',
    't9': 'Regroupement 3',
    't10': 'Fin descente 4',
}

# Chaque sous-phase = (clé_debut, clé_fin, nom_affiché, fonction de détail
# affichée dans le titre du graphique correspondant). Une sous-phase par
# graphique : c'est ce qui évite de mélanger montée et maintien dans un
# même graphique (voir hypothèse 11).
_SOUS_PHASES = [
    ('t0', 't1', "Montée 1", lambda c: f"vitesse {fmt_vitesse_h(c['vitesse_montee_1'])}"),
    ('t1', 't2', "Regroupement 1", lambda c: f"durée {fmt_duree(c['duree_regroupement_1_min'])}"),
    ('t2', 't3', "Maintien 1", lambda c: f"durée {fmt_duree(c['duree_maintien_1_min'])}"),
    ('t3', 't4', "Montée 2", lambda c: f"vitesse {fmt_vitesse_h(c['vitesse_montee_2'])}"),
    ('t4', 't5', "Regroupement 2", lambda c: f"durée {fmt_duree(c['duree_regroupement_2_min'])}"),
    ('t5', 't6', "Maintien 2", lambda c: (f"durée {fmt_duree(c['duree_maintien_2_min'])}, "
                                            f"max {fmt_temperature(c['temp_maintien_2'])}")),
    ('t6', 't7', "Sortie 2", lambda c: f"durée {fmt_duree(c['duree_sortie_2_min'])}"),
    ('t7', 't8', "Descente 3", lambda c: f"vitesse {fmt_vitesse_h(c['vitesse_descente_3'])}"),
    ('t8', 't9', "Regroupement 3", lambda c: f"durée {fmt_duree(c['duree_regroupement_3_min'])}"),
    ('t9', 't10', "Descente 4", lambda c: f"vitesse {fmt_vitesse_h(c['vitesse_descente_4'])}"),
]


def _tracer_courbes_fenetre(ax, dates, voies, t_min, t_max):
    """Trace, dans la fenêtre [t_min, t_max], toutes les voies retenues en
    gris (une seule entrée de légende)."""
    premiere_grise = True
    for nom, valeurs in voies.items():
        xs, ys = [], []
        for t, v in zip(dates, valeurs):
            if t_min <= t <= t_max:
                xs.append(t)
                ys.append(v)
        if xs:
            label = "Voies retenues" if premiere_grise else None
            ax.plot(xs, ys, color=_COULEUR_GRISE, linewidth=0.9, alpha=0.85, zorder=1, label=label)
            premiere_grise = False


def _annoter_point(ax, instant, valeur, texte, couleur, x_range, y_range, label=None):
    """Place un marqueur + une étiquette pour un point d'intérêt. La
    position de l'étiquette (haut/bas, gauche/droite) est choisie
    automatiquement selon la position du point dans (x_range, y_range),
    pour rester à l'intérieur du graphique et limiter le risque de
    chevauchement entre étiquettes voisines. x_range peut contenir des
    datetime ou des floats matplotlib (ex. ax.get_xlim()) : les deux sont
    normalisés avant comparaison."""
    def _num(x):
        return mdates.date2num(x) if isinstance(x, datetime) else x

    x_min, x_max = x_range
    y_min, y_max = y_range
    a_gauche = _num(instant) <= (_num(x_min) + _num(x_max)) / 2
    au_dessus = valeur <= (y_min + y_max) / 2
    dx = 8 if a_gauche else -8
    dy = 8 if au_dessus else -8
    ha = 'left' if a_gauche else 'right'
    va = 'bottom' if au_dessus else 'top'
    ax.plot([instant], [valeur], marker='o', color=couleur, markersize=6,
             markeredgecolor='white', markeredgewidth=0.7, zorder=6, label=label,
             clip_on=True)
    ax.annotate(texte, xy=(instant, valeur), xytext=(dx, dy), textcoords='offset points',
                 color=couleur, fontsize=7.5, fontweight='bold', ha=ha, va=va, clip_on=True,
                 bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                           edgecolor=couleur, linewidth=1, alpha=0.9))


def _tracer_zoom_phase(ax, dates, voies, config, points, titre):
    """points : liste de tuples (instant, valeur, couleur, label, ligne_horizontale, texte)."""
    if not points or any(p[0] is None or p[1] is None for p in points):
        ax.set_title(f"{titre} : seuils non atteints", fontweight='bold', pad=8, fontsize=9)
        ax.text(0.5, 0.5, "Seuils non franchis", ha='center', va='center',
                 transform=ax.transAxes, color='dimgray', fontsize=8)
        ax.set_xticks([])
        ax.set_yticks([])
        return

    marge_t = timedelta(minutes=config['zoom_marge_temps'])
    marge_T = config['zoom_marge_temperature']
    instants = [p[0] for p in points]
    valeurs = [p[1] for p in points]
    t_min, t_max = min(instants) - marge_t, max(instants) + marge_t
    y_min, y_max = min(valeurs) - marge_T, max(valeurs) + marge_T
    x_range, y_range = (t_min, t_max), (y_min, y_max)

    _tracer_courbes_fenetre(ax, dates, voies, t_min, t_max)

    for instant, valeur, couleur, label, ligne_horizontale, texte in points:
        if ligne_horizontale:
            ax.axhline(valeur, color=couleur, linestyle='--', linewidth=1.2)
        _annoter_point(ax, instant, valeur, texte, couleur, x_range, y_range, label=label)

    ax.set_xlim(t_min, t_max)
    ax.set_ylim(y_min, y_max)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
    for lbl in ax.get_xticklabels():
        lbl.set_rotation(20)
        lbl.set_ha('right')
    ax.tick_params(labelsize=7)
    ax.set_ylabel("T (°C)", fontsize=8)
    ax.grid(True, alpha=0.3)
    ax.set_title(titre, fontweight='bold', pad=8, fontsize=9)


def _construire_points_phase(cycle, cle_debut, cle_fin):
    points = []
    for cle in (cle_debut, cle_fin):
        instant = cycle[cle]
        seuil = cycle[f'seuil_{cle}']
        texte = _texte_point(instant, seuil, cycle[f'voie_{cle}'])
        points.append((instant, seuil, COULEURS_POINTS[cle], LABELS_POINTS[cle], True, texte))
    return points


def tracer_graphique(fichier, titre, dates, voies, cycle, config):
    fig = plt.figure(figsize=tuple(config['figure_taille']))
    n_phases = len(_SOUS_PHASES)
    n_lignes = (n_phases + 1) // 2  # 2 colonnes de graphiques de zoom
    gs = fig.add_gridspec(n_lignes, 3, width_ratios=[1.6, 1, 1], hspace=0.9, wspace=0.3)
    ax_gauche = fig.add_subplot(gs[:, 0])

    # --- Graphique global (gauche, pleine hauteur) ---
    premiere_grise = True
    for nom, valeurs in voies.items():
        label = "Voies retenues" if premiere_grise else None
        ax_gauche.plot(dates, valeurs, color=_COULEUR_GRISE, linewidth=0.9,
                        alpha=0.85, label=label, zorder=1)
        premiere_grise = False

    if cycle['temp_initiale'] is not None and dates:
        x_range_gauche = ax_gauche.get_xlim()
        y_range_gauche = ax_gauche.get_ylim()
        _annoter_point(ax_gauche, dates[0], cycle['temp_initiale'],
                        f"{dates[0].strftime('%H:%M:%S')}\n{fmt_temperature(cycle['temp_initiale'])}",
                        _COULEUR_INITIALE, x_range_gauche, y_range_gauche,
                        label="Température initiale (moyenne)")

    for cle in _ORDRE_POINTS:
        instant = cycle[cle]
        if instant is not None:
            ax_gauche.axvline(instant, color=COULEURS_POINTS[cle], linestyle='--',
                                linewidth=1.2, alpha=0.8, label=LABELS_POINTS[cle])

    ax_gauche.set_xlabel("Temps")
    ax_gauche.set_ylabel("Température (°C)")
    ax_gauche.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m/%d %H:%M:%S'))
    for lbl in ax_gauche.get_xticklabels():
        lbl.set_rotation(30)
        lbl.set_ha('right')
    ax_gauche.grid(True, alpha=0.3)
    ax_gauche.legend(loc='upper left', ncol=2, fontsize=6.5, framealpha=0.9)

    # --- Une sous-phase par graphique, répartie sur 2 colonnes ---
    for idx, (cle_debut, cle_fin, nom_phase, detail_fn) in enumerate(_SOUS_PHASES):
        ligne = idx % n_lignes
        colonne = 1 + idx // n_lignes
        ax = fig.add_subplot(gs[ligne, colonne])

        points = _construire_points_phase(cycle, cle_debut, cle_fin)
        if nom_phase == "Maintien 2" and cycle['instant_maintien_2'] is not None \
                and cycle['temp_maintien_2'] is not None:
            points.append((
                cycle['instant_maintien_2'], cycle['temp_maintien_2'], _COULEUR_MAX_MAINTIEN,
                "Max maintien 2", True,
                _texte_point(cycle['instant_maintien_2'], cycle['temp_maintien_2'], cycle['voie_maintien_2'])
            ))

        titre_phase = f"{nom_phase} — {detail_fn(cycle)}"
        _tracer_zoom_phase(ax, dates, voies, config, points, titre_phase)

    fig.suptitle(titre or os.path.basename(fichier), fontsize=14, fontweight='bold', y=0.995)
    fig.text(0.5, 0.975,
              f"T° initiale (moy.) : {fmt_temperature(cycle['temp_initiale'])}   |   "
              f"T° maintien 2 (max) : {fmt_temperature(cycle['temp_maintien_2'])}",
              ha='center', fontsize=9, color='dimgray')

    fig.subplots_adjust(top=0.93, bottom=0.10, left=0.05, right=0.97)
    return fig


# =========================================================================
# Fichier de résultats (formalisme strict, voir hypothèse 8)
# =========================================================================

# (position_tabulation, libellé) pour chacune des 12 valeurs, dans l'ordre.
# 'avant' -> "\tvaleur" ; 'apres' -> "valeur\t" (reproduit exactement le
# formalisme de l'exemple fourni par l'utilisateur ; fautes de frappe
# "maintient"/"vitese" corrigées en "maintien"/"vitesse").
_FORMALISME_RESULTATS = [
    ('avant', "température initiale"),
    ('avant', "vitesse de montée 1"),
    ('avant', "durée de regroupement 1"),
    ('avant', "durée du maintien 1"),
    ('apres', "Vitesse de montée 2"),
    ('apres', "Durée de regroupement 2"),
    ('apres', "Température de maintien 2"),
    ('apres', "Durée du maintien 2"),
    ('apres', "Durée de la sortie 2"),
    ('apres', "Vitesse de descente 3"),
    ('avant', "Durée du regroupement 3"),
    ('avant', "vitesse de descente 4"),
]


def ecrire_resultats(fichier_sortie, cycle):
    valeurs = [
        fmt_temperature(cycle['temp_initiale']),
        fmt_vitesse_h(cycle['vitesse_montee_1']),
        fmt_duree(cycle['duree_regroupement_1_min']),
        fmt_duree(cycle['duree_maintien_1_min']),
        fmt_vitesse_h(cycle['vitesse_montee_2']),
        fmt_duree(cycle['duree_regroupement_2_min']),
        fmt_temperature(cycle['temp_maintien_2']),
        fmt_duree(cycle['duree_maintien_2_min']),
        fmt_duree(cycle['duree_sortie_2_min']),
        fmt_vitesse_h(cycle['vitesse_descente_3']),
        fmt_duree(cycle['duree_regroupement_3_min']),
        fmt_vitesse_h(cycle['vitesse_descente_4']),
    ]

    lignes_libelles = []
    lignes_valeurs = []
    for (position, libelle), valeur in zip(_FORMALISME_RESULTATS, valeurs):
        if position == 'avant':
            lignes_libelles.append(f"\t{libelle}")
            lignes_valeurs.append(f"\t{valeur}")
        else:
            lignes_libelles.append(f"{libelle}\t")
            lignes_valeurs.append(f"{valeur}\t")

    with open(fichier_sortie, 'w', encoding='utf-8', newline='') as f:
        f.write("\n".join(lignes_libelles + lignes_valeurs) + "\n")


# =========================================================================
# Traitement d'un fichier
# =========================================================================

def analyser_fichier(fichier, config):
    print(f"📄 Lecture de {fichier}...")
    titre, dates, valeurs_brutes = lire_donnees_excel(fichier)
    print(f"   {len(dates)} échantillons, {len(valeurs_brutes)} voies détectées.")

    voies = filtrer_voies(valeurs_brutes, config['CH_ignore_dT'])
    voies_ignorees = sorted(set(valeurs_brutes) - set(voies))
    if voies_ignorees:
        print(f"   ⏭️  Voies ignorées (-OVER ou ΔT ≤ {config['CH_ignore_dT']}°) : "
              f"{', '.join(voies_ignorees)}")
    if not voies:
        print(f"   ⚠️ Aucune voie exploitable pour {fichier}, fichier ignoré.")
        return
    print(f"   ✅ Voies retenues : {', '.join(voies)}")

    cycle = calculer_cycle(dates, voies, config)
    print(f"   T° initiale (moyenne) = {fmt_temperature(cycle['temp_initiale'])}")

    etapes = [
        ('t0', 'début montée'), ('t1', 'fin montée 1'), ('t2', 'regroupement 1'),
        ('t3', 'maintien 1'), ('t4', 'fin montée 2'), ('t5', 'regroupement 2'),
        ('t6', 'maintien 2'), ('t7', 'sortie 2'), ('t8', 'fin descente 3'),
        ('t9', 'regroupement 3'), ('t10', 'fin descente 4'),
    ]
    complet = True
    for cle, nom_etape in etapes:
        if cycle[cle] is None:
            print(f"   ⚠️ Étape non atteinte : {nom_etape} ({cle}).")
            complet = False
            break
        print(f"   ⏱️  {nom_etape} ({cle}) = {cycle[cle]} — voie : {cycle[f'voie_{cle}']}")
    if complet:
        print("   ✅ Cycle complet détecté (t0 → t10).")

    fig = tracer_graphique(fichier, titre, dates, voies, cycle, config)

    base, _ = os.path.splitext(fichier)
    fichier_image = f"{base}_{config['suffixe']}.{FORMAT_IMAGE}"
    fig.savefig(fichier_image, dpi=config['dpi'], bbox_inches='tight')
    plt.close(fig)
    print(f"✅ Graphique généré -> {fichier_image}")

    fichier_resultats = f"{base}_{config['suffixe_resultats']}.{FORMAT_RESULTATS}"
    ecrire_resultats(fichier_resultats, cycle)
    print(f"✅ Résultats générés -> {fichier_resultats}")


# =========================================================================
# Configuration (YAML) et arguments en ligne de commande
# =========================================================================

DEFAULTS = {
    'CH_ignore_dT': 20.0,
    'temp_delta_ambiant': 10.0,
    'temp_palier_1': 450.0,
    'temp_delta_1': 15.0,
    'temp_palier_2': 550.0,
    'temp_delta_2': 15.0,
    'temp_palier_3': 300.0,
    'temp_delta_3': 15.0,
    'temp_palier_4': 50.0,
    'zoom_marge_temperature': 20.0,
    'zoom_marge_temps': 5.0,
    'suffixe': 'Analyse',
    'suffixe_resultats': 'Resultats',
    'dpi': 150,
    'figure_taille': [20, 14],
}


def lire_parametres_yaml(fichier):
    if not os.path.exists(fichier):
        raise FileNotFoundError(f"Le fichier {fichier} est introuvable.")
    with open(fichier, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f) or {}


def construire_config(args):
    fichier_config = args.config or DEFAULT_CONFIG_FILE
    yaml_params = {}
    if os.path.exists(fichier_config):
        print(f"⚙️ Lecture des paramètres depuis {fichier_config}...")
        yaml_params = lire_parametres_yaml(fichier_config)
    elif args.config:
        raise FileNotFoundError(f"Le fichier {fichier_config} est introuvable.")
    else:
        print(f"⚙️ Aucun fichier {fichier_config} trouvé, utilisation des valeurs par défaut.")

    def valeur(cli_val, cle_yaml):
        if cli_val is not None:
            return cli_val
        return yaml_params.get(cle_yaml, DEFAULTS[cle_yaml])

    config = {
        'CH_ignore_dT': float(valeur(args.ch_ignore_dt, 'CH_ignore_dT')),
        'temp_delta_ambiant': float(valeur(args.temp_delta_ambiant, 'temp_delta_ambiant')),
        'temp_palier_1': float(valeur(args.temp_palier_1, 'temp_palier_1')),
        'temp_delta_1': float(valeur(args.temp_delta_1, 'temp_delta_1')),
        'temp_palier_2': float(valeur(args.temp_palier_2, 'temp_palier_2')),
        'temp_delta_2': float(valeur(args.temp_delta_2, 'temp_delta_2')),
        'temp_palier_3': float(valeur(args.temp_palier_3, 'temp_palier_3')),
        'temp_delta_3': float(valeur(args.temp_delta_3, 'temp_delta_3')),
        'temp_palier_4': float(valeur(args.temp_palier_4, 'temp_palier_4')),
        'zoom_marge_temperature': float(
            yaml_params.get('zoom_marge_temperature', DEFAULTS['zoom_marge_temperature'])
        ),
        'zoom_marge_temps': float(
            yaml_params.get('zoom_marge_temps', DEFAULTS['zoom_marge_temps'])
        ),
        'suffixe': str(yaml_params.get('suffixe', DEFAULTS['suffixe'])),
        'suffixe_resultats': str(yaml_params.get('suffixe_resultats', DEFAULTS['suffixe_resultats'])),
        'dpi': int(yaml_params.get('dpi', DEFAULTS['dpi'])),
        'figure_taille': yaml_params.get('figure_taille', DEFAULTS['figure_taille']),
    }

    return config


def main():
    parser = argparse.ArgumentParser(
        description="Analyse du cycle de détensionnement (montée / regroupement / maintien x2 + "
                    "refroidissement x2) d'un ou plusieurs fichiers Excel et génération "
                    "d'images + fichiers de résultats."
    )
    parser.add_argument("fichiers", nargs="*", help="Un ou plusieurs fichiers Excel (.xlsx) à traiter")
    parser.add_argument("--config", type=str, help="Chemin du fichier YAML de configuration")
    parser.add_argument("--ch-ignore-dt", dest="ch_ignore_dt", type=float,
                         help="Seuil d'amplitude (°C) en dessous duquel une voie est ignorée")
    parser.add_argument("--temp-delta-ambiant", dest="temp_delta_ambiant", type=float,
                         help="Écart (°C) au-dessus de la température initiale marquant le début de la montée")
    parser.add_argument("--temp-palier-1", dest="temp_palier_1", type=float, help="Température du palier 1")
    parser.add_argument("--temp-delta-1", dest="temp_delta_1", type=float, help="Demi-tolérance (°C) du palier 1")
    parser.add_argument("--temp-palier-2", dest="temp_palier_2", type=float, help="Température du palier 2")
    parser.add_argument("--temp-delta-2", dest="temp_delta_2", type=float, help="Demi-tolérance (°C) du palier 2")
    parser.add_argument("--temp-palier-3", dest="temp_palier_3", type=float, help="Température du palier 3 (refroidissement)")
    parser.add_argument("--temp-delta-3", dest="temp_delta_3", type=float, help="Demi-tolérance (°C) du palier 3")
    parser.add_argument("--temp-palier-4", dest="temp_palier_4", type=float, help="Température de sortie finale (palier 4)")

    args = parser.parse_args()

    try:
        config = construire_config(args)
    except (FileNotFoundError, ValueError) as e:
        print(f"❌ Erreur de configuration : {e}")
        sys.exit(1)

    print("=== Paramètres d'analyse ===")
    for cle in ('CH_ignore_dT', 'temp_delta_ambiant', 'temp_palier_1', 'temp_delta_1',
                'temp_palier_2', 'temp_delta_2', 'temp_palier_3', 'temp_delta_3',
                'temp_palier_4', 'zoom_marge_temperature', 'zoom_marge_temps',
                'suffixe', 'suffixe_resultats'):
        print(f"{cle} = {config[cle]}")
    print()

    if args.fichiers:
        print(f"🔍 {len(args.fichiers)} fichier(s) Excel fourni(s) :")
        for fichier in args.fichiers:
            try:
                analyser_fichier(fichier, config)
            except Exception as e:
                print(f"❌ Erreur lors du traitement de {fichier} : {e}")
    else:
        print("Traitement des fichiers .xlsx (hors fichiers déjà suffixés) du répertoire actuel.")
        suffixe_bas = f"_{config['suffixe']}".lower()
        suffixe_resultats_bas = f"_{config['suffixe_resultats']}".lower()
        fichiers_xlsx = [
            f for f in os.listdir('.')
            if f.lower().endswith('.xlsx') and not f.lower()[:-5].endswith(suffixe_bas)
            and not f.lower()[:-5].endswith(suffixe_resultats_bas)
        ]
        if not fichiers_xlsx:
            print("⚠️ Aucun fichier .xlsx trouvé dans le répertoire actuel.")
            return

        print(f"🔍 {len(fichiers_xlsx)} fichier(s) .xlsx trouvé(s) :")
        for f in fichiers_xlsx:
            print(f" → {f}")
        print("\n⏳ Traitement en cours...\n")

        for fichier in fichiers_xlsx:
            try:
                analyser_fichier(fichier, config)
            except Exception as e:
                print(f"❌ Erreur lors du traitement de {fichier} : {e}")

        print("\n✅ Tous les fichiers .xlsx ont été traités.")


if __name__ == "__main__":
    main()
