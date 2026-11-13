"""
Agent de sourcing build-up : fontaines a eau (France).

Pipeline en 4 etapes, inspire de la methode "Rebuild / Maintain / Source" :
  1. SIGNAL    : avis d'attribution de marches publics (BOAMP, open data) sur la
                 location et la maintenance de fontaines a eau -> qui gagne des
                 contrats recurrents, ou, et quand.
  2. TRI       : on retire les grands groupes et les entites du groupe acquereur,
                 on garde les independants.
  3. ENRICHIR  : registre national des entreprises (API Recherche d'entreprises) :
                 SIREN, chiffre d'affaires, resultat, effectif, dirigeants.
  4. SCORER    : grille ponderee modifiable dans Excel, l'humain valide.

Usage :
  python sourcing_agent.py --cache          # rejoue avec les donnees en cache (data/)
  python sourcing_agent.py --live           # interroge les API publiques (internet requis)
  python sourcing_agent.py --live --out ma_liste.xlsx

Aucune cle d'API n'est necessaire : les deux sources sont publiques et gratuites.
"""
import argparse
import json
import re
import time
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
DATA = HERE / "data"

# --------------------------------------------------------------------------
# 0. These d'investissement (a ajuster)
# --------------------------------------------------------------------------
THESE = {
    "secteur": "Location, installation et maintenance de fontaines a eau (B2B et acheteurs publics)",
    "ca_min_france": 1_000_000,
    "ca_max": 10_000_000,
    "mots_cles_boamp": ["fontaines a eau", "fontaine a eau"],
}

# Entites du groupe acquereur : jamais des cibles
# Remplacer par les noms des societes de votre groupe acquereur.
GROUPE_ACQUEREUR = ["FILIALE ACQUEREUR"]

# Grands groupes et acteurs hors cible (taille ou metier)
GRANDS_GROUPES = {
    "ELIS": "Elis (groupe cote, services aux entreprises)",
    "MAJ": "Elis (filiale MAJ)",
    "RLST": "Elis (filiale RLST)",
    "INITIAL": "Rentokil Initial (groupe cote)",
    "CULLIGAN": "Culligan (groupe international)",
    "CHATEAU": "Chateau d'Eau (groupe Culligan)",
    "WATERLOGIC": "Waterlogic (groupe Culligan)",
    "PIERRETTE": "Pierrette T.B.A. (grande entreprise, 68 M EUR de CA, location de linge)",
    "MAXICOFFEE": "Maxicoffee (distribution automatique, metier adjacent)",
    "SOFODA": "Sofoda-Vitdis (distribution automatique, metier adjacent)",
    "CUP SERVICE": "Cup Service (distribution automatique, metier adjacent)",
}

NAF_PERTINENTS = {"77.39Z", "77.29Z", "46.69C", "46.34Z", "33.12Z", "28.25Z", "47.99B"}

EFFECTIF = {"00": "0", "01": "1-2", "02": "3-5", "03": "6-9", "11": "10-19", "12": "20-49",
            "21": "50-99", "22": "100-199", "31": "200-249", "32": "250-499", "41": "500-999"}


def norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.upper().replace("'", " ")).strip()


def classer(nom: str):
    n = norm(nom)
    for g in GROUPE_ACQUEREUR:
        if norm(g) in n:
            return "Groupe acquereur", "Entite du groupe acquereur : exclue"
    for k, lib in GRANDS_GROUPES.items():
        if norm(k) in n:
            return "Grand groupe / hors cible", lib
    return "Independant", ""


# --------------------------------------------------------------------------
# 1. SIGNAL : marches publics (BOAMP)
# --------------------------------------------------------------------------
BOAMP_API = ("https://boamp-datadila.opendatasoft.com/api/explore/v2.1/catalog/"
             "datasets/boamp/records")


def boamp_live():
    import requests
    out = []
    for kw in THESE["mots_cles_boamp"]:
        offset = 0
        while True:
            params = {
                "where": f'search(objet,"{kw}") and nature="ATTRIBUTION"',
                "select": "idweb,objet,nomacheteur,dateparution,code_departement,titulaire",
                "order_by": "dateparution desc", "limit": 100, "offset": offset,
            }
            r = requests.get(BOAMP_API, params=params, timeout=30)
            r.raise_for_status()
            res = r.json().get("results", [])
            for x in res:
                if x.get("titulaire"):
                    dep = (x.get("code_departement") or [""])[0]
                    out.append([x["idweb"], x["dateparution"], x["nomacheteur"], dep,
                                sorted(set(x["titulaire"]))])
            if len(res) < 100 or offset >= 900:
                break
            offset += 100
    seen, uniq = set(), []
    for row in out:
        if row[0] not in seen:
            seen.add(row[0]); uniq.append(row)
    return uniq


def boamp_cache():
    return json.loads((DATA / "boamp_winners.json").read_text(encoding="utf-8"))


# --------------------------------------------------------------------------
# 3. ENRICHIR : registre national des entreprises
# --------------------------------------------------------------------------
RE_API = "https://recherche-entreprises.api.gouv.fr/search"


def registre_live(nom):
    import requests
    q = re.sub(r"\b(SAS|SARL|SA|SOCIETE)\b", "", norm(nom)).split("/")[0].strip()
    params = {"q": q, "per_page": 10, "minimal": "true", "include": "siege,dirigeants,finances",
              "etat_administratif": "A"}
    r = requests.get(RE_API, params=params, timeout=30)
    if r.status_code == 429:
        time.sleep(2); r = requests.get(RE_API, params=params, timeout=30)
    r.raise_for_status()
    for e in r.json().get("results", []):
        if e.get("activite_principale") in NAF_PERTINENTS:
            fin = e.get("finances") or {}
            an = max(fin) if fin else ""
            dirs = "; ".join(
                f"{d.get('prenoms','').title()} {d.get('nom','').title()} ({d.get('qualite','')}, {d.get('annee_de_naissance','?')})"
                if d.get("type_dirigeant") == "personne physique"
                else f"{d.get('qualite','')} : {d.get('denomination','')}"
                for d in e.get("dirigeants", []) if "Commissaire" not in (d.get("qualite") or ""))
            return {"siren": e["siren"], "nom": e["nom_complet"], "ville": e["siege"].get("libelle_commune"),
                    "dep": e["siege"].get("departement"), "naf": e["activite_principale"],
                    "effectif": e.get("tranche_effectif_salarie"), "ca": (fin.get(an) or {}).get("ca") or None,
                    "ca_annee": an, "resultat_net": (fin.get(an) or {}).get("resultat_net"),
                    "dirigeants": dirs, "creation": (e.get("date_creation") or "")[:4],
                    "etablissements_ouverts": e.get("nombre_etablissements_ouverts"),
                    "note": "Correspondance automatique sur le nom et le code NAF : a verifier."}
    return None


def registre_cache():
    return json.loads((DATA / "registry_cache.json").read_text(encoding="utf-8"))


# --------------------------------------------------------------------------
# 4. SCORER + Excel
# --------------------------------------------------------------------------
# Notes 0 a 5 proposees par l'analyste (a challenger), avec justification.
NOTES_ANALYSTE = {
    "AXO": ([5, 5, 3, 4, 2, 3],
            "Loueur-mainteneur multi-sites ; 8 marches publics gagnes en 2023-2026 dans 6 regions ; "
            "resultat net ~6 % du CA ; synergie possible avec la gamme de l'acquereur ; holding au capital a comprendre."),
    "BEHRING": ([4, 3, 1, 4, 4, 4],
                "Grossiste-installateur, clients hospitaliers ; perte nette de 1,3 M EUR en 2025 a expliquer ; "
                "dirigeants nes autour de 1960 : la question de la transmission merite d'etre posee ; profil plus risque."),
    "AQUAOCEANE / M2P": ([4, 3, 2, 3, 3, 4],
                         "Petit acteur normand (6 a 9 salaries) titulaire de marches CPAM et France Travail ; "
                         "densification Hauts-de-France et Normandie ; CA a obtenir."),
}
CRITERES = ["Fit strategique", "Qualite clients et revenus", "Rentabilite et cash",
            "Synergies", "Transactionabilite / fondateur", "Prix et levier"]
POIDS = [0.25, 0.20, 0.15, 0.15, 0.15, 0.10]

ARIAL = "Arial"
H_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
BLUE = Font(name=ARIAL, size=10, color="0000FF")
BLACK = Font(name=ARIAL, size=10)
GREEN = Font(name=ARIAL, size=10, color="008000")
BOLD = Font(name=ARIAL, size=10, bold=True)
HDR = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
WRAP = Alignment(wrap_text=True, vertical="top")


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.alignment, c.border = HDR, H_FILL, Alignment(wrap_text=True, vertical="center"), THIN
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 32


def build_excel(signaux, registre, out_path, mode):
    wb = Workbook()

    # ---------- Signaux ----------
    ws_s = wb.active
    ws_s.title = "Signaux"
    cols = ["Ref. BOAMP", "Date", "Acheteur public", "Dep.", "Titulaire", "Type d'acteur", "Commentaire"]
    header(ws_s, 1, cols, [12, 11, 42, 6, 30, 24, 46])
    r = 2
    for idweb, date, acheteur, dep, titulaires in signaux:
        for t in titulaires:
            typ, com = classer(t)
            vals = [idweb, date, acheteur, dep, norm(t), typ, com]
            for j, v in enumerate(vals, 1):
                c = ws_s.cell(row=r, column=j, value=v)
                c.font, c.border = BLUE, THIN
            r += 1
    last_sig = r - 1
    ws_s.freeze_panes = "A2"
    ws_s.auto_filter.ref = f"A1:G{last_sig}"
    sig_rng = f"Signaux!$E$2:$E${last_sig}"

    # ---------- Long list ----------
    ws = wb.create_sheet("Long list", 0)
    ws["A1"] = "Long list scoree : independants de la location-maintenance de fontaines a eau (France)"
    ws["A1"].font = Font(name=ARIAL, size=13, bold=True)
    ws["A2"] = ("Notes de 0 a 5 en jaune : proposition de l'analyste, a challenger. Poids modifiables en ligne 4. "
                "Le score et le rang se recalculent.")
    ws["A2"].font = Font(name=ARIAL, size=9, italic=True)

    # poids
    ws["L3"] = "Poids"; ws["L3"].font = BOLD
    first_score_col = 12  # L
    for i, (crit, p) in enumerate(zip(CRITERES, POIDS)):
        c = ws.cell(row=4, column=first_score_col + i, value=p)
        c.font, c.fill, c.number_format, c.border = BLUE, INPUT_FILL, "0%", THIN
    sum_cell = ws.cell(row=4, column=first_score_col + 6,
                       value=f"=SUM({get_column_letter(first_score_col)}4:{get_column_letter(first_score_col+5)}4)")
    sum_cell.font, sum_cell.number_format = BLACK, "0%"
    ws.cell(row=3, column=first_score_col + 6, value="Total poids").font = BOLD

    hdr = ["Cible", "SIREN", "Ville", "Dep.", "Code NAF", "Effectif", "Annee",
           "CA (EUR)", "Resultat net (EUR)", "Marge nette", "Marches publics gagnes (2022-2026)"] + CRITERES + \
          ["Score /5", "Rang", "Dirigeants (registre)", "Lecture de l'analyste", "Statut", "Prochaine action"]
    header(ws, 5, hdr, [20, 11, 14, 6, 9, 9, 7, 13, 14, 10, 13] + [11] * 6 + [9, 6, 38, 60, 14, 34])

    cibles = [k for k in NOTES_ANALYSTE if k in registre] if mode == "cache" else \
        [k for k in registre if classer(k)[0] == "Independant"]
    first = 6
    for i, key in enumerate(cibles):
        rr = first + i
        e = registre[key]
        notes, lecture = NOTES_ANALYSTE.get(key, ([3] * 6, "A analyser."))
        row = [e["nom"], e["siren"], e["ville"], e["dep"], e["naf"], EFFECTIF.get(e["effectif"] or "", e["effectif"]),
               e["ca_annee"] or "n.d.", e["ca"], e["resultat_net"]]
        for j, v in enumerate(row, 1):
            c = ws.cell(row=rr, column=j, value=v)
            c.font, c.border = BLUE, THIN
        ws.cell(row=rr, column=8).number_format = '#,##0;(#,##0);"-"'
        ws.cell(row=rr, column=9).number_format = '#,##0;(#,##0);"-"'
        if e["ca"] is None:
            ws.cell(row=rr, column=8, value="non publie").font = BLUE
        c = ws.cell(row=rr, column=10, value=f'=IFERROR(I{rr}/H{rr},"n.d.")')
        c.font, c.number_format, c.border = BLACK, "0.0%;(0.0%);-", THIN
        # marches gagnes : on compte les lignes Signaux dont le titulaire contient le mot cle
        mot = norm(key).split("/")[0].split()[0]
        c = ws.cell(row=rr, column=11, value=f'=COUNTIF({sig_rng},"*{mot}*")')
        c.font, c.border = GREEN, THIN
        c.comment = Comment(f"Compte les avis BOAMP dont le titulaire contient '{mot}' (onglet Signaux).", "Agent")
        for k, n in enumerate(notes):
            c = ws.cell(row=rr, column=first_score_col + k, value=n)
            c.font, c.fill, c.border = BLUE, INPUT_FILL, THIN
        L, Q = get_column_letter(first_score_col), get_column_letter(first_score_col + 5)
        c = ws.cell(row=rr, column=18, value=f"=SUMPRODUCT({L}{rr}:{Q}{rr},${L}$4:${Q}$4)")
        c.font, c.number_format, c.border = BOLD, "0.00", THIN
        last = first + len(cibles) - 1
        c = ws.cell(row=rr, column=19, value=f"=RANK(R{rr},$R${first}:$R${last})")
        c.font, c.border = BLACK, THIN
        for j, v in [(20, e["dirigeants"]), (21, lecture + " " + e["note"]), (22, "A qualifier"),
                     (23, "Valider le perimetre avec l'acquereur, puis lire les comptes deposes")]:
            c = ws.cell(row=rr, column=j, value=v)
            c.font, c.alignment, c.border = BLUE if j < 22 else BLACK, WRAP, THIN
        ws.row_dimensions[rr].height = 75
    dv = DataValidation(type="list", formula1='"A qualifier,Approche,NDA,LOI,Ecartee"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"V{first}:V{first + len(cibles) - 1}")

    n = first + len(cibles) + 1
    ws.cell(row=n, column=1, value="Titulaires independants a identifier dans le registre (SIREN non trouve automatiquement)").font = BOLD
    indep = Counter()
    for _, _, _, _, tit in signaux:
        for t in tit:
            if classer(t)[0] == "Independant" and norm(t) not in {norm(k) for k in cibles}:
                indep[norm(t)] += 1
    for i, (nom, cnt) in enumerate(sorted(indep.items(), key=lambda x: -x[1])):
        ws.cell(row=n + 1 + i, column=1, value=nom).font = BLUE
        ws.cell(row=n + 1 + i, column=2, value=f"{cnt} marche(s)").font = BLACK
    ws.freeze_panes = "B6"

    # ---------- Concurrence ----------
    wc = wb.create_sheet("Concurrence")
    header(wc, 1, ["Acteur", "Type", "Attributions gagnees", "Part des attributions"], [34, 26, 18, 20])
    com_rng = f"Signaux!$G$2:$G${last_sig}"
    typ_rng = f"Signaux!$F$2:$F${last_sig}"
    acteurs = [  # (libelle, type, formule de comptage)
        ("Elis (Elis, MAJ, RLST)", "Grand groupe", f'=COUNTIF({com_rng},"Elis*")'),
        ("Culligan (Culligan, Chateau d'Eau, Waterlogic)", "Grand groupe", f'=COUNTIF({com_rng},"*Culligan*")'),
        ("Rentokil Initial", "Grand groupe", f'=COUNTIF({com_rng},"Rentokil*")'),
        ("Filiale de l'acquereur", "Groupe acquereur", f'=COUNTIF({typ_rng},"Groupe acquereur")'),
        ("AXO", "Independant", f'=COUNTIF({sig_rng},"AXO*")'),
        ("Aquafontaine et Aquaoceane", "Independants", f'=COUNTIF({sig_rng},"*AQUA*")'),
        ("Behring", "Independant", f'=COUNTIF({sig_rng},"BEHRING*")'),
    ]
    n_act = len(acteurs)
    tot = n_act + 4
    for i, (a, t, f) in enumerate(acteurs, 2):
        wc.cell(row=i, column=1, value=a).font = BLUE
        wc.cell(row=i, column=2, value=t).font = BLUE
        wc.cell(row=i, column=3, value=f).font = GREEN
    i_aut = n_act + 2
    wc.cell(row=i_aut, column=1, value="Autres (independants et metiers adjacents)").font = BLUE
    wc.cell(row=i_aut, column=2, value="Divers").font = BLUE
    wc.cell(row=i_aut, column=3, value=f"=C{tot}-SUM(C2:C{i_aut - 1})").font = BLACK
    for i in range(2, i_aut + 1):
        c = wc.cell(row=i, column=4, value=f"=C{i}/$C${tot}")
        c.font, c.number_format = BLACK, "0.0%"
    wc.cell(row=tot, column=1, value="Total des attributions (une ligne par titulaire)").font = BOLD
    wc.cell(row=tot, column=3, value=f"=COUNTA({sig_rng})").font = BOLD
    c = wc.cell(row=tot, column=4, value=f"=SUM(D2:D{i_aut})")
    c.font, c.number_format = BOLD, "0.0%"
    wc.cell(row=tot + 2, column=1,
            value="Lecture : un avis peut avoir plusieurs titulaires (lots). Les marches publics ne sont qu'un echantillon "
                  "du marche ; les parts ne sont pas des parts de marche.").font = Font(name=ARIAL, size=9, italic=True)

    # ---------- Lisez-moi ----------
    wr = wb.create_sheet("Lisez-moi", 0)
    wr.column_dimensions["A"].width = 120
    lignes = [
        ("Agent de sourcing build-up : prototype", Font(name=ARIAL, size=14, bold=True)),
        (f"These : {THESE['secteur']}. Taille visee : CA de 1 a 10 M EUR.", BLACK),
        ("", BLACK),
        ("Comment l'agent travaille", BOLD),
        ("1. Signal : il lit les avis d'attribution de marches publics (BOAMP, open data) sur la location et la maintenance "
         "de fontaines a eau. Un marche gagne est la preuve d'un contrat recurrent, d'une zone couverte et d'une capacite de service.", BLACK),
        ("2. Tri : il retire les grands groupes (Elis, Culligan et ses marques) et les entites du groupe acquereur.", BLACK),
        ("3. Enrichissement : il interroge le registre national des entreprises (API Recherche d'entreprises) : SIREN, "
         "chiffre d'affaires, resultat net, effectif, dirigeants et leur annee de naissance.", BLACK),
        ("4. Score : grille ponderee (fit, clients, rentabilite, synergies, transactionabilite, prix). "
         "Les notes sont des propositions : l'humain valide.", BLACK),
        ("", BLACK),
        ("Sources", BOLD),
        ("BOAMP open data : https://boamp-datadila.opendatasoft.com/explore/dataset/boamp/", BLACK),
        ("API Recherche d'entreprises : https://recherche-entreprises.api.gouv.fr/docs/", BLACK),
        ("Fiches entreprises : https://annuaire-entreprises.data.gouv.fr/ (remplacer par le SIREN dans l'adresse /entreprise/SIREN)", BLACK),
        ("", BLACK),
        ("Limites", BOLD),
        ("Les marches publics ne montrent qu'une partie du marche : les acteurs 100 % B2B prives n'y apparaissent pas. "
         "Prochaine source a brancher : annuaires professionnels et sites web (cartographie des zones d'intervention).", BLACK),
        ("Le chiffre d'affaires du registre est celui des comptes deposes ; il peut etre absent (comptes confidentiels). "
         "Le resultat net n'est pas l'EBITDA : il faut lire la liasse.", BLACK),
        ("La correspondance titulaire -> SIREN est automatique : chaque ligne doit etre verifiee avant tout contact.", BLACK),
        ("Aucun contact n'a ete pris avec ces entreprises.", BLACK),
        ("", BLACK),
        ("Legende", BOLD),
        ("Bleu : donnee saisie ou importee. Noir : formule. Vert : formule qui lit un autre onglet. "
         "Fond jaune : hypotheses a modifier (poids et notes).", BLACK),
        (f"Mode d'execution : {mode}. Donnees extraites le 1er novembre 2026.", BLACK),
    ]
    for i, (t, f) in enumerate(lignes, 1):
        c = wr.cell(row=i, column=1, value=t)
        c.font, c.alignment = f, Alignment(wrap_text=True, vertical="top")

    # ---------- Exclus ----------
    we = wb.create_sheet("Exclus")
    header(we, 1, ["Acteur", "Raison de l'exclusion", "Detail"], [24, 34, 90])
    excl = [("FILIALE ACQUEREUR", "Entite du groupe acquereur", "Jamais une cible"),
            ("ELIS / MAJ / RLST", "Grand groupe", "Groupe cote, services aux entreprises"),
            ("CULLIGAN / CHATEAU D'EAU / WATERLOGIC", "Grand groupe", "Leader du marche, acquereur concurrent"),
            ("INITIAL", "Grand groupe", "Rentokil Initial"),
            ("PIERRETTE T.B.A.", "Trop grand / metier adjacent", registre.get("PIERRETTE - T.B.A", {}).get("note", "")),
            ("MAXICOFFEE, SOFODA-VITDIS, CUP SERVICE", "Metier adjacent",
             "Distribution automatique : plutot des acquereurs concurrents que des cibles")]
    for i, row in enumerate(excl, 2):
        for j, v in enumerate(row, 1):
            c = we.cell(row=i, column=j, value=v)
            c.font, c.alignment, c.border = BLUE, WRAP, THIN

    for w in wb.worksheets:
        w.sheet_view.showGridLines = False
    wb.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--live", action="store_true")
    ap.add_argument("--cache", action="store_true")
    ap.add_argument("--out", default="Sourcing_fontaines_longlist.xlsx")
    a = ap.parse_args()
    mode = "live" if a.live else "cache"

    signaux = boamp_live() if a.live else boamp_cache()
    print(f"[1] {len(signaux)} avis d'attribution avec titulaire")

    titulaires = Counter(t for s in signaux for t in s[4])
    indep = [t for t in titulaires if classer(t)[0] == "Independant"]
    print(f"[2] {len(titulaires)} titulaires distincts, dont {len(indep)} independants")

    if a.live:
        registre = {}
        for t in indep:
            res = registre_live(t)
            if res:
                registre[t] = res
            time.sleep(0.3)
    else:
        registre = registre_cache()
    print(f"[3] {len(registre)} entreprises enrichies via le registre")

    out = build_excel(signaux, registre, HERE / a.out, mode)
    print(f"[4] Long list ecrite : {out}")


if __name__ == "__main__":
    main()
