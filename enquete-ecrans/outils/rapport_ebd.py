#!/usr/bin/env python3
"""Rapport de l'enquête Opération Écran, boulot, dodo, LVH Sofia.
Usage, python3 rapport_ebd.py data.json dossier_sortie
Produit Rapport_Enquete_Ecran_boulot_dodo.docx, la version PDF et synthese_direction.txt"""
import json, sys, os, re, subprocess, datetime, unicodedata
from collections import Counter, OrderedDict
from zoneinfo import ZoneInfo
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from scipy.stats import spearmanr
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC, OUT = sys.argv[1], sys.argv[2]
os.makedirs(OUT, exist_ok=True)
FIG = os.path.join(OUT, "_fig"); os.makedirs(FIG, exist_ok=True)
_HERE = os.path.dirname(os.path.abspath(__file__))
def logo(name):
    for c in (f"/home/claude/assets/logo_{name}.png", os.path.join(_HERE, "..", "assets", f"logo-{name}.png")):
        if os.path.exists(c): return c
    raise FileNotFoundError(name)
NAVY, RED, GREY, SOFT = "16256B", "C8102E", "5A6072", "F1F3F8"
TZ = ZoneInfo("Europe/Sofia")
NB = " "

for f in font_manager.findSystemFonts():
    if "Carlito" in f: font_manager.fontManager.addfont(f)
plt.rcParams["font.family"] = "Carlito"

D = json.load(open(SRC, encoding="utf8"))
E, P = D.get("Eleves", []), D.get("Parents", [])
NOW = datetime.datetime.now(TZ)

ORD = {
 "niveauE": ["6e","5e","4e","3e","2de","1re","Terminale"],
 "niveauP": ["Maternelle","Élémentaire, du CP au CM2","6e","5e","4e","3e","2de","1re","Terminale"],
 "temps": ["Moins d'1 h","1 à 2 h","2 à 3 h","3 à 4 h","4 à 5 h","Plus de 5 h","Je ne sais pas"],
 "sommeil": ["Moins de 6 h","6 à 7 h","7 à 8 h","8 à 9 h","Plus de 9 h"],
 "coucher": ["Avant 21 h","Entre 21 h et 22 h","Entre 22 h et 23 h","Entre 23 h et minuit","Après minuit"],
 "freq": ["Jamais","Rarement","Parfois","Souvent","Toujours","Très souvent"],
 "accord": ["Oui","Plutôt oui","Plutôt non","Non"],
 "effet": ["Pas du tout","Un peu","Assez","Beaucoup"],
 "resultats": ["Très satisfaisants","Satisfaisants","Moyens","Insuffisants","Je ne sais pas"],
 "dernier": ["Dans les 15 minutes avant de dormir","Entre 15 et 30 minutes avant","Entre 30 minutes et 1 heure avant","Plus d'une heure avant"],
 "regleLycee": ["Utilisable librement, sauf en cours","Utilisable seulement à certains moments ou dans certains lieux","Rangé toute la journée, éteint ou au casier","Sans avis"],
 "regleMaison": ["Oui, et elles sont respectées","Oui, mais pas toujours respectées","Non"],
}
NUM = {
 "temps": {"Moins d'1 h":0.5,"1 à 2 h":1.5,"2 à 3 h":2.5,"3 à 4 h":3.5,"4 à 5 h":4.5,"Plus de 5 h":5.5},
 "sommeil": {"Moins de 6 h":5.5,"6 à 7 h":6.5,"7 à 8 h":7.5,"8 à 9 h":8.5,"Plus de 9 h":9.5},
 "freq": {"Jamais":0,"Rarement":1,"Parfois":2,"Souvent":3,"Toujours":4,"Très souvent":4},
 "resultats": {"Très satisfaisants":4,"Satisfaisants":3,"Moyens":2,"Insuffisants":1},
 "oui": {"Oui":1,"Non":0},
 "regleMaison": {"Oui, et elles sont respectées":2,"Oui, mais pas toujours respectées":1,"Non":0},
}
MIN_CORR = 15

# ---------- calculs ----------
def vals(rows, key):
    return [str(r.get(key, "")).strip() for r in rows if str(r.get(key, "")).strip()]
def counts(rows, key, multi=False, order=None):
    c = Counter(); n = 0
    for r in rows:
        v = str(r.get(key, "")).strip()
        if not v: continue
        n += 1
        for x in (v.split(" | ") if multi else [v]): c[x] += 1
    keys = list(c.keys())
    if order: keys.sort(key=lambda k: order.index(k) if k in order else 99)
    else: keys.sort(key=lambda k: -c[k])
    return n, [(k, c[k]) for k in keys]
def share(rows, key, targets):
    n, it = counts(rows, key)
    k = sum(v for x, v in it if x in targets)
    return k, n
def mean(rows, key, m):
    v = [m[r.get(key)] for r in rows if r.get(key) in m]
    return (sum(v)/len(v), len(v)) if v else (None, 0)
def wellbeing(r):
    f = NUM["freq"]; p = []
    for k in ("be_repose","be_concentration","be_humeur"):
        if r.get(k) in f: p.append(f[r[k]])
    for k in ("be_stress","be_seul"):
        if r.get(k) in f: p.append(4 - f[r[k]])
    return sum(p)/len(p)*25 if len(p) >= 3 else None
def norm_code(c):
    c = unicodedata.normalize("NFD", str(c or "").lower())
    c = "".join(ch for ch in c if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9а-я]", "", c)
def dyads():
    by = {}
    for p in P:
        c = norm_code(p.get("code_famille"))
        if len(c) >= 3: by.setdefault(c, []).append(p)
    out = []
    for e in E:
        c = norm_code(e.get("code_famille"))
        if len(c) >= 3 and c in by:
            for p in by[c]: out.append((e, p))
    return out
DY = dyads()

def corr(pairs):
    pairs = [(a, b) for a, b in pairs if a is not None and b is not None]
    n = len(pairs)
    if n < MIN_CORR: return None, None, n
    x, y = zip(*pairs)
    if len(set(x)) < 2 or len(set(y)) < 2: return None, None, n
    r, p = spearmanr(x, y)
    return float(r), float(p), n
g = lambda key, m: (lambda r: m.get(r.get(key)))
def pairs(rows, fx, fy): return [(fx(r), fy(r)) for r in rows]
CORRS = [
 ("Élèves","Temps d'écran un jour de classe et durée de sommeil", pairs(E, g("temps_semaine",NUM["temps"]), g("duree_sommeil",NUM["sommeil"]))),
 ("Élèves","Temps d'écran et concentration en classe", pairs(E, g("temps_semaine",NUM["temps"]), g("be_concentration",NUM["freq"]))),
 ("Élèves","Temps d'écran et résultats perçus", pairs(E, g("temps_semaine",NUM["temps"]), g("resultats",NUM["resultats"]))),
 ("Élèves","Temps d'écran et indice de bien-être", pairs(E, g("temps_semaine",NUM["temps"]), wellbeing)),
 ("Élèves","Durée de sommeil et concentration en classe", pairs(E, g("duree_sommeil",NUM["sommeil"]), g("be_concentration",NUM["freq"]))),
 ("Élèves","Téléphone dans la chambre et durée de sommeil", pairs(E, g("tel_chambre",NUM["oui"]), g("duree_sommeil",NUM["sommeil"]))),
 ("Élèves","Téléphone pendant les devoirs et résultats perçus", pairs(E, g("tel_devoirs",NUM["freq"]), g("resultats",NUM["resultats"]))),
 ("Élèves","Temps d'écran des parents perçu et temps d'écran de l'élève", pairs(E, g("parents_temps",NUM["temps"]), g("temps_semaine",NUM["temps"]))),
 ("Élèves","Parents absorbés par leur téléphone et bien-être de l'élève", pairs(E, g("parents_technoference",NUM["freq"]), wellbeing)),
 ("Élèves","Règles à la maison et temps d'écran", pairs(E, g("regles_maison",NUM["regleMaison"]), g("temps_semaine",NUM["temps"]))),
 ("Parents","Temps d'écran du parent et temps d'écran estimé de l'enfant", pairs(P, g("temps_parent",NUM["temps"]), g("temps_semaine",NUM["temps"]))),
 ("Parents","Temps d'écran du parent et résultats de l'enfant", pairs(P, g("temps_parent",NUM["temps"]), g("enfant_resultats",NUM["resultats"]))),
 ("Parents","Temps d'écran du parent et fatigue de l'enfant le matin", pairs(P, g("temps_parent",NUM["temps"]), g("enfant_fatigue",NUM["freq"]))),
 ("Parents","Écrans pendant les repas et tensions familiales", pairs(P, g("ecrans_repas",NUM["freq"]), g("tensions",NUM["freq"]))),
 ("Familles appariées","Temps d'écran du parent et temps d'écran déclaré par l'élève", [(NUM["temps"].get(p.get("temps_parent")), NUM["temps"].get(e.get("temps_semaine"))) for e, p in DY]),
 ("Familles appariées","Temps d'écran du parent et résultats déclarés par l'élève", [(NUM["temps"].get(p.get("temps_parent")), NUM["resultats"].get(e.get("resultats"))) for e, p in DY]),
 ("Familles appariées","Temps d'écran du parent et sommeil de l'élève", [(NUM["temps"].get(p.get("temps_parent")), NUM["sommeil"].get(e.get("duree_sommeil"))) for e, p in DY]),
]
CORR_RES = [(s, l) + corr(pp) for s, l, pp in CORRS]

# ---------- formats ----------
def pc(k, n): return (round(k / n * 100, 1) if n else 0)
def fp(x): return (f"{x:.1f}".replace(".", ",").replace(",0", "") + NB + "%")
def fpk(k, n):
    s = fp(pc(k, n))
    return s + (f" ({k} sur {n})" if n < 30 else "")
def fh(x): return f"{x:.1f}".replace(".", ",").replace(",0", "")
def hrs(x): return fh(x) + (" heures" if x >= 2 else " heure")
def frac_txt(k, n):
    p = pc(k, n)
    if n == 0: return "aucun"
    if p >= 75: return "une très large majorité"
    if p > 55: return "une majorité"
    if p >= 45: return "environ la moitié"
    if p >= 30: return "près d'un tiers" if p < 37 else "plus d'un tiers"
    if p >= 20: return "environ un sur quatre" if p < 28 else "un peu plus d'un quart"
    if p > 0: return "une minorité"
    return "aucun"
def force(r, p):
    if r is None: return "Effectif insuffisant", "B0B6C6"
    a = abs(r)
    t = "Lien négligeable" if a < 0.1 else "Lien faible" if a < 0.3 else "Lien modéré" if a < 0.5 else "Lien fort"
    if a >= 0.1: t += ", positif" if r > 0 else ", négatif"
    if p is not None and p >= 0.05 and a >= 0.1: t += ", non significatif"
    col = "B0B6C6" if a < 0.1 else "9AA5D6" if a < 0.3 else "E07A8E" if a < 0.5 else RED
    return t, col

nE, nP, nD = len(E), len(P), len(DY)
PROVISOIRE = nE < 30 or nP < 20

# ---------- chiffres clés ----------
def key_figures():
    out = []
    k, n = share(E, "temps_semaine", ["3 à 4 h","4 à 5 h","Plus de 5 h"])
    m, _ = mean(E, "temps_semaine", NUM["temps"])
    if n and k == 0: out.append(f"Un jour de classe, aucun des {n} élèves ne déclare plus de 3 heures d'écran en dehors du travail scolaire" + (f", pour une moyenne estimée à {hrs(m)}." if m is not None else "."))
    elif n: out.append(f"Un jour de classe, {fpk(k,n)} des élèves déclarent plus de 3 heures d'écran en dehors du travail scolaire" + (f", pour une moyenne estimée à {hrs(m)}." if m is not None else "."))
    mw, _ = mean(E, "temps_weekend", NUM["temps"])
    if mw is not None: out.append(f"Le week-end, le temps d'écran moyen estimé atteint {hrs(mw)} par jour.")
    k, n = share(E, "tel_chambre", ["Oui"]); k2, n2 = share(E, "dernier_ecran", ORD["dernier"][:2])
    if n: out.append(f"{fpk(k,n)} des élèves gardent leur téléphone dans leur chambre la nuit, et {fpk(k2,n2)} regardent un écran dans la demi-heure qui précède le sommeil.")
    k, n = share(E, "duree_sommeil", ["Moins de 6 h","6 à 7 h","7 à 8 h"])
    if n: out.append(f"{fpk(k,n)} des élèves dorment moins de 8 heures une nuit de semaine, en dessous du repère de 8 à 10 heures recommandé à l'adolescence.")
    k, n = share(E, "reveil_nuit", ["Parfois","Souvent","Toujours"])
    if n: out.append(f"{fpk(k,n)} se réveillent au moins parfois la nuit pour consulter leur téléphone.")
    k, n = share(E, "tel_devoirs", ["Souvent","Toujours"]); k2, n2 = share(E, "tel_classe", ["Parfois","Souvent","Toujours"])
    if n: out.append(f"{fpk(k,n)} ont souvent ou toujours leur téléphone à portée de main pendant les devoirs, et {fpk(k2,n2)} disent l'utiliser au moins parfois en classe sans autorisation.")
    k, n = share(E, "trop_ecrans", ["Oui","Plutôt oui"]); k2, n2 = share(E, "essai_reduire", ["Oui, et j'y suis arrivé·e","Oui, mais sans y arriver"])
    if n: out.append(f"{fpk(k,n)} des élèves estiment eux-mêmes passer trop de temps sur les écrans, et {fpk(k2,n2)} ont déjà essayé de le réduire.")
    k, n = share(E, "parents_technoference", ["Souvent","Toujours"])
    if n: out.append(f"{fpk(k,n)} des élèves disent que leurs parents regardent souvent ou toujours leur téléphone quand ils leur parlent.")
    mp, _ = mean(P, "temps_parent", NUM["temps"]); mpe, _ = mean(P, "temps_semaine", NUM["temps"])
    if mp is not None: out.append(f"Les parents déclarent en moyenne {hrs(mp)} d'écran par jour en dehors de leur travail" + (f", et estiment celui de leur enfant à {hrs(mpe)} un jour de classe." if mpe is not None else "."))
    k, n = share(P, "tensions", ["Souvent","Très souvent"])
    if n: out.append(f"Pour {fpk(k,n)} des parents, les écrans sont souvent ou très souvent une source de tension à la maison.")
    if nD: out.append(f"{nD} famille{'s' if nD>1 else ''} {'ont' if nD>1 else 'a'} utilisé le même code famille, ce qui permet de rapprocher les réponses d'un parent et de son enfant sans connaître leur identité.")
    return out

def rule_top(rows):
    n, it = counts(rows, "regle_lycee", order=ORD["regleLycee"])
    it = [x for x in it if x[0] != "Sans avis"]
    if not it: return None, 0, n
    k, v = max(it, key=lambda x: x[1])
    return k, v, n
RULE_SHORT = {"Utilisable librement, sauf en cours":"un usage libre en dehors des cours",
              "Utilisable seulement à certains moments ou dans certains lieux":"un usage limité à certains moments ou certains lieux",
              "Rangé toute la journée, éteint ou au casier":"un téléphone rangé toute la journée"}

def synthese():
    lines = []
    lines.append(f"{nE} élèves et {nP} parents ont répondu à l'enquête anonyme, en français, en bulgare ou en anglais" + (f", et {nD} famille{'s ont' if nD>1 else ' a'} pu être appariée{'s' if nD>1 else ''} grâce au code commun." if nD else "."))
    m, n = mean(E, "temps_semaine", NUM["temps"]); k, n3 = share(E, "temps_semaine", ["3 à 4 h","4 à 5 h","Plus de 5 h"])
    if n: lines.append(f"Un jour de classe, les élèves passent en moyenne environ {hrs(m)} sur les écrans hors travail scolaire, " + (f"et {fpk(k,n3)} dépassent 3 heures." if k else "et aucun ne déclare plus de 3 heures."))
    kc, nc = share(E, "tel_chambre", ["Oui"]); ks, ns = share(E, "duree_sommeil", ["Moins de 6 h","6 à 7 h","7 à 8 h"])
    if nc: lines.append(f"La nuit reste un point d'attention, {fpk(kc,nc)} gardent le téléphone dans leur chambre et {fpk(ks,ns)} dorment moins de 8 heures en semaine.")
    kcl, ncl = share(E, "tel_classe", ["Parfois","Souvent","Toujours"])
    rE, vE, nRE = rule_top(E); rP, vP, nRP = rule_top(P)
    if rE or rP:
        s = f"Au lycée, {fpk(kcl,ncl)} des élèves utilisent au moins parfois leur téléphone en classe sans autorisation. "
        if rE and rP and rE == rP:
            s += f"Pour la règle à venir, élèves et parents privilégient {RULE_SHORT.get(rE, rE)} ({fp(pc(vE,nRE))} des élèves, {fp(pc(vP,nRP))} des parents)."
        else:
            parts = []
            if rE: parts.append(f"les élèves privilégient {RULE_SHORT.get(rE, rE)} ({fp(pc(vE,nRE))})")
            if rP: parts.append(f"les parents {RULE_SHORT.get(rP, rP)} ({fp(pc(vP,nRP))})")
            s += "Pour la règle à venir, " + " et ".join(parts) + "."
        lines.append(s)
    lines.append("Une politique d'usage commune au lycée et aux familles, centrée sur le téléphone la nuit et un cadre clair en journée, construite avec les délégués élèves, s'appuierait sur ces constats" + (", à confirmer sur un plus grand nombre de réponses." if PROVISOIRE else "."))
    return lines[:5]

# ---------- graphiques ----------
SHORT = {"Utilisable librement, sauf en cours":"Libre, sauf en cours","Utilisable seulement à certains moments ou dans certains lieux":"Limité à certains moments",
 "Rangé toute la journée, éteint ou au casier":"Rangé toute la journée","Dans les 15 minutes avant de dormir":"Moins de 15 min avant",
 "Entre 15 et 30 minutes avant":"15 à 30 min avant","Entre 30 minutes et 1 heure avant":"30 min à 1 h avant","Plus d'une heure avant":"Plus d'1 h avant",
 "Je n'ai pas de téléphone":"Pas de téléphone","Entre 21 h et 22 h":"21 h à 22 h","Entre 22 h et 23 h":"22 h à 23 h","Entre 23 h et minuit":"23 h à minuit"}
PAL = ["#16256B","#C8102E","#2E86AB","#E0A100","#6B4FA0","#3C8D5A","#E07A8E","#7A869A"]
def seq(n):
    a, b = (0.79, 0.82, 0.94), (0.086, 0.145, 0.42)
    return [tuple(a[i] + (b[i]-a[i]) * (j/(n-1) if n > 1 else 1) for i in range(3)) for j in range(n)]
def pie(name, title, rows, key, order=None, ordinal=False):
    n, it = counts(rows, key, order=order)
    fig, ax = plt.subplots(figsize=(3.3, 2.15), dpi=220)
    if not n:
        ax.text(0.5, 0.5, "Pas encore de réponse", ha="center", va="center", fontsize=9, color="#5A6072"); ax.axis("off")
    else:
        cols = seq(len(it)) if ordinal else PAL[:len(it)]
        ax.pie([v for _, v in it], colors=cols, startangle=90, counterclock=False,
               wedgeprops=dict(width=0.42, edgecolor="white", linewidth=1.2))
        ax.text(0, 0.06, str(n), ha="center", va="center", fontsize=13, fontweight="bold", color="#16256B")
        ax.text(0, -0.2, "réponses", ha="center", va="center", fontsize=6.5, color="#5A6072")
        labels = [f"{SHORT.get(k, k)}  {fp(pc(v, n)).replace(NB,' ')}" for k, v in it]
        ax.legend(labels=[l if len(l) < 46 else l[:43] + "…" for l in labels], loc="center left", bbox_to_anchor=(1.0, 0.5),
                  frameon=False, fontsize=6.6, handlelength=0.9, handleheight=0.9)
        ax.set_aspect("equal")
    ax.set_title(title, fontsize=8.6, fontweight="bold", color="#1b2033", loc="left", pad=4)
    fig.subplots_adjust(left=0.0, right=0.52, top=0.86, bottom=0.04)
    path = os.path.join(FIG, name + ".png"); fig.savefig(path, facecolor="white"); plt.close(fig)
    return path
def hbar(name, title, rows, key):
    n, it = counts(rows, key, multi=True)
    fig, ax = plt.subplots(figsize=(6.6, 0.32 * max(len(it), 2) + 0.55), dpi=220)
    if n:
        labs = [k for k, _ in it][::-1]; ps = [pc(v, n) for _, v in it][::-1]
        ax.barh(labs, ps, color="#16256B", height=0.6)
        for i, p in enumerate(ps): ax.text(p + 1, i, fp(p).replace(NB, " "), va="center", fontsize=7.5, color="#16256B", fontweight="bold")
        ax.set_xlim(0, 115); ax.tick_params(axis="y", labelsize=7.5, length=0); ax.set_xticks([])
        for s in ax.spines.values(): s.set_visible(False)
    else: ax.axis("off")
    ax.set_title(f"{title} ({n} répondants, plusieurs réponses possibles)", fontsize=8.6, fontweight="bold", loc="left")
    fig.tight_layout(); path = os.path.join(FIG, name + ".png"); fig.savefig(path, facecolor="white"); plt.close(fig)
    return path

# ---------- docx helpers ----------
doc = Document()
st = doc.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
st.paragraph_format.space_after = Pt(4); st.paragraph_format.line_spacing = 1.1
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(1.9); sec.top_margin = Cm(1.6); sec.bottom_margin = Cm(1.6)
sec.header_distance = Cm(0.8); sec.footer_distance = Cm(0.8)
W = sec.page_width - sec.left_margin - sec.right_margin

def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr(); sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), color); tcPr.append(sh)
def borders(cell, **kw):
    tcPr = cell._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
    for side in ("top","left","bottom","right"):
        el = OxmlElement(f"w:{side}")
        if side in kw:
            sz, col = kw[side]; el.set(qn("w:val"), "single"); el.set(qn("w:sz"), str(sz)); el.set(qn("w:color"), col)
        else: el.set(qn("w:val"), "nil")
        b.append(el)
    tcPr.append(b)
def cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr(); m = OxmlElement("w:tcMar")
    for k, v in (("top",top),("bottom",bottom),("left",left),("right",right)):
        e = OxmlElement(f"w:{k}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); m.append(e)
    tcPr.append(m)
def fixed(table, widths):
    table.autofit = False
    grid = table._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths): gc.set(qn("w:w"), str(int(Emu(w).twips)))
    tblPr = table._tbl.tblPr; lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    for row in table.rows:
        for i, w in enumerate(widths): row.cells[i].width = w
def run(p, text, bold=False, size=None, color=None, italic=False):
    r = p.add_run(text); r.bold = bold; r.italic = italic
    if size: r.font.size = Pt(size)
    if color: r.font.color.rgb = RGBColor.from_string(color)
    return r
def para(text="", size=None, color=None, bold=False, align=None, after=4, italic=False):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(after)
    if align: p.alignment = align
    if text: run(p, text, bold, size, color, italic)
    return p
def h1(text):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    run(p, text, True, 14, NAVY)
    pPr = p._p.get_or_add_pPr(); bd = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
    bt.set(qn("w:val"), "single"); bt.set(qn("w:sz"), "12"); bt.set(qn("w:space"), "1"); bt.set(qn("w:color"), RED)
    bd.append(bt); pPr.append(bd)
    return p
def bullet(text, size=10.5):
    p = doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(3)
    run(p, text, size=size); return p
def boxed(lines, title, fill=SOFT, bar=RED):
    t = doc.add_table(rows=1, cols=1); fixed(t, [W]); c = t.rows[0].cells[0]
    shade(c, fill); borders(c, left=(36, bar)); cell_margins(c, 140, 140, 220, 180)
    p = c.paragraphs[0]; run(p, title, True, 11.5, NAVY); p.paragraph_format.space_after = Pt(5)
    for i, l in enumerate(lines, 1):
        q = c.add_paragraph(); q.paragraph_format.space_after = Pt(3)
        run(q, f"{i}. ", True, 10.5, RED); run(q, l, size=10.5)
    para("", after=2)
def add_field(p, instr):
    r = p.add_run(); f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    r._r.append(f1); r._r.append(it); r._r.append(f2); r.font.size = Pt(8.5); r.font.color.rgb = RGBColor.from_string(GREY)

# en-tête et pied de page
hd = sec.header; ht = hd.add_table(rows=1, cols=2, width=W); fixed(ht, [W/2, W/2])
c0, c1 = ht.rows[0].cells
c0.paragraphs[0].add_run().add_picture(logo("lvh"), height=Cm(0.85))
pr = c1.paragraphs[0]; pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT; pr.add_run().add_picture(logo("aefe"), height=Cm(0.85))
for c in (c0, c1): borders(c, bottom=(12, RED)); cell_margins(c, 0, 60, 0, 0)
hd.paragraphs[0].paragraph_format.space_after = Pt(0)
fp_ = sec.footer.paragraphs[0]; fp_.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(fp_, "Grégory Cron · laclassedegreg@gmail.com · Lycée français Victor Hugo, Sofia · page ", size=8.5, color=GREY)
add_field(fp_, "PAGE"); run(fp_, " sur ", size=8.5, color=GREY); add_field(fp_, "NUMPAGES")

# ---------- contenu ----------
def ts(dt): return dt.strftime("%d/%m/%Y").lstrip("0") if False else f"{dt.day} {['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'][dt.month-1]} {dt.year}"
stamps = []
for r in E + P:
    try: stamps.append(datetime.datetime.fromisoformat(str(r.get("horodatage")).replace("Z", "+00:00")).astimezone(TZ))
    except Exception: pass
first, last = (min(stamps), max(stamps)) if stamps else (NOW, NOW)

band = doc.add_table(rows=1, cols=1); fixed(band, [W]); bc = band.rows[0].cells[0]
shade(bc, NAVY); cell_margins(bc, 200, 200, 260, 260)
p = bc.paragraphs[0]; run(p, "OPÉRATION ÉCRAN, BOULOT, DODO", True, 20, "FFFFFF"); p.paragraph_format.space_after = Pt(2)
q = bc.add_paragraph(); run(q, "Écrans, sommeil, travail scolaire et bien-être au Lycée français Victor Hugo de Sofia", size=12, color="FFFFFF"); q.paragraph_format.space_after = Pt(2)
q = bc.add_paragraph(); run(q, "Rapport d'enquête à l'attention de la direction", True, 11, "F6C3CC")
para("", after=2)
p = para(after=2); run(p, "Enquête menée par les élèves de 2nde A et 2nde B en sciences économiques et sociales, sous la direction de Grégory Cron. ", size=10, color=GREY)
run(p, f"Données arrêtées le {ts(NOW)} à {NOW.hour}h{NOW.minute:02d}.", True, 10, GREY)
if PROVISOIRE:
    t = doc.add_table(rows=1, cols=1); fixed(t, [W]); c = t.rows[0].cells[0]; shade(c, "FDE8EC"); borders(c, left=(36, RED)); cell_margins(c, 100, 100, 200, 160)
    p = c.paragraphs[0]; run(p, "Version provisoire. ", True, 10, RED)
    run(p, f"La collecte est en cours et les effectifs restent faibles ({nE} élèves, {nP} parents). Les pourcentages peuvent encore évoluer nettement, et les liens statistiques ne sont calculés qu'à partir de {MIN_CORR} réponses comparables.", size=10)
    para("", after=2)

boxed(synthese(), "Synthèse pour la direction en cinq points")

h1("1. Méthode")
nl, il = counts(E + P, "langue")
LG = {"fr": "français", "bg": "bulgare", "en": "anglais"}
lang_txt = ", ".join(f"{LG.get(k, k)} {v}" for k, v in il)
_, ni = counts(E, "niveau", order=ORD["niveauE"])
niv_txt = ", ".join(f"{k} ({v})" for k, v in ni) or "non renseigné"
para(f"Le questionnaire en ligne, anonyme et trilingue, a été diffusé par Pronote à l'ensemble des élèves du secondaire et des familles de la maternelle à la Terminale. Il comprend un parcours élève de 34 questions et un parcours parent de 29 questions, sur les écrans, le sommeil, le travail scolaire, le bien-être, les pratiques familiales et la règle souhaitée au lycée. Aucune donnée personnelle n'est recueillie.")
para(f"Les réponses analysées ont été enregistrées entre le {ts(first)} et le {ts(last)}. {nE} élèves et {nP} parents ont répondu, en {lang_txt}. Répartition des élèves par niveau, {niv_txt}.")
para(f"Un code famille facultatif, choisi ensemble par un parent et son enfant, permet de rapprocher leurs réponses sans lever l'anonymat. {nD} paire{'s' if nD>1 else ''} parent et enfant {'ont' if nD>1 else 'a'} pu être constituée{'s' if nD>1 else ''} ainsi.")
para("Les temps d'écran et de sommeil sont déclarés par tranches. Les moyennes sont calculées à partir du centre de chaque tranche, et les liens entre deux réponses sont mesurés par le coefficient de corrélation de rang de Spearman.", color=GREY, size=9.5, italic=True)

h1("2. Chiffres clés")
for x in key_figures(): bullet(x)

def grid(paths, cols=2):
    t = doc.add_table(rows=(len(paths)+cols-1)//cols, cols=cols); fixed(t, [W/cols]*cols)
    for i, pth in enumerate(paths):
        c = t.rows[i//cols].cells[i % cols]; c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].add_run().add_picture(pth, width=W/cols - Cm(0.3)); cell_margins(c, 40, 40, 20, 20)
    para("", after=2)
h1("3. Les élèves en images")
grid([pie("e_temps","Temps d'écran un jour de classe", E, "temps_semaine", ORD["temps"], True),
      pie("e_we","Temps d'écran un jour de week-end", E, "temps_weekend", ORD["temps"], True),
      pie("e_sommeil","Durée de sommeil en semaine", E, "duree_sommeil", ORD["sommeil"], True),
      pie("e_coucher","Heure du coucher en semaine", E, "coucher", ORD["coucher"], True),
      pie("e_chambre","Téléphone dans la chambre la nuit", E, "tel_chambre"),
      pie("e_dernier","Dernier écran avant de dormir", E, "dernier_ecran", ORD["dernier"], True),
      pie("e_classe","Téléphone en classe sans autorisation", E, "tel_classe", ORD["freq"], True),
      pie("e_trop","Pense passer trop de temps sur les écrans", E, "trop_ecrans", ORD["accord"], True)])
doc.add_picture(hbar("e_act", "Ce que font les élèves sur les écrans", E, "activites"), width=W)

h1("4. Les parents en images")
grid([pie("p_temps","Temps d'écran du parent hors travail", P, "temps_parent", ORD["temps"], True),
      pie("p_enf","Temps d'écran estimé de l'enfant un jour de classe", P, "temps_semaine", ORD["temps"], True),
      pie("p_tens","Écrans source de tension à la maison", P, "tensions", ORD["freq"], True),
      pie("p_repas","Écrans pendant les repas en famille", P, "ecrans_repas", ORD["freq"], True)])
doc.add_picture(hbar("p_regles", "Règles mises en place à la maison", P, "regles"), width=W)

h1("5. La règle souhaitée au lycée")
grid([pie("r_e","Élèves", E, "regle_lycee", ORD["regleLycee"]), pie("r_p","Parents", P, "regle_lycee", ORD["regleLycee"])])
na, _ = counts(P, "attentes", multi=True)
if na: doc.add_picture(hbar("p_att", "Ce que les parents attendent du lycée", P, "attentes"), width=W)

h1("6. Ce qui va ensemble")
para("Le tableau mesure la force du lien entre deux réponses, de moins 1 à plus 1. Un coefficient proche de zéro signale une absence de lien, un coefficient proche de 1 ou de moins 1 un lien fort. Un lien décrit une tendance dans l'échantillon, il ne démontre pas qu'une variable cause l'autre.", size=10)
t = doc.add_table(rows=1, cols=5); widths = [Cm(2.6), Cm(7.6), Cm(1.6), Cm(1.8), W - Cm(13.6)]; fixed(t, widths)
for i, h in enumerate(["Source","Relation étudiée","Effectif","Coefficient","Lecture"]):
    c = t.rows[0].cells[i]; shade(c, NAVY); c.paragraphs[0].text = ""; run(c.paragraphs[0], h, True, 9, "FFFFFF"); cell_margins(c, 60, 60, 80, 80)
for s, l, r, p, n in CORR_RES:
    row = t.add_row().cells
    txt, col = force(r, p)
    for i, v in enumerate([s, l, str(n), "·" if r is None else f"{r:+.2f}".replace(".", ","), txt]):
        c = row[i]; c.width = widths[i]; c.paragraphs[0].text = ""; cell_margins(c, 40, 40, 80, 80)
        rr = run(c.paragraphs[0], v, size=8.8, bold=(i == 4), color=("FFFFFF" if i == 4 and col in (RED,) else None))
        if i == 4: shade(c, col if r is not None else "ECEEF3")
        borders(c, bottom=(4, "DDE1EC"))
para("", after=2)
sig = [x for x in CORR_RES if x[2] is not None and x[3] is not None and x[3] < 0.05 and abs(x[2]) >= 0.3]
if sig:
    para("Liens les plus nets, statistiquement significatifs au seuil de 5 %", bold=True, color=NAVY)
    for s, l, r, p, n in sorted(sig, key=lambda x: -abs(x[2])):
        bullet(f"{l} ({s.lower()}), lien {'positif' if r>0 else 'négatif'} de {abs(r):.2f}".replace(".", ",") + f" sur {n} réponses.")
else:
    para("À ce stade, aucun lien ne ressort de façon statistiquement significative" + (", faute d'un nombre suffisant de réponses." if PROVISOIRE else "."), italic=True, color=GREY)

h1("7. Familles appariées, le regard croisé")
if nD:
    comp = [(NUM["temps"].get(p.get("temps_semaine")), NUM["temps"].get(e.get("temps_semaine"))) for e, p in DY]
    comp = [(a, b) for a, b in comp if a is not None and b is not None]
    under = sum(1 for a, b in comp if a < b); same = sum(1 for a, b in comp if a == b); over = len(comp) - under - same
    if len(comp) == 1:
        verbe = "sous-estime" if under else ("situe dans la même tranche" if same else "surestime")
        para(f"Une paire parent et enfant a utilisé le même code. Sur cette seule paire, le parent {verbe} le temps d'écran déclaré par son enfant un jour de classe.")
    else:
        para(f"{nD} paires parent et enfant ont utilisé le même code. Sur les {len(comp)} paires comparables, le parent sous-estime le temps d'écran déclaré par son enfant dans {under} cas, le situe dans la même tranche dans {same} cas et le surestime dans {over} cas.")
    if len(comp) < MIN_CORR: para("Ce sous-échantillon est encore trop réduit pour en tirer une tendance, il sera exploité si le nombre de paires augmente.", italic=True, color=GREY)
else:
    para("Aucune famille n'a encore utilisé le code commun. Ce volet sera exploité si des paires parent et enfant apparaissent dans les réponses.", italic=True, color=GREY)

def verb(rows, key):
    out = []
    for v in vals(rows, key):
        v = re.sub(r"\s+", " ", v).strip()
        if len(v) > 320: v = v[:300].rsplit(" ", 1)[0] + " […]"
        out.append(v)
    return out
VE = verb(E, "aide") + verb(E, "idee"); VP = verb(P, "remarque")
if VE or VP:
    h1("8. Paroles d'élèves et de parents")
    para("Extraits des réponses libres, reproduits tels quels.", size=9.5, color=GREY, italic=True)
    for v in VE[:8]: bullet(f"Élève, « {v} »", size=10)
    for v in VP[:6]: bullet(f"Parent, « {v} »", size=10)

h1("9. Limites de l'enquête")
for x in ["Les répondants se sont portés volontaires, les familles les plus attentives à la question sont sans doute surreprésentées et l'échantillon n'est pas représentatif au sens strict.",
          "Les temps d'écran et de sommeil sont déclarés, donc approximatifs, et la désirabilité sociale peut conduire à les sous-estimer.",
          "L'âge, le milieu social et la langue parlée à la maison influencent à la fois les usages et les résultats scolaires. Un lien observé peut donc venir d'un troisième facteur.",
          f"La période de collecte est courte et les effectifs de certains croisements restent faibles, ce qui appelle la prudence."]:
    bullet(x)

h1("10. Pistes pour une politique d'usage du téléphone")
kc, nc = share(E, "tel_chambre", ["Oui"]); kd, nd = share(E, "dernier_ecran", ORD["dernier"][:2])
kcl, ncl = share(E, "tel_classe", ["Parfois","Souvent","Toujours"])
rE, vE, nRE = rule_top(E); rP, vP, nRP = rule_top(P)
kt, nt = share(E, "parents_technoference", ["Souvent","Toujours"])
pistes = [
 ("Le téléphone la nuit.", f"C'est le levier le plus directement lié au repos des élèves. Avec {fp(pc(kc,nc))} des élèves qui gardent le téléphone dans leur chambre et {fp(pc(kd,nd))} qui regardent un écran dans la demi-heure avant de dormir, une campagne commune avec les familles sur le téléphone hors de la chambre et une heure de coupure le soir peut s'appuyer sur ces données."),
 ("Un cadre clair en journée.", f"{fp(pc(kcl,ncl))} des élèves reconnaissent utiliser au moins parfois leur téléphone en classe sans autorisation. " + (f"La formule la plus souvent choisie est {RULE_SHORT.get(rE, rE)} chez les élèves et {RULE_SHORT.get(rP, rP)} chez les parents. " if rE and rP else "") + "La règle gagnera à distinguer le collège et le lycée, et les temps de cours, de pause et de restauration."),
 ("Une règle construite avec les élèves.", "Une charte du numérique discutée en conseil de vie collégienne et en conseil de la vie lycéenne, puis présentée au conseil d'établissement, sera mieux comprise et mieux respectée qu'une interdiction imposée."),
 ("L'exemple des adultes.", f"{fp(pc(kt,nt))} des élèves disent que leurs parents regardent souvent leur téléphone quand ils leur parlent. Une réunion d'information pour les familles peut aborder les usages des adultes, à la maison comme au lycée."),
 ("Mesurer les effets.", "Reconduire la même enquête en fin d'année permettra d'évaluer la politique mise en place, avec les mêmes questions et le même code famille."),
]
for t_, x in pistes:
    p = doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(4)
    run(p, t_ + " ", True, 10.5, NAVY); run(p, x, size=10.5)

para("", after=6)
p = para(after=0); run(p, "Grégory Cron", True, 11, NAVY)
para("Professeur d'HGEMC, HGGSP et SES, référent EDD", size=10, after=0)
para("Lycée français Victor Hugo de Sofia", size=10, after=6)
para("Tableau de bord en ligne des résultats, actualisé en temps réel, laclassedegreg.github.io/enquete-ecrans/resultats.html", size=9, color=GREY, italic=True)

name = "Rapport_Enquete_Ecran_boulot_dodo"
dp = os.path.join(OUT, name + ".docx"); doc.save(dp)
subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", OUT, dp], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
open(os.path.join(OUT, "synthese_direction.txt"), "w", encoding="utf8").write("\n".join(f"{i}. {l}" for i, l in enumerate(synthese(), 1)))
print(json.dumps({"eleves": nE, "parents": nP, "paires": nD, "provisoire": PROVISOIRE}))
