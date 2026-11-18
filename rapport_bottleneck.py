"""
rapport_bottleneck.py
Génère le rapport de synthèse « santé financière » pour les dirigeants
(PDF + Excel) à partir des objets du notebook.

Dépendances : pip install reportlab xlsxwriter pillow
"""
import os
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (Image, KeepTogether, PageBreak, Paragraph,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

ORANGE = colors.Color(255 / 255, 114 / 255, 0)
TEAL = colors.Color(51 / 255, 165 / 255, 182 / 255)
GRIS = colors.Color(0.95, 0.95, 0.95)

COLONNES_REQUISES = ["product_id", "post_title", "product_type", "price",
                     "purchase_price", "stock_quantity", "total_sales", "CA",
                     "rotation_stock", "valorisation_stock_euros", "taux_marge"]

# Description « langage dirigeant » de chaque type d'anomalie (modifiable)
DESCRIPTIONS_ANOMALIES = {
    "Stock négatif": "Quantité en stock inférieure à 0 : erreur de saisie ou de "
                     "mouvement de stock. La valorisation du stock est faussée.",
    "Prix négatif": "Prix de vente inférieur à 0 : erreur de saisie, le produit "
                    "ne peut pas être vendu correctement.",
    "Statut de stock incohérent": "Le statut affiché (en stock / rupture) ne correspond "
                                  "pas à la quantité réelle : la disponibilité annoncée "
                                  "aux clients est fausse.",
    "Publiés en rupture": "Produits visibles sur le site alors que le stock est épuisé : "
                          "risque de ventes perdues et de clients déçus.",
    "Stock non publié": "Produits en stock mais absents du site : capital immobilisé "
                        "qui ne peut pas se vendre en ligne.",
    "Marge négative": "Prix de vente inférieur au prix d'achat : l'entreprise perd de "
                      "l'argent à chaque vente.",
}

NOMS_COLONNES = {
    "product_id": "ID produit", "post_title": "Produit", "product_type": "Catégorie",
    "price": "Prix (€)", "purchase_price": "Prix achat (€)",
    "stock_quantity": "Stock", "stock_status": "Statut", "onsale_web": "Publié web",
    "taux_marge": "Marge (%)", "erreur_type": "Type d'incohérence",
    "id_web": "ID web", "total_sales": "Ventes oct.",
}


# ----------------------------------------------------------------------------
# Utilitaires
# ----------------------------------------------------------------------------
def _eur(x, d=0):
    return f"{x:,.{d}f}".replace(",", " ") + " €"


def _nb(x, d=0):
    return f"{x:,.{d}f}".replace(",", " ")


def _txt(x):
    """Texte compatible polices PDF standard (supprime les emojis)."""
    if pd.isna(x):
        return ""
    return str(x).encode("cp1252", errors="ignore").decode("cp1252")


def _part_refs_pour_80(df, col):
    """% des références (ayant des ventes) qui font 80 % du total de `col`."""
    s = df.loc[df[col] > 0, col].sort_values(ascending=False)
    cumul = s.cumsum() / s.sum()
    return ((cumul < 0.8).sum() + 1) / len(s) * 100


def _verifier(df):
    manquantes = [c for c in COLONNES_REQUISES if c not in df.columns]
    if manquantes:
        raise ValueError(f"Colonnes absentes de df_final : {manquantes}")


# ----------------------------------------------------------------------------
# Calculs : chiffres clés + commentaires automatiques
# ----------------------------------------------------------------------------
def calculer_kpis(df):
    # total_sales est vide (NaN) pour les produits absents du web : on le compte comme 0 vente
    jamais = df[df["total_sales"].fillna(0) == 0]
    return [
        ("Chiffre d'affaires d'octobre", _eur(df["CA"].sum())),
        ("Valeur du stock (prix d'achat)", _eur(df["valorisation_stock_euros"].sum())),
        ("Bouteilles en stock", _nb(df["stock_quantity"].sum())),
        ("Taux de marge moyen (sur coût HT)", f"{df['taux_marge'].mean():.1f} %"),
        (f"Références sans vente en octobre ({_eur(jamais['valorisation_stock_euros'].sum())} de stock)",
         f"{len(jamais)} / {len(df)}"),
        ("Produits vendus à perte", _nb((df["taux_marge"] < 0).sum())),
    ]


def commentaires_auto(df):
    """Commentaires calculés sur les données (à surcharger si besoin)."""
    top_ca = df.nlargest(20, "CA")
    top_vol = df.nlargest(20, "total_sales")
    flop = df.nlargest(20, "rotation_stock")
    # Comme dans le graphique du notebook : moyenne par catégorie sur les marges positives
    marge_cat = df[df["taux_marge"] > 0].groupby("product_type")["taux_marge"].mean().sort_values()
    commun = len(set(top_ca["product_id"]) & set(top_vol["product_id"]))
    pct80 = _part_refs_pour_80(df, "CA")

    return {
        "pareto_ca": (
            f"{pct80:.0f} % des références vendues génèrent 80 % du chiffre d'affaires. "
            + ("L'activité repose sur un nombre restreint de produits : forte dépendance."
               if pct80 < 30 else
               "Le chiffre d'affaires est réparti sur une large part du catalogue : "
               "la dépendance à quelques produits est limitée.")),
        "top20_ca": (
            f"Les 20 produits qui rapportent le plus représentent {top_ca['CA'].sum() / df['CA'].sum() * 100:.0f} % "
            f"du chiffre d'affaires ({_eur(top_ca['CA'].sum())}). Ce sont les références à "
            "surveiller en priorité (disponibilité et approvisionnement)."),
        "top20_volume": (
            f"{commun} produit(s) figurent à la fois dans le Top 20 en valeur et dans le Top 20 en volume. "
            + ("Les produits qui se vendent le plus ne sont donc pas ceux qui rapportent le plus."
               if commun <= 5 else "Les produits les plus vendus sont aussi parmi les plus rentables en valeur.")),
        "rotation": (
            f"Les 20 produits qui tournent le moins immobilisent "
            f"{_eur(flop['valorisation_stock_euros'].sum())}, soit "
            f"{flop['valorisation_stock_euros'].sum() / df['valorisation_stock_euros'].sum() * 100:.0f} % "
            "de la valeur du stock. Au rythme d'octobre, certains stocks représentent "
            f"jusqu'à {flop['rotation_stock'].max():.0f} mois de ventes."),
        "marge": (
            f"Catégorie la plus rentable : {marge_cat.index[-1]} ({marge_cat.iloc[-1]:.0f} %). "
            f"Catégorie la moins rentable : {marge_cat.index[0]} ({marge_cat.iloc[0]:.0f} %). "
            f"{(df['taux_marge'] < 0).sum()} produit(s) sont vendus à perte."),
    }


# ----------------------------------------------------------------------------
# PDF
# ----------------------------------------------------------------------------
def _styles():
    ss = getSampleStyleSheet()
    return {
        "titre": ParagraphStyle("titre", parent=ss["Title"], fontName="Helvetica-Bold",
                                fontSize=22, textColor=ORANGE, spaceAfter=6),
        "sous_titre": ParagraphStyle("sous_titre", parent=ss["Normal"], fontSize=11,
                                     textColor=colors.grey, alignment=TA_CENTER, spaceAfter=14),
        "h1": ParagraphStyle("h1", parent=ss["Heading1"], fontName="Helvetica-Bold",
                             fontSize=15, textColor=TEAL, spaceBefore=10, spaceAfter=6),
        "h2": ParagraphStyle("h2", parent=ss["Heading2"], fontName="Helvetica-Bold",
                             fontSize=12, textColor=colors.black, spaceBefore=8, spaceAfter=3,
                             keepWithNext=1),
        "txt": ParagraphStyle("txt", parent=ss["Normal"], fontSize=10, leading=14),
        "txt_suite": ParagraphStyle("txt_suite", parent=ss["Normal"], fontSize=10, leading=14,
                                    spaceAfter=4, keepWithNext=1),
        "cell": ParagraphStyle("cell", parent=ss["Normal"], fontSize=7.5, leading=9),
        "cell_h": ParagraphStyle("cell_h", parent=ss["Normal"], fontSize=7.5, leading=9,
                                 fontName="Helvetica-Bold", textColor=colors.white),
        "kpi_v": ParagraphStyle("kpi_v", parent=ss["Normal"], fontName="Helvetica-Bold",
                                fontSize=13, textColor=ORANGE, alignment=TA_CENTER, leading=16),
        "kpi_l": ParagraphStyle("kpi_l", parent=ss["Normal"], fontSize=8.5,
                                textColor=colors.black, alignment=TA_CENTER, leading=10),
    }


def _bloc_kpis(kpis, st):
    lignes = []
    for i in range(0, len(kpis), 3):
        trio = kpis[i:i + 3]
        lignes.append([[Paragraph(_txt(v), st["kpi_v"]), Paragraph(_txt(l), st["kpi_l"])]
                       for l, v in trio])
    t = Table(lignes, colWidths=[5.8 * cm] * 3, rowHeights=[2.3 * cm] * len(lignes))
    t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, TEAL), ("INNERGRID", (0, 0), (-1, -1), 0.5, TEAL),
        ("BACKGROUND", (0, 0), (-1, -1), GRIS), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def _bloc_graphique(g, st, largeur_max=17 * cm, hauteur_max=10 * cm, entete=None):
    """(Titre de section) + titre + image + commentaire, insécables."""
    elems = ([entete] if entete is not None else []) + [Paragraph(_txt(g["titre"]), st["h2"])]
    if g.get("chemin") and os.path.exists(g["chemin"]):
        w, h = ImageReader(g["chemin"]).getSize()
        ratio = min(largeur_max / w, hauteur_max / h)
        elems.append(Image(g["chemin"], width=w * ratio, height=h * ratio))
    else:
        elems.append(Paragraph(f"<i>(graphique introuvable : {escape(str(g.get('chemin')))})</i>",
                               st["txt"]))
        print(f"⚠️ Graphique introuvable : {g.get('chemin')}")
    elems += [Spacer(1, 4), Paragraph(_txt(g.get("commentaire", "")), st["txt"]), Spacer(1, 8)]
    return KeepTogether(elems)


def _table_anomalies(df, st, largeur=17.4 * cm):
    cols = [c for c in df.columns if c != "erreur_type"]
    entetes = [Paragraph(_txt(NOMS_COLONNES.get(c, c)), st["cell_h"]) for c in cols]
    lignes = [entetes]
    for _, r in df.iterrows():
        ligne = []
        for c in cols:
            v = r[c]
            if isinstance(v, float) and not pd.isna(v):
                v = f"{v:.2f}"
            ligne.append(Paragraph(escape(_txt(v)), st["cell"]))
        lignes.append(ligne)
    # le nom du produit prend plus de place s'il existe
    poids = [3.0 if c == "post_title" else 1.0 for c in cols]
    largeurs = [largeur * p / sum(poids) for p in poids]
    t = Table(lignes, colWidths=largeurs, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), TEAL),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GRIS]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def _pied_de_page(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    canvas.drawString(1.8 * cm, 1.0 * cm, "Bottleneck - Rapport de santé financière")
    canvas.drawRightString(A4[0] - 1.8 * cm, 1.0 * cm, f"Page {doc.page}")
    canvas.restoreState()


def _construire_pdf(chemin_pdf, titre, periode, kpis, sections, recommandations, anomalies):
    st = _styles()
    doc = SimpleDocTemplate(chemin_pdf, pagesize=A4, leftMargin=1.8 * cm, rightMargin=1.8 * cm,
                            topMargin=1.8 * cm, bottomMargin=1.8 * cm, title=titre)
    e = [Paragraph(_txt(titre), st["titre"]),
         Paragraph(_txt(periode), st["sous_titre"]),
         Paragraph("1. Chiffres clés", st["h1"]), _bloc_kpis(kpis, st), Spacer(1, 10)]

    for i, (nom_section, graphs) in enumerate(sections, start=2):
        titre_section = Paragraph(f"{i}. {_txt(nom_section)}", st["h1"])
        for k, g in enumerate(graphs):
            e.append(_bloc_graphique(g, st, entete=titre_section if k == 0 else None))

    n = len(sections) + 2
    # --- Anomalies par type, liste complète des produits ---
    e += [PageBreak(), Paragraph(f"{n}. Anomalies à corriger", st["h1"])]
    total = sum(len(d) for d in anomalies.values())
    e.append(Paragraph(f"{total} ligne(s) à vérifier, regroupées par type d'anomalie. "
                       "Les mêmes listes figurent dans le fichier Excel.", st["txt"]))
    resume = [[Paragraph("Type d'anomalie", st["cell_h"]), Paragraph("Nombre", st["cell_h"])]]
    resume += [[Paragraph(_txt(k), st["cell"]), Paragraph(str(len(d)), st["cell"])]
               for k, d in anomalies.items()]
    t = Table(resume, colWidths=[12 * cm, 3 * cm])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), TEAL),
                           ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey)]))
    e += [Spacer(1, 6), t]

    for k, d in anomalies.items():
        bloc = [Paragraph(f"{_txt(k)} ({len(d)})", st["h2"]),
                Paragraph(_txt(DESCRIPTIONS_ANOMALIES.get(k, "")), st["txt_suite"])]
        if d.empty:
            bloc.append(Paragraph("<i>Aucune anomalie détectée.</i>", st["txt"]))
            e.append(KeepTogether(bloc))
        else:
            e += bloc + [_table_anomalies(d, st)]
        e.append(Spacer(1, 8))

    # --- Recommandations (fin du rapport) ---
    e.append(KeepTogether([Paragraph(f"{n + 1}. Recommandations", st["h1"])] +
                          [x for r in recommandations
                           for x in (Paragraph("• " + _txt(r), st["txt"]), Spacer(1, 3))]))

    doc.build(e, onFirstPage=_pied_de_page, onLaterPages=_pied_de_page)


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
def _nom_feuille(nom):
    for c in '[]:*?/\\':
        nom = nom.replace(c, "-")
    return nom[:31]


COLS_CENTRE = {"product_id", "id_web", "onsale_web", "stock_status", "total_sales"}
COLS_ENTIER = {"stock_quantity", "Nombre"}
COLS_DECIMAL = {"price", "purchase_price", "taux_marge"}


def _py(v):
    """Convertit une valeur pandas/numpy en valeur Python (None si vide)."""
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v.item() if hasattr(v, "item") else v


def _formats_excel(wb):
    base = {"font_name": "Calibri", "font_size": 10, "border": 1,
            "border_color": "#BFBFBF", "valign": "vcenter"}

    def f(**k):
        return wb.add_format({**base, **k})

    return {
        "entete": f(bold=True, font_color="#FFFFFF", bg_color="#33A5B6", align="center",
                    text_wrap=True, border_color="#33A5B6"),
        "gauche": f(align="left", text_wrap=True),
        "centre": f(align="center"),
        "id": f(align="center", num_format="0"),
        "entier": f(align="center", num_format="#,##0"),
        "decimal": f(align="right", num_format="#,##0.00"),
        "kpi_lib": f(bold=True, bg_color="#F2F2F2", align="left", text_wrap=True),
        "kpi_val": f(bold=True, font_color="#FF7200", align="left"),
        "titre": wb.add_format({"bold": True, "font_size": 16, "font_color": "#FF7200"}),
        "sous_titre": wb.add_format({"italic": True, "font_color": "#7F7F7F"}),
        "section": wb.add_format({"bold": True, "font_size": 12, "font_color": "#33A5B6"}),
        "alt": wb.add_format({"bg_color": "#F5F9FA"}),
    }


def _ecrire_tableau(wb, ws, df, F, l0=0, largeur_max=55, hauteur_entete=30):
    """Tableau encadré avec filtres, entêtes colorés, colonnes alignées, lignes alternées."""
    cols_orig = list(df.columns)
    noms = [NOMS_COLONNES.get(c, c) for c in cols_orig]
    if df.empty:
        for j, n in enumerate(noms):
            ws.write(l0, j, n, F["entete"])
            ws.set_column(j, j, max(len(str(n)) + 4, 14))
        ws.write(l0 + 1, 0, "Aucune anomalie détectée.", F["gauche"])
        return
    colonnes = []
    for c, n in zip(cols_orig, noms):
        if c in COLS_DECIMAL:
            fmt = F["decimal"]
        elif c in COLS_ENTIER:
            fmt = F["entier"]
        elif c in ("product_id", "id_web"):
            fmt = F["id"]
        elif c in COLS_CENTRE:
            fmt = F["centre"]
        else:
            fmt = F["gauche"]
        colonnes.append({"header": str(n), "format": fmt, "header_format": F["entete"]})
    data = [[_py(v) for v in row] for row in df.itertuples(index=False, name=None)]
    derniere = l0 + len(data)
    ws.add_table(l0, 0, derniere, len(colonnes) - 1,
                 {"columns": colonnes, "data": data, "style": None, "autofilter": True})
    ws.conditional_format(l0 + 1, 0, derniere, len(colonnes) - 1,
                          {"type": "formula", "criteria": "=MOD(ROW(),2)=0", "format": F["alt"]})
    for j, (c, n) in enumerate(zip(cols_orig, noms)):
        longueur = max(len(str(n)), df[c].astype(str).str.len().max())
        ws.set_column(j, j, min(max(longueur + 3, 12), largeur_max))
    ws.set_row(l0, hauteur_entete)
    ws.freeze_panes(l0 + 1, 0)
    ws.repeat_rows(l0)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.hide_gridlines(2)


def _construire_excel(chemin_excel, kpis, sections, anomalies,
                      titre="Bottleneck - Rapport de santé financière", periode=""):
    with pd.ExcelWriter(chemin_excel, engine="xlsxwriter") as w:
        wb = w.book
        F = _formats_excel(wb)

        # --- Feuille Synthèse : chiffres clés + commentaires -------------------
        ws = wb.add_worksheet("Synthèse")
        w.sheets["Synthèse"] = ws
        ws.hide_gridlines(2)
        ws.set_column(0, 0, 48)
        ws.set_column(1, 1, 95)
        ws.write(0, 0, titre, F["titre"])
        ws.write(1, 0, periode, F["sous_titre"])
        ws.write(3, 0, "Chiffres clés", F["section"])
        ws.write(4, 0, "Indicateur", F["entete"])
        ws.write(4, 1, "Valeur", F["entete"])
        r = 5
        for lib, val in kpis:
            ws.write(r, 0, lib, F["kpi_lib"])
            ws.write(r, 1, val, F["kpi_val"])
            ws.set_row(r, 22)
            r += 1
        r += 1
        ws.write(r, 0, "Commentaires", F["section"])
        r += 1
        ws.write(r, 0, "Thème", F["entete"])
        ws.write(r, 1, "Commentaire", F["entete"])
        r += 1
        for _, graphs in sections:
            for g in graphs:
                ws.write(r, 0, g["titre"], F["kpi_lib"])
                ws.write(r, 1, g.get("commentaire", ""), F["gauche"])
                ws.set_row(r, 48)
                r += 1
        ws.set_landscape()
        ws.fit_to_pages(1, 0)

        # --- Feuille Graphiques -----------------------------------------------
        wg = wb.add_worksheet("Graphiques")
        w.sheets["Graphiques"] = wg
        wg.hide_gridlines(2)
        ligne = 0
        for _, graphs in sections:
            for g in graphs:
                if g.get("chemin") and os.path.exists(g["chemin"]):
                    wg.write(ligne, 0, g["titre"], F["section"])
                    wg.insert_image(ligne + 1, 0, g["chemin"], {"x_scale": 0.35, "y_scale": 0.35})
                    ligne += 32

        # --- Résumé des anomalies ---------------------------------------------
        nom_resume = "Anomalies - résumé"
        wr = wb.add_worksheet(nom_resume)
        w.sheets[nom_resume] = wr
        resume = pd.DataFrame({"Type d'anomalie": list(anomalies),
                               "Nombre": [len(d) for d in anomalies.values()],
                               "Description": [DESCRIPTIONS_ANOMALIES.get(k, "") for k in anomalies]})
        _ecrire_tableau(wb, wr, resume, F, largeur_max=100)
        wr.set_column(0, 0, 32)
        wr.set_column(1, 1, 12)
        wr.set_column(2, 2, 100)

        # --- Une feuille par type d'anomalie ----------------------------------
        for k, d in anomalies.items():
            nom = _nom_feuille(k)
            wk = wb.add_worksheet(nom)
            w.sheets[nom] = wk
            _ecrire_tableau(wb, wk, d.reset_index(drop=True), F)


# ----------------------------------------------------------------------------
# Fonction principale
# ----------------------------------------------------------------------------
def generer_rapport(df_final, anomalies, graphiques, recommandations=None,
                    commentaires=None, dossier="Rapport",
                    titre="Bottleneck - Rapport de santé financière",
                    periode="Données au 31 octobre"):
    """
    df_final        : DataFrame final du notebook (avec CA, taux_marge, rotation_stock...)
    anomalies       : dict {"Stock négatif": df, ...} (ordre = ordre d'affichage)
    graphiques      : dict {"pareto_ca": chemin_png, "top20_ca": ..., "top20_volume": ...,
                            "rotation": ..., "marge": ...}
    recommandations : liste de phrases (sinon texte par défaut)
    commentaires    : dict pour remplacer un commentaire automatique (mêmes clés que graphiques)
    """
    _verifier(df_final)
    os.makedirs(dossier, exist_ok=True)

    kpis = calculer_kpis(df_final)
    textes = commentaires_auto(df_final)
    textes.update(commentaires or {})

    def g(cle, titre_g):
        return {"titre": titre_g, "chemin": graphiques.get(cle), "commentaire": textes[cle]}

    sections = [
        ("Où se fait le chiffre d'affaires", [
            g("pareto_ca", "Concentration du chiffre d'affaires"),
            g("top20_ca", "Top 20 des produits en chiffre d'affaires"),
            g("top20_volume", "Top 20 des produits en volume"),
        ]),
        ("Stock et trésorerie immobilisée", [g("rotation", "Produits à plus faible rotation")]),
        ("Rentabilité", [g("marge", "Taux de marge par catégorie")]),
    ]

    if recommandations is None:
        recommandations = [
            "Sécuriser l'approvisionnement des produits du Top 20 en chiffre d'affaires.",
            "Réduire le stock des références à très faible rotation (promotions, déstockage).",
            "Revoir le prix des produits vendus à perte.",
            "Corriger les anomalies listées dans la section précédente (stocks, prix, statuts, publication web).",
        ]

    chemin_pdf = os.path.join(dossier, "rapport_sante_financiere.pdf")
    chemin_xlsx = os.path.join(dossier, "rapport_sante_financiere.xlsx")
    _construire_pdf(chemin_pdf, titre, periode, kpis, sections, recommandations, anomalies)
    _construire_excel(chemin_xlsx, kpis, sections, anomalies, titre, periode)
    print(f"✅ Rapport PDF   : {chemin_pdf}")
    print(f"✅ Rapport Excel : {chemin_xlsx}")
    return chemin_pdf, chemin_xlsx