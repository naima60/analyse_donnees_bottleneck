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
    jamais = df[df["total_sales"] == 0]
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
    marge_cat = df.groupby("product_type")["taux_marge"].mean().sort_values()
    commun = len(set(top_ca["product_id"]) & set(top_vol["product_id"]))

    return {
        "pareto_ca": (
            f"{_part_refs_pour_80(df, 'CA'):.0f} % des références vendues génèrent 80 % du "
            "chiffre d'affaires : l'activité repose sur un nombre restreint de produits."),
        "top20_ca": (
            f"Les 20 premiers produits représentent {top_ca['CA'].sum() / df['CA'].sum() * 100:.0f} % "
            f"du chiffre d'affaires ({_eur(top_ca['CA'].sum())}). Toute rupture "
            "d'approvisionnement sur ces références a un impact direct sur les ventes."),
        "top20_volume": (
            f"Seuls {commun} produit(s) figurent à la fois dans le Top 20 en valeur et dans le Top 20 "
            "en volume : les produits qui se vendent le plus ne sont pas ceux qui rapportent le plus."),
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
                             fontSize=12, textColor=colors.black, spaceBefore=8, spaceAfter=3),
        "txt": ParagraphStyle("txt", parent=ss["Normal"], fontSize=10, leading=14),
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


def _bloc_graphique(g, st, largeur_max=17 * cm, hauteur_max=11.5 * cm):
    """Titre + image + commentaire, insécables."""
    elems = [Paragraph(_txt(g["titre"]), st["h2"])]
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
    cols = list(df.columns)
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
        e.append(Paragraph(f"{i}. {_txt(nom_section)}", st["h1"]))
        for g in graphs:
            e.append(_bloc_graphique(g, st))

    n = len(sections) + 2
    e.append(Paragraph(f"{n}. Recommandations", st["h1"]))
    for r in recommandations:
        e.append(Paragraph("• " + _txt(r), st["txt"]))
        e.append(Spacer(1, 3))

    # --- Annexe : anomalies par type, liste complète des produits ---
    e += [PageBreak(), Paragraph(f"{n + 1}. Annexe - Anomalies à corriger", st["h1"])]
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
                Paragraph(_txt(DESCRIPTIONS_ANOMALIES.get(k, "")), st["txt"]), Spacer(1, 4)]
        if d.empty:
            bloc.append(Paragraph("<i>Aucune anomalie détectée.</i>", st["txt"]))
            e.append(KeepTogether(bloc))
        else:
            e += bloc + [_table_anomalies(d, st)]
        e.append(Spacer(1, 8))

    doc.build(e, onFirstPage=_pied_de_page, onLaterPages=_pied_de_page)


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
def _nom_feuille(nom):
    for c in '[]:*?/\\':
        nom = nom.replace(c, "-")
    return nom[:31]


def _construire_excel(chemin_excel, kpis, sections, anomalies):
    with pd.ExcelWriter(chemin_excel, engine="xlsxwriter") as w:
        wb = w.book
        f_titre = wb.add_format({"bold": True, "font_size": 14, "font_color": "#FF7200"})
        f_txt = wb.add_format({"text_wrap": True, "valign": "top"})
        f_entete = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#33A5B6"})

        # Feuille Synthèse : chiffres clés + commentaires
        ws = wb.add_worksheet("Synthèse")
        w.sheets["Synthèse"] = ws
        ws.set_column(0, 0, 42)
        ws.set_column(1, 1, 60)
        ws.write(0, 0, "Chiffres clés", f_titre)
        for i, (lib, val) in enumerate(kpis, start=2):
            ws.write(i, 0, lib)
            ws.write(i, 1, val)
        r = len(kpis) + 4
        ws.write(r, 0, "Commentaires", f_titre)
        for _, graphs in sections:
            for g in graphs:
                r += 1
                ws.write(r, 0, g["titre"], f_txt)
                ws.write(r, 1, g.get("commentaire", ""), f_txt)
                ws.set_row(r, 60)

        # Feuille Graphiques : images
        wg = wb.add_worksheet("Graphiques")
        w.sheets["Graphiques"] = wg
        ligne = 0
        for _, graphs in sections:
            for g in graphs:
                if g.get("chemin") and os.path.exists(g["chemin"]):
                    wg.write(ligne, 0, g["titre"], f_titre)
                    wg.insert_image(ligne + 1, 0, g["chemin"], {"x_scale": 0.35, "y_scale": 0.35})
                    ligne += 32

        # Résumé des anomalies + une feuille par type
        pd.DataFrame({"Type d'anomalie": list(anomalies),
                      "Nombre": [len(d) for d in anomalies.values()],
                      "Description": [DESCRIPTIONS_ANOMALIES.get(k, "") for k in anomalies]}
                     ).to_excel(w, sheet_name="Anomalies - résumé", index=False)
        w.sheets["Anomalies - résumé"].set_column(0, 0, 30)
        w.sheets["Anomalies - résumé"].set_column(1, 1, 10)
        w.sheets["Anomalies - résumé"].set_column(2, 2, 100)

        for k, d in anomalies.items():
            nom = _nom_feuille(k)
            d.rename(columns=NOMS_COLONNES).to_excel(w, sheet_name=nom, index=False)
            feuille = w.sheets[nom]
            for j, c in enumerate(d.rename(columns=NOMS_COLONNES).columns):
                largeur = max(len(str(c)), d[d.columns[j]].astype(str).str.len().max()
                              if len(d) else 0)
                feuille.set_column(j, j, min(largeur + 2, 60))
                feuille.write(0, j, c, f_entete)
            feuille.freeze_panes(1, 0)


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
            "Corriger les anomalies listées en annexe (stocks, prix, statuts, publication web).",
        ]

    chemin_pdf = os.path.join(dossier, "rapport_sante_financiere.pdf")
    chemin_xlsx = os.path.join(dossier, "rapport_sante_financiere.xlsx")
    _construire_pdf(chemin_pdf, titre, periode, kpis, sections, recommandations, anomalies)
    _construire_excel(chemin_xlsx, kpis, sections, anomalies)
    print(f"✅ Rapport PDF   : {chemin_pdf}")
    print(f"✅ Rapport Excel : {chemin_xlsx}")
    return chemin_pdf, chemin_xlsx
