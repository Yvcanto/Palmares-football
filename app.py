"""
Tableau de bord interactif - Palmarès Football FM26
======================================================
Lance avec :  streamlit run app.py

Le fichier Excel 'Palmares_Football_FM26_Complet_V3.xlsx' doit se trouver
dans le même dossier que ce script (sinon un champ d'upload apparaît).
"""

import os

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from data_utils import (
    COLOR_LEGEND,
    CONTINENT_COLORS,
    CONTINENTAL_COMPETITIONS,
    HEADER_COLOR,
    SHEET_C1,
    SHEET_FINALISTES,
    classify_color,
    load_all,
)

# ----------------------------------------------------------------------
# Configuration générale
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="Palmarès Football FM",
    page_icon="⚽",
    layout="wide",
)

DEFAULT_XLSX_NAME = "Palmares_Football_FM26_Complet_V3.xlsx"
HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSX_PATH = os.path.join(HERE, DEFAULT_XLSX_NAME)


@st.cache_data(show_spinner="Lecture du classeur Excel...")
def get_data(file_bytes_or_path):
    sheets, unified = load_all(file_bytes_or_path)
    return sheets, unified


# ----------------------------------------------------------------------
# Chargement du fichier (chemin par défaut, sinon upload)
# ----------------------------------------------------------------------
xlsx_path = None
if os.path.exists(DEFAULT_XLSX_PATH):
    xlsx_path = DEFAULT_XLSX_PATH
else:
    st.sidebar.warning("Fichier Excel introuvable à côté du script.")
    uploaded = st.sidebar.file_uploader(
        "Charger Palmares_Football_FM26_Complet_V3.xlsx", type=["xlsx"]
    )
    if uploaded is not None:
        xlsx_path = uploaded

if xlsx_path is None:
    st.title("⚽ Palmarès Football FM")
    st.info("Veuillez charger le fichier Excel dans la barre latérale pour démarrer.")
    st.stop()

sheets, df = get_data(xlsx_path)

df_national = df[df["Categorie"] == "Championnat national"].copy()
df_c1 = df[(df["Categorie"] == "Continental - Vainqueurs") & (df["Competition"] == "Ligue des Champions (UEFA)")].copy()
df_finalistes = df[(df["Categorie"] == "Continental - Finalistes") & (df["Competition"] == "Ligue des Champions (UEFA)")].copy()

# ----------------------------------------------------------------------
# Fonctions d'affichage (fidélité au formatage Excel)
# ----------------------------------------------------------------------

def render_styled_table(sub_df: pd.DataFrame, columns: list[tuple[str, str]]):
    """
    Affiche un DataFrame sous forme de tableau HTML reproduisant les
    couleurs de surlignage définies dans le fichier Excel d'origine,
    ainsi que la police rouge marquant les clubs non jouables en FM26.

    columns : liste de tuples (nom_colonne_df, libellé_affiché)
    """
    header_cells = "".join(f"<th>{label}</th>" for _, label in columns)
    rows_html = []
    for _, row in sub_df.iterrows():
        color_code = row.get("_color", "NONE")
        meta = COLOR_LEGEND[classify_color(color_code)]
        bg = meta["hex"]
        font_color = "#FFFFFF" if color_code == "FF0000" else "#000000"
        weight = "bold" if color_code == "FF0000" else "normal"
        non_jouable = bool(row.get("_non_jouable", False))
        cells = []
        for col, _ in columns:
            val = row.get(col, "")
            val = "" if pd.isna(val) else val
            if col == "Club" and non_jouable:
                cells.append(f'<td><span style="color:#B91C1C;font-weight:bold;">{val}</span></td>')
            else:
                cells.append(f"<td>{val}</td>")
        rows_html.append(
            f'<tr style="background-color:{bg};color:{font_color};'
            f'font-weight:{weight};">' + "".join(cells) + "</tr>"
        )

    html = f"""
    <div style="overflow-x:auto;">
    <table style="width:100%;border-collapse:collapse;font-size:14px;">
      <thead>
        <tr style="background-color:{HEADER_COLOR};color:#FFFFFF;font-weight:bold;">
          {header_cells}
        </tr>
      </thead>
      <tbody>
        {''.join(rows_html)}
      </tbody>
    </table>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)


def show_color_legend(show_red_name_convention=True):
    swatches = "".join(
        f'<span style="display:inline-block;margin-right:16px;">'
        f'<span style="display:inline-block;width:14px;height:14px;'
        f'background-color:{meta["hex"]};border:1px solid #999;'
        f'margin-right:6px;vertical-align:middle;"></span>'
        f'{meta["label"]}</span>'
        for key, meta in COLOR_LEGEND.items() if key != "NONE"
    )
    if show_red_name_convention:
        swatches += (
            '<span style="display:inline-block;">'
            '<span style="color:#B91C1C;font-weight:bold;margin-right:6px;">Nom du club</span>'
            '= non jouable en FM26</span>'
        )
    st.markdown(f"<div style='margin-bottom:8px;'>{swatches}</div>", unsafe_allow_html=True)


# ----------------------------------------------------------------------
# En-tête
# ----------------------------------------------------------------------
st.title("⚽ Palmarès Football FM — Tableau de bord")

tab_carte, tab_compare, tab_stats, tab_pays = st.tabs(
    ["🗺️ Carte interactive", "⚖️ Comparateur de clubs",
     "🏆 Statistiques globales", "🔎 Détail par compétition"]
)

# ========================================================================
# ONGLET 1 — CARTE INTERACTIVE
# ========================================================================
with tab_carte:
    st.subheader("Carte interactive")

    view_mode = st.radio(
        "Afficher",
        ["Championnats nationaux", "Compétitions continentales"],
        horizontal=True,
    )

    def format_division(r):
        """Ex: 'Ligue 1 (D1)' — combine le nom de la division et son niveau hiérarchique."""
        niveau = r.get("Niveau_division")
        if pd.isna(niveau):
            return r["Division_actuelle"]
        return f"{r['Division_actuelle']} (D{int(niveau)})"

    def wrap_long_text(text, max_len=45):
        """Insère un <br> près du milieu d'un texte trop long (ex: le
        libellé du statut orange, particulièrement long), pour que la
        fiche reste compacte plutôt que de s'étaler sur une seule ligne
        très large."""
        if len(text) <= max_len:
            return text
        mid = len(text) // 2
        # Cherche l'espace le plus proche du milieu pour couper un mot entier
        left_space = text.rfind(" ", 0, mid)
        right_space = text.find(" ", mid)
        if left_space == -1:
            cut = right_space
        elif right_space == -1:
            cut = left_space
        else:
            cut = left_space if (mid - left_space) <= (right_space - mid) else right_space
        if cut == -1:
            return text
        return text[:cut] + "<br>" + text[cut + 1:]

    def make_hover_text(r):
        txt = (
            f"<b>{r['Club']}</b><br>"
            f"Division : {format_division(r)}<br>"
            f"Dernier titre : {r['Annee_dernier_titre']}<br>"
            f"Titres : {r['Nb_titres']}"
        )
        if r.get("_non_jouable"):
            txt += "<br><span style='color:#B91C1C'><b>Non jouable en FM26</b></span>"
        return txt

    def apply_jitter(points_df, radius=0.05):
        """
        Quand plusieurs clubs partagent exactement les mêmes coordonnées
        (ex: un même stade), les répartit légèrement en cercle autour du
        point d'origine pour qu'ils restent distincts et cliquables sur la
        carte, plutôt que de se superposer parfaitement.
        """
        import math
        points_df = points_df.copy()
        groups = points_df.groupby(
            [points_df["Latitude"].round(4), points_df["Longitude"].round(4)]
        ).indices
        for _, idx_positions in groups.items():
            n = len(idx_positions)
            if n <= 1:
                continue
            row_indices = points_df.index[list(idx_positions)]
            for i, ridx in enumerate(row_indices):
                angle = 2 * math.pi * i / n
                points_df.loc[ridx, "Latitude"] += radius * math.sin(angle)
                points_df.loc[ridx, "Longitude"] += radius * math.cos(angle)
        return points_df

    # --------------------------------------------------------------
    # MODE "Compétitions continentales" : carte directe et navigable
    # de la compétition choisie, avec en option les finalistes
    # malheureux en triangles rouges.
    # --------------------------------------------------------------
    if view_mode == "Compétitions continentales":
        competition_choice = st.selectbox(
            "Compétition", list(CONTINENTAL_COMPETITIONS.keys())
        )
        show_finalistes = st.checkbox(
            "Afficher aussi les finalistes malheureux (triangles rouges)",
            value=True,
        )
        show_color_legend(show_red_name_convention=False)
        st.caption(
            "🔺 Triangle rouge = finaliste n'ayant jamais remporté cette "
            "compétition (à ne pas confondre avec le rond rouge 'club disparu')."
        )

        winners = df[
            (df["Categorie"] == "Continental - Vainqueurs")
            & (df["Competition"] == competition_choice)
        ].copy()
        winners["_prio"] = (winners["_color"].map(classify_color) != "NONE").astype(int)
        winners = winners.sort_values("_prio")  # points colorés dessinés en dernier = au-dessus
        winners["MarkerColor"] = winners["_color"].map(
            lambda c: COLOR_LEGEND[classify_color(c)]["marker"]
        )
        winners["OutlineColor"] = winners["_non_jouable"].map(
            lambda x: "#B91C1C" if x else "#444444"
        )
        winners["HoverText"] = winners.apply(make_hover_text, axis=1)
        winners = apply_jitter(winners)

        geo_fig = go.Figure()
        geo_fig.add_trace(go.Scattergeo(
            lat=winners["Latitude"], lon=winners["Longitude"],
            text=winners["HoverText"], hoverinfo="text",
            mode="markers", name="Vainqueurs",
            marker=dict(
                size=11, symbol="circle",
                color=winners["MarkerColor"],
                line=dict(width=1.6, color=winners["OutlineColor"]),
            ),
        ))

        if show_finalistes:
            finalistes = df[
                (df["Categorie"] == "Continental - Finalistes")
                & (df["Competition"] == competition_choice)
            ].copy()
            finalistes["HoverText"] = finalistes.apply(
                lambda r: (
                    f"<b>{r['Club']}</b><br>"
                    f"Division : {format_division(r)}<br>"
                    f"Finales perdues : {r['Nb_finales_perdues']}"
                ),
                axis=1,
            )
            finalistes = apply_jitter(finalistes)
            geo_fig.add_trace(go.Scattergeo(
                lat=finalistes["Latitude"], lon=finalistes["Longitude"],
                text=finalistes["HoverText"], hoverinfo="text",
                mode="markers", name="Finalistes malheureux",
                marker=dict(
                    size=11, symbol="triangle-up",
                    color="#DC2626",
                    line=dict(width=1, color="#444444"),
                ),
            ))

        all_lats = pd.concat([winners["Latitude"]] + ([finalistes["Latitude"]] if show_finalistes else []))
        all_lons = pd.concat([winners["Longitude"]] + ([finalistes["Longitude"]] if show_finalistes else []))
        # Bornes calculées dynamiquement selon la compétition (Europe, Amérique
        # du Sud, Afrique, Asie/Océanie ou Amérique du Nord) plutôt que fixées
        # sur l'Europe — chaque compétition obtient ainsi un cadrage adapté.
        geo_fig.update_geos(
            lataxis_range=[max(all_lats.min() - 3, -60), min(all_lats.max() + 3, 80)],
            lonaxis_range=[max(all_lons.min() - 3, -180), min(all_lons.max() + 3, 180)],
            visible=True,
            showcountries=True, countrycolor="#999999",
            showland=True, landcolor="#f2f2f2",
            showocean=True, oceancolor="#dce6f2",
            showlakes=False,
        )
        geo_fig.update_layout(
            margin=dict(l=0, r=0, t=10, b=0), height=560,
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.caption("Déplacez-vous et zoomez librement sur la carte. Survolez un point pour voir le détail du club.")
        st.plotly_chart(geo_fig, use_container_width=True, key=f"continental_map_{competition_choice}")

    # --------------------------------------------------------------
    # MODE "Championnats nationaux" : carte du monde par continent,
    # remplacée par la carte du pays au clic (avec bouton retour).
    # --------------------------------------------------------------
    else:
        source_df = df_national

        if "nat_selected_iso3" not in st.session_state:
            st.session_state.nat_selected_iso3 = None

        # -------- VUE PAYS (si un pays a été sélectionné) --------
        if st.session_state.nat_selected_iso3:
            iso3 = st.session_state.nat_selected_iso3
            candidates = sorted(
                source_df.loc[source_df["Iso3"] == iso3, "Pays"].unique()
            )

            if st.button("← Retour à la carte du monde"):
                st.session_state.nat_selected_iso3 = None
                st.rerun()

            if len(candidates) > 1:
                chosen_country = st.selectbox(
                    "Plusieurs championnats partagent ce territoire — choisissez lequel afficher :",
                    candidates,
                )
            else:
                chosen_country = candidates[0]

            st.markdown(f"### {chosen_country} — championnat national")
            show_color_legend(show_red_name_convention=False)
            st.caption("⭕ Contour rouge épais = division trop basse pour être jouable dans FM26.")

            clubs_df = source_df[source_df["Pays"] == chosen_country].copy()
            clubs_df["_prio"] = (clubs_df["_color"].map(classify_color) != "NONE").astype(int)
            clubs_df = clubs_df.sort_values("_prio")  # points colorés dessinés en dernier = au-dessus
            clubs_df["MarkerColor"] = clubs_df["_color"].map(
                lambda c: COLOR_LEGEND[classify_color(c)]["marker"]
            )
            clubs_df["OutlineColor"] = clubs_df["_non_jouable"].map(
                lambda x: "#B91C1C" if x else "#444444"
            )
            clubs_df["OutlineWidth"] = clubs_df["_non_jouable"].map(
                lambda x: 3 if x else 1
            )
            clubs_df["HoverText"] = clubs_df.apply(make_hover_text, axis=1)
            clubs_df = apply_jitter(clubs_df)

            geo_fig = go.Figure(go.Scattergeo(
                lat=clubs_df["Latitude"],
                lon=clubs_df["Longitude"],
                text=clubs_df["HoverText"],
                hoverinfo="text",
                mode="markers",
                marker=dict(
                    size=11,
                    color=clubs_df["MarkerColor"],
                    line=dict(width=clubs_df["OutlineWidth"], color=clubs_df["OutlineColor"]),
                ),
            ))
            geo_fig.update_geos(
                fitbounds="locations",
                visible=True,
                showcountries=True, countrycolor="#999999",
                showland=True, landcolor="#f2f2f2",
                showocean=True, oceancolor="#dce6f2",
                showlakes=False,
            )
            geo_fig.update_layout(margin=dict(l=0, r=0, t=10, b=0), height=560)
            st.caption("Survolez un point pour voir le détail du club.")
            st.plotly_chart(geo_fig, use_container_width=True, key=f"countrymap_{chosen_country}")

        # -------- VUE MONDE (par défaut) --------
        else:
            SMALL_TERRITORIES = ["Gibraltar", "Hong Kong", "Singapour"]
            available_small = [c for c in SMALL_TERRITORIES if c in source_df["Pays"].unique()]

            caption_text = (
                "Cliquez sur un pays pour zoomer et voir chaque club positionné "
                "sur la carte. Écosse, Pays de Galles et Irlande du Nord n'ayant "
                "pas de code ISO propre, ils sont affichés sur le même polygone "
                "que le Royaume-Uni (Angleterre) sur la carte du monde uniquement "
                "— leurs palmarès restent séparés dans les données et sur la "
                "carte pays."
            )
            if available_small:
                caption_text += (
                    " Pour les petits territoires difficiles à cliquer précisément "
                    "(Gibraltar, Hong Kong, Singapour), utilisez le menu "
                    "déroulant ci-dessous."
                )
            st.caption(caption_text)

            if available_small:
                direct_country = st.selectbox(
                    "Petit territoire",
                    ["— Sélectionner un petit territoire —"] + available_small,
                    key="direct_country_select",
                    label_visibility="collapsed",
                )
                if direct_country != "— Sélectionner un petit territoire —":
                    st.session_state.nat_selected_iso3 = source_df.loc[
                        source_df["Pays"] == direct_country, "Iso3"
                    ].iloc[0]
                    st.rerun()

            # Agrégation par code ISO3 (plusieurs 'Pays' peuvent partager un ISO3)
            agg_rows = []
            for iso3, grp in source_df.groupby("Iso3"):
                pays_list = sorted(grp["Pays"].unique())
                top3 = (
                    grp.groupby("Club")["Nb_titres"].max()
                    .sort_values(ascending=False)
                    .head(3)
                )
                top3_txt = "<br>".join(f"• {c} ({int(t)} titres)" for c, t in top3.items())
                agg_rows.append({
                    "Iso3": iso3,
                    "Pays_label": " / ".join(pays_list),
                    "Continent": grp["Continent"].iloc[0],
                    "Nb_clubs": grp["Club"].nunique(),
                    "Nb_titres_total": int(grp["Nb_titres"].sum()),
                    "Top3": top3_txt,
                })
            map_df = pd.DataFrame(agg_rows)

            fig = px.choropleth(
                map_df,
                locations="Iso3",
                color="Continent",
                hover_name="Pays_label",
                custom_data=["Top3", "Nb_clubs", "Nb_titres_total"],
                color_discrete_map=CONTINENT_COLORS,
                category_orders={"Continent": list(CONTINENT_COLORS.keys())},
            )
            fig.update_traces(
                hovertemplate=(
                    "<b>%{hovertext}</b><br>"
                    "Clubs recensés : %{customdata[1]}<br>"
                    "Titres cumulés : %{customdata[2]}<br>"
                    "<b>Clubs dominants :</b><br>%{customdata[0]}"
                    "<extra></extra>"
                )
            )
            fig.update_layout(
                margin=dict(l=0, r=0, t=10, b=0), height=560,
                legend_title_text="Confédération",
            )

            event = st.plotly_chart(
                fig, use_container_width=True, on_select="rerun", key="worldmap_national"
            )

            if event and event.get("selection", {}).get("points"):
                clicked_iso3 = event["selection"]["points"][0].get("location")
                if clicked_iso3:
                    st.session_state.nat_selected_iso3 = clicked_iso3
                    st.rerun()

# ----------------------------------------------------------------------
# Fonctions génériques de graphiques : le survol fonctionne sur TOUTE la
# largeur/hauteur de la ligne (y compris près du nom du club, sur l'axe),
# pas uniquement sur la barre ou le point lui-même. Une piste invisible
# est superposée à la barre/au point visible et c'est elle seule qui
# porte l'infobulle, pour n'avoir jamais qu'une seule fiche affichée à
# la fois, correctement positionnée (pas de rotation ni de superposition).
# ----------------------------------------------------------------------
def make_hbar_with_hover(labels, values, colors, card_texts, xaxis_title, height=None, x_range=None):
    labels = list(labels)
    max_val = x_range[1] if x_range else max(1, max(values) if len(values) else 1)
    fig = go.Figure()
    # Piste invisible pleine largeur (déclenche l'infobulle où que l'on
    # survole la ligne) + hovermode="y" (affiche l'infobulle à la position
    # réelle de la piste — ici à droite — plutôt qu'à l'endroit précis du
    # survol) : les deux combinés donnent "survol n'importe où sur la
    # ligne, fiche toujours au même endroit".
    fig.add_trace(go.Bar(
        x=[max_val] * len(labels), y=labels, orientation="h",
        marker=dict(color="rgba(0,0,0,0)"),
        customdata=list(card_texts), hovertemplate="%{customdata}<extra></extra>",
        showlegend=False,
    ))
    fig.add_trace(go.Bar(
        x=list(values), y=labels, orientation="h",
        marker=dict(color=list(colors)),
        hoverinfo="skip", showlegend=False,
    ))
    fig.update_layout(
        barmode="overlay", height=height or max(320, 38 * len(labels)),
        xaxis_title=xaxis_title, yaxis_title="",
        yaxis=dict(categoryorder="array", categoryarray=labels, autorange="reversed", automargin=True),
        hovermode="y",
    )
    # Graduations toujours horizontales (jamais en diagonale) : si l'échelle
    # est grande, on n'écrit qu'un chiffre sur deux, mais un trait fin reste
    # visible à chaque cran intermédiaire pour garder la précision visuelle.
    # Trait plus marqué sur les crans numérotés (pairs), plus discret sur
    # les crans intermédiaires non numérotés (impairs).
    label_step = x_range[2] if x_range and len(x_range) > 2 else 1
    grid_step = label_step
    if max_val > 15:
        label_step = grid_step * 2
    fig.update_xaxes(
        range=[0, max_val],
        tickangle=0,
        dtick=label_step,
        tick0=0,
        minor=dict(dtick=grid_step, showgrid=True, gridcolor="rgba(148,163,184,0.15)"),
        showgrid=True, gridcolor="rgba(148,163,184,0.35)",
    )
    return fig


def make_vbar_with_hover(labels, values, colors, card_texts, yaxis_title, height=380, tickvals=None, ticktext=None):
    labels = list(labels)
    max_val = max(1, max(values) if len(values) else 1)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=labels, y=[max_val] * len(labels),
        marker=dict(color="rgba(0,0,0,0)"),
        customdata=list(card_texts), hovertemplate="%{customdata}<extra></extra>",
        showlegend=False,
    ))
    fig.add_trace(go.Bar(
        x=labels, y=list(values),
        marker=dict(color=list(colors)),
        hoverinfo="skip", showlegend=False,
    ))
    yaxis_config = dict(title=yaxis_title)
    if tickvals is not None:
        yaxis_config.update(tickmode="array", tickvals=tickvals, ticktext=ticktext)
    # Au-delà de 10 clubs, un nom sur deux est poussé sur une ligne
    # supplémentaire (via <br>, nativement géré par Plotly) pour éviter
    # que les noms ne se chevauchent — sans passer par une trace de texte
    # personnalisée, qui cassait le survol.
    stagger = len(labels) > 10
    xaxis_ticktext = [("<br>" + l if stagger and i % 2 else l) for i, l in enumerate(labels)]
    fig.update_layout(
        barmode="overlay", height=height + (30 if stagger else 0),
        xaxis_title="",
        xaxis=dict(
            categoryorder="array", categoryarray=labels, automargin=True, tickangle=0,
            tickmode="array", tickvals=labels, ticktext=xaxis_ticktext,
        ),
        yaxis=yaxis_config,
        hoverdistance=100,
    )
    return fig


def make_timeline_with_hover(labels, years, colors, card_texts, height=None):
    labels = list(labels)
    years_clean = [y for y in years if pd.notna(y)]
    x_min = min(years_clean) if years_clean else 2000
    x_max = max(years_clean) if years_clean else 2026
    pad = max(1, (x_max - x_min) * 0.08)
    fig = go.Figure()

    # Piste invisible pleine largeur : UNE SEULE trace consolidée (segments
    # séparés par des None), pas une trace par club — avec hovermode="y",
    # plusieurs traces séparées déclenchaient chacune leur propre fiche
    # simultanément (plusieurs fiches empilées en diagonale). Une seule
    # trace élimine ce problème, comme pour les barres horizontales.
    line_x, line_y, line_customdata = [], [], []
    for label, year, card in zip(labels, years, card_texts):
        if pd.isna(year):
            continue
        line_x += [x_min - pad, x_max + pad, None]
        line_y += [label, label, None]
        line_customdata += [card, card, None]
    fig.add_trace(go.Scatter(
        x=line_x, y=line_y, mode="lines",
        line=dict(color="rgba(0,0,0,0)", width=22),
        customdata=line_customdata, hovertemplate="%{customdata}<extra></extra>",
        showlegend=False,
    ))

    fig.add_trace(go.Scatter(
        x=years, y=labels, mode="markers",
        marker=dict(size=16, color=list(colors), line=dict(width=1, color="#444444")),
        hoverinfo="skip", showlegend=False,
    ))

    fig.update_layout(
        height=height or max(320, 32 * len(labels)),
        xaxis_title="Année du dernier titre", yaxis_title="",
        xaxis=dict(range=[x_min - pad, x_max + pad]),
        yaxis=dict(categoryorder="array", categoryarray=labels, autorange="reversed", automargin=True),
        hovermode="y",
    )
    return fig

# ========================================================================
# ONGLET 2 — COMPARATEUR DE CLUBS
# ========================================================================
with tab_compare:
    st.subheader("Comparateur de clubs")
    st.caption(
        "Comparaison restreinte aux clubs mis en avant par un surlignage "
        "dans le fichier Excel (les autres clubs sont exclus pour garder "
        "une comparaison lisible)."
    )

    compare_mode = st.radio(
        "Type de comparaison",
        [
            "🔵 Clubs à la division la plus basse (bleu + orange)",
            "🟢 Clubs au titre le plus ancien (vert + orange)",
        ],
    )

    df_national_c = df_national.copy()
    df_national_c["Classified"] = df_national_c["_color"].map(classify_color)

    if compare_mode.startswith("🔵"):
        pool = df_national_c[df_national_c["Classified"].isin(["0070C0", "FF9900"])].copy()
        criterion_note = "Regroupe les clubs surlignés en bleu et en orange (l'orange combinant les deux critères)."
    else:
        pool = df_national_c[df_national_c["Classified"].isin(["00B050", "FF9900"])].copy()
        criterion_note = "Regroupe les clubs surlignés en vert et en orange (l'orange combinant les deux critères)."

    continent_choice = st.selectbox(
        "Continent",
        ["Tous les continents"] + sorted(pool["Continent"].dropna().unique().tolist()),
    )
    if continent_choice != "Tous les continents":
        pool = pool[pool["Continent"] == continent_choice].copy()

    st.caption(f"{criterion_note} — {pool['Club'].nunique()} clubs éligibles.")

    def filter_block(label, options, key_prefix, invert=False, value_labels=None):
        op = st.selectbox(
            label,
            ["Peu importe", "Inférieur à", "Supérieur à", "Entre"],
            key=f"{key_prefix}_op",
        )
        if op == "Peu importe":
            return None
        fmt = (lambda v: value_labels[v]) if value_labels else (lambda v: str(v))
        if op == "Entre":
            c1, c2 = st.columns(2)
            v1 = c1.selectbox("De", options, index=0, format_func=fmt, key=f"{key_prefix}_v1")
            v2 = c2.selectbox("À", options, index=len(options) - 1, format_func=fmt, key=f"{key_prefix}_v2")
            lo, hi = min(v1, v2), max(v1, v2)
            return ("between", lo, hi)
        val = st.selectbox("Valeur", options, format_func=fmt, key=f"{key_prefix}_v1")
        if op == "Inférieur à":
            actual_op = "gt" if invert else "lt"
        else:  # Supérieur à
            actual_op = "lt" if invert else "gt"
        return (actual_op, val)

    def apply_filter(series, filt):
        if filt is None:
            return pd.Series(True, index=series.index)
        if filt[0] == "lt":
            return series < filt[1]
        if filt[0] == "gt":
            return series > filt[1]
        return (series >= filt[1]) & (series <= filt[2])

    year_options = sorted(pool["Annee_dernier_titre"].dropna().unique().astype(int).tolist())
    niveau_options = list(range(1, 9))  # 1 à 8 fixe, même si tous les niveaux ne sont pas encore utilisés (pays/clubs futurs)
    titres_options = sorted(pool["Nb_titres"].dropna().unique().astype(int).tolist())
    niveau_labels = {
        v: f"Niveau {v}" + (" (le plus haut)" if v == min(niveau_options)
                             else " (le plus bas)" if v == max(niveau_options) else "")
        for v in niveau_options
    }

    st.markdown("#### Critères de sélection")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**Année du dernier titre**")
        filt_year = filter_block("Filtre année", year_options, "year")
    with col2:
        st.markdown("**Niveau de division**")
        st.caption("« Inférieur » = division plus basse dans la hiérarchie (niveau numérique plus élevé).")
        filt_niveau = filter_block("Filtre niveau", niveau_options, "niveau", invert=True, value_labels=niveau_labels)
    with col3:
        st.markdown("**Nombre de titres nationaux**")
        filt_titres = filter_block("Filtre titres", titres_options, "titres")

    if filt_year is None and filt_niveau is None and filt_titres is None:
        st.info("Choisissez au moins un critère ci-dessus pour afficher des clubs à comparer.")
        compare_df = pool.iloc[0:0]  # vide
    else:
        mask = (
            apply_filter(pool["Annee_dernier_titre"], filt_year)
            & apply_filter(pool["Niveau_division"], filt_niveau)
            & apply_filter(pool["Nb_titres"], filt_titres)
        )
        compare_df = pool[mask].copy()

    compare_df = compare_df.drop_duplicates(subset="Club").set_index("Club")

    st.divider()

    if compare_df.empty:
        if filt_year is not None or filt_niveau is not None or filt_titres is not None:
            st.info("Aucun club ne correspond à ces critères.")
    else:
        st.caption(f"{len(compare_df)} club(s) correspondant aux critères.")
        continental_winners = df[df["Categorie"] == "Continental - Vainqueurs"]
        compare_df["Titres_Continental"] = compare_df.index.map(
            lambda club: int(continental_winners.loc[continental_winners["Club"] == club, "Nb_titres"].sum())
        )
        compare_df["CouleurBarre"] = compare_df["Classified"].map(
            lambda c: COLOR_LEGEND[c]["marker"]
        )
        compare_df = compare_df.sort_values("Niveau_division")
        # Nom de club affiché avec son pays entre parenthèses, pour lever
        # toute ambiguïté quand deux clubs de pays différents ont un nom proche.
        def make_compare_card_text(r):
            lignes = [
                f"<b>{r.name}</b>",
                f"Pays : {r['Pays']}",
                f"Division : {format_division(r)}",
                f"Dernier titre : {r['Annee_dernier_titre']}",
                f"Titres nationaux : {r['Nb_titres']}",
                f"Titres continentaux : {r['Titres_Continental']}",
                wrap_long_text(COLOR_LEGEND[r["Classified"]]["label"]),
            ]
            return "<br>".join(lignes)

        compare_df["CardText"] = compare_df.apply(make_compare_card_text, axis=1)

        st.caption("Survolez la ligne d'un club (partout, y compris son nom) pour voir sa fiche complète.")

        st.markdown("#### Niveau de division actuel (1 = plus haut niveau du pays)")
        NIVEAU_MAX = 8  # échelle fixe (1 = élite ... 8 = niveau le plus bas du système)
        compare_df["NiveauHauteur"] = NIVEAU_MAX + 1 - compare_df["Niveau_division"]
        fig_niveau = make_vbar_with_hover(
            compare_df.index.tolist(), compare_df["NiveauHauteur"].tolist(),
            compare_df["CouleurBarre"].tolist(), compare_df["CardText"].tolist(),
            yaxis_title="Niveau de division",
            tickvals=list(range(1, NIVEAU_MAX + 1)),
            ticktext=[str(NIVEAU_MAX + 1 - v) for v in range(1, NIVEAU_MAX + 1)],
        )
        st.plotly_chart(fig_niveau, use_container_width=True)

        st.markdown("#### Année du dernier titre national (frise chronologique)")
        timeline_df = compare_df.sort_values("Annee_dernier_titre")
        fig_timeline = make_timeline_with_hover(
            timeline_df.index.tolist(), timeline_df["Annee_dernier_titre"].tolist(),
            timeline_df["CouleurBarre"].tolist(), timeline_df["CardText"].tolist(),
        )
        st.plotly_chart(fig_timeline, use_container_width=True)

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### Nombre de titres nationaux")
            nat_sorted = compare_df.sort_values("Nb_titres")
            fig_nat = make_hbar_with_hover(
                nat_sorted.index.tolist(), nat_sorted["Nb_titres"].tolist(),
                nat_sorted["CouleurBarre"].tolist(), nat_sorted["CardText"].tolist(),
                xaxis_title="Titres nationaux",
            )
            st.plotly_chart(fig_nat, use_container_width=True)
        with col2:
            st.markdown("#### Nombre de titres continentaux")
            cont_sorted = compare_df.sort_values("Titres_Continental")
            max_continental = max(1, int(compare_df["Titres_Continental"].max()))
            fig_c1 = make_hbar_with_hover(
                cont_sorted.index.tolist(), cont_sorted["Titres_Continental"].tolist(),
                cont_sorted["CouleurBarre"].tolist(), cont_sorted["CardText"].tolist(),
                xaxis_title="Titres continentaux", x_range=(0, max_continental, 1),
            )
            st.plotly_chart(fig_c1, use_container_width=True)

# ========================================================================
# ONGLET 3 — STATISTIQUES GLOBALES
# ========================================================================
with tab_stats:
    st.subheader("Club le plus titré par pays")

    stat_mode = st.radio(
        "Afficher",
        ["Championnat national", "Compétition continentale", "Les deux"],
        horizontal=True,
    )

    competition_for_stats = None
    if stat_mode == "Compétition continentale":
        competition_for_stats = st.selectbox(
            "Compétition", list(CONTINENTAL_COMPETITIONS.keys()), key="stats_competition"
        )

    COLOR_CHAMP = "#2563EB"  # bleu — championnat national
    COLOR_LDC = "#D97706"    # orange — compétition(s) continentale(s)

    # Titres nationaux par club+pays (le plus titré de chaque pays)
    nat_best = (
        df_national.loc[df_national.groupby("Pays")["Nb_titres"].idxmax()]
        [["Pays", "Club", "Nb_titres", "Division_actuelle", "Niveau_division", "Annee_dernier_titre"]]
        .rename(columns={"Nb_titres": "Titres_nat"})
    )

    if stat_mode == "Les deux":
        # Vue d'ensemble : les 5 compétitions continentales cumulées, pour
        # que tous les pays ayant un titre continental (peu importe lequel)
        # apparaissent — restreindre à une seule compétition viderait trop
        # de pays et nuirait à la lisibilité de cette vue globale.
        continental_df = df[df["Categorie"] == "Continental - Vainqueurs"]
        continental_label = "Compétitions continentales (5 cumulées)"
    elif competition_for_stats:
        continental_df = df[
            (df["Categorie"] == "Continental - Vainqueurs")
            & (df["Competition"] == competition_for_stats)
        ]
        continental_label = competition_for_stats
        # Titres par club (peut y avoir plusieurs clubs vainqueurs par pays ; on garde le plus titré)
        c1_best = (
            continental_df.loc[continental_df.groupby("Pays")["Nb_titres"].idxmax()]
            [["Pays", "Club", "Nb_titres", "Division_actuelle", "Niveau_division", "Annee_dernier_titre"]]
            .rename(columns={"Nb_titres": "Titres_c1", "Club": "Club_c1"})
        )

    if stat_mode == "Championnat national":
        chart_df = nat_best.copy()
        chart_df["Label"] = chart_df["Pays"] + " — " + chart_df["Club"]
        chart_df = chart_df.sort_values("Titres_nat")
        chart_df["CardText"] = chart_df.apply(
            lambda r: (
                f"<b>{r['Club']}</b><br>Pays : {r['Pays']}<br>"
                f"Division : {format_division(r)}<br>"
                f"Dernier titre : {r['Annee_dernier_titre']}<br>"
                f"Titres nationaux : {r['Titres_nat']}"
            ),
            axis=1,
        )
        fig = make_hbar_with_hover(
            chart_df["Label"].tolist(), chart_df["Titres_nat"].tolist(),
            [COLOR_CHAMP] * len(chart_df), chart_df["CardText"].tolist(),
            xaxis_title="Titres nationaux", height=max(500, 38 * len(chart_df)),
        )
        st.caption("Survolez la ligne d'un pays (partout, y compris son nom) pour voir la fiche complète.")
        st.plotly_chart(fig, use_container_width=True)

    elif stat_mode == "Compétition continentale":
        chart_df = c1_best.rename(columns={"Club_c1": "Club"}).copy()
        chart_df["Label"] = chart_df["Pays"] + " — " + chart_df["Club"]
        chart_df = chart_df.sort_values("Titres_c1")
        chart_df["CardText"] = chart_df.apply(
            lambda r: (
                f"<b>{r['Club']}</b><br>Pays : {r['Pays']}<br>"
                f"Division : {format_division(r)}<br>"
                f"Dernier titre : {r['Annee_dernier_titre']}<br>"
                f"Titres {competition_for_stats} : {r['Titres_c1']}"
            ),
            axis=1,
        )
        fig = make_hbar_with_hover(
            chart_df["Label"].tolist(), chart_df["Titres_c1"].tolist(),
            [COLOR_LDC] * len(chart_df), chart_df["CardText"].tolist(),
            xaxis_title=f"Titres — {competition_for_stats}", height=max(400, 38 * len(chart_df)),
        )
        st.caption("Survolez la ligne d'un pays (partout, y compris son nom) pour voir la fiche complète.")
        st.plotly_chart(fig, use_container_width=True)

    else:  # Les deux — club le plus titré (national + compétition choisie) par pays, barre empilée
        c1_by_club = continental_df.groupby(["Pays", "Club"])["Nb_titres"].sum().rename("Titres_c1_club")
        combo = df_national.groupby(["Pays", "Club"])["Nb_titres"].sum().rename("Titres_nat_club").reset_index()
        combo = combo.merge(c1_by_club, on=["Pays", "Club"], how="left")
        combo["Titres_c1_club"] = combo["Titres_c1_club"].fillna(0)
        combo["Total"] = combo["Titres_nat_club"] + combo["Titres_c1_club"]
        chart_df = combo.loc[combo.groupby("Pays")["Total"].idxmax()].copy()
        # Récupération du détail (division, niveau, dernier titre national)
        # perdu lors de l'agrégation, via jointure sur le profil national du club.
        details = df_national.drop_duplicates(subset=["Pays", "Club"])[
            ["Pays", "Club", "Division_actuelle", "Niveau_division", "Annee_dernier_titre"]
        ]
        chart_df = chart_df.merge(details, on=["Pays", "Club"], how="left")
        chart_df["Label"] = chart_df["Pays"] + " — " + chart_df["Club"]
        chart_df = chart_df.sort_values("Total")
        chart_df["CardText"] = chart_df.apply(
            lambda r: (
                f"<b>{r['Club']}</b><br>Pays : {r['Pays']}<br>"
                f"Division : {format_division(r)}<br>"
                f"Dernier titre national : {r['Annee_dernier_titre']}<br>"
                f"Titres nationaux : {int(r['Titres_nat_club'])}<br>"
                f"Titres continentaux : {int(r['Titres_c1_club'])}"
            ),
            axis=1,
        )

        fig = go.Figure()
        # Piste invisible pleine largeur + hovermode="y" (voir
        # make_hbar_with_hover pour l'explication détaillée) : survol
        # possible partout sur la ligne, fiche toujours affichée au même
        # endroit (à droite), sans jamais recouvrir le nom du pays.
        max_total = max(1, chart_df["Total"].max())
        labels_stack = chart_df["Label"].tolist()
        fig.add_trace(go.Bar(
            x=[max_total] * len(chart_df), y=labels_stack, orientation="h",
            marker=dict(color="rgba(0,0,0,0)"),
            customdata=chart_df["CardText"], hovertemplate="%{customdata}<extra></extra>",
            showlegend=False,
        ))
        fig.add_trace(go.Bar(
            x=chart_df["Titres_nat_club"], y=labels_stack, orientation="h", base=0,
            name="Championnat national", marker=dict(color=COLOR_CHAMP, line=dict(width=1.5, color="white")),
            text=chart_df["Titres_nat_club"], textposition="inside",
            hoverinfo="skip",
        ))
        fig.add_trace(go.Bar(
            x=chart_df["Titres_c1_club"], y=labels_stack, orientation="h",
            base=chart_df["Titres_nat_club"],
            name=continental_label, marker=dict(color=COLOR_LDC, line=dict(width=1.5, color="white")),
            text=chart_df["Titres_c1_club"], textposition="inside",
            hoverinfo="skip",
        ))
        fig.update_layout(
            barmode="overlay", height=max(500, 38 * len(chart_df)),
            xaxis_title=f"Titres (championnat + {continental_label})",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            yaxis=dict(categoryorder="array", categoryarray=labels_stack, autorange="reversed", automargin=True),
            hovermode="y",
        )
        label_step = 2 if max_total > 15 else 1
        fig.update_xaxes(
            range=[0, max_total], tickangle=0, dtick=label_step, tick0=0,
            minor=dict(dtick=1, showgrid=True, gridcolor="rgba(148,163,184,0.15)"),
            showgrid=True, gridcolor="rgba(148,163,184,0.35)",
        )
        st.caption("Survolez la ligne d'un pays (partout, y compris son nom) pour voir la fiche complète.")
        st.plotly_chart(fig, use_container_width=True)

# ========================================================================
# ONGLET 4 — DÉTAIL PAR PAYS / FEUILLE (formatage Excel respecté)
# ========================================================================
with tab_pays:
    st.subheader("Explorer une feuille du classeur")
    all_sheet_names = list(CONTINENTAL_COMPETITIONS.keys()) + sorted(df_national["Pays"].unique())
    choice = st.selectbox("Choisir une feuille", all_sheet_names)

    show_color_legend()
    is_continental = choice in CONTINENTAL_COMPETITIONS
    if is_continental:
        sub = df[
            (df["Categorie"] == "Continental - Vainqueurs") & (df["Competition"] == choice)
        ].sort_values("Nb_titres", ascending=False)
    else:
        sub = df_national[df_national["Pays"] == choice].sort_values("Nb_titres", ascending=False)

    if not sub.empty:
        sub_charts = sub.copy()
        # Gris neutre pour les clubs non surlignés (au lieu du quasi-noir
        # utilisé sur la carte, invisible en thème sombre sur des barres).
        CHART_NEUTRAL_COLOR = "#94A3B8"
        sub_charts["CouleurBarre"] = sub_charts["_color"].map(
            lambda c: CHART_NEUTRAL_COLOR if classify_color(c) == "NONE" else COLOR_LEGEND[classify_color(c)]["marker"]
        )

        continental_winners = df[df["Categorie"] == "Continental - Vainqueurs"]
        if not is_continental:
            sub_charts["Titres_Continental"] = sub_charts["Club"].map(
                lambda club: int(continental_winners.loc[continental_winners["Club"] == club, "Nb_titres"].sum())
            )

        def make_card_text(r):
            """Fiche complète du club, affichée au survol de n'importe quel graphique."""
            lignes = [f"<b>{r['Club']}</b>"]
            if is_continental:
                lignes.append(f"Pays : {r['Pays']}")
            lignes.append(f"Division : {format_division(r)}")
            lignes.append(f"Dernier titre : {r['Annee_dernier_titre']}")
            if is_continental:
                lignes.append(f"Titres {choice} : {r['Nb_titres']}")
            else:
                lignes.append(f"Titres nationaux : {r['Nb_titres']}")
                lignes.append(f"Titres continentaux : {r['Titres_Continental']}")
            lignes.append(wrap_long_text(r["Statut"]))
            if r.get("_non_jouable"):
                lignes.append("<span style='color:#B91C1C'><b>Non jouable en FM26</b></span>")
            return "<br>".join(lignes)

        sub_charts["CardText"] = sub_charts.apply(make_card_text, axis=1)

        st.caption("Survolez la ligne d'un club (partout, y compris son nom) pour voir sa fiche complète.")

        st.markdown("#### Niveau de division actuel (1 = plus haut niveau du pays)")
        st.caption(
            "Les clubs disparus n'ont pas de division actuelle et n'apparaissent "
            "pas sur ce graphique — ils restent visibles sur les autres."
        )
        NIVEAU_MAX = 8  # échelle fixe (1 = élite ... 8 = niveau le plus bas du système)
        niveau_sorted = sub_charts.dropna(subset=["Niveau_division"]).sort_values("Niveau_division").copy()
        niveau_sorted["NiveauHauteur"] = NIVEAU_MAX + 1 - niveau_sorted["Niveau_division"]
        fig_niveau = make_vbar_with_hover(
            niveau_sorted["Club"].tolist(), niveau_sorted["NiveauHauteur"].tolist(),
            niveau_sorted["CouleurBarre"].tolist(), niveau_sorted["CardText"].tolist(),
            yaxis_title="Niveau de division",
            tickvals=list(range(1, NIVEAU_MAX + 1)),
            ticktext=[str(NIVEAU_MAX + 1 - v) for v in range(1, NIVEAU_MAX + 1)],
        )
        st.plotly_chart(fig_niveau, use_container_width=True)

        st.markdown("#### Année du dernier titre (frise chronologique)")
        timeline_df = sub_charts.dropna(subset=["Annee_dernier_titre"]).sort_values("Annee_dernier_titre")
        fig_timeline = make_timeline_with_hover(
            timeline_df["Club"].tolist(), timeline_df["Annee_dernier_titre"].tolist(),
            timeline_df["CouleurBarre"].tolist(), timeline_df["CardText"].tolist(),
        )
        st.plotly_chart(fig_timeline, use_container_width=True)

        if is_continental:
            st.markdown(f"#### Nombre de titres — {choice}")
            titres_sorted = sub_charts.sort_values("Nb_titres")
            fig_titres = make_hbar_with_hover(
                titres_sorted["Club"].tolist(), titres_sorted["Nb_titres"].tolist(),
                titres_sorted["CouleurBarre"].tolist(), titres_sorted["CardText"].tolist(),
                xaxis_title="Titres",
            )
            st.plotly_chart(fig_titres, use_container_width=True)
        else:
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### Nombre de titres nationaux")
                nat_sorted = sub_charts.sort_values("Nb_titres")
                fig_nat = make_hbar_with_hover(
                    nat_sorted["Club"].tolist(), nat_sorted["Nb_titres"].tolist(),
                    nat_sorted["CouleurBarre"].tolist(), nat_sorted["CardText"].tolist(),
                    xaxis_title="Titres nationaux",
                )
                st.plotly_chart(fig_nat, use_container_width=True)
            with col2:
                st.markdown("#### Nombre de titres continentaux")
                cont_sorted = sub_charts.sort_values("Titres_Continental")
                max_cont = max(1, int(sub_charts["Titres_Continental"].max()))
                fig_cont = make_hbar_with_hover(
                    cont_sorted["Club"].tolist(), cont_sorted["Titres_Continental"].tolist(),
                    cont_sorted["CouleurBarre"].tolist(), cont_sorted["CardText"].tolist(),
                    xaxis_title="Titres continentaux", x_range=(0, max_cont, 1),
                )
                st.plotly_chart(fig_cont, use_container_width=True)
