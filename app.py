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
    page_title="Palmarès Football FM26",
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
    st.title("⚽ Palmarès Football FM26")
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


def show_color_legend():
    swatches = "".join(
        f'<span style="display:inline-block;margin-right:16px;">'
        f'<span style="display:inline-block;width:14px;height:14px;'
        f'background-color:{meta["hex"]};border:1px solid #999;'
        f'margin-right:6px;vertical-align:middle;"></span>'
        f'{meta["label"]}</span>'
        for key, meta in COLOR_LEGEND.items() if key != "NONE"
    )
    swatches += (
        '<span style="display:inline-block;">'
        '<span style="color:#B91C1C;font-weight:bold;margin-right:6px;">Nom du club</span>'
        '= non jouable en FM26</span>'
    )
    st.markdown(f"<div style='margin-bottom:8px;'>{swatches}</div>", unsafe_allow_html=True)


# ----------------------------------------------------------------------
# En-tête
# ----------------------------------------------------------------------
st.title("⚽ Palmarès Football FM26 — Tableau de bord")
st.caption(
    f"{df['Source'].nunique()} feuilles agrégées · "
    f"{len(df_national)} lignes de palmarès nationaux · "
    f"{len(df_c1)} vainqueurs de C1 · {len(df_finalistes)} finalistes malheureux"
)

tab_carte, tab_compare, tab_stats, tab_pays = st.tabs(
    ["🗺️ Carte interactive", "⚖️ Comparateur de clubs",
     "🏆 Statistiques globales", "🔎 Détail par pays"]
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

    def make_hover_text(r):
        txt = (
            f"<b>{r['Club']}</b><br>"
            f"Division : {r['Division_actuelle']}<br>"
            f"Dernier titre : {r['Annee_dernier_titre']}<br>"
            f"Titres : {r['Nb_titres']}<br>"
            f"{r['Statut']}"
        )
        if r.get("_non_jouable"):
            txt += "<br><span style='color:#B91C1C'><b>Non jouable en FM26</b></span>"
        return txt

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
        show_color_legend()
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
                    f"Division : {r['Division_actuelle']}<br>"
                    f"Finales perdues : {r['Nb_finales_perdues']}"
                ),
                axis=1,
            )
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
        st.plotly_chart(geo_fig, use_container_width=True, key=f"continental_map_{competition_choice}")
        st.caption("Déplacez-vous et zoomez librement sur la carte. Survolez un point pour voir le détail du club.")

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
            show_color_legend()
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
            st.plotly_chart(geo_fig, use_container_width=True, key=f"countrymap_{chosen_country}")
            st.caption("Survolez un point pour voir le détail du club.")

        # -------- VUE MONDE (par défaut) --------
        else:
            st.caption(
                "Cliquez sur un pays pour zoomer et voir chaque club positionné "
                "sur la carte. Écosse, Pays de Galles et Irlande du Nord n'ayant "
                "pas de code ISO propre, ils sont affichés sur le même polygone "
                "que le Royaume-Uni (Angleterre) sur la carte du monde uniquement "
                "— leurs palmarès restent séparés dans les données et sur la "
                "carte pays."
            )

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

            st.info("Cliquez sur un pays de la carte pour afficher ses clubs positionnés géographiquement.")

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
        compare_df["ClubLabel"] = compare_df.index + " (" + compare_df["Pays"] + ")"
        label_color_map = dict(zip(compare_df["ClubLabel"], compare_df["CouleurBarre"]))

        st.markdown("#### Niveau de division actuel (1 = plus haut niveau du pays)")
        fig_niveau = px.bar(
            compare_df.reset_index(), x="ClubLabel", y="Niveau_division",
            color="ClubLabel", color_discrete_map=label_color_map,
        )
        fig_niveau.update_layout(height=380, showlegend=False, yaxis_title="Niveau de division", xaxis_title="")
        st.plotly_chart(fig_niveau, use_container_width=True)

        st.markdown("#### Année du dernier titre national (frise chronologique)")
        timeline_df = compare_df.reset_index().sort_values("Annee_dernier_titre")
        fig_timeline = px.scatter(
            timeline_df, x="Annee_dernier_titre", y="ClubLabel",
            color="ClubLabel", color_discrete_map=label_color_map,
        )
        fig_timeline.update_traces(marker=dict(size=14, line=dict(width=1, color="#444444")))
        fig_timeline.update_layout(
            height=max(320, 28 * len(timeline_df)), showlegend=False,
            xaxis_title="Année du dernier titre", yaxis_title="",
        )
        st.plotly_chart(fig_timeline, use_container_width=True)

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### Nombre de titres nationaux")
            fig_nat = px.bar(
                compare_df.reset_index().sort_values("Nb_titres"),
                x="Nb_titres", y="ClubLabel", orientation="h",
                color="ClubLabel", color_discrete_map=label_color_map,
            )
            fig_nat.update_layout(height=max(320, 26 * len(compare_df)), showlegend=False, yaxis_title="")
            st.plotly_chart(fig_nat, use_container_width=True)
        with col2:
            st.markdown("#### Nombre de titres continentaux")
            fig_c1 = px.bar(
                compare_df.reset_index().sort_values("Titres_Continental"),
                x="Titres_Continental", y="ClubLabel", orientation="h",
                color="ClubLabel", color_discrete_map=label_color_map,
            )
            fig_c1.update_layout(height=max(320, 26 * len(compare_df)), showlegend=False, yaxis_title="")
            st.plotly_chart(fig_c1, use_container_width=True)

        st.divider()
        st.dataframe(
            compare_df[["Pays", "Division_actuelle", "Niveau_division", "Annee_dernier_titre", "Nb_titres", "Titres_Continental"]]
            .rename(columns={
                "Division_actuelle": "Division actuelle", "Niveau_division": "Niveau",
                "Annee_dernier_titre": "Dernier titre", "Nb_titres": "Titres nationaux",
                "Titres_Continental": "Titres continentaux",
            }),
            use_container_width=True,
        )

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
        [["Pays", "Club", "Nb_titres"]]
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
            [["Pays", "Club", "Nb_titres"]]
            .rename(columns={"Nb_titres": "Titres_c1", "Club": "Club_c1"})
        )

    if stat_mode == "Championnat national":
        chart_df = nat_best.copy()
        chart_df["Label"] = chart_df["Pays"] + " — " + chart_df["Club"]
        chart_df = chart_df.sort_values("Titres_nat")
        fig = go.Figure(go.Bar(
            x=chart_df["Titres_nat"], y=chart_df["Label"], orientation="h",
            marker_color=COLOR_CHAMP,
            text=chart_df["Titres_nat"], textposition="outside",
        ))
        fig.update_layout(height=max(500, 24 * len(chart_df)), xaxis_title="Titres nationaux")
        st.plotly_chart(fig, use_container_width=True)

    elif stat_mode == "Compétition continentale":
        chart_df = c1_best.rename(columns={"Club_c1": "Club"}).copy()
        chart_df["Label"] = chart_df["Pays"] + " — " + chart_df["Club"]
        chart_df = chart_df.sort_values("Titres_c1")
        fig = go.Figure(go.Bar(
            x=chart_df["Titres_c1"], y=chart_df["Label"], orientation="h",
            marker_color=COLOR_LDC,
            text=chart_df["Titres_c1"], textposition="outside",
        ))
        fig.update_layout(height=max(400, 24 * len(chart_df)), xaxis_title=f"Titres — {competition_for_stats}")
        st.plotly_chart(fig, use_container_width=True)

    else:  # Les deux — club le plus titré (national + compétition choisie) par pays, barre empilée
        c1_by_club = continental_df.groupby(["Pays", "Club"])["Nb_titres"].sum().rename("Titres_c1_club")
        combo = df_national.groupby(["Pays", "Club"])["Nb_titres"].sum().rename("Titres_nat_club").reset_index()
        combo = combo.merge(c1_by_club, on=["Pays", "Club"], how="left")
        combo["Titres_c1_club"] = combo["Titres_c1_club"].fillna(0)
        combo["Total"] = combo["Titres_nat_club"] + combo["Titres_c1_club"]
        chart_df = combo.loc[combo.groupby("Pays")["Total"].idxmax()].copy()
        chart_df["Label"] = chart_df["Pays"] + " — " + chart_df["Club"]
        chart_df = chart_df.sort_values("Total")

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=chart_df["Titres_nat_club"], y=chart_df["Label"], orientation="h",
            name="Championnat national", marker=dict(color=COLOR_CHAMP, line=dict(width=1.5, color="white")),
            text=chart_df["Titres_nat_club"], textposition="inside",
        ))
        fig.add_trace(go.Bar(
            x=chart_df["Titres_c1_club"], y=chart_df["Label"], orientation="h",
            name=continental_label, marker=dict(color=COLOR_LDC, line=dict(width=1.5, color="white")),
            text=chart_df["Titres_c1_club"], textposition="inside",
        ))
        fig.update_layout(
            barmode="stack", height=max(500, 24 * len(chart_df)),
            xaxis_title=f"Titres (championnat + {continental_label})",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig, use_container_width=True)

# ========================================================================
# ONGLET 4 — DÉTAIL PAR PAYS / FEUILLE (formatage Excel respecté)
# ========================================================================
with tab_pays:
    st.subheader("Explorer une feuille du classeur")
    all_sheet_names = [SHEET_C1] + sorted(df_national["Pays"].unique())
    choice = st.selectbox("Choisir une feuille", all_sheet_names)

    show_color_legend()
    if choice == SHEET_C1:
        sub = df_c1.sort_values("Nb_titres", ascending=False)
        render_styled_table(sub, [
            ("Club", "Club"),
            ("Division_actuelle", "Division actuelle"),
            ("Niveau_division", "Niveau"),
            ("Annee_dernier_titre", "Dernier titre"),
            ("Nb_titres", "Titres"),
            ("Pays", "Pays"),
        ])
    else:
        sub = df_national[df_national["Pays"] == choice].sort_values("Nb_titres", ascending=False)
        render_styled_table(sub, [
            ("Club", "Club"),
            ("Division_actuelle", "Division actuelle"),
            ("Niveau_division", "Niveau"),
            ("Annee_dernier_titre", "Dernier titre"),
            ("Nb_titres", "Titres"),
        ])
