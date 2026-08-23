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
    HEADER_COLOR,
    SHEET_C1,
    SHEET_FINALISTES,
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
df_c1 = df[df["Categorie"] == "Ligue des Champions"].copy()
df_finalistes = df[df["Categorie"] == "Finalistes C1"].copy()

# ----------------------------------------------------------------------
# Fonctions d'affichage (fidélité au formatage Excel)
# ----------------------------------------------------------------------

def render_styled_table(sub_df: pd.DataFrame, columns: list[tuple[str, str]]):
    """
    Affiche un DataFrame sous forme de tableau HTML reproduisant les
    couleurs de surlignage définies dans le fichier Excel d'origine.

    columns : liste de tuples (nom_colonne_df, libellé_affiché)
    """
    header_cells = "".join(f"<th>{label}</th>" for _, label in columns)
    rows_html = []
    for _, row in sub_df.iterrows():
        color_code = row.get("_color", "NONE")
        meta = COLOR_LEGEND.get(color_code, COLOR_LEGEND["NONE"])
        bg = meta["hex"]
        font_color = "#FFFFFF" if color_code == "FF0000" else "#000000"
        weight = "bold" if color_code == "FF0000" else "normal"
        cells = []
        for col, _ in columns:
            val = row.get(col, "")
            val = "" if pd.isna(val) else val
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

tab_carte, tab_compare, tab_stats, tab_pays, tab_data = st.tabs(
    ["🗺️ Carte interactive", "⚖️ Comparateur de clubs",
     "🏆 Statistiques globales", "🔎 Détail par pays", "📄 Données brutes"]
)

# ========================================================================
# ONGLET 1 — CARTE INTERACTIVE
# ========================================================================
with tab_carte:
    st.subheader("Carte interactive")

    view_mode = st.radio(
        "Afficher",
        ["Championnats nationaux", "Ligue des Champions"],
        horizontal=True,
    )

    def make_hover_text(r):
        return (
            f"<b>{r['Club']}</b><br>"
            f"Division : {r['Division_actuelle']}<br>"
            f"Dernier titre : {r['Annee_dernier_titre']}<br>"
            f"Titres : {r['Nb_titres']}<br>"
            f"{r['Statut']}"
        )

    # --------------------------------------------------------------
    # MODE "Ligue des Champions" : carte d'Europe directe, navigable,
    # avec en option les finalistes malheureux en triangles rouges.
    # --------------------------------------------------------------
    if view_mode == "Ligue des Champions":
        show_finalistes = st.checkbox(
            "Afficher aussi les finalistes malheureux (triangles rouges)",
            value=True,
        )
        show_color_legend()
        st.caption(
            "🔺 Triangle rouge = finaliste n'ayant jamais remporté la C1 "
            "(à ne pas confondre avec le rond rouge 'club disparu')."
        )

        winners = df_c1.copy()
        winners["MarkerColor"] = winners["_color"].map(
            lambda c: COLOR_LEGEND.get(c, COLOR_LEGEND["NONE"])["marker"]
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
                line=dict(width=1, color="#444444"),
            ),
        ))

        if show_finalistes:
            finalistes = df_finalistes.copy()
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
        geo_fig.update_geos(
            lataxis_range=[max(all_lats.min()-3, -60), min(all_lats.max()+3, 75)],
            lonaxis_range=[max(all_lons.min()-3, -30), min(all_lons.max()+3, 45)],
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
        st.plotly_chart(geo_fig, use_container_width=True, key="c1_europe_map")
        st.caption("Déplacez-vous et zoomez librement sur la carte. Survolez un point pour voir le détail du club.")

    # --------------------------------------------------------------
    # MODE "Championnats nationaux" : carte du monde par continent,
    # puis clic sur un pays pour zoomer sur ses clubs.
    # --------------------------------------------------------------
    else:
        source_df = df_national
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
            margin=dict(l=0, r=0, t=10, b=0), height=480,
            legend_title_text="Confédération",
        )

        event = st.plotly_chart(
            fig, use_container_width=True, on_select="rerun", key="worldmap_national"
        )

        selected_iso3 = None
        if event and event.get("selection", {}).get("points"):
            selected_iso3 = event["selection"]["points"][0].get("location")

        st.divider()
        if selected_iso3:
            candidates = sorted(map_df.loc[map_df["Iso3"] == selected_iso3, "Pays_label"]
                                 .iloc[0].split(" / "))
            if len(candidates) > 1:
                chosen_country = st.selectbox(
                    "Plusieurs championnats partagent ce territoire — choisissez lequel afficher :",
                    candidates,
                )
            else:
                chosen_country = candidates[0]

            st.markdown(f"### {chosen_country} — championnat national")
            show_color_legend()

            clubs_df = source_df[source_df["Pays"] == chosen_country].copy()
            clubs_df["MarkerColor"] = clubs_df["_color"].map(
                lambda c: COLOR_LEGEND.get(c, COLOR_LEGEND["NONE"])["marker"]
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
                    line=dict(width=1, color="#444444"),
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
            geo_fig.update_layout(margin=dict(l=0, r=0, t=10, b=0), height=480)
            st.plotly_chart(geo_fig, use_container_width=True, key=f"countrymap_{chosen_country}")
            st.caption("Survolez un point pour voir le détail du club.")
        else:
            st.info("Cliquez sur un pays de la carte pour afficher ses clubs positionnés géographiquement.")

# ========================================================================
# ONGLET 2 — COMPARATEUR DE CLUBS
# ========================================================================
with tab_compare:
    st.subheader("Comparer le nombre de titres nationaux entre plusieurs clubs")

    all_clubs = sorted(df_national["Club"].unique())
    default_selection = all_clubs[:0]
    selected_clubs = st.multiselect(
        "Sélectionnez des clubs à comparer", all_clubs, default=default_selection
    )
    include_c1 = st.checkbox("Inclure les titres de Ligue des Champions dans la comparaison", value=True)

    if selected_clubs:
        nat_titles = (
            df_national[df_national["Club"].isin(selected_clubs)]
            .groupby("Club")["Nb_titres"].sum()
        )
        rows = []
        for club in selected_clubs:
            row = {"Club": club, "Titres nationaux": int(nat_titles.get(club, 0))}
            if include_c1:
                c1_val = df_c1.loc[df_c1["Club"] == club, "Nb_titres"]
                row["Titres Ligue des Champions"] = int(c1_val.sum()) if not c1_val.empty else 0
            rows.append(row)
        comp_df = pd.DataFrame(rows)

        value_cols = ["Titres nationaux"] + (["Titres Ligue des Champions"] if include_c1 else [])
        melted = comp_df.melt(id_vars="Club", value_vars=value_cols,
                               var_name="Compétition", value_name="Titres")
        fig_cmp = px.bar(
            melted, x="Club", y="Titres", color="Compétition",
            barmode="group", text="Titres",
        )
        fig_cmp.update_layout(height=450, legend_title="")
        st.plotly_chart(fig_cmp, use_container_width=True)
        st.dataframe(comp_df.set_index("Club"), use_container_width=True)
    else:
        st.info("Choisissez au moins un club dans la liste ci-dessus.")

# ========================================================================
# ONGLET 3 — STATISTIQUES GLOBALES
# ========================================================================
with tab_stats:
    st.subheader("Clubs les plus titrés toutes compétitions confondues (C1 + championnats nationaux)")

    nat_sum = df_national.groupby("Club")["Nb_titres"].sum().rename("Titres_nationaux")
    c1_sum = df_c1.groupby("Club")["Nb_titres"].sum().rename("Titres_C1")
    total = pd.concat([nat_sum, c1_sum], axis=1).fillna(0)
    total["Total"] = total["Titres_nationaux"] + total["Titres_C1"]
    total = total.sort_values("Total", ascending=False)

    top_n = st.slider("Nombre de clubs à afficher", 5, 30, 15)
    top_total = total.head(top_n).reset_index().rename(columns={"index": "Club"})

    fig_top = px.bar(
        top_total.sort_values("Total"),
        x="Total", y="Club", orientation="h",
        hover_data=["Titres_nationaux", "Titres_C1"],
        color="Total", color_continuous_scale="Oranges",
    )
    fig_top.update_layout(height=max(400, 28 * top_n), showlegend=False,
                           coloraxis_showscale=False)
    st.plotly_chart(fig_top, use_container_width=True)

    st.dataframe(
        top_total.rename(columns={
            "Titres_nationaux": "Titres nationaux",
            "Titres_C1": "Titres Ligue des Champions",
        }).set_index("Club"),
        use_container_width=True,
    )

    st.divider()
    st.subheader("Finalistes malheureux de la Ligue des Champions")
    show_color_legend()
    render_styled_table(
        df_finalistes.sort_values("Nb_finales_perdues", ascending=False),
        [
            ("Club", "Club"),
            ("Division_actuelle", "Division actuelle"),
            ("Niveau_division", "Niveau"),
            ("Nb_finales_perdues", "Finales perdues"),
        ],
    )

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

# ========================================================================
# ONGLET 5 — DONNÉES BRUTES (DataFrame unifié complet)
# ========================================================================
with tab_data:
    st.subheader("DataFrame unifié (toutes feuilles agrégées)")
    st.caption("Colonne 'Source' = nom de la feuille Excel d'origine.")
    display_cols = [
        "Source", "Categorie", "Club", "Pays", "Division_actuelle",
        "Niveau_division", "Annee_dernier_titre", "Nb_titres",
        "Nb_finales_perdues", "Statut",
    ]
    st.dataframe(df[display_cols], use_container_width=True, height=500)

    csv = df[display_cols].to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Télécharger le DataFrame unifié (CSV)",
        data=csv,
        file_name="palmares_unifie.csv",
        mime="text/csv",
    )
