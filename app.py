import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(
    page_title="Traçabilité OF / Articles / Moules / Machines",
    page_icon="🏭",
    layout="wide"
)

# =========================
# Style
# =========================
st.markdown("""
<style>
.main-title {
    font-size: 2.4rem;
    font-weight: 700;
    margin-bottom: 0.2rem;
}
.sub-title {
    color: #666;
    margin-bottom: 1.2rem;
}
.block-container {
    padding-top: 1.5rem;
}
div[data-testid="stMetricValue"] {
    font-size: 1.6rem;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">Traçabilité OF / Articles / Moules / Machines</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Application avec chargement sécurisé du fichier Excel par l’utilisateur.</div>',
    unsafe_allow_html=True
)

# =========================
# Fonctions utilitaires
# =========================
def nettoyer_texte(serie):
    return serie.astype(str).str.strip()

def compter_series_par_machine(df, colonne_element):
    df_temp = df.copy()
    df_temp = df_temp.sort_values(by=["Machine", "Date", "OF"]).reset_index(drop=True)

    precedent_machine = df_temp["Machine"].shift(1)
    precedent_element = df_temp[colonne_element].shift(1)

    df_temp["Nouvelle_serie"] = (
        (df_temp["Machine"] != precedent_machine) |
        (df_temp[colonne_element] != precedent_element)
    )

    debuts_series = df_temp[df_temp["Nouvelle_serie"]].copy()

    resultat = (
        debuts_series.groupby([colonne_element, "Machine"])
        .size()
        .reset_index(name="Nombre series")
    )

    return resultat

@st.cache_data
def charger_et_nettoyer_donnees(fichier_excel):
    df = pd.read_excel(fichier_excel, engine="openpyxl")

    # =========================
    # 1er nettoyage métier
    # =========================
    df["Presse_str"] = (
        df["Presse"]
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )

    df["N_OF_str"] = (
        df["N_OF"]
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )

    machines_a_supprimer = ["3", "253", "692"]
    df = df[~df["Presse_str"].isin(machines_a_supprimer)].copy()
    df = df[~df["N_OF_str"].str.startswith("9000", na=False)].copy()

    # Supprimer lignes sans machine ou sans moule
    df = df.dropna(subset=["Presse", "SP_OF.SP_OUTIL.REF_OUTIL"]).copy()

    # =========================
    # Base de recherche
    # =========================
    base_recherche = df[
        [
            "N_OF",
            "Article",
            "LIB_ARTICLE",
            "SP_OF.SP_OUTIL.REF_OUTIL",
            "DHD_OF",
            "Presse"
        ]
    ].copy()

    base_recherche.columns = [
        "OF",
        "Code article",
        "Libellé article",
        "Moule",
        "Date",
        "Machine"
    ]

    for col in ["OF", "Code article", "Libellé article", "Moule"]:
        base_recherche[col] = nettoyer_texte(base_recherche[col])

    base_recherche["Machine"] = (
        base_recherche["Machine"]
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )

    base_recherche = base_recherche.replace(r'^\s*$', pd.NA, regex=True)

    # Supprimer valeurs parasites
    base_recherche = base_recherche[
        ~base_recherche["Moule"].astype(str).str.strip().str.lower().isin(["aucun", "aucune"])
    ].copy()

    base_recherche = base_recherche[
        ~base_recherche["Code article"].astype(str).str.strip().str.lower().eq("fictif")
    ].copy()

    # Date réelle pour le tri
    base_recherche["Date"] = pd.to_datetime(base_recherche["Date"], errors="coerce")

    # Supprimer lignes incomplètes
    base_recherche = base_recherche.dropna(
        subset=["OF", "Code article", "Libellé article", "Moule", "Date", "Machine"]
    ).copy()

    # Colonne d'affichage
    base_recherche["Date affichage"] = base_recherche["Date"].dt.strftime("%Y-%m-%d")

    base_recherche = base_recherche.drop_duplicates().reset_index(drop=True)

    return base_recherche

# =========================
# Chargement utilisateur
# =========================
st.sidebar.header("Chargement du fichier")

fichier_excel = st.sidebar.file_uploader(
    "Charger le fichier Excel",
    type=["xlsx"]
)

if fichier_excel is None:
    st.info("Charge le fichier Excel dans la barre latérale pour utiliser l'application.")
    st.stop()

try:
    base_recherche = charger_et_nettoyer_donnees(fichier_excel)
except Exception as e:
    st.error("Erreur lors du chargement ou du traitement du fichier.")
    st.exception(e)
    st.stop()

# =========================
# Filtres globaux
# =========================
st.sidebar.header("Filtres globaux")

articles = sorted([
    x for x in base_recherche["Code article"].dropna().unique().tolist()
    if x and str(x).lower() != "nan"
])

machines = sorted([
    x for x in base_recherche["Machine"].dropna().unique().tolist()
    if x and str(x).lower() != "nan"
])

moules = sorted([
    x for x in base_recherche["Moule"].dropna().unique().tolist()
    if x and str(x).lower() != "nan"
])

filtre_article = st.sidebar.selectbox("Code article", ["Tous"] + articles)
filtre_machine = st.sidebar.selectbox("Machine", ["Toutes"] + machines)
filtre_moule = st.sidebar.selectbox("Moule", ["Tous"] + moules)

base_filtre = base_recherche.copy()

if filtre_article != "Tous":
    base_filtre = base_filtre[base_filtre["Code article"] == filtre_article]

if filtre_machine != "Toutes":
    base_filtre = base_filtre[base_filtre["Machine"] == filtre_machine]

if filtre_moule != "Tous":
    base_filtre = base_filtre[base_filtre["Moule"] == filtre_moule]

base_filtre = base_filtre.reset_index(drop=True)

# =========================
# Indicateurs globaux
# =========================
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Enregistrements", len(base_filtre))
c2.metric("OF", base_filtre["OF"].nunique())
c3.metric("Articles", base_filtre["Code article"].nunique())
c4.metric("Moules", base_filtre["Moule"].nunique())
c5.metric("Machines", base_filtre["Machine"].nunique())

st.divider()

# =========================
# Indicateurs métier par séries
# =========================
series_moules_global = compter_series_par_machine(base_filtre, "Moule")
series_articles_global = compter_series_par_machine(base_filtre, "Code article")
series_of_global = compter_series_par_machine(base_filtre, "OF")

moule_top = (
    series_moules_global.groupby("Moule")["Nombre series"]
    .sum()
    .reset_index(name="Nombre")
    .sort_values(by="Nombre", ascending=False)
    .reset_index(drop=True)
)

article_top = (
    series_articles_global.groupby("Code article")["Nombre series"]
    .sum()
    .reset_index(name="Nombre")
    .sort_values(by="Nombre", ascending=False)
    .reset_index(drop=True)
)

of_top = (
    series_of_global.groupby("OF")["Nombre series"]
    .sum()
    .reset_index(name="Nombre")
    .sort_values(by="Nombre", ascending=False)
    .reset_index(drop=True)
)

moule_top_val = moule_top.iloc[0]["Moule"] if len(moule_top) > 0 else "-"
moule_top_n = int(moule_top.iloc[0]["Nombre"]) if len(moule_top) > 0 else 0

article_top_val = article_top.iloc[0]["Code article"] if len(article_top) > 0 else "-"
article_top_n = int(article_top.iloc[0]["Nombre"]) if len(article_top) > 0 else 0

of_top_val = of_top.iloc[0]["OF"] if len(of_top) > 0 else "-"
of_top_n = int(of_top.iloc[0]["Nombre"]) if len(of_top) > 0 else 0

st.subheader("Indicateurs métier")
k1, k2, k3 = st.columns(3)
k1.metric("Moule le plus monté", moule_top_val, delta=f"{moule_top_n} fois")
k2.metric("Article le plus utilisé", article_top_val, delta=f"{article_top_n} fois")
k3.metric("OF le plus fréquent", of_top_val, delta=f"{of_top_n} fois")

st.divider()

# =========================
# Onglets de recherche
# =========================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Recherche par article",
    "Recherche par OF",
    "Recherche par moule",
    "Recherche par machine",
    "Analyse par machine",
    "Diagrammes globaux",
    "Base complète"
])

with tab1:
    st.subheader("Recherche par code article")
    article_input = st.text_input("Entrer un code article", key="article_input")

    if article_input:
        resultat = base_recherche[
            base_recherche["Code article"].astype(str).str.contains(article_input.strip(), case=False, na=False)
        ][["Code article", "Libellé article", "OF", "Date affichage", "Moule", "Machine"]].drop_duplicates()

        resultat = resultat.sort_values(by=["Date affichage", "OF"]).reset_index(drop=True)
        st.write(f"Résultats : {len(resultat)} ligne(s)")
        st.dataframe(resultat, use_container_width=True)

with tab2:
    st.subheader("Recherche par OF")
    of_input = st.text_input("Entrer un OF", key="of_input")

    if of_input:
        resultat = base_recherche[
            base_recherche["OF"].astype(str).str.strip() == of_input.strip()
        ][["OF", "Code article", "Libellé article", "Date affichage", "Moule", "Machine"]].drop_duplicates()

        resultat = resultat.sort_values(by=["Date affichage"]).reset_index(drop=True)
        st.write(f"Résultats : {len(resultat)} ligne(s)")
        st.dataframe(resultat, use_container_width=True)

with tab3:
    st.subheader("Recherche par moule")
    moule_input = st.text_input("Entrer un moule", key="moule_input")

    if moule_input:
        resultat = base_recherche[
            base_recherche["Moule"].astype(str).str.contains(moule_input.strip(), case=False, na=False)
        ][["Moule", "Machine", "OF", "Code article", "Libellé article", "Date affichage"]].drop_duplicates()

        resultat = resultat.sort_values(by=["Machine", "Date affichage"]).reset_index(drop=True)
        st.write(f"Résultats : {len(resultat)} ligne(s)")
        st.dataframe(resultat, use_container_width=True)

with tab4:
    st.subheader("Recherche par machine")
    machine_input = st.text_input("Entrer une machine", key="machine_input")

    if machine_input:
        resultat = base_recherche[
            base_recherche["Machine"].astype(str).str.strip() == machine_input.strip()
        ][["Machine", "Moule", "OF", "Code article", "Libellé article", "Date affichage"]].drop_duplicates()

        resultat = resultat.sort_values(by=["Moule", "Date affichage"]).reset_index(drop=True)
        st.write(f"Résultats : {len(resultat)} ligne(s)")
        st.dataframe(resultat, use_container_width=True)

with tab5:
    st.subheader("Analyse : combien de séries sur chaque machine")

    type_recherche = st.selectbox(
        "Choisir le type de recherche",
        ["Article", "OF", "Moule"],
        key="type_analyse"
    )

    valeur_recherche = st.text_input("Entrer la valeur à analyser", key="valeur_analyse")

    if valeur_recherche:
        if type_recherche == "Article":
            resultat = base_recherche[
                base_recherche["Code article"].astype(str).str.contains(valeur_recherche.strip(), case=False, na=False)
            ].copy()
            colonne = "Code article"

        elif type_recherche == "OF":
            resultat = base_recherche[
                base_recherche["OF"].astype(str).str.strip() == valeur_recherche.strip()
            ].copy()
            colonne = "OF"

        else:
            resultat = base_recherche[
                base_recherche["Moule"].astype(str).str.contains(valeur_recherche.strip(), case=False, na=False)
            ].copy()
            colonne = "Moule"

        if len(resultat) == 0:
            st.warning("Aucun résultat trouvé.")
        else:
            resume_machine = compter_series_par_machine(resultat, colonne)
            resume_machine = resume_machine.rename(columns={"Nombre series": "Nombre de fois"})
            st.dataframe(resume_machine, use_container_width=True)

with tab6:
    st.subheader("Diagrammes globaux")

    # Listes complètes
    liste_complete_moules = pd.DataFrame({
        "Moule": sorted(base_recherche["Moule"].dropna().unique())
    })

    liste_complete_articles = pd.DataFrame({
        "Code article": sorted(base_recherche["Code article"].dropna().unique())
    })

    liste_complete_of = pd.DataFrame({
        "OF": sorted(base_recherche["OF"].dropna().unique())
    })

    # Séries par machine
    detail_machine_moule = compter_series_par_machine(base_filtre, "Moule")
    detail_machine_article = compter_series_par_machine(base_filtre, "Code article")
    detail_machine_of = compter_series_par_machine(base_filtre, "OF")

    # Totaux
    compte_moules = (
        detail_machine_moule.groupby("Moule")["Nombre series"]
        .sum()
        .reset_index(name="Nombre montage")
    )

    compte_articles = (
        detail_machine_article.groupby("Code article")["Nombre series"]
        .sum()
        .reset_index(name="Nombre utilisation")
    )

    compte_of = (
        detail_machine_of.groupby("OF")["Nombre series"]
        .sum()
        .reset_index(name="Nombre occurrence")
    )

    all_moules = liste_complete_moules.merge(compte_moules, on="Moule", how="left")
    all_articles = liste_complete_articles.merge(compte_articles, on="Code article", how="left")
    all_of = liste_complete_of.merge(compte_of, on="OF", how="left")

    all_moules["Nombre montage"] = all_moules["Nombre montage"].fillna(0).astype(int)
    all_articles["Nombre utilisation"] = all_articles["Nombre utilisation"].fillna(0).astype(int)
    all_of["Nombre occurrence"] = all_of["Nombre occurrence"].fillna(0).astype(int)

    # Détail machines
    detail_machine_moule["Texte machine"] = (
        detail_machine_moule["Machine"].astype(str)
        + " : "
        + detail_machine_moule["Nombre series"].astype(int).astype(str)
        + " fois"
    )
    resume_machine_moule = (
        detail_machine_moule.groupby("Moule")["Texte machine"]
        .apply(lambda x: " | ".join(x))
        .reset_index(name="Detail machines")
    )
    all_moules = all_moules.merge(resume_machine_moule, on="Moule", how="left")
    all_moules["Detail machines"] = all_moules["Detail machines"].fillna("Aucune machine")

    detail_machine_article["Texte machine"] = (
        detail_machine_article["Machine"].astype(str)
        + " : "
        + detail_machine_article["Nombre series"].astype(int).astype(str)
        + " fois"
    )
    resume_machine_article = (
        detail_machine_article.groupby("Code article")["Texte machine"]
        .apply(lambda x: " | ".join(x))
        .reset_index(name="Detail machines")
    )
    all_articles = all_articles.merge(resume_machine_article, on="Code article", how="left")
    all_articles["Detail machines"] = all_articles["Detail machines"].fillna("Aucune machine")

    detail_machine_of["Texte machine"] = (
        detail_machine_of["Machine"].astype(str)
        + " : "
        + detail_machine_of["Nombre series"].astype(int).astype(str)
        + " fois"
    )
    resume_machine_of = (
        detail_machine_of.groupby("OF")["Texte machine"]
        .apply(lambda x: " | ".join(x))
        .reset_index(name="Detail machines")
    )
    all_of = all_of.merge(resume_machine_of, on="OF", how="left")
    all_of["Detail machines"] = all_of["Detail machines"].fillna("Aucune machine")

    all_moules = all_moules.sort_values(by="Nombre montage", ascending=False).reset_index(drop=True)
    all_articles = all_articles.sort_values(by="Nombre utilisation", ascending=False).reset_index(drop=True)
    all_of = all_of.sort_values(by="Nombre occurrence", ascending=False).reset_index(drop=True)

    st.markdown("### Top moules")
    chart_moules = alt.Chart(all_moules).mark_bar().encode(
        x=alt.X("Moule:N", sort="-y", title="Numero moule"),
        y=alt.Y("Nombre montage:Q", title="Nombre montage"),
        tooltip=[
            alt.Tooltip("Moule:N", title="Moule"),
            alt.Tooltip("Nombre montage:Q", title="Nombre montage"),
            alt.Tooltip("Detail machines:N", title="Machines")
        ]
    ).properties(height=400)
    st.altair_chart(chart_moules, use_container_width=True)

    st.markdown("### Top articles")
    chart_articles = alt.Chart(all_articles).mark_bar().encode(
        x=alt.X("Code article:N", sort="-y", title="Code article"),
        y=alt.Y("Nombre utilisation:Q", title="Nombre utilisation"),
        tooltip=[
            alt.Tooltip("Code article:N", title="Code article"),
            alt.Tooltip("Nombre utilisation:Q", title="Nombre utilisation"),
            alt.Tooltip("Detail machines:N", title="Machines")
        ]
    ).properties(height=400)
    st.altair_chart(chart_articles, use_container_width=True)

    st.markdown("### Top OF")
    chart_of = alt.Chart(all_of).mark_bar().encode(
        x=alt.X("OF:N", sort="-y", title="OF"),
        y=alt.Y("Nombre occurrence:Q", title="Nombre occurrence"),
        tooltip=[
            alt.Tooltip("OF:N", title="OF"),
            alt.Tooltip("Nombre occurrence:Q", title="Nombre occurrence"),
            alt.Tooltip("Detail machines:N", title="Machines")
        ]
    ).properties(height=400)
    st.altair_chart(chart_of, use_container_width=True)

with tab7:
    st.subheader("Base complète")
    base_affichage = base_filtre[
        ["OF", "Code article", "Libellé article", "Date affichage", "Machine", "Moule"]
    ].copy()

    base_affichage = base_affichage.sort_values(by=["Date affichage", "OF"]).reset_index(drop=True)
    st.dataframe(base_affichage, use_container_width=True, height=500)

    csv = base_affichage.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="Télécharger la base filtrée en CSV",
        data=csv,
        file_name="base_recherche_filtre.csv",
        mime="text/csv"
    )

st.divider()
st.caption("Le fichier Excel est chargé par l’utilisateur dans l’interface et n’est pas inclus dans le dépôt GitHub.")
