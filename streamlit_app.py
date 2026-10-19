import streamlit as st
import pandas as pd
import os
import plotly.graph_objects as go
import plotly.express as px
from pathlib import Path

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FAAB PROJET | Premium Management Tool",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── PREMIUM CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap');
html, body, [class*="css"] { font-family: 'Outfit', sans-serif; }

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #064e3b 0%, #065f46 100%);
    color: white;
}

.main-title {
    font-weight: 800;
    background: linear-gradient(90deg, #10b981, #3b82f6, #6366f1);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 2.2rem;
    text-align: center;
    margin-bottom: 1rem;
}

.header-card {
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
    border: 1px solid #334155;
    border-left: 5px solid #10b981;
    border-radius: 12px;
    padding: 20px 24px;
    margin-bottom: 1.5rem;
    color: #f8fafc;
}
.header-card h2 { margin: 0 0 4px 0; color: #10b981; font-size: 1.3rem; }
.header-card small { color: #94a3b8; font-size: 0.8rem; }
.header-meta { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; margin-top: 14px; }
.header-meta-item small { color: #64748b; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; }
.header-meta-item b { color: #f1f5f9; font-size: 0.9rem; display: block; }

.glass-header {
    background: rgba(16,185,129,0.08);
    border-left: 5px solid #10b981;
    padding: 10px 16px;
    border-radius: 0 8px 8px 0;
    margin: 1.2rem 0 0.8rem 0;
    font-weight: 700;
    color: #065f46;
    font-size: 1rem;
}

.stat-card {
    background: linear-gradient(135deg, #0f172a, #1e293b);
    border: 1px solid #334155;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
    color: white;
}
.stat-card .val { font-size: 2rem; font-weight: 800; }
.stat-card .lbl { font-size: 0.72rem; color: #94a3b8; text-transform: uppercase; margin-top: 4px; }

.status-livré    { background:#dcfce7; color:#166534; padding:2px 8px; border-radius:4px; font-size:0.72rem; font-weight:700; }
.status-encours  { background:#dbeafe; color:#1e40af; padding:2px 8px; border-radius:4px; font-size:0.72rem; font-weight:700; }
.status-retard   { background:#fee2e2; color:#991b1b; padding:2px 8px; border-radius:4px; font-size:0.72rem; font-weight:700; }
.status-attente  { background:#f1f5f9; color:#475569; padding:2px 8px; border-radius:4px; font-size:0.72rem; font-weight:700; }

.stButton > button {
    width: 100%; text-align: left !important;
    background: rgba(255,255,255,0.05); color: #ecf3f0 !important;
    border: 1px solid rgba(255,255,255,0.05);
    padding: 0.5rem 1rem; border-radius: 8px;
    transition: all 0.2s ease; margin-bottom: 0.2rem; font-size: 0.82rem;
}
.stButton > button:hover {
    background: #10b981 !important; border-color: #10b981 !important;
    color: white !important; transform: translateX(4px);
}

div[data-testid="stMetricValue"] { color: #10b981; font-weight: 800; }
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ─────────────────────────────────────────────────────────────────
KIT_DIR = Path("KIT CHEF PROJET")
if not KIT_DIR.exists():
    KIT_DIR = Path(".")

@st.cache_data
def scan_kit(directory):
    files_data = []
    exts = ('.xlsx', '.xlsm', '.xlsb', '.xls')
    if not directory.exists(): return []
    for root, _, filenames in os.walk(directory):
        for fname in filenames:
            if fname.endswith(exts) and not fname.startswith('~$'):
                fpath = Path(root) / fname
                cat = Path(root).name
                cat = "📁 GÉNÉRAL" if cat == directory.name else f"📂 {cat.upper()}"
                files_data.append({"name": fname, "path": fpath, "category": cat})
    return sorted(files_data, key=lambda x: (x["category"], x["name"]))

def render_project_header(title, subtitle=""):
    st.markdown(f"""
    <div class="header-card">
        <h2>{title}</h2>
        <small>{subtitle}</small>
        <div class="header-meta">
            <div class="header-meta-item"><small>Chef de projet</small><b>Mme Dupont</b></div>
            <div class="header-meta-item"><small>Nom / Code</small><b>SAP30 / CMRL</b></div>
            <div class="header-meta-item"><small>Organisation</small><b>Confassis S.A.</b></div>
            <div class="header-meta-item"><small>Version</small><b>V2.3</b></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ─── VUE : SUIVI DES LIVRABLES ───────────────────────────────────────────────
def render_suivi_livrables():
    render_project_header("23 — Suivi des Livrables", "Registre et suivi des livrables du projet")

    col1, col2, col3, col4 = st.columns(4)
    for col, val, lbl, color in [
        (col1, "43", "Total Livrables", "#10b981"),
        (col2, "28", "Livrés ✅", "#10b981"),
        (col3, "12", "En Cours 🔵", "#3b82f6"),
        (col4, "3",  "En Retard 🔴", "#ef4444"),
    ]:
        col.markdown(f"""
        <div class="stat-card">
            <div class="val" style="color:{color}">{val}</div>
            <div class="lbl">{lbl}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    tab_reg, tab_param = st.tabs(["📋 Livrables", "⚙️ Paramétrage"])

    LIVRABLES = [
        {"#": 1,  "Livrable": "Etude d'opportunité",                 "Type": "Support décisionnel", "Responsable": "Chantal / Oliver",  "Etat": "Livré",         "Date livraison prévue": "9/4/2022",  "Validateur": "Comité de direction"},
        {"#": 2,  "Livrable": "Charte Projet",                        "Type": "Initiation",           "Responsable": "Chantal / Oliver",  "Etat": "Livré",         "Date livraison prévue": "29/4/2022", "Validateur": "Comité de direction"},
        {"#": 3,  "Livrable": "Macro planning",                       "Type": "Initiation",           "Responsable": "Chantal / Oliver",  "Etat": "Livré",         "Date livraison prévue": "29/4/2022", "Validateur": "Comité de direction"},
        {"#": 4,  "Livrable": "Support réunion de lancement",         "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "Livré",         "Date livraison prévue": "29/4/2022", "Validateur": "Sponsor"},
        {"#": 5,  "Livrable": "Expression des besoins",               "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "En cours",      "Date livraison prévue": "29/6/2022", "Validateur": "Sponsor"},
        {"#": 6,  "Livrable": "Adéquation exigences",                 "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "Retard",        "Date livraison prévue": "29/6/2022", "Validateur": "Sponsor"},
        {"#": 7,  "Livrable": "Cahier des charges",                   "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "Non commencé",  "Date livraison prévue": "30/6/2022", "Validateur": "Comité de direction"},
        {"#": 8,  "Livrable": "Plan de projet",                       "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "Non commencé",  "Date livraison prévue": "30/6/2022", "Validateur": "Comité de direction"},
        {"#": 9,  "Livrable": "Plan de charge",                       "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "Non commencé",  "Date livraison prévue": "29/6/2022", "Validateur": "Comité de direction"},
        {"#": 10, "Livrable": "Registre des problèmes et risques",    "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "En cours",      "Date livraison prévue": "29/6/2022", "Validateur": "Sponsor"},
        {"#": 11, "Livrable": "Registre des parties prenantes",       "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "En cours",      "Date livraison prévue": "29/6/2022", "Validateur": "Sponsor"},
        {"#": 12, "Livrable": "Matrice des compétences",              "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "En cours",      "Date livraison prévue": "29/6/2022", "Validateur": "Sponsor"},
        {"#": 13, "Livrable": "Plan de communication",                "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "En cours",      "Date livraison prévue": "29/6/2022", "Validateur": "Comité de direction"},
        {"#": 14, "Livrable": "Suivi des actions",                    "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "En cours",      "Date livraison prévue": "29/6/2022", "Validateur": "Sponsor"},
        {"#": 15, "Livrable": "Suivi des livrables",                  "Type": "Planification",        "Responsable": "Chantal / Oliver",  "Etat": "En cours",      "Date livraison prévue": "29/6/2022", "Validateur": "Sponsor"},
        {"#": 16, "Livrable": "Registre des demandes de changement",  "Type": "Réalisation",          "Responsable": "Chantal / Coordinateurs", "Etat": "Non commencé", "Date livraison prévue": "15/7/2022", "Validateur": "Sponsor / Communication"},
        {"#": 17, "Livrable": "Tableau de bord",                      "Type": "Rapports",             "Responsable": "Chantal",           "Etat": "Non commencé",  "Date livraison prévue": "15/7/2022", "Validateur": "Sponsor / Communication"},
        {"#": 18, "Livrable": "Rapport de suivi de projet",           "Type": "Rapports",             "Responsable": "Chantal",           "Etat": "Non commencé",  "Date livraison prévue": "15/7/2022", "Validateur": "Sponsor / Communication"},
        {"#": 19, "Livrable": "Rapport Flash",                        "Type": "Rapports",             "Responsable": "Chantal",           "Etat": "Non commencé",  "Date livraison prévue": "15/7/2022", "Validateur": "Sponsor / Communication"},
    ]

    df = pd.DataFrame(LIVRABLES)

    STATUS_COLOR = {
        "Livré":         "status-livré",
        "En cours":      "status-encours",
        "Retard":        "status-retard",
        "Non commencé":  "status-attente",
    }

    def _style(row):
        etat = str(row.get("Etat", ""))
        if etat == "Livré":        return ["background-color: #f0fdf4"]*len(row)
        if etat == "En cours":     return ["background-color: #eff6ff"]*len(row)
        if etat == "Retard":       return ["background-color: #fef2f2"]*len(row)
        return ["background-color: #f9fafb"]*len(row)

    with tab_reg:
        filter_etat = st.selectbox("Filtrer par état :", ["Tous", "Livré", "En cours", "Retard", "Non commencé"])
        df_filtered = df if filter_etat == "Tous" else df[df["Etat"] == filter_etat]
        st.dataframe(df_filtered.style.apply(_style, axis=1), use_container_width=True, hide_index=True)
        st.caption(f"{len(df_filtered)} livrable(s) affiché(s)")

    with tab_param:
        st.markdown('<div class="glass-header">⚙️ Paramétrage des types de livrables</div>', unsafe_allow_html=True)
        param_df = pd.DataFrame({
            "Type de livrable": ["Support décisionnel", "Initiation", "Planification", "Réalisation", "Rapports", "PoC", "Tests", "Déploiement"],
            "Couleur": ["#10b981", "#3b82f6", "#f59e0b", "#8b5cf6", "#64748b", "#ec4899", "#f97316", "#14b8a6"],
            "Actif": [True]*8
        })
        st.dataframe(param_df, use_container_width=True, hide_index=True)

# ─── VUE : AMDEC ─────────────────────────────────────────────────────────────
def render_amdec():
    render_project_header("26 — AMDEC", "Méthode AMDEC : Evaluation des modes de défaillance")

    tab_ech, tab_log, tab_mat, tab_rep = st.tabs([
        "📐 Échelles de pondération",
        "💻 Logiciel",
        "🛠️ Matériel",
        "📊 Répartition des tests"
    ])

    with tab_ech:
        st.markdown('<div class="glass-header">📐 Échelles de pondération AMDEC</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Gravité")
            st.dataframe(pd.DataFrame({
                "Niveau": [1,2,3,4,5],
                "Libellé": ["Négligeable","Mineur","Modéré","Majeur","Critique"],
                "Description": [
                    "Impact quasi nul sur l'activité",
                    "Gène légère, contournement possible",
                    "Impact significatif, ralentissement",
                    "Blocage partiel de l'activité",
                    "Blocage total, perte de données"
                ]
            }), use_container_width=True, hide_index=True)
        with col2:
            st.subheader("Fréquence / Détection")
            st.dataframe(pd.DataFrame({
                "Niveau": [1,2,3,4,5],
                "Fréquence": ["Très rare","Rare","Occasionnel","Fréquent","Très fréquent"],
                "Détection": ["Certain","Probable","Possible","Difficile","Quasi impossible"],
            }), use_container_width=True, hide_index=True)

    AMDEC_LOGICIEL = [
        {"Système": "Editer des factures", "Sous-système": "Domaine facturation", "Composant": "Service gestion des ventes", "Fonctions": "Charger les données produit et vente", "Mode de défaillance": "Service inaccessible", "Causes": "Service externe en maintenance", "Effets": "Consolidation des données impossible", "G": 4, "F": 3, "D": 3, "Criticité": 36, "Actions correctives": "Monitoring + test service + test d'intégration"},
        {"Système": "Editer des factures", "Sous-système": "Domaine facturation", "Composant": "Repository de données", "Fonctions": "Consolider les données", "Mode de défaillance": "Erreur dans la consolidation", "Causes": "Bug logiciel", "Effets": "Consolidation des données impossible", "G": 4, "F": 2, "D": 3, "Criticité": 24, "Actions correctives": "Test unitaire"},
        {"Système": "Editer des factures", "Sous-système": "Domaine facturation", "Composant": "Service de calcul des prix", "Fonctions": "Calculer les prix", "Mode de défaillance": "Erreurs dans le calcul des prix", "Causes": "Bug logiciel", "Effets": "Erreurs de facturation", "G": 4, "F": 2, "D": 1, "Criticité": 8, "Actions correctives": "Test unitaire"},
        {"Système": "Editer des factures", "Sous-système": "Domaine facturation", "Composant": "Templating", "Fonctions": "Fusionner les données dans le modèle", "Mode de défaillance": "Mauvais mapping modèle/données", "Causes": "Bug logiciel", "Effets": "Erreurs de facturation", "G": 3, "F": 2, "D": 2, "Criticité": 12, "Actions correctives": "Tests d'intégration"},
        {"Système": "Editer des factures", "Sous-système": "Domaine facturation", "Composant": "Templating", "Fonctions": "Fusionner les données dans le modèle", "Mode de défaillance": "Volumétrie trop grande pour le modèle", "Causes": "Bug logiciel", "Effets": "Problème de mise en page", "G": 2, "F": 3, "D": 1, "Criticité": 6, "Actions correctives": "Test de volumétrie + test service"},
        {"Système": "Editer des factures", "Sous-système": "Domaine facturation", "Composant": "Service d'impression PDF", "Fonctions": "Editer le PDF", "Mode de défaillance": "Moteur d'impression non opérationnel", "Causes": "Bug logiciel", "Effets": "Pas de sortie PDF", "G": 3, "F": 2, "D": 1, "Criticité": 6, "Actions correctives": "Test service"},
    ]

    def _color_criticite(val):
        if isinstance(val, (int, float)):
            if val >= 20: return "background-color: #fee2e2; color: #991b1b; font-weight: 800;"
            if val >= 10: return "background-color: #fff7ed; color: #9a3412; font-weight: 700;"
            return "background-color: #f0fdf4; color: #166534;"
        return ""

    with tab_log:
        st.markdown('<div class="glass-header">💻 AMDEC Logiciel</div>', unsafe_allow_html=True)
        df_log = pd.DataFrame(AMDEC_LOGICIEL)
        st.dataframe(
            df_log.style.applymap(_color_criticite, subset=["Criticité"]),
            use_container_width=True, hide_index=True
        )
        st.markdown("**Légende Criticité :** 🔴 ≥20 (Critique) | 🟠 10-19 (Elevé) | 🟢 <10 (Acceptable)")

    AMDEC_MAT = [
        {"Système": "Infrastructure", "Composant": "Serveur applicatif", "Mode": "Panne disque", "G": 5, "F": 2, "D": 2, "Criticité": 20, "Actions": "Redondance RAID + monitoring"},
        {"Système": "Infrastructure", "Composant": "Réseau LAN", "Mode": "Coupure réseau", "G": 4, "F": 2, "D": 3, "Criticité": 24, "Actions": "Lien de secours 4G + VPN"},
        {"Système": "Poste client", "Composant": "PC utilisateur", "Mode": "Crash OS", "G": 2, "F": 3, "D": 1, "Criticité": 6, "Actions": "Sauvegarde locale + helpdesk"},
        {"Système": "Infrastructure", "Composant": "Imprimante", "Mode": "Bourrage papier", "G": 1, "F": 4, "D": 1, "Criticité": 4, "Actions": "Maintenance préventive"},
    ]

    with tab_mat:
        st.markdown('<div class="glass-header">🛠️ AMDEC Matériel</div>', unsafe_allow_html=True)
        df_mat = pd.DataFrame(AMDEC_MAT)
        st.dataframe(
            df_mat.style.applymap(_color_criticite, subset=["Criticité"]),
            use_container_width=True, hide_index=True
        )

    with tab_rep:
        st.markdown('<div class="glass-header">📊 Répartition des Tests par Criticité</div>', unsafe_allow_html=True)
        all_data = AMDEC_LOGICIEL + AMDEC_MAT
        crits = [r.get("Criticité", 0) for r in all_data]
        labels = ["Critique (≥20)", "Elevé (10-19)", "Acceptable (<10)"]
        values = [sum(1 for c in crits if c >= 20), sum(1 for c in crits if 10 <= c < 20), sum(1 for c in crits if c < 10)]
        fig = go.Figure(go.Pie(
            labels=labels, values=values,
            hole=0.5,
            marker_colors=["#ef4444", "#f97316", "#10b981"],
            textfont_size=14
        ))
        fig.update_layout(height=350, margin=dict(l=20,r=20,t=30,b=20))
        st.plotly_chart(fig, use_container_width=True)

# ─── VUE : PLAN DE TEST ──────────────────────────────────────────────────────
def render_plan_test():
    render_project_header("Plan de Test", "Suivi des tests d'intégration et de consolidation")

    STATS = {"Conforme": 173, "Non conforme": 8, "Anomalie": 1, "En attente": 0, "Pas applicable": 3}
    total = sum(STATS.values())

    col_gauge, col_stats = st.columns([1, 2])
    with col_gauge:
        pct = round(STATS["Conforme"] / total * 100)
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=pct,
            title={"text": "Taux de Conformité (%)"},
            number={"suffix": "%"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": "#10b981"},
                "steps": [
                    {"range": [0, 60], "color": "#fee2e2"},
                    {"range": [60, 85], "color": "#fff7ed"},
                    {"range": [85, 100], "color": "#f0fdf4"},
                ],
            }
        ))
        fig.update_layout(height=280, margin=dict(l=20,r=20,t=40,b=20))
        st.plotly_chart(fig, use_container_width=True)

    with col_stats:
        st.markdown('<div class="glass-header">📊 Indicateurs Tests</div>', unsafe_allow_html=True)
        cols = st.columns(len(STATS))
        colors = {"Conforme": "#10b981", "Non conforme": "#ef4444", "Anomalie": "#f97316", "En attente": "#64748b", "Pas applicable": "#a855f7"}
        for col, (k, v) in zip(cols, STATS.items()):
            col.markdown(f"""
            <div class="stat-card">
                <div class="val" style="color:{colors[k]}">{v}</div>
                <div class="lbl">{k}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("---")
    tab_int, tab_cons = st.tabs(["🔄 Tests d'Intégration", "🏗️ Tests de Consolidation"])

    TESTS_DATA = {
        "Domaine": ["FINANCES"]*6 + ["VENTES"]*4 + ["ACHATS"]*3 + ["LOGISTIQUE"]*2,
        "Général/Spécifique": ["Général"]*3 + ["Immobilisations"]*3 + ["Général"]*4 + ["Général"]*3 + ["Général"]*2,
        "Scénario de test": [
            "Facturation inter-compagnies", "Clôture", "Rapprochement des comptes",
            "Saisir mouvement d'immobilisation", "Amortissements", "Annuler poste immo",
            "Ventes standards", "Unités organisationnelles", "Prise en commande", "Facturation",
            "Gestion d'approvisionnement", "Achats avec contrats", "Sous-contractants",
            "Gestion des expéditions", "Suivi de transport",
        ],
        "Résultat attendu": ["X"]*15,
        "Conforme": ["X","","X","X","X","","X","X","","X","X","","X","X",""] ,
        "Non conforme": ["","X","","","","X","","","X","","","X","","","X"],
        "Statut": ["✅","❌","✅","✅","✅","❌","✅","✅","❌","✅","✅","❌","✅","✅","❌"],
    }
    df_tests = pd.DataFrame(TESTS_DATA)

    def _style_tests(row):
        if row["Statut"] == "✅": return ["background-color: #f0fdf4"]*len(row)
        if row["Statut"] == "❌": return ["background-color: #fef2f2"]*len(row)
        return [""]*len(row)

    with tab_int:
        search = st.text_input("🔍 Rechercher un scénario :", key="search_int")
        df_f = df_tests[df_tests["Scénario de test"].str.contains(search, case=False)] if search else df_tests
        st.dataframe(df_f.style.apply(_style_tests, axis=1), use_container_width=True, hide_index=True)

    with tab_cons:
        st.info("Les tests de consolidation valident l'ensemble des scénarios après intégration complète.")
        df_cons = df_tests.copy()
        df_cons["Statut"] = ["✅"]*12 + ["🔵"]*3
        def _style_cons(row):
            if row["Statut"] == "✅": return ["background-color: #f0fdf4"]*len(row)
            return ["background-color: #eff6ff"]*len(row)
        st.dataframe(df_cons.style.apply(_style_cons, axis=1), use_container_width=True, hide_index=True)

# ─── VUE : REGISTRE DES CHANGEMENTS ─────────────────────────────────────────
def render_changements():
    render_project_header("Registre des Demandes de Changement", "Gestion des évolutions du périmètre projet")

    tab_reg, tab_param = st.tabs(["📋 Demandes de changement", "⚙️ Paramétrage"])

    CHANGEMENTS = [
        {"#": 1, "Date": "15/07/2022", "Demandeur": "Dupont Chantal", "Description": "Ajout module gestion des avances",              "Impact": "Moyen",  "Priorité": "Haute",  "Statut": "Approuvé",    "Décision": "Intégré dans la release 2"},
        {"#": 2, "Date": "22/07/2022", "Demandeur": "Eric Duteil",    "Description": "Modification des droits d'accès logistique",    "Impact": "Faible", "Priorité": "Normale", "Statut": "Approuvé",    "Décision": "Appliqué en environnement test"},
        {"#": 3, "Date": "05/08/2022", "Demandeur": "Monique Davant", "Description": "Ajout d'un rapport de rapprochement bancaire",  "Impact": "Faible", "Priorité": "Basse",   "Statut": "En attente",  "Décision": "Analyse en cours"},
        {"#": 4, "Date": "15/08/2022", "Demandeur": "Jean Albert",    "Description": "Changement de la politique de taxe TVA",        "Impact": "Elevé",  "Priorité": "Haute",   "Statut": "Refusé",      "Décision": "Hors périmètre"},
        {"#": 5, "Date": "01/09/2022", "Demandeur": "Denise Attali",  "Description": "Optimisation du flux de validation commandes",  "Impact": "Moyen",  "Priorité": "Normale", "Statut": "En analyse",  "Décision": "Chiffrage en cours"},
    ]

    df = pd.DataFrame(CHANGEMENTS)

    def _style_chg(row):
        if row["Statut"] == "Approuvé":  return ["background-color: #f0fdf4"]*len(row)
        if row["Statut"] == "Refusé":    return ["background-color: #fef2f2"]*len(row)
        if row["Statut"] == "En attente": return ["background-color: #eff6ff"]*len(row)
        return ["background-color: #fefce8"]*len(row)

    with tab_reg:
        col1, col2 = st.columns([3, 1])
        with col2:
            filtre = st.selectbox("Filtrer statut :", ["Tous", "Approuvé", "En attente", "En analyse", "Refusé"])
        df_f = df if filtre == "Tous" else df[df["Statut"] == filtre]
        st.dataframe(df_f.style.apply(_style_chg, axis=1), use_container_width=True, hide_index=True)
        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Total demandes", len(df))
        col_b.metric("Approuvées", len(df[df["Statut"] == "Approuvé"]))
        col_c.metric("En attente", len(df[df["Statut"].isin(["En attente", "En analyse"])]))

    with tab_param:
        st.markdown('<div class="glass-header">⚙️ Statuts disponibles</div>', unsafe_allow_html=True)
        st.dataframe(pd.DataFrame({
            "Statut": ["En attente", "En analyse", "Approuvé", "Refusé"],
            "Couleur assignée": ["Bleu", "Jaune", "Vert", "Rouge"],
            "Description": [
                "Demande reçue, non encore traitée",
                "En cours d'évaluation d'impact",
                "Validée par le comité projet",
                "Rejetée, hors périmètre ou trop risqué"
            ]
        }), use_container_width=True, hide_index=True)

# ─── VUE : SUIVI DES ACTIONS ─────────────────────────────────────────────────
def render_suivi_actions():
    render_project_header("Suivi des Actions", "Liste des actions opérationnelles et leur avancement")

    tab_act, tab_param = st.tabs(["✅ Actions", "⚙️ Paramétrage"])

    ACTIONS = [
        {"#": 1,  "Date": "01/07/2022", "Action": "Organiser réunion de lancement",                   "Responsable": "Chantal Dupont",  "Priorité": "Haute",   "Statut": "Terminé",     "% Avt": 100, "Échéance": "15/07/2022"},
        {"#": 2,  "Date": "15/07/2022", "Action": "Valider le cahier des charges avec le sponsor",    "Responsable": "Olivier Martin",  "Priorité": "Haute",   "Statut": "Terminé",     "% Avt": 100, "Échéance": "30/07/2022"},
        {"#": 3,  "Date": "01/08/2022", "Action": "Finaliser AMDEC logiciel",                        "Responsable": "Eric Duteil",     "Priorité": "Haute",   "Statut": "En cours",    "% Avt": 75,  "Échéance": "30/08/2022"},
        {"#": 4,  "Date": "15/08/2022", "Action": "Organiser session de formation utilisateurs clés", "Responsable": "Monique Davant",  "Priorité": "Normale", "Statut": "En cours",    "% Avt": 40,  "Échéance": "15/09/2022"},
        {"#": 5,  "Date": "01/09/2022", "Action": "Exécuter tests d'intégration module Finance",      "Responsable": "Denise Attali",   "Priorité": "Haute",   "Statut": "En cours",    "% Avt": 60,  "Échéance": "30/09/2022"},
        {"#": 6,  "Date": "15/09/2022", "Action": "Exécuter tests d'intégration module Ventes",      "Responsable": "Jean Albert",     "Priorité": "Haute",   "Statut": "À faire",     "% Avt": 0,   "Échéance": "31/10/2022"},
        {"#": 7,  "Date": "01/10/2022", "Action": "Backup complet de la base de données prod",       "Responsable": "Daniel Tech",     "Priorité": "Haute",   "Statut": "À faire",     "% Avt": 0,   "Échéance": "15/10/2022"},
        {"#": 8,  "Date": "01/10/2022", "Action": "Rédiger rapport de suivi v2",                     "Responsable": "Chantal Dupont",  "Priorité": "Normale", "Statut": "À faire",     "% Avt": 0,   "Échéance": "31/10/2022"},
    ]

    df = pd.DataFrame(ACTIONS)

    def _style_act(row):
        if row["Statut"] == "Terminé":  return ["background-color: #f0fdf4"]*len(row)
        if row["Statut"] == "En cours": return ["background-color: #eff6ff"]*len(row)
        if row["Statut"] == "Bloqué":   return ["background-color: #fef2f2"]*len(row)
        return ["background-color: #fefce8"]*len(row)

    with tab_act:
        col1, col2, col3 = st.columns(3)
        col1.metric("Total", len(df))
        col2.metric("Terminées ✅", len(df[df["Statut"] == "Terminé"]))
        col3.metric("En cours 🔵", len(df[df["Statut"] == "En cours"]))

        filtre = st.selectbox("Filtrer statut :", ["Tous", "Terminé", "En cours", "À faire", "Bloqué"], key="act_filter")
        df_f = df if filtre == "Tous" else df[df["Statut"] == filtre]
        st.dataframe(df_f.style.apply(_style_act, axis=1), use_container_width=True, hide_index=True)

        # Graphique barres de progression
        st.markdown('<div class="glass-header">📊 Avancement des actions</div>', unsafe_allow_html=True)
        fig = px.bar(df, x="Action", y="% Avt",
            color="% Avt",
            color_continuous_scale=["#fef2f2","#fff7ed","#f0fdf4"],
            labels={"% Avt": "Avancement (%)", "Action": ""},
            height=350,
        )
        fig.update_layout(xaxis_tickangle=-30, showlegend=False, margin=dict(l=20,r=20,t=20,b=100))
        st.plotly_chart(fig, use_container_width=True)

    with tab_param:
        st.markdown('<div class="glass-header">⚙️ Configuration des priorités et statuts</div>', unsafe_allow_html=True)
        st.dataframe(pd.DataFrame({
            "Statut": ["Terminé", "En cours", "À faire", "Bloqué"],
            "Icône": ["✅", "🔵", "⬜", "🔴"],
            "Priorité": ["Haute", "Normale", "Basse", "—"],
        }), use_container_width=True, hide_index=True)

# ─── SIDEBAR + ROUTING ───────────────────────────────────────────────────────
def main():
    st.markdown('<h1 class="main-title">FAAB PROJET 📋</h1>', unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("## 🗂️ Navigation")
        st.markdown("---")
        section = st.radio("", [
            "🏠 Accueil",
            "── PLANIFICATION ──",
            "── EXÉCUTION ──",
            "── KIT CHEF PROJET ──",
        ], label_visibility="collapsed")
        st.markdown("---")

    # ── EXÉCUTION section ─────────────────────────────────────────────────────
    if section == "── EXÉCUTION ──":
        with st.sidebar:
            exec_module = st.radio("Module Exécution :", [
                "📋 Suivi des Livrables",
                "🛡️ AMDEC",
                "🧪 Plan de Tests",
                "🔄 Registre Changements",
                "✅ Suivi des Actions",
            ])

        if   exec_module == "📋 Suivi des Livrables":  render_suivi_livrables()
        elif exec_module == "🛡️ AMDEC":                render_amdec()
        elif exec_module == "🧪 Plan de Tests":         render_plan_test()
        elif exec_module == "🔄 Registre Changements":  render_changements()
        elif exec_module == "✅ Suivi des Actions":      render_suivi_actions()

    elif section == "── KIT CHEF PROJET ──":
        kit_files = scan_kit(KIT_DIR)
        with st.sidebar:
            cats = sorted(set(f["category"] for f in kit_files))
            active = st.session_state.get("active")
            for c in cats:
                with st.expander(c, expanded=bool(active and active["category"] == c)):
                    for f in [x for x in kit_files if x["category"] == c]:
                        if st.button(f"📄 {f['name']}", key=str(f["path"])):
                            st.session_state["active"] = f
                            st.rerun()

        active = st.session_state.get("active")
        if active:
            try:
                xl = pd.ExcelFile(active["path"])
                sheets = xl.sheet_names
                render_project_header(active["name"], str(active["path"]))
                tabs = st.tabs([f"📄 {s}" for s in sheets])
                for t, s in zip(tabs, sheets):
                    with t:
                        df = xl.parse(s)
                        if df.empty:
                            st.info("Feuille vide.")
                            continue
                        fn = active["name"].lower()
                        sn = s.lower()
                        if "raci" in fn or "raci" in sn:
                            def _sr(v):
                                v = str(v).upper().strip()
                                if v == "R": return "background-color:#dcfce7;color:#166534;font-weight:800;text-align:center"
                                if v == "A": return "background-color:#dbeafe;color:#1e40af;font-weight:800;text-align:center"
                                if v == "C": return "background-color:#ffedd5;color:#9a3412;font-weight:800;text-align:center"
                                if v == "I": return "background-color:#fef9c3;color:#854d0e;font-weight:800;text-align:center"
                                return ""
                            st.dataframe(df.style.applymap(_sr), use_container_width=True, hide_index=True)
                        else:
                            st.dataframe(df, use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"Erreur : {e}")
        else:
            st.info("Sélectionnez un fichier dans la barre latérale.")

    else:
        # Accueil
        st.markdown("""
        <div style="text-align:center; padding: 80px 20px; color:#64748b;">
            <div style="font-size:5rem; margin-bottom:20px;">📋</div>
            <h2 style="color:#065f46;">Système de Gestion FAAB PROJET</h2>
            <p>Bienvenue. Sélectionnez un module dans le menu de navigation.</p>
            <div style="display:flex; justify-content:center; gap:16px; margin-top:2rem; flex-wrap:wrap;">
                <div style="background:#f0fdf4;border:1px solid #10b981;border-radius:10px;padding:20px;width:160px;">
                    <div style="font-size:2rem;">📋</div><b>Exécution</b>
                    <p style="font-size:0.75rem;color:#64748b;">Suivi, Tests, AMDEC</p>
                </div>
                <div style="background:#eff6ff;border:1px solid #3b82f6;border-radius:10px;padding:20px;width:160px;">
                    <div style="font-size:2rem;">🗂️</div><b>Kit Chef Projet</b>
                    <p style="font-size:0.75rem;color:#64748b;">Vos fichiers Excel</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.sidebar.markdown("---")
    st.sidebar.caption("FAAB PROJET v3.0 | Développé avec ❤️")

if __name__ == "__main__":
    main()
