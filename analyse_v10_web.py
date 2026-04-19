import streamlit as st
import pandas as pd
import re
from datetime import datetime

# Configuration de la page
st.set_page_config(page_title="Analyse V10", layout="wide")

st.title("Analyse V10")

# --- BARRE LATÉRALE D'IMPORT ---
st.sidebar.header("1. Chargement des fichiers")
file_v10 = st.sidebar.file_uploader("Historique V10 (CSV)", type=["csv"])
file_plume = st.sidebar.file_uploader("Suivi Plume (Excel/CSV)", type=["xlsx", "csv"])

if st.sidebar.button("Lancer l'Analyse Globale"):
    if file_v10 is not None:
        try:
            # Lecture V10
            df_v10 = pd.read_csv(file_v10, sep=';', encoding='latin-1')
            df_v10.columns = [c.strip() for c in df_v10.columns]
            df_v10['dt'] = pd.to_datetime(df_v10['Date de création'] + ' ' + df_v10['Heure de création'], dayfirst=True)
            df_v10 = df_v10.sort_values('dt')

            states = {}
            inc_pat = r'(INC\d+)'
            maintenant = datetime.now()
            
            for _, row in df_v10.iterrows():
                site = str(row['Produit'])
                comm = str(row.get('Commentaire', ''))
                ack = str(row.get("Heure d'acquittement", ''))
                text = f"{comm} {ack}"
                
                is_m_exit = "Sortir de maintenance" in text
                found_inc = re.search(inc_pat, text, re.IGNORECASE)
                is_m_entry = ("Mettre en maintenance" in text) or (found_inc is not None)

                is_t_exit = "Sortir de travaux" in text
                is_t_entry = ("Mettre en travaux" in text) or ("En Travaux" in text)

                if site not in states:
                    states[site] = {'maint': False, 'travaux': False, 'inc': None, 'reason': '', 'date_trav': None}
                
                if is_m_exit: 
                    states[site]['maint'] = False
                elif is_m_entry:
                    states[site]['maint'] = True
                    if found_inc: states[site]['inc'] = found_inc.group(1).upper()

                if is_t_exit: 
                    states[site]['travaux'] = False
                    states[site]['date_trav'] = None
                elif is_t_entry:
                    states[site]['travaux'] = True
                    states[site]['reason'] = ack if "travaux" in ack.lower() else comm
                    if states[site]['date_trav'] is None:
                        states[site]['date_trav'] = row['dt']

            # --- TRAITEMENT MAINTENANCE ---
            df_final_maint = pd.DataFrame()
            if file_plume:
                df_plume = pd.read_excel(file_plume) if file_plume.name.endswith('.xlsx') else pd.read_csv(file_plume)
                df_plume.columns = [c.strip() for c in df_plume.columns]
                maint_list = [{'Site': s, 'INC_V10': v['inc']} for s, v in states.items() if v['maint'] and v['inc']]
                if maint_list:
                    df_m = pd.DataFrame(maint_list)
                    col_inc_p = 'Numéro' if 'Numéro' in df_plume.columns else df_plume.columns[0]
                    merged = pd.merge(df_m, df_plume, left_on='INC_V10', right_on=col_inc_p, how='inner')
                    df_final_maint = merged[merged['État'].isin(['Résolu', 'Fermé'])].copy()
                    df_final_maint['Statut Prynvision'] = "En maintenance"
                    df_final_maint = df_final_maint[['Site', 'INC_V10', 'État', 'Statut Prynvision', 'Affecté à']]
                    df_final_maint.columns = ['Code et Nom du Site', 'N° INC', 'Statut Plume', 'Statut Prynvision', 'Affecté à']

            # --- TRAITEMENT TRAVAUX ---
            travaux_rows = []
            for s, v in states.items():
                if v['travaux']:
                    diff = (maintenant - v['date_trav']).days
                    travaux_rows.append({
                        'Code et Nom du Site': s,
                        'Date Mise en Travaux': v['date_trav'],
                        'Statut Prynvision': 'En Travaux',
                        'Raison (Commentaire V10)': v['reason'],
                        'jours': diff
                    })
            df_final_trav = pd.DataFrame(travaux_rows)

            # --- AFFICHAGE DES RÉSULTATS ---
            st.success(f"Analyse terminée ! {len(df_final_maint)} maintenance(s) à clore et {len(df_final_trav)} site(s) en travaux.")
            
            tab1, tab2 = st.tabs(["Sites en Maintenance", "Sites en Travaux"])

            with tab1:
                if not df_final_maint.empty:
                    st.dataframe(df_final_maint, use_container_width=True)
                    csv_m = df_final_maint.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
                    st.download_button("Export Maintenance", csv_m, "export_maint.csv", "text/csv")
                else:
                    st.info("Aucune anomalie de maintenance.")

            with tab2:
                if not df_final_trav.empty:
                    # Fonction pour colorer en rouge si > 20 jours
                    def color_red(row):
                        return ['color: red' if row.jours > 20 else '' for _ in row]
                    
                    st.dataframe(df_final_trav.style.apply(color_red, axis=1), use_container_width=True)
                    csv_t = df_final_trav.drop(columns=['jours']).to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
                    st.download_button("Export Travaux", csv_t, "export_travaux.csv", "text/csv")
                else:
                    st.info("Aucun site en travaux.")

        except Exception as e:
            st.error(f"Erreur : {e}")
    else:
        st.warning("Veuillez charger le fichier V10.")