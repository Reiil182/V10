import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

class AnalyseV10App:
    def __init__(self, root):
        self.root = root
        self.root.title("Analyse V10")
        self.root.geometry("1150x750")
        self.root.configure(padx=20, pady=20)

        self.path_v10 = None
        self.path_plume = None
        self.df_anomalies = pd.DataFrame()
        self.df_travaux = pd.DataFrame()

        # --- SECTION IMPORT (HAUT) ---
        tk.Label(root, text="Analyse V10", font=("Arial", 18, "bold")).pack(pady=10)
        
        frame_btns = tk.Frame(root)
        frame_btns.pack(pady=10)

        self.btn_v10 = tk.Button(frame_btns, text="1. Importer Historique V10 (CSV)", command=self.load_v10, width=40)
        self.btn_v10.grid(row=0, column=0, padx=10, pady=5)

        self.btn_plume = tk.Button(frame_btns, text="2. Importer Historique Plume (Excel/CSV)", command=self.load_plume, width=40)
        self.btn_plume.grid(row=0, column=1, padx=10, pady=5)

        self.btn_analyze = tk.Button(root, text="Lancer l'Analyse Globale", bg="#2ecc71", fg="white", 
                                     font=("Arial", 11, "bold"), command=self.process_all, width=35, height=2)
        self.btn_analyze.pack(pady=20)

        # --- SECTION ONGLETS ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10)

        # Onglet 1 : Maintenance
        self.tab_maint = tk.Frame(self.notebook)
        self.notebook.add(self.tab_maint, text=" Sites en Maintenance ")
        self.setup_tab_maintenance()

        # Onglet 2 : Travaux
        self.tab_travaux = tk.Frame(self.notebook)
        self.notebook.add(self.tab_travaux, text=" Sites en Travaux ")
        self.setup_tab_travaux()

    def setup_tab_maintenance(self):
        self.lbl_maint = tk.Label(self.tab_maint, text="Anomalies : Maintenance V10 active vs Plume Résolu/Fermé", font=("Arial", 10, "italic"))
        self.lbl_maint.pack(pady=10)
        
        cols = ('Code et Nom du Site', 'N° INC', 'Statut Plume', 'Statut Prynvision', 'Affecté à')
        self.tree_maint = self.create_tree(self.tab_maint, cols)
        
        self.btn_exp_maint = tk.Button(self.tab_maint, text="Export", bg="#3498db", fg="white", 
                                       font=("Arial", 10, "bold"), command=self.export_maint, width=20)
        self.btn_exp_maint.pack(pady=15)
        self.btn_exp_maint.config(state="disabled")

    def setup_tab_travaux(self):
        self.lbl_trav = tk.Label(self.tab_travaux, text="Sites actuellement en travaux (Alerte rouge si > 20 jours)", font=("Arial", 10, "italic"))
        self.lbl_trav.pack(pady=10)
        
        cols = ('Code et Nom du Site', 'Date Mise en Travaux', 'Statut Prynvision', 'Raison (Commentaire V10)')
        self.tree_trav = self.create_tree(self.tab_travaux, cols)
        
        # Configuration de la couleur rouge pour les délais dépassés
        self.tree_trav.tag_configure('alerte', foreground='red')
        
        self.btn_exp_trav = tk.Button(self.tab_travaux, text="Export", bg="#3498db", fg="white", 
                                      font=("Arial", 10, "bold"), command=self.export_trav, width=20)
        self.btn_exp_trav.pack(pady=15)
        self.btn_exp_trav.config(state="disabled")

    def create_tree(self, parent, cols):
        tree_frame = tk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        
        tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=12)
        for c in cols:
            tree.heading(c, text=c, command=lambda _c=c, _t=tree: self.sort_column(_t, _c, False))
            tree.column(c, width=350 if "Raison" in c else 180, anchor="center")
            
        scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        return tree

    def sort_column(self, tree, col, reverse):
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.sort_column(tree, col, not reverse))

    def load_v10(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path:
            self.path_v10 = path
            self.btn_v10.config(text="V10 : " + path.split('/')[-1], fg="#27ae60")

    def load_plume(self):
        path = filedialog.askopenfilename(filetypes=[("Excel or CSV", "*.xlsx *.csv")])
        if path:
            self.path_plume = path
            self.btn_plume.config(text="Plume : " + path.split('/')[-1], fg="#27ae60")

    def process_all(self):
        if not self.path_v10:
            messagebox.showwarning("Attention", "Le fichier Historique V10 est obligatoire.")
            return
        
        try:
            for t in [self.tree_maint, self.tree_trav]:
                for item in t.get_children(): t.delete(item)

            df_v10 = pd.read_csv(self.path_v10, sep=';', encoding='latin-1')
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
                
                # Logique Maintenance
                is_m_exit = "Sortir de maintenance" in text
                found_inc = re.search(inc_pat, text, re.IGNORECASE)
                is_m_entry = ("Mettre en maintenance" in text) or (found_inc is not None)

                # Logique Travaux
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
                    # On garde la date de la première mise en travaux détectée
                    if states[site]['date_trav'] is None:
                        states[site]['date_trav'] = row['dt']

            # --- ONGLET MAINTENANCE ---
            count_maint = 0
            if self.path_plume:
                df_plume = pd.read_excel(self.path_plume) if self.path_plume.endswith('.xlsx') else pd.read_csv(self.path_plume)
                df_plume.columns = [c.strip() for c in df_plume.columns]
                maint_list = [{'Site': s, 'INC_V10': v['inc']} for s, v in states.items() if v['maint'] and v['inc']]
                if maint_list:
                    df_m = pd.DataFrame(maint_list)
                    col_inc_p = 'Numéro' if 'Numéro' in df_plume.columns else df_plume.columns[0]
                    merged = pd.merge(df_m, df_plume, left_on='INC_V10', right_on=col_inc_p, how='inner')
                    anom = merged[merged['État'].isin(['Résolu', 'Fermé'])].copy()
                    if not anom.empty:
                        anom['Statut Prynvision'] = "En maintenance"
                        self.df_anomalies = anom[['Site', 'INC_V10', 'État', 'Statut Prynvision', 'Affecté à']]
                        self.df_anomalies.columns = ['Code et Nom du Site', 'N° INC', 'Statut Plume', 'Statut Prynvision', 'Affecté à']
                        for _, r in self.df_anomalies.iterrows(): self.tree_maint.insert("", tk.END, values=list(r))
                        count_maint = len(self.df_anomalies)
                        self.btn_exp_maint.config(state="normal")

            # --- ONGLET TRAVAUX ---
            count_trav = 0
            travaux_rows = []
            for s, v in states.items():
                if v['travaux']:
                    diff_jours = (maintenant - v['date_trav']).days
                    tag = 'alerte' if diff_jours > 20 else ''
                    date_str = v['date_trav'].strftime('%d/%m/%Y')
                    travaux_rows.append((s, date_str, 'En Travaux', v['reason'], tag))
            
            if travaux_rows:
                self.df_travaux = pd.DataFrame([r[:-1] for r in travaux_rows], 
                                              columns=['Code et Nom du Site', 'Date Mise en Travaux', 'Statut Prynvision', 'Raison (Commentaire V10)'])
                for r in travaux_rows:
                    self.tree_trav.insert("", tk.END, values=r[:-1], tags=(r[-1],))
                count_trav = len(self.df_travaux)
                self.btn_exp_trav.config(state="normal")

            # --- POP-UP RÉCAPITULATIF ---
            messagebox.showinfo("Résultat de l'analyse", 
                                f"Analyse terminée !\n\n"
                                f"• Sites en Maintenance à clore : {count_maint}\n"
                                f"• Sites en Travaux en cours : {count_trav}")

        except Exception as e:
            messagebox.showerror("Erreur", f"Détails : {str(e)}")

    def export_maint(self):
        self.save_csv(self.df_anomalies, "Export_Maintenance")

    def export_trav(self):
        self.save_csv(self.df_travaux, "Export_Travaux")

    def save_csv(self, df, default_name):
        if df.empty: return
        path = filedialog.asksaveasfilename(initialfile=default_name, defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            df.to_csv(path, index=False, sep=';', encoding='utf-8-sig')
            messagebox.showinfo("Succès", "Export terminé.")

if __name__ == "__main__":
    root = tk.Tk()
    app = AnalyseV10App(root)
    root.mainloop()