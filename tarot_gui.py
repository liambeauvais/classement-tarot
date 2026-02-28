import os
import sys
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, Tuple, List

# Gestion des chemins pour PyInstaller
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Import local après la définition de resource_path
try:
    from tarot_rankings import run
except ImportError as e:
    print(f"Erreur d'importation: {e}")
    run = None

class TarotRankingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Classement Tarot")
        self.root.geometry("800x600")
        
        # Variables
        self.excel_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.export_pdf = tk.BooleanVar(value=True)
        self.export_csv = tk.BooleanVar(value=True)
        self.day_var = tk.StringVar(value="Après-midi")
        self.month_var = tk.StringVar(value="")
        self.error_detection = tk.BooleanVar(value=False)
        
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Sélection du fichier Excel
        ttk.Label(main_frame, text="Fichier Excel:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Parcourir...", command=self.browse_excel).grid(row=0, column=2, padx=5, pady=5)
        
        # Sélection du répertoire de sortie
        ttk.Label(main_frame, text="Dossier de sortie:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Parcourir...", command=self.browse_output_dir).grid(row=1, column=2, padx=5, pady=5)
        
        # Options d'exportation
        ttk.Label(main_frame, text="Options d'exportation:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Checkbutton(main_frame, text="Exporter en PDF", variable=self.export_pdf).grid(row=2, column=1, sticky=tk.W, pady=5)
        ttk.Checkbutton(main_frame, text="Exporter en CSV", variable=self.export_csv).grid(row=3, column=1, sticky=tk.W, pady=5)
        ttk.Checkbutton(main_frame, text="Détection d'erreur", variable=self.error_detection).grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # Sélection de la période
        ttk.Label(main_frame, text="Période du tournoi:").grid(row=5, column=0, sticky=tk.W, pady=5)
        day_frame = ttk.Frame(main_frame)
        day_frame.grid(row=5, column=1, sticky=tk.W, pady=5)
        ttk.Radiobutton(day_frame, text="Après-midi", variable=self.day_var, value="Après-midi").pack(side=tk.LEFT)
        ttk.Radiobutton(day_frame, text="Soir", variable=self.day_var, value="Soir").pack(side=tk.LEFT, padx=10)
        
        # Sélection du mois
        ttk.Label(main_frame, text="Mois (optionnel):").grid(row=6, column=0, sticky=tk.W, pady=5)
        month_frame = ttk.Frame(main_frame)
        month_frame.grid(row=6, column=1, sticky=tk.W, pady=5)
        months = ["", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        month_combo = ttk.Combobox(month_frame, textvariable=self.month_var, values=months, width=15, state="readonly")
        month_combo.pack(side=tk.LEFT)
        month_combo.set("")
        
        # Bouton de génération
        ttk.Button(main_frame, text="Générer le classement", command=self.generate_ranking).grid(row=7, column=0, columnspan=3, pady=20)
        
        # Zone de statut
        self.status_var = tk.StringVar()
        self.status_var.set("Prêt")
        ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).grid(
            row=8, column=0, columnspan=3, sticky=tk.EW, pady=10)
        
        # Zone d'aperçu
        ttk.Label(main_frame, text="Aperçu du classement:").grid(row=9, column=0, sticky=tk.NW, pady=5)
        self.preview_text = tk.Text(main_frame, height=15, width=70, state=tk.DISABLED)
        self.preview_text.grid(row=10, column=0, columnspan=3, sticky=tk.NSEW, pady=5)
        
        # Configuration du redimensionnement
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(10, weight=1)
    
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="Sélectionner le fichier Excel",
            filetypes=(("Fichiers Excel", "*.xlsx *.xlsm"), ("Tous les fichiers", "*.*"))
        )
        if filename:
            self.excel_path.set(filename)
            # Définir le dossier de sortie par défaut
            if not self.output_dir.get():
                self.output_dir.set(os.path.dirname(filename) or ".")
    
    def browse_output_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir.set(directory)
    
    def update_status(self, message: str):
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def generate_ranking(self):
        excel_path = self.excel_path.get()
        output_dir = self.output_dir.get()
        
        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel valide")
            return
        
        if not output_dir:
            messagebox.showerror("Erreur", "Veuillez sélectionner un dossier de sortie")
            return
        
        try:
            self.update_status("Génération en cours...")
            self.root.update()  # Force la mise à jour de l'interface
            
            # Vérifier que le répertoire de sortie existe
            os.makedirs(output_dir, exist_ok=True)
            
            # Vérifier les autorisations d'écriture
            test_file = os.path.join(output_dir, "test_write.tmp")
            try:
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'écrire dans le dossier de sortie: {e}")
                self.update_status("Erreur d'écriture")
                return
            
            # Appel de la fonction principale du script existant
            if run is None:
                messagebox.showerror("Erreur", "Le module tarot_rankings n'a pas pu être chargé")
                return
                
            run(
                excel_path=os.path.abspath(excel_path),
                out_dir=os.path.abspath(output_dir),
                want_pdf=self.export_pdf.get(),
                want_csv=self.export_csv.get(),
                day=self.day_var.get(),
                month=self.month_var.get(),
                error_detection=self.error_detection.get()
            )
            
            # Mettre à jour le titre du PDF
            title = f"Challenge {self.day_var.get()}"
            
            # Mettre à jour l'aperçu
            self.update_preview(output_dir)
            
            # Ouvrir le dossier de sortie
            try:
                os.startfile(os.path.abspath(output_dir))
            except:
                pass
                
            messagebox.showinfo("Succès", "Le classement a été généré avec succès!")
            self.update_status("Terminé")
            
        except Exception as e:
            error_msg = f"Une erreur est survenue: {str(e)}\n\nDétails techniques:\n{traceback.format_exc()}"
            messagebox.showerror("Erreur", error_msg)
            self.update_status("Erreur")
    
    def update_preview(self, output_dir: str):
        # Chercher le dernier fichier CSV généré
        csv_files = [f for f in os.listdir(output_dir) if f.endswith('.csv')]
        if not csv_files:
            return
            
        latest_csv = max(csv_files, key=lambda f: os.path.getmtime(os.path.join(output_dir, f)))
        csv_path = os.path.join(output_dir, latest_csv)
        
        # Lire les premières lignes du CSV pour l'aperçu
        preview_lines = []
        try:
            with open(csv_path, 'r', encoding='utf-8') as f:
                preview_lines = [line.strip() for i, line in enumerate(f) if i < 11]  # 10 premières lignes max
        except Exception:
            return
        
        # Mettre à jour la zone d'aperçu
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.insert(tk.END, "\n".join(preview_lines))
        self.preview_text.config(state=tk.DISABLED)

def main():
    root = tk.Tk()
    app = TarotRankingApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
