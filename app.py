import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import font
import threading as th
import pandas as pd
from docx import Document
import time
import os
import re

import utils


# FENÊTRE DE DÉMARRAGE
class Demarrage(tk.Toplevel):
    def __init__(self, master, on_finish):
        super().__init__(master)
        self.geometry("400x100+568+322")
        self.resizable(False, False)
        self.configure(background="white")
        self.on_finish = on_finish

        self.protocol("WM_DELETE_WINDOW", self.kill)

        self.label = ttk.Label(self)
        self.label.pack(expand=tk.YES)

        thread = th.Thread(target=self.lancer_application, daemon=True)
        thread.start()

    def recuperer_mandats(self):
        self.label.config(text="Récupération des mandats...")
        time.sleep(1)
        chemin = "C:\\Users\\nathan.dimartino\\Documents\\Archivage affaires\\H"
        self.liste_mandats = []

        for affaire in [d for d in os.listdir(chemin) if os.path.isdir(os.path.join(chemin, d))]:
            for mandat in os.listdir(os.path.join(chemin, affaire)):
                if not mandat.startswith("_"):
                    self.liste_mandats.append(mandat)

        time.sleep(1)
        
    def recuperer_mandats_a_archiver(self):
        self.label.config(text="Récupération des mandats prêts à être archivés...")
        time.sleep(1)
        chemin = "C:\\Users\\nathan.dimartino\\Documents\\Archivage affaires\\H"
        self.liste_mandats_a_archiver = utils.liste_mandat_a_archiver(chemin=chemin)
        nb = len(self.liste_mandats_a_archiver)
        self.label.config(text=f"{nb} mandat{'s' if nb != 1 else ''} prêt{'s' if nb != 1 else ''} à être archivé{'s' if nb != 1 else ''}.")

        time.sleep(1)
        self.label.config(text="Lancement de l'application...")

    def lancer_application(self):

        self.recuperer_mandats()
        self.recuperer_mandats_a_archiver()

        time.sleep(1)

        self.destroy()
        self.on_finish(self.liste_mandats, self.liste_mandats_a_archiver)

    def kill(self):
        self.destroy()
        self.master.destroy()


# APPLICATION PRINCIPALE
class Application(tk.Toplevel):
    def __init__(self, master, liste_mandats, liste_mandats_a_archiver):
        super().__init__(master)
        self.title("QMT Archivage")
        self.state("zoomed")
        self.minsize(480, 360)
        self.configure(background="white")

        self.protocol("WM_DELETE_WINDOW", self.kill)

        self.racine = Racine(self, liste_mandats, liste_mandats_a_archiver)
        self.racine.pack(expand=True, fill="both")

    def kill(self):
        self.destroy()
        self.master.destroy()


# CONTENU PRINCIPAL
class Racine(ttk.Frame):
    def __init__(self, master, liste_mandats, liste_mandats_a_archiver):
        super().__init__(master)
        self.master = master
        self.liste_mandats = liste_mandats
        self.liste_mandats_a_archiver = liste_mandats_a_archiver
        self.barre_menu()
        self.page()

    def barre_menu(self):

        self.menu = tk.Menu(self.master)

        # Fichier
        self.menu_fichier = tk.Menu(self.menu, tearoff=0)

        # Ajouter des mandats à éditer depuis un fichier
        self.menu_fichier.add_command(
            label="Ajouter des mandats à éditer depuis un fichier", 
            command=self.charger, 
            accelerator="Ctrl+O",
        )
        self.bind_all("<Control-o>", lambda event: self.charger())

        # ---
        # Quitter
        self.menu_fichier.add_separator()
        self.menu_fichier.add_command(
            label="Quitter", 
            command=self.kill, 
            accelerator="Ctrl+Q",
        )
        self.bind_all("<Control-q>", lambda event: self.kill())

        self.menu.add_cascade(
            label="Fichier",
            menu=self.menu_fichier,
        )

        # Édition
        self.menu_edition = tk.Menu(self.menu, tearoff=0)

        # Archiver les mandats
        self.menu_edition.add_command(
            label="Archiver les mandats", 
            command=self.a_archiver, 
            accelerator="Ctrl+R",
        )
        self.bind_all("<Control-r>", lambda event: self.a_archiver())

        # Ne plus archiver les mandats
        self.menu_edition.add_command(
            label="Ne plus archiver les mandats", 
            command=self.pas_a_archiver, 
            accelerator="Ctrl+Shift+R",
        )
        self.bind_all("<Control-Shift-KeyPress-R>", lambda event: self.pas_a_archiver())

        self.menu.add_cascade(
            label="Édition",
            menu=self.menu_edition,
        )

        # Sélection
        self.menu_selection = tk.Menu(self.menu, tearoff=0)

        # Ajouter des mandats à la sélection
        self.menu_selection.add_command(
            label="Ajouter des mandats à la sélection depuis un fichier",
            command=self.ajouter_selection,
            accelerator="Ctrl+Shift+O"
        )
        self.bind_all("<Control-Shift-KeyPress-O>", lambda event: self.ajouter_selection())

        # Supprimer la sélection
        self.menu_selection.add_command(
            label="Supprimer la sélection",
            command=self.supprimer_selection,
            accelerator="Ctrl+Shift+D"
        )
        self.bind_all("<Control-Shift-KeyPress-D>", lambda event: self.supprimer_selection())

        self.menu.add_cascade(
            label="Sélection",
            menu=self.menu_selection,
        )

        # Exécution
        self.menu_execution = tk.Menu(self.menu, tearoff=0)

        # Lancer l'archivage
        self.menu_execution.add_command(
            label="Lancer l'archivage", 
            command=self.archiver, 
            accelerator="Ctrl+Entrer",
        )
        self.bind_all("<Control-Return>", lambda event: self.archiver())

        # Arrêter l'archivage
        self.menu_execution.add_command(
            label="Arrêter l'archivage", 
            command=self.stop, 
            accelerator="Ctrl+S",
        )
        self.bind_all("<Control-s>", lambda event: self.stop())

        # Arrêt d'urgence
        self.menu_execution.add_command(
            label="Arrêt d'urgence", 
            command=self.kill, 
            accelerator="Ctrl+K",
        )
        self.bind_all("<Control-k>", lambda event: self.kill())

        self.menu.add_cascade(
            label="Exécution",
            menu=self.menu_execution,
        )

        self.master.config(menu=self.menu)

    # Fonctions utiles
    def charger(self):
        chemin = filedialog.askopenfile(defaultextension="*.xlsx", filetypes=[
            ("Fichiers Excel", "*.xlsx"),
            ("Fichiers Excel", "*.xlsm"),
            ("Fichiers CSV", "*.csv"),
            ("Fichiers CSV", "*.CSV"),
            ("Fichiers texte", "*.txt"),
            ("Fichiers Word", "*.docx")
        ])

        if not chemin:
            return
        chemin = chemin.name

        extension = chemin.split(".")[-1]
        separateur = None
        en_tete = False
        colonne = None
        
        match extension:
            case x if x in ["xlsx", "xlsm"]: 
                try:
                    df = pd.read_excel(chemin)
                    colonne = 0

                    for i in range(len(df.columns)):
                        if df.columns[i].startswith(r"\b\d{4}-\d{2}\b"):
                            colonne = i
                            break
                        elif df.iloc[1, i].startswith(r"\b\d{4}-\d{2}\b"):
                            en_tete = True
                            colonne = i
                            break

                except:
                    colonne = 0

            case x if x in ["csv", "CSV"]:
                try:
                    df = pd.read_csv(chemin)
                    colonne = 0
                    separateur = ","

                    for i in range(len(df.columns)):
                        if df.columns[i].startswith(r"\b\d{4}-\d{2}\b"):
                            colonne = i
                            break
                        elif df.iloc[1, i].startswith(r"\b\d{4}-\d{2}\b"):
                            en_tete = True
                            colonne = i
                            break
                        
                except:
                    colonne = 0
                    separateur = ","

            case "docx":
                try:
                    doc = Document(chemin)
                    texte = ""
                    for para in doc.paragraphs:
                        texte = texte + "SAUT_DE_LIGNE" + para.text

                    separateurs = re.findall(r"\d{4}-\d{2}(.*?)(?=\d{4}-\d{2})", texte)
                    if len(separateurs) == 1 or len(set(separateurs)) == 1:
                        separateur = separateurs[0]
                
                except:
                    pass

            case "txt":
                try:
                    with open(chemin, "r", encoding="utf-8") as fichier:
                        lignes = fichier.readlines()
                    texte = ""
                    for ligne in lignes:
                        texte = texte + "SAUT_DE_LIGNE" + ligne.replace("\n", "")

                    separateurs = re.findall(r"\d{4}-\d{2}(.*?)(?=\d{4}-\d{2})", texte)
                    if len(separateurs) == 1 or len(set(separateurs)) == 1:
                        separateur = separateurs[0]
                
                except:
                    pass

        PopUpCharger(self.master, on_finish=None, geometry="450x300+543+222", chemin=chemin, separateur=separateur, en_tete=en_tete, colonne=colonne)


    def kill(self):
        self.master.kill()

    def a_archiver(self):
        pass

    def pas_a_archiver(self):
        pass

    def ajouter_selection(self):
        liste_mandats = self.charger()

    def supprimer_selection(self):
        pass

    def archiver(self):
        print("Archivage")

    def stop(self):
        pass

    def precedent_mandat(self, event):
        match event.keysym:
            case "Up":
                sens = -1
            case "Down":
                sens = 1
            case _:
                return

        try:
            _ = self.precedentes_recherches[self.indice_recherches + sens]
            self.indice_recherches += sens
            self.barre_de_recherche.delete(0, tk.END)
            self.barre_de_recherche.insert(0, self.precedentes_recherches[self.indice_recherches])

        except IndexError:
            pass

    def afficher_message_temporaire(self, texte, couleur="red", duree=2000, wait=3000):
        top = tk.Toplevel(self)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        top.attributes("-alpha", 1.0)
        top.configure(bg=couleur)

        frame = tk.Frame(
            top,
            bg=couleur,
        )
        label = tk.Text(
            frame, 
            bg=couleur, 
            fg="white", 
            font=("Segoe UI", 9, "bold"), 
            width=50, height=min(len(texte.split("\n")), 15), 
            bd=0
        )
        label.insert('1.0', texte)
        label.config(state="disabled")
        label.pack(padx=10, pady=5)
        frame.pack()

        # ➜ Mise à jour du placement principal
        self.update_idletasks()
        top.update_idletasks()

        # ➜ Centrage par rapport à la fenêtre principale sur le bon écran
        x = self.winfo_rootx() + 20
        y = self.winfo_rooty() + 20
        top.geometry(f"+{x}+{y}")

        top.after(wait, lambda: self._fade_out(top, 1.0, duree))

    def _fade_out(self, window, alpha, duree):
        steps = 200
        delay = duree // steps

        def fade():
            nonlocal alpha
            alpha -= 1 / steps
            if alpha <= 0:
                window.destroy()
            else:
                window.attributes("-alpha", alpha)
                window.after(delay, fade)

        fade()

    def ajouter_ligne_a_editer(self, mandat):
        frame = ttk.Frame(self.frame_liste_mandats_a_editer)
        frame.grid_columnconfigure(0, weight=0)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=0)

        var_selection = tk.BooleanVar()
        checkbox = ttk.Checkbutton(frame, variable=var_selection)
        checkbox.grid(row=0, column=0, padx=5)

        label = ttk.Label(frame, text=mandat)
        label.grid(row=0, column=1, sticky="w")

        btn_supprimer = ttk.Button(frame, text="x", width=3, command=lambda: self.supprimer_ligne(frame, (mandat, var_selection)))
        btn_supprimer.grid(row=0, column=2, padx=5)

        frame.pack(fill=tk.X, side=tk.TOP, anchor=tk.N, pady=2)

        self.liste_mandats_a_editer.append((mandat, var_selection))

    def supprimer_ligne(self, frame, mandat_tuple):
        frame.destroy()
        if mandat_tuple in self.liste_mandats_a_editer:
            self.liste_mandats_a_editer.remove(mandat_tuple)
            self.liste_mandats.append(mandat_tuple[0])

    def ajouter_mandat_a_editer(self, mandat=None):
        if mandat is None:
            mandat = self.barre_de_recherche.get()
        mandat = mandat.strip()
        mandat = mandat.replace("\n", "")

        if mandat in self.liste_mandats:
            self.ajouter_ligne_a_editer(mandat)
            self.barre_de_recherche.delete(0, tk.END)
            self.precedentes_recherches.append(mandat)
            self.indice_recherches = 0
            self.liste_mandats.remove(mandat)

        candidats = [m for m in self.liste_mandats if mandat in m]

        if mandat == "":
            pass
        elif len(candidats) == 0:
            self.afficher_message_temporaire(f"Aucun mandat correspondant à \"{mandat}\".")
        else:
            self.afficher_message_temporaire(f"Aucun mandat correspondant à \"{mandat}\".\nEssayer plutôt :\n- {"\n- ".join([f"\"{m}\"" for m in candidats])}")


    def page(self):
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.en_tete().grid(row=0, column=0, columnspan=2, sticky="ew")
        self.colonne_gauche().grid(row=1, column=0, sticky="news")
        self.colonne_droite().grid(row=1, column=1, sticky="news")
        self.pied_de_page().grid(row=2, column=0, columnspan=2, sticky="ew")

    def en_tete(self):
        frame = ttk.Frame(self)
        label_en_tete = ttk.Label(frame, text="QMT Archivage", font=("Helvetica", 16))
        label_en_tete.pack(padx=10, pady=10)
        return frame

    def colonne_gauche(self):
        frame = ttk.Frame(self)
        label_colonne_gauche = ttk.Label(frame, text="Mandats à  archiver")
        label_colonne_gauche.pack(padx=10, pady=10)

        self.label_nombre_mandats = ttk.Label(frame, text=f"{len(self.liste_mandats_a_archiver)} mandat{"s" if len(self.liste_mandats_a_archiver) else ""} prêt{"s" if len(self.liste_mandats_a_archiver) else ""} à être archivés.")
        self.label_nombre_mandats.pack(padx=10, pady=10)

        self.frame_liste_mandats_a_archiver = ttk.Frame(frame)
        self.frame_liste_mandats_a_archiver.pack(expand=tk.YES, fill=tk.BOTH, padx=10, pady=10)

        ligne = 0
        for mandat, _ in self.liste_mandats_a_archiver:
            ttk.Label(self.frame_liste_mandats_a_archiver, text=mandat).grid(row=ligne, column=0, sticky="w")
            ligne += 1
        return frame

    def colonne_droite(self):
        frame = ttk.Frame(self)
        label_colonne_droite = ttk.Label(frame, text="Éditer des mandats")
        label_colonne_droite.pack(padx=10, pady=10)

        frame_barre_de_recherche = ttk.Frame(frame)
        frame_barre_de_recherche.grid_columnconfigure(0, weight=1)
        frame_barre_de_recherche.grid_columnconfigure(1, weight=0)
        frame_barre_de_recherche.grid_columnconfigure(2, weight=0)
        frame_barre_de_recherche.pack(fill=tk.X, padx=10, pady=10)

        self.barre_de_recherche = ttk.Entry(frame_barre_de_recherche)
        self.barre_de_recherche.grid(row=0, column=0, sticky="ew")
        self.barre_de_recherche.bind("<Return>", lambda event: self.ajouter_mandat_a_editer())
        self.barre_de_recherche.bind("<Up>", self.precedent_mandat)
        self.barre_de_recherche.bind("<Down>", self.precedent_mandat)
        self.barre_de_recherche.bind("<Key>", lambda e: setattr(self, "indice_recherches", 0))

        self.bouton_recherche = ttk.Button(frame_barre_de_recherche, command=self.ajouter_mandat_a_editer, text="+")
        self.bouton_recherche.grid(row=0, column=1)
        self.precedentes_recherches = []

        self.bouton_charger = ttk.Button(frame_barre_de_recherche, text="+", command=self.charger)
        self.bouton_charger.grid(row=0, column=2)

        self.frame_liste_mandats_a_editer = ttk.Frame(frame)
        self.frame_liste_mandats_a_editer.pack_propagate(False)
        self.frame_liste_mandats_a_editer.pack(expand=tk.YES, fill=tk.BOTH, padx=10, pady=10, anchor=tk.N)
        self.liste_mandats_a_editer = []

        return frame

    def pied_de_page(self):
        frame = ttk.Frame(self)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=1)

        self.bouton_archiver = ttk.Button(frame, text="Archiver", command=self.archiver)
        self.bouton_archiver.grid(row=0, column=0, sticky="news", padx=10, pady=10)
        self.bouton_stop = ttk.Button(frame, text="Stop", command=self.stop, state="disabled")
        self.bouton_stop.grid(row=0, column=1, sticky="news", padx=10, pady=10)
        self.bouton_kill = ttk.Button(frame, text="Arrêt d'urgence", command=self.kill, state="disabled")
        self.bouton_kill.grid(row=0, column=2, sticky="news", padx=10, pady=10)
        return frame


class PopUp(tk.Toplevel):
    def __init__(self, master, on_finish, geometry):
        super().__init__(master)
        self.geometry(geometry)
        self.resizable(False, False)
        self.configure(background="white")
        self.on_finish = on_finish

        self.grab_set()
        self.page()

    def page(self):
        pass


class PopUpCharger(PopUp):
    def __init__(self, master, on_finish, geometry, chemin, separateur, en_tete, colonne):
        self.chemin = chemin
        self.separateur = separateur
        self.avec_en_tete = en_tete
        self.colonne = colonne

        self.extension = self.chemin.split(".")[-1]

        super().__init__(master, on_finish, geometry)

    def page(self):
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)
        self.grid_columnconfigure(0, weight=1)

        frame_en_tete = self.en_tete()
        frame_en_tete.grid(row=0, column=0, sticky="news")
        frame_pied_de_page = self.pied_de_page()
        frame_pied_de_page.grid(row=2, column=0, sticky="news")

        if self.extension in ["xlsx", "xlsm"]:
            frame_corps = self.page_excel()
        elif self.extension in ["csv", "CSV"]:
            frame_corps = self.page_csv()
        else:
            frame_corps = self.page_texte()

        frame_corps.grid(row=1, column=0, sticky="new", padx=10, pady=10)

    def en_tete(self):
        frame = ttk.Frame(self)
        label_en_tete = ttk.Label(frame, text="Charger un fichier", font=("Helvetica", 16))
        label_en_tete.pack(padx=10, pady=10)
        return frame

    def page_excel(self):
        frame = ttk.Frame(self)
        return frame
    
    def page_csv(self):
        frame = ttk.Frame(self)
        return frame
    
    def page_texte(self):
        frame = ttk.Frame(self)
        frame.grid_rowconfigure(0, weight=0)
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_rowconfigure(2, weight=1)
        frame.grid_columnconfigure(0, weight=0)
        frame.grid_columnconfigure(1, weight=0)
        frame.grid_columnconfigure(2, weight=1)

        self.label_nombre_mandats = ttk.Label(frame)
        self.label_nombre_mandats.grid(row=0, column=0, columnspan=3, padx=10, pady=(0, 10))

        thread = th.Thread(target=self.detection_des_mandats, daemon=True)
        thread.start()

        label_chemin = ttk.Label(frame, text="Fichier")
        label_chemin.grid(row=1, column=0, sticky="w")
        ttk.Label(frame, text=":").grid(row=1, column=1, sticky="ew", padx=10)

        bouton_chemin = tk.Button(frame, text=self.chemin.split("/")[-1], anchor="w", command=lambda chemin=self.chemin: os.startfile(chemin))
        bouton_chemin.grid(row=1, column=2, sticky="ew")

        label_separateur = ttk.Label(frame, text="Séparateur")
        label_separateur.grid(row=2, column=0, sticky="w")
        ttk.Label(frame, text=":").grid(row=2, column=1, sticky="ew", padx=10)

        self.barre_separateur = ttk.Entry(frame)
        self.barre_separateur.insert(0, ("Saut de ligne" if self.separateur == "SAUT_DE_LIGNE" else "Espace" if self.separateur == " " else self.separateur if self.separateur else ""))
        self.barre_separateur.bind("<KeyRelease>", lambda event: self.modifier_barre(self.barre_separateur))
        self.barre_separateur.grid(row=2, column=2, sticky="ew")

        if self.barre_separateur.get() != "":
            self.bouton_charger.config(state="normal")
            self.bouton_tester.config(state="normal")

        return frame
    
    def pied_de_page(self):
        frame = ttk.Frame(self)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=1)
        frame.grid_columnconfigure(3, weight=1)

        self.bouton_charger = ttk.Button(frame, text="Charger", state="disabled")
        self.bouton_charger.grid(row=0, column=0, sticky="news", padx=(10, 5), pady=10)
        self.bouton_tester = ttk.Button(frame, text="Tester", command=self.detection_des_mandats, state="disabled")
        self.bouton_tester.grid(row=0, column=1, sticky="news", padx=5, pady=10)
        self.bouton_choisir_autre_fichier = ttk.Button(frame, text="Choisir un autre fichier")
        self.bouton_choisir_autre_fichier.grid(row=0, column=2, sticky="news", padx=5, pady=10)
        self.bouton_annuler = ttk.Button(frame, text="Annuler", command=self.destroy)
        self.bouton_annuler.grid(row=0, column=3, sticky="news", padx=(5, 10), pady=10)

        return frame
    
    def detection_des_mandats(self):

        try:
            separateur = self.barre_separateur.get()
            self.separateur = "SAUT_DE_LIGNE" if separateur.strip().lower() == "saut de ligne" else " " if separateur.strip().lower() == "espace" else separateur
        except:
            pass

        self.label_nombre_mandats.config(text="Détection des mandats...")
        self.liste_mandats = utils.charger_liste_mandats(self.chemin, ("\n" if self.separateur == "SAUT_DE_LIGNE" else self.separateur), self.en_tete, self.colonne)
        self.label_nombre_mandats.config(text=f"{len(self.liste_mandats)} mandat{"s" if len(self.liste_mandats) > 1 else ""} détecté{"s" if len(self.liste_mandats) > 1 else ""}.")

    def modifier_barre(self, barre):

        if barre.get() != "":
            self.bouton_charger.config(state="normal")
            self.bouton_tester.config(state="normal")
        else:
            self.bouton_charger.config(state="disabled")
            self.bouton_tester.config(state="disabled")

        if barre.get().strip().lower() == "saut de ligne":
            barre.delete(0, tk.END)
            barre.insert(0, "Saut de ligne")
        elif barre.get().strip().lower() == "espace":
            barre.delete(0, tk.END)
            barre.insert(0, "Espace")



# LANCEMENT DE L'APPLICATION
if __name__ == "__main__":

    root = tk.Tk()
    root.withdraw()

    def lancer_app(mandats, mandats_a_archiver):
        app = Application(root, mandats, mandats_a_archiver)

    splash = Demarrage(master=root, on_finish=lancer_app)
    root.mainloop()