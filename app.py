import tkinter as tk
import tkinter.ttk as ttk
import threading as th

import utils
import time


class Demarrage(tk.Tk):

    def __init__(self):

        super().__init__()
        self.geometry("400x100+568+322")
        self.configure(background="white")

        self.label = ttk.Label(master=self)
        self.label["text"] = "Récupération des mandats prêts à être archivés..."
        self.label.pack(expand=tk.YES)

        thread = th.Thread(target=self.recuperer_mandats_a_archiver)
        thread.daemon = True
        thread.start()

    def recuperer_mandats_a_archiver(self):
        """
        Récupère la liste des mandats à archiver et la sauvegarde.
        """

        time.sleep(1)
        chemin="C:\\Users\\nathan.dimartino\\Documents\\Archivage affaires\\H"
        self.liste_mandats_a_archiver = utils.liste_mandat_a_archiver(chemin=chemin)
        self.label["text"] = f"{len(self.liste_mandats_a_archiver)} mandat{"" if len(self.liste_mandats_a_archiver) < 2 else "s"} prêt{"" if len(self.liste_mandats_a_archiver) < 2 else "s"} à être archiver."
        
        time.sleep(1)
        self.label["text"] = "Lancement de l'application..."

        time.sleep(1)
        self.after(0, self.lancer_application)

    def lancer_application(self):
        self.destroy()
        # Lance Application dans le thread principal
        app = Application(liste_mandats_a_archiver=self.liste_mandats_a_archiver)
        app.mainloop()
        

class Application(tk.Tk):

    def __init__(self, liste_mandats_a_archiver):

        super().__init__()
        self.wm_state(newstate="zoomed")
        self.minsize(480, 360)
        self.configure(background="white")

        self.liste_mandats_a_archiver=liste_mandats_a_archiver

        # Crée la barre de menu
        self.barre_menu()

        # Crée le contenu
        self.page()

    # Barre de menu
    def barre_menu(self):

        self.menu = tk.Menu(self)

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
        self.bind_all("<Control-Enter>", lambda event: self.archiver())

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

        self.config(menu=self.menu)

    # Fonctions utiles
    def charger(self):
        pass

    def kill(self):
        pass

    def a_archiver(self):
        pass

    def pas_a_archiver(self):
        pass

    def ajouter_selection(self):
        pass

    def supprimer_selection(self):
        pass

    def archiver(self):
        pass

    def stop(self):
        pass

    # Création de la page et de son contenu
    def page(self):
        """
        Crée la page avec son contenu, en faisant appel aux fonctions en_tete, colonne_gauche, colonne_droite et pied_de_page.
        """

        # Création du frame principal
        self.frame = ttk.Frame(master=self)
        self.frame.pack(expand=tk.YES, fill=tk.BOTH)

        self.frame.grid_rowconfigure(0, weight=0)
        self.frame.grid_rowconfigure(1, weight=1)
        self.frame.grid_rowconfigure(2, weight=0)

        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_columnconfigure(1, weight=1)

        # Création de l'en-tête
        self.en_tete_frame = self.en_tete()
        self.en_tete_frame.grid(row=0, column=0, columnspan=2)

        # Création de la colonne de gauche
        self.colonne_gauche_frame = self.colonne_gauche()
        self.colonne_gauche_frame.grid(row=1, column=0)

        # Création de la colonne de droite
        self.colonne_droite_frame = self.colonne_droite()
        self.colonne_droite_frame.grid(row=1, column=1)

        # Création du pied-de-page
        self.pied_de_page_frame = self.pied_de_page()
        self.pied_de_page_frame.grid(row=2, column=0, columnspan=2)

    def en_tete(self):
        """
        crée le contenu de l'en-tête
        """

        frame = ttk.Frame(master=self.frame)
        label = ttk.Label(master=frame)
        label["text"] = "QMT Archivage"
        label.pack(expand=tk.YES, fill=tk.BOTH)

        return frame

    def colonne_gauche(self):
        """
        crée le contenu de la colonne de gauche
        """

        frame = ttk.Frame(master=self.frame)
        label = ttk.Label(master=frame)
        label["text"] = "QMT Archivage"
        label.pack(expand=tk.YES, fill=tk.BOTH)

        return frame

    def colonne_droite(self):
        """
        crée le contenu de la colonne de droite
        """

        frame = ttk.Frame(master=self.frame)
        label = ttk.Label(master=frame)
        label["text"] = "QMT Archivage"
        label.pack(expand=tk.YES, fill=tk.BOTH)

        return frame

    def pied_de_page(self):
        pass

Demarrage().mainloop()