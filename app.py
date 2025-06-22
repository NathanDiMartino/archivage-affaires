import tkinter as tk
import tkinter.ttk as ttk


class Application(tk.Tk):

    def __init__(self):

        super().__init__()
        self.wm_state(newstate="zoomed")
        self.minsize(480, 360)
        self.configure(background="white")

        self.barre_menu()

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


Application().mainloop()