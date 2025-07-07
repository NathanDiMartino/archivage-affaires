import tkinter as tk
from tkinter import ttk

class PopUp(tk.Toplevel):
    def __init__(self, master, racine, on_finish, largeur, hauteur):

        super().__init__(master)
        self.master = master
        self.update_idletasks()
        self.master.update_idletasks()  # très important

        x = self.master.winfo_rootx() + self.master.winfo_width() // 2 - largeur // 2
        y = self.master.winfo_rooty() + self.master.winfo_height() // 2 - hauteur // 2
        geometry = f"{largeur}x{hauteur}+{x}+{y}"

        self.geometry(geometry)
        self.resizable(False, False)
        self.configure(background="white")
        self.on_finish = on_finish
        self.racine = racine

        self.grab_set()
        self.page()

    def page(self):
        pass


class MenuDeroulant(ttk.Label):
    def __init__(self, master, items, taille, command=None, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        self.items = items
        self.taille = taille
        self.command = command

        self.bind("<Button-1>", self.show_items)
        self.bind("<FocusIn>", self.show_items)
        self.bind("<FocusOut>", self.hide_items)
        self.bind_all("<Escape>", lambda e: self.hide_items())

        self.dropdown = None
        self.listbox = None
        self.selection_index = None

    def get(self):
        return self["text"]
    
    def set_items(self, nouveaux_items):
        self.items = nouveaux_items

    def show_items(self, event=None):
        if self.dropdown:
            return

        self.update_idletasks()

        # Création du Toplevel
        self.dropdown = tk.Toplevel(self)
        self.dropdown.wm_overrideredirect(True)
        self.dropdown.attributes("-topmost", True)

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        height = min(len(self.items), self.taille) * 20
        self.dropdown.geometry(f"{self.winfo_width()}x{height}+{x}+{y}")

        # Scrollbar + Listbox
        scrollbar = tk.Scrollbar(self.dropdown, orient="vertical")
        self.listbox = tk.Listbox(
            self.dropdown, height=min(len(self.items), self.taille),
            yscrollcommand=scrollbar.set
        )
        scrollbar.config(command=self.listbox.yview)

        for item in self.items:
            self.listbox.insert(tk.END, item)

        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.listbox.bind("<<ListboxSelect>>", self.select_item)

    def hide_items(self, event=None):
        if self.dropdown:
            self.dropdown.destroy()
            self.dropdown = None
            self.listbox = None

    def select_item(self, event=None):
        if self.listbox:
            selection = self.listbox.get(self.listbox.curselection())
            self.selection_index = self.listbox.curselection()[0]
            self.config(text=selection)
            self.hide_items()
            if self.command:
                try:
                    self.command(selection)
                except:
                    self.command()

    def get_index(self):
        return self.selection_index
    

class Spinner(tk.Canvas):
    def __init__(self, parent, radius=8, speed=50, **kwargs):
        super().__init__(parent, width=radius*2, height=radius*2, highlightthickness=0, **kwargs)
        self.radius = radius
        self.angle = 0
        self.speed = speed
        self.arc = self.create_arc(2, 2, radius*2-2, radius*2-2, start=0, extent=90, style="arc", width=2)
        self.animating = False

    def start(self):
        if not self.animating:
            self.animating = True
            self._animate()

    def stop(self):
        self.animating = False

    def _animate(self):
        if self.animating:
            self.angle = (self.angle + 10) % 360
            self.itemconfig(self.arc, start=self.angle)
            self.after(self.speed, self._animate)


class ChargementButton(ttk.Frame):
    def __init__(self, master, text="", padding=5, height=3, width=3, command=None, state="normal", **kwargs):
        super().__init__(master)
        self.spinner = Spinner(self)
        self.label = ttk.Label(self, text=text, anchor="center")
        self.command = command
        self._state = state
        self.padding = padding
        self.height = height
        self.width = width

        # Place le spinner à gauche
        self.spinner.pack(side="left", padx=(self.padding, self.width), pady=self.height)
        # Place le label qui prend tout le reste
        self.label.pack(side="left", expand=True, fill="both", padx=self.width, pady=self.height)

        self.bind("<Button-1>", lambda e: self.invoke())
        self.label.bind("<Button-1>", lambda e: self.invoke())
        self.spinner.bind("<Button-1>", lambda e: self.invoke())

        self.configure(cursor="hand2")
        self.set_state(self._state)

        self.hide()

    def set_state(self, state):
        self._state = state
        if state == "disabled":
            self.label.state(["disabled"])
            self.configure(cursor="arrow")
        else:
            self.label.state(["!disabled"])
            self.configure(cursor="hand2")

    def show(self):
        self.label.pack_forget()
        self.spinner.pack(side="left", padx=(self.padding, self.width), pady=self.height)
        self.label.pack(side="left", expand=True, fill="both", padx=self.width, pady=self.height)

    def hide(self):
        self.spinner.pack_forget()

    def start(self):
        self.show()
        self.spinner.start()

    def stop(self):
        self.spinner.stop()
        self.hide()

    def configure(self, **kwargs):
        if "text" in kwargs:
            self.label.config(text=kwargs.pop("text"))
        if "state" in kwargs:
            self.set_state(kwargs.pop("state"))
        return super().configure(**kwargs)

    config = configure

    def invoke(self):
        if self._state != "disabled" and self.command:
            self.command()