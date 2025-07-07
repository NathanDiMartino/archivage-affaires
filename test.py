import tkinter as tk
from tkinter import HORIZONTAL

def update_value(val):
    label.config(text=f"Valeur : {val}")

# Création de la fenêtre principale
root = tk.Tk()
root.title("Potentiomètre Tkinter")

# Ajout d'un widget Scale (curseur horizontal)
potentiometer = tk.Scale(root, from_=0, to=100, orient=HORIZONTAL, command=update_value)
potentiometer.pack(pady=20)

# Affichage de la valeur actuelle
label = tk.Label(root, text="Valeur : 0")
label.pack()

# Lancement de la boucle principale
root.mainloop()