import os
import ctypes
import shutil
import zipfile
import time
import argparse
import datetime


def barre_de_progression(pourcentage, curseur, longueur=100, temps_depart=None, affichage_temps_restant=False):
    """
    Affiche une barre de progression dans la console.
    
    :param pourcentage: Le pourcentage d'avancement (0 à 100).
    :param longueur: La longueur de la barre de progression.
    :param temps_restant: Temps restant estimé en secondes (optionnel).
    """

    barre = "█" * int(pourcentage / 100 * longueur)
    espace = " " * (longueur - len(barre))

    # Calcul du temps restant si le temps de départ est fourni
    if temps_depart is not None and affichage_temps_restant:
        temps_ecoule = time.time() - temps_depart
        temps_restant = temps_ecoule * (100 - pourcentage) / pourcentage if pourcentage > 0 else None

    if affichage_temps_restant:
        print(f"\r{curseur} |{barre}{espace}| {pourcentage:.2f}% - Temps restant : {int(round(temps_restant)) // 60} min {temps_restant - (int(round(temps_restant)) // 60) * 60:.0f} s", end=30 * " ")
    else:
        print(f"\r{curseur} |{barre}{espace}| {pourcentage:.2f}%", end=30 * " ")


def archivage(chemin, chemin_archive):
    """
    Retourne une liste des mandats prêts pour archivage.
    Un mandat est considérée prêt à être archivé si l'attribut "Le dossier est prêt à être archivé" est coché pour le dossier.

    :param chemin: Chemin du dossier contenant les affaires
    :param chemin_archive: Chemin du dossier d'archivage.
    """

    # Préparation du fichier de log
    log_file = os.path.join(chemin, f"log_ajouter_dossier_archivage {datetime.datetime.today().strftime("%d.%m.%Y - %H.%M.%S")}.txt")
    log = f"Début de la préparation des mandats pour archivage - {datetime.datetime.today().strftime("le %d/%m/%Y à %H:%M:%S")}\n\n"

    debut = time.time()

    # Définition de l'attribut à vérifier
    ATTRIBUT_PRET_ARCHIVAGE = 0x20

    if not os.path.exists(chemin):
        log += f"Erreur : Le chemin spécifié n'existe pas : {chemin}\n"

        # Écriture du log dans le fichier
        with open(log_file, "w") as f:
            f.write(log)

        raise FileNotFoundError(f"Le chemin spécifié n'existe pas : {chemin}")
    
    if not os.path.exists(chemin_archive):
        log += f"Erreur : Le chemin d'archivage spécifié n'existe pas : {chemin_archive}\n"

        # Écriture du log dans le fichier
        with open(log_file, "w") as f:
            f.write(log)

        raise FileNotFoundError(f"Le chemin spécifié n'existe pas : {chemin_archive}")

    # Liste des mandats archivés
    mandats_archives = []
    # Liste des mandats pour lesquels l'archivage a échoué
    mandats_non_archives = []

    # Parcours des affaires et des mandats
    for affaire in os.listdir(chemin):

        # Flag pour ne pas avoir à recréer l'archive des articles
        chemin_zip = ""

        # Vérification que c'est un dossier
        if not os.path.isdir(os.path.join(chemin, affaire)):
            continue

        # On récupère le dossier Articles (s'il existe)
        dossier_articles = [dossier for dossier in os.listdir(os.path.join(chemin, affaire)) if dossier.startswith("_")]

        if dossier_articles and len(dossier_articles) == 1:
            dossier_articles = os.path.join(chemin, affaire, dossier_articles[0])
        else:
            dossier_articles = "___xxxxx___" # Pour être sûr que le dossier n'existe pas

        for mandat in os.listdir(os.path.join(chemin, affaire)):

            # Vérification que c'est un dossier
            if not os.path.isdir(os.path.join(chemin, affaire, mandat)) or mandat.startswith("_"):
                continue

            else:
                try:
                    debut_archivage_tot = time.time()

                    dossier_mandat = os.path.join(chemin, affaire, mandat)

                    # Vérification de l'attribut "Le dossier est prêt à être archivé"
                    attributs = ctypes.windll.kernel32.GetFileAttributesW(dossier_mandat)
                    if attributs & ATTRIBUT_PRET_ARCHIVAGE:
                        # Log l'action effectuée
                        log += f"Le mandat {mandat} de l'affaire {affaire} est prêt à être archivé : {dossier_mandat}\n"

                        print(f"\nLe mandat {mandat} de l'affaire {affaire} est prêt à être archivé.")
                        print(f"Début de l'archivage...")

                        # Copie dans Z:
                        log  += f"Copie du mandat {mandat} vers {chemin_archive}{affaire}\\{mandat}\n"

                        print(f"\nCopie du mandat {mandat} vers {chemin_archive}{affaire}\\{mandat}...")

                        destination = os.path.join(chemin_archive, affaire, mandat)
                        os.makedirs(os.path.dirname(destination), exist_ok=True)
                        shutil.copytree(dossier_mandat, destination, dirs_exist_ok=True)
                        barre_de_progression(100, "✓", affichage_temps_restant=False)

                        # Log la copie
                        log += f"Copie du mandat {mandat} vers {chemin_archive}{affaire}\\{mandat} terminée.\n"
                        print(f"\nCopie du mandat {mandat} vers {chemin_archive}{affaire}\\{mandat} terminée.")

                        # Ajout de l'archive du dossier Articles s'il existe
                        if os.path.exists(dossier_articles):
                            source_zip = dossier_articles
                            nom_zip = f"_{affaire.split(' ')[0]}-AA (Articles).zip"

                            # Localisation du dossier de destination
                            destination_zip = os.path.join(chemin_archive, affaire, mandat, "000_Zip articles livres", nom_zip)

                            for dossier in os.listdir(os.path.join(chemin_archive, affaire, mandat)):
                                if dossier.startswith("000"):
                                    destination_zip = os.path.join(chemin_archive, affaire, mandat, dossier, nom_zip)
                                    break

                            if not os.path.exists(destination_zip):
                                # Création du dossier de destination s'il n'existe pas
                                os.makedirs(os.path.dirname(destination_zip), exist_ok=True)
                            
                            log += f"Création de l'archive {nom_zip} pour le dossier Articles du mandat {mandat}...\n"
                            print(f"\nCréation de l'archive {nom_zip} pour le dossier Articles du mandat {mandat}...")

                            # Création de l'archive du dossier Articles seulement si elle n'est pas déjà créée
                            if chemin_zip == "":

                                # Récupération du nombre de fichiers à archiver
                                nombre_fichiers = sum(len(files) for _, _, files in os.walk(source_zip))

                                debut_archivage = time.time()

                                with zipfile.ZipFile(destination_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                                    i = 1
                                    for root, dirs, files in os.walk(source_zip):
                                        for file in files:

                                            # Affichage de la progression
                                            curseurs = "◴◷◶◵"
                                            barre_de_progression(i / nombre_fichiers * 100, curseurs[i % 4], temps_depart=debut_archivage, affichage_temps_restant=True)
                                            i += 1
                                            
                                            filepath = os.path.join(root, file)

                                            # On enlève le chemin racine pour garder une structure relative
                                            arcname = os.path.relpath(filepath, start=source_zip)
                                            zipf.write(filepath, arcname)

                                barre_de_progression(100, "✓", affichage_temps_restant=False)

                                # Log la création de l'archive
                                log += f"Création de l'archive {nom_zip} pour le dossier Articles du mandat {mandat} terminée.\n"
                                print(f"\nCréation de l'archive {nom_zip} pour le dossier Articles du mandat {mandat} terminée.")

                                # On garde le chemin de l'archive pour la copie
                                chemin_zip = destination_zip

                            else:

                                shutil.copy(chemin_zip, destination_zip)
                                barre_de_progression(100, "✓", affichage_temps_restant=False)

                                # Log la copie de l'archive
                                log += f"Copie de l'archive {nom_zip} pour le dossier Articles du mandat {mandat} terminée.\n"
                                print(f"\nCopie de l'archive {nom_zip} pour le dossier Articles du mandat {mandat} terminée.")

                        else:
                            log += f"Aucun dossier Articles trouvé pour le mandat {mandat}. Aucune archive créée.\n"
                            print(f"\nAucun dossier Articles trouvé pour le mandat {mandat}. Aucune archive créée.")

                        # Suppression du dossier du mandat
                        log += f"Suppression du dossier du mandat {mandat}...\n"
                        print(f"\nSuppression du dossier du mandat {mandat}...")
                        shutil.rmtree(dossier_mandat)

                        fin_archivage = time.time()
                        log += f"Archivage du mandat {mandat} de l'affaire {affaire} terminé (temps écoulé : {int(round(fin_archivage - debut_archivage_tot)) // 60} min {(fin_archivage - debut_archivage_tot) - (int(round(fin_archivage - debut_archivage_tot)) // 60) * 60:.0f} s).\n"
                        print(f"\nFin de l'archivage (temps écoulé : {int(round(fin_archivage - debut_archivage_tot)) // 60} min {(fin_archivage - debut_archivage_tot) - (int(round(fin_archivage - debut_archivage_tot)) // 60) * 60:.0f} s).\n")
                        mandats_archives.append(dossier_mandat)

                except Exception as e:
                    # Log l'erreur et continue avec le prochain mandat
                    log += f"Erreur lors de l'archivage du mandat {mandat} de l'affaire {affaire} : {e}\n"
                    print(f"Erreur lors de l'archivage du mandat {mandat} de l'affaire {affaire} : {e}")
                    mandats_non_archives.append(mandat)

    fin = time.time()
    
    log += f"\nFin de la préparation des mandats pour archivage - {datetime.datetime.today().strftime("le %d/%m/%Y à %H:%M:%S")}\n"
    log += f"Temps écoulé : {int(round(fin - debut)) // 3600} h {(fin - debut) - (int(round(fin - debut)) // 3600) * 60:.0f} min {(fin - debut) - ((fin - debut) - (int(round(fin - debut)) // 3600) * 60) * 60:.0f} s.\n"
    # Écriture du log dans le fichier
    with open(log_file, "w") as f:
        f.write(log)
    print(f"\nFin de l'archivage de tous les mandats\nTemps écoulé : {int(round(fin - debut)) // 3600} h {(fin - debut) - (int(round(fin - debut)) // 3600) * 60:.0f} min {(fin - debut) - ((fin - debut) - (int(round(fin - debut)) // 3600) * 60) * 60:.0f}.\n")

    return mandats_archives, mandats_non_archives


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="""
    Retourne une liste des mandats prêts pour archivage.
    Un mandat est considérée prêt à être archivé si l'attribut "Le dossier est prêt à être archivé" est coché pour le dossier.
                                     
    :param chemin: Chemin du dossier contenant les affaires (par défaut H:\\).
    :param chemin_archive: Chemin du dossier d'archivage (par défaut Z:\\Affaires\\).
                                     
    Usage : python executer_archivage.py [chemin] [chemin_archive]
    """)
    
    parser.add_argument("--chemin", type=str, default="H:\\", help="Chemin du dossier contenant les affaires (par défaut H:\\)")
    parser.add_argument("--chemin_archive", type=str, default="Z:\\Affaires\\", help="Chemin du dossier où archiver les mandats (par défaut Z:\\Affaires\\)")
    args = parser.parse_args()

    mandats_archives, mandats_non_archives = archivage(args.chemin, args.chemin_archive)
    
    print("Mandats archivagés :", end=" ")
    print(f"({len(mandats_archives)})\n\"{'\", \n\"'.join(mandats_archives)}\"\n" if mandats_archives else "Aucun mandat archivé.\n")
    print("Mandats non archivés :", end=" ")
    print(f"({len(mandats_non_archives)})\n\"{'\", \n\"'.join(mandats_non_archives)}\"\n" if mandats_non_archives else "Aucun mandat non archivé.\n")
