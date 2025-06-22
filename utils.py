import os
import filecmp
import ctypes
import shutil
import zipfile
import time
import datetime


def liste_mandat_a_archiver(chemin):
    """
    Retourne une liste des mandats prêts pour archivage.
    Un mandat est considérée prêt à être archivé si l'attribut "Le dossier est prêt à être archivé" est coché pour le dossier.

    :param chemin: Chemin du dossier contenant les affaires
    """

    # Définition de l'attribut à vérifier
    ATTRIBUT_PRET_ARCHIVAGE = 0x20

    if not os.path.exists(chemin):
        raise FileNotFoundError(f"Le chemin spécifié n'existe pas : {chemin}")

    # Liste des mandats à archivéer
    mandats_a_archiver = []

    # Parcours des affaires et des mandats
    for affaire in os.listdir(chemin):

        # Vérification que c'est un dossier
        if not os.path.isdir(os.path.join(chemin, affaire)):
            continue

        # Parcours des mandats
        for mandat in os.listdir(os.path.join(chemin, affaire)):

            # Vérification que c'est un dossier
            if not os.path.isdir(os.path.join(chemin, affaire, mandat)) or mandat.startswith("_"):
                continue

            dossier_mandat = os.path.join(chemin, affaire, mandat)

            # Vérification de l'attribut "Le dossier est prêt à être archivé"
            attributs = ctypes.windll.kernel32.GetFileAttributesW(dossier_mandat)
            if attributs & ATTRIBUT_PRET_ARCHIVAGE:
                mandats_a_archiver.append((mandat, dossier_mandat))

    return mandats_a_archiver


def comparer_deux_dossiers(dossier1, dossier2):
    """
    Renvoie un booléen qui indique si les deux dossiers sont indetiques

    :param dossier1: Chemin du premier dossier à comparer
    :param dossier2: Chemin du deuxième dossier à comparer
    """
    
    # Vérification en partant du dossier1
    for chemin1, sous_dossiers1, fichiers1 in os.walk(dossier1):
        # Construire le chemin correspondant dans le deuxième dossier
        chemin2 = chemin1.replace(dossier1, dossier2, 1)
        
        # Vérifier les fichiers
        for fichier in fichiers1:
            fichier1 = os.path.join(chemin1, fichier)
            fichier2 = os.path.join(chemin2, fichier)
            
            if not os.path.exists(fichier2) or not filecmp.cmp(fichier1, fichier2, shallow=False):
                return False
        
        # Vérifier les sous-dossiers
        for sous_dossier in sous_dossiers1:
            sous_dossier2 = os.path.join(chemin2, sous_dossier)
            if not os.path.exists(sous_dossier2):
                return False
    
    # Vérification en partant du dossier2
    for chemin2, sous_dossiers2, fichiers2 in os.walk(dossier2):
        # Construire le chemin correspondant dans le deuxième dossier
        chemin1 = chemin2.replace(dossier2, dossier1, 1)
        
        # Vérifier les fichiers
        for fichier in fichiers2:
            fichier2 = os.path.join(chemin2, fichier)
            fichier1 = os.path.join(chemin1, fichier)
            
            if not os.path.exists(fichier1) or not filecmp.cmp(fichier2, fichier1, shallow=False):
                return False
        
        # Vérifier les sous-dossiers
        for sous_dossier in sous_dossiers2:
            sous_dossier1 = os.path.join(chemin1, sous_dossier)
            if not os.path.exists(sous_dossier1):
                return False
    
    return True


def archivage(liste_mandats, chemin_archive):
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
    
    if not os.path.exists(chemin_archive):
        log += f"Erreur : Le chemin d'archivage spécifié n'existe pas : {chemin_archive}\n"

        # Écriture du log dans le fichier
        with open(log_file, "w") as f:
            f.write(log)

        raise FileNotFoundError(f"Le chemin spécifié n'existe pas : {chemin_archive}")

    # Liste des affaires déjà visitées pour stocker le zip
    affaires = {dossier_mandat[1].split("\\")[-2]: "" for dossier_mandat in liste_mandats}
    # Liste des mandats pour lesquels l'archivage a échoué
    mandats_non_archives = []

    # Parcours des mandats
    for mandat, dossier_mandat in liste_mandats:

        debut_archivage_tot = time.time()

        affaire = dossier_mandat[1].split("\\")[-2]
        chemin = dossier_mandat[1].split(affaire + "\\" + mandat)[0]

        # On récupère le dossier Articles (s'il existe)
        dossier_articles = [dossier for dossier in os.listdir(os.path.join(chemin, affaire)) if dossier.startswith("_")]

        if dossier_articles and len(dossier_articles) == 1:
            dossier_articles = os.path.join(os.path.join(chemin, affaire), dossier_articles[0])
        else:
            dossier_articles = "___xxxxx___" # Pour être sûr que le dossier n'existe pas

        try:

            print(f"Début de l'archivage...")

            # Copie dans le dossier d'archivage:
            log  += f"Copie du mandat {mandat} vers {chemin_archive}{affaire}\\{mandat}\n"

            print(f"\nCopie du mandat {mandat} vers {chemin_archive}{affaire}\\{mandat}...")

            destination = os.path.join(chemin_archive, affaire, mandat)
            os.makedirs(os.path.dirname(destination), exist_ok=True)
            shutil.copytree(dossier_mandat, destination, dirs_exist_ok=True)

            # Log la copie
            log += f"Copie du mandat {mandat} vers {chemin_archive}{affaire}\\{mandat} terminée.\n"
            print(f"\nCopie du mandat {mandat} vers {chemin_archive}{affaire}\\{mandat} terminée.")

            # Vérification stricte
            if not comparer_deux_dossiers(dossier_mandat, destination):
                raise ValueError(f"La copie du mandat {mandat} n'a pas été une copie parfaite.")

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
                if affaires[affaire] == "":

                    # Récupération du nombre de fichiers à archiver
                    nombre_fichiers = sum(len(files) for _, _, files in os.walk(source_zip))

                    with zipfile.ZipFile(destination_zip, 'w', zipfile.ZIP_DEFLATED) as zipf: 

                        for root, dirs, files in os.walk(source_zip):
                            for file in files:
            
                                filepath = os.path.join(root, file)

                                # On enlève le chemin racine pour garder une structure relative
                                arcname = os.path.relpath(filepath, start=source_zip)
                                zipf.write(filepath, arcname)

                    # Log la création de l'archive
                    log += f"Création de l'archive {nom_zip} pour le dossier Articles du mandat {mandat} terminée.\n"
                    print(f"\nCréation de l'archive {nom_zip} pour le dossier Articles du mandat {mandat} terminée.")

                    # On garde le chemin de l'archive pour la copie
                    chemin_zip = destination_zip

                else:

                    shutil.copy(chemin_zip, destination_zip)

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

    return mandats_non_archives


def a_archiver(liste_mandats, chemin):
    """
    Pour tous les mandats dans liste_mandats, sélectionne le dossier du mandat et coche l'attribut "Le dossier est prêt à être archiver"

    :param liste_mandats: Liste des mandats à préparer pour archivage.
    :param chemin: Chemin du dossier contenant les affaires.
    """

    # Préparation du fichier de log
    log_file = os.path.join(chemin, f"log_ajouter_dossier_archivage {datetime.datetime.today().strftime("%d.%m.%Y - %H.%M.%S")}.txt")
    log = f"Début de la préparation des mandats pour archivage - {datetime.datetime.today().strftime("le %d/%m/%Y à %H:%M:%S")}\n\n"

    # Définition de l'attribut à vérifier
    ATTRIBUT_PRET_ARCHIVAGE = 0x20

    if not os.path.exists(chemin):
        log += f"Erreur : Le chemin spécifié n'existe pas : {chemin}\n"

        # Écriture du log dans le fichier
        with open(log_file, "w") as f:
            f.write(log)

        raise FileNotFoundError(f"Le chemin spécifié n'existe pas : {chemin}")

    # Récupère la liste des mandats non archivés
    mandats_non_archives = []

    # Parcours des mandats
    for mandat in liste_mandats:
        dossier_mandat = os.path.join(chemin, mandat.split("-")[0], mandat)

        # Vérification que le dossier du mandat existe
        if not os.path.exists(dossier_mandat):
            # Log l'erreur et continue avec le prochain mandat
            log += f"Le dossier du mandat {mandat} n'existe pas : {dossier_mandat}\n"

            print(f"Le dossier du mandat {mandat} n'existe pas : {dossier_mandat}")
            mandats_non_archives.append(mandat)
            continue

        # Vérification de l'attribut "Le dossier est prêt à être archivé"
        attributs = ctypes.windll.kernel32.GetFileAttributesW(dossier_mandat)

        if not (attributs & ATTRIBUT_PRET_ARCHIVAGE):
            # Ajout de l'attribut "Le dossier est prêt à être archivé"
            ctypes.windll.kernel32.SetFileAttributesW(dossier_mandat, attributs | ATTRIBUT_PRET_ARCHIVAGE)
            
            # Log l'action effectuée
            log += f"Le mandat {mandat} est maintenant prêt à être archivé : {dossier_mandat}\n"

            print(f"Le mandat {mandat} est maintenant prêt à être archivé.")

        else:
            # Log que le mandat est déjà prêt à être archivé
            log += f"Le mandat {mandat} est déjà prêt à être archivé : {dossier_mandat}\n"

            print(f"Le mandat {mandat} est déjà prêt à être archivé.")

    if len(mandats_non_archives) == 0:
        # Log que tous les mandats sont prêts à être archivés
        log += "\nTous les mandats spécifiés sont maintenant prêts à être archivés.\n"

    if mandats_non_archives:
        # Log les mandats qui n'ont pas pu être préparés pour archivage
        log += f"\nLes mandats suivants n'ont pas pu être préparés pour archivage : \n{', '.join(mandats_non_archives)}\n"

    # Écriture du log dans le fichier
    with open(log_file, "w") as f:
        f.write(log)

    return mandats_non_archives


def non_a_archiver(liste_mandats, chemin):
    """
    Pour tous les mandats dans liste_mandats, sélectionne le dossier du mandat et décoche l'attribut "Le dossier est prêt à être archiver"

    :param liste_mandats: Liste des mandats à ne pas préparer pour archivage.
    :param chemin: Chemin du dossier contenant les affaires.
    """

    # Préparation du fichier de log
    log_file = os.path.join(chemin, f"log_ajouter_dossier_archivage {datetime.datetime.today().strftime("%d.%m.%Y - %H.%M.%S")}.txt")
    log = f"Début de la préparation des mandats pour archivage - {datetime.datetime.today().strftime("le %d/%m/%Y à %H:%M:%S")}\n\n"

    # Définition de l'attribut à vérifier
    ATTRIBUT_PRET_ARCHIVAGE = 0x20

    if not os.path.exists(chemin):
        log += f"Erreur : Le chemin spécifié n'existe pas : {chemin}\n"

        # Écriture du log dans le fichier
        with open(log_file, "w") as f:
            f.write(log)

        raise FileNotFoundError(f"Le chemin spécifié n'existe pas : {chemin}")

    # Récupère la liste des mandats non archivés
    mandats_archives = []

    # Parcours des mandats
    for mandat in liste_mandats:
        dossier_mandat = os.path.join(chemin, mandat.split("-")[0], mandat)

        # Vérification que le dossier du mandat existe
        if not os.path.exists(dossier_mandat):
            # Log l'erreur et continue avec le prochain mandat
            log += f"Le dossier du mandat {mandat} n'existe pas : {dossier_mandat}\n"

            print(f"Le dossier du mandat {mandat} n'existe pas : {dossier_mandat}")
            mandats_archives.append(mandat)
            continue

        # Vérification de l'attribut "Le dossier est prêt à être archivé"
        attributs = ctypes.windll.kernel32.GetFileAttributesW(dossier_mandat)

        if not (attributs & ATTRIBUT_PRET_ARCHIVAGE):
            # Ajout de l'attribut "Le dossier est prêt à être archivé"
            ctypes.windll.kernel32.SetFileAttributesW(dossier_mandat, attributs | ATTRIBUT_PRET_ARCHIVAGE)
            
            # Log l'action effectuée
            log += f"Le mandat {mandat} est maintenant prêt à être archivé : {dossier_mandat}\n"

            print(f"Le mandat {mandat} est maintenant prêt à être archivé.")

        else:
            # Log que le mandat est déjà prêt à être archivé
            log += f"Le mandat {mandat} est déjà prêt à être archivé : {dossier_mandat}\n"

            print(f"Le mandat {mandat} est déjà prêt à être archivé.")

    if len(mandats_archives) == 0:
        # Log que tous les mandats sont prêts à être archivés
        log += "\nTous les mandats spécifiés ne sont maintenant plus prêts à être archivés.\n"

    if mandats_archives:
        # Log les mandats qui n'ont pas pu être préparés pour archivage
        log += f"\nLes mandats suivants sont toujours prêts pour archivage : \n{', '.join(mandats_archives)}\n"

    # Écriture du log dans le fichier
    with open(log_file, "w") as f:
        f.write(log)

    return mandats_archives