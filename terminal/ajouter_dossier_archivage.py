import os
import ctypes
import argparse
import datetime


def pret_a_archiver(liste_mandats, chemin):
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


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="""
    Pour tous les mandats dans liste_mandats, sélectionne le dossier du mandat et coche l'attribut "Le dossier est prêt à être archiver"
                                     
    :param liste_mandats: Liste des mandats à préparer pour archivage.
    :param chemin: Chemin du dossier contenant les affaires (par défaut H:\\).
                                     
    Usage : python ajouter_dossier_archivage.py mandat1 mandat2 ... mandatN [chemin]
    """)
    parser.add_argument("liste_mandats", type=str, nargs='+', help="Liste des mandats à préparer pour archivage")
    parser.add_argument("--chemin", type=str, default="H:\\", help="Chemin du dossier contenant les affaires (par défaut H:\\)")
    args = parser.parse_args()

    mandats_non_archives = pret_a_archiver(args.liste_mandats, args.chemin)

    if len(mandats_non_archives) > 0:
        print("Certains mandats n'ont pas pu être préparés pour archivage :")
        for mandat in mandats_non_archives:
            print(f"- {mandat}")
    else:
        print("Tous les mandats spécifées sont maintenant prêts à être archivés.")
