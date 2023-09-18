import tkinter as tk
from tkinter import filedialog
import threading
import logging
logging.getLogger().setLevel(logging.DEBUG)
from wordRapport import Rapporteur


class InterfaceGraphique:
    def __init__(self):
        self.fenetre = tk.Tk()
        self.fenetre.title("Création rapport")
        self.fenetre.geometry("800x600")
        self.fichier_selectionne = tk.StringVar()
        self.etat_case = tk.BooleanVar()
        self.case_var = tk.BooleanVar()  # Ajout de la variable case_var

        self.creer_interface()  # Déplacez la création de l'interface ici

    def creer_interface(self):
        self.nom_entree = tk.Entry(self.fenetre)  # Champ de saisie de nom
        self.nom_entree.pack(pady=20)

        bouton = tk.Button(self.fenetre, text="Créer le rapport", command=self.afficher_nom)
        bouton.pack(pady=10)

        case = tk.Checkbutton(self.fenetre, text="Présence d'une référence", variable=self.case_var, command=self.case_cocher)
        case.pack()

        trouver_fichier_bouton = tk.Button(self.fenetre, text="Trouver un fichier", command=self.trouver_fichier)
        trouver_fichier_bouton.pack(pady=20)

        fichier_label = tk.Label(self.fenetre, textvariable=self.fichier_selectionne, wraplength=400)
        fichier_label.pack()

        case_label = tk.Label(self.fenetre, textvariable=self.etat_case, wraplength=400)
        case_label.pack()

    def trouver_fichier(self):
        chemin_fichier = filedialog.askopenfilename()
        self.fichier_selectionne.set(chemin_fichier)
    
    def get_fichier_selectionne(self):
        return self.fichier_selectionne.get()

    def case_cocher(self):
        self.etat_case.set(self.case_var.get())

    def afficher_nom(self):
        nom_saisi = self.nom_entree.get()
        fichier_saisi = self.get_fichier_selectionne()
        case_cochee = self.etat_case.get()

        # Lancer la fonction sur un nouveau thread
        thread = threading.Thread(target=self.fonction_sur_thread, args=(case_cochee, nom_saisi, fichier_saisi))
        thread.start()

    def fonction_sur_thread(self, case_cochee, nom, fichier):
        try:
            # La logique que vous souhaitez exécuter sur le nouveau thread
            ref_txt = "oui"
            if not case_cochee : ref_txt = "non"
            logging.info(f"Présence d'une référence : {ref_txt}")
            logging.info(f"Nom du rapport : {nom}")
            logging.info(f"Fichier sélectionné : {fichier}")
            rapporteur = Rapporteur(nom)
            rapporteur.load_from_excel(fichier)

            ecart_moyen = False
            ecart_rel = False
            if case_cochee:
                ecart_moyen = True
                ecart_rel = True

            rapporteur.rapporter(intro=True,limite_Q=True,ecart_rel=ecart_rel,taux_de_fonc=True,ecart_moyen=ecart_moyen,stats_entre_capteur=True)
            logging.info("Fin de la création du rapport")
        except Exception as e:
            logging.debug(e)
            logging.error('Une erreur est survenue lors de la création du rapport')

    def run(self):
        self.fenetre.mainloop()

if __name__ == "__main__":
    app = InterfaceGraphique()
    app.run()
