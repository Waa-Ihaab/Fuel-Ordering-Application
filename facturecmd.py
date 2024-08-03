import re
import shutil
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
from datetime import datetime
import os
import hashlib
import openpyxl
import sqlite3
from tkinter import Menu
from tkcalendar import DateEntry



# Liste pour stocker les commandes
commandes = []
numero_facture = 1


def creer_table_sites():
    # Créer une table pour stocker les informations des sites (si elle n'existe pas déjà)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sites (
            Site INTEGER PRIMARY KEY AUTOINCREMENT,
            nom TEXT,
            prix_litre REAL
        )
    ''')
    conn.commit()


# Créer une connexion à la base de données SQLite
conn = sqlite3.connect('enr.db')
cursor = conn.cursor()
def creer_table_historique_commandes():
    # Créer une table pour stocker les données historiques des commandes (si elle n'existe pas déjà)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historique_commandes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site TEXT,
            periode_du TEXT,
            periode_au TEXT,
            quantite REAL,
            prix_litre REAL,
            montant_ttc REAL
        )
    ''')
    conn.commit()

# Créer la table 'historique_commandes' si elle n'existe pas déjà
creer_table_historique_commandes()

def enregistrer_commande_historique(site, periode_du, periode_au, quantite, prix_litre, montant_ttc):
    # Enregistrer la commande dans la table historique_commandes
    cursor.execute('INSERT INTO historique_commandes (site, periode_du, periode_au, quantite, prix_litre, montant_ttc) VALUES (?, ?, ?, ?, ?, ?)',
                   (site, periode_du, periode_au, quantite, prix_litre, montant_ttc))
    conn.commit()

# Créer la table 'sites' si elle n'existe pas déjà
creer_table_sites()
def obtenir_sites():
    # Récupérer la liste des sites à partir de la base de données
    cursor.execute('SELECT nom FROM sites')
    sites = cursor.fetchall()
    return [site[0] for site in sites]

def ajouter_site(nom_site, prix_litre):
    # Vérifier d'abord si le site existe déjà dans la base de données
    sites_existant = obtenir_sites()
    if nom_site in sites_existant:
        pass
    else:
        # Ajouter un nouveau site avec son prix du litre dans la table sites
        cursor.execute('INSERT INTO sites (nom, prix_litre) VALUES (?, ?)', (nom_site, prix_litre))
        conn.commit()

# Ajouter les sites "Nador" et "Mohamedia" avec leurs prix respectifs
ajouter_site('Nador', 15.44)
ajouter_site('Mohamedia', 14.88)
def obtenir_prix_litre(nom_site):
    # Récupérer le prix du litre d'un site à partir de la base de données
    cursor.execute('SELECT prix_litre FROM sites WHERE nom = ?', (nom_site,))
    prix_litre = cursor.fetchone()
    return prix_litre[0] if prix_litre else None


def obtenir_sites():
    # Récupérer la liste des sites à partir de la base de données
    cursor.execute('SELECT nom FROM sites')
    sites = cursor.fetchall()
    return [site[0] for site in sites]


def obtenir_dernier_numero_facture():
    # Obtenez le chemin complet du dossier Téléchargements pour l'utilisateur actuel
    dossier_telechargements = os.path.join(os.path.expanduser("~"), "Downloads")
    # Recherchez les fichiers facture existants dans le dossier Téléchargements
    fichiers_facture = [fichier for fichier in os.listdir(dossier_telechargements) if re.match(r'^facture_\d{4}\.xlsx$', fichier)]
    if fichiers_facture:
        # Triez les fichiers par numéro de facture (en supposant qu'ils sont nommés avec le format facture_XXXX.xlsx)
        fichiers_facture.sort(key=lambda x: int(re.search(r'^facture_(\d{4})\.xlsx$', x).group(1)))
        # Obtenez le numéro de facture le plus élevé
        dernier_numero = int(re.search(r'^facture_(\d{4})\.xlsx$', fichiers_facture[-1]).group(1))
        return dernier_numero
    else:
        # Aucun fichier facture trouvé, commencez à partir du numéro 1
        return 1


def enregistrer_facture(factures):
    global numero_facture
    numero_facture_str = str(numero_facture).zfill(4)  # Formatage du numéro de facture avec des zéros à gauche
    numero_facture += 1

    # Récupérer la date d'enregistrement
    date_enregistrement = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Combine 'Site', 'Période du', and 'Au' into a single 'Designation' column
    factures = [(f"Site:{site}   \nPériode du : {periode_du}   Au : {periode_au}", float(quantite), prix_litre, montant_ttc) for (site, periode_du, periode_au, quantite, prix_litre, montant_ttc) in factures]

    df = pd.DataFrame(factures, columns=['Designation', 'Quantité (litres)', 'Prix TTC', 'Montant TTC'])
    with pd.ExcelWriter(f'facture_{numero_facture_str}.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook  = writer.book
        sheet = writer.sheets['Sheet1']
        sheet.column_dimensions['A'].width = 80
    # Calculer les totaux des quantités en litres et des montants TTC
    total_quantites = sum(row[1] for row in factures)
    total_montant_ttc = sum(float(row[3]) for row in factures)
    # Ajouter une ligne pour afficher la TVA 10%
    
    # Ajouter une ligne vide
    df.loc[len(df)] = ['', '', '', '']

    df.loc[len(df)] = ['Total:', total_quantites, '', total_montant_ttc]
    # Ajouter une ligne vide
    df.loc[len(df)] = ['', '', '', '']

    # Calculer le montant HT
    montant_ht = total_montant_ttc / 1.1

    # Calculer la TVA 10%
    tva_10 = total_montant_ttc - montant_ht

    # Ajouter une ligne pour afficher le montant HT
    df.loc[len(df)] = ['Montant HT:', '', '', montant_ht]

    # Ajouter une ligne pour afficher la TVA 10%
    df.loc[len(df)] = ['TVA 10%:', '', '', tva_10]

    # Ajouter une ligne pour afficher le montant TTC
    df.loc[len(df)] = ['Montant TTC:', '', '', total_montant_ttc]

    # Ajouter une ligne vide
    df.loc[len(df)] = ['', '', '', '']

    # Ajouter une ligne pour afficher le numéro de facture
    df.loc[len(df)] = ['Numéro de Facture:', '', '', numero_facture_str]

    # Ajouter une ligne pour afficher la date d'enregistrement de la facture
    df.loc[len(df)] = ['Date:', '', '', date_enregistrement]
    # # Save the workbook
    # Save the DataFrame to an Excel file
    with pd.ExcelWriter(f'facture_{numero_facture_str}.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        
        # Access the workbook and sheet
        workbook = writer.book
        sheet = writer.sheets['Sheet1']

        # Adjust column sizes (example: set the width of column A to 40)
        sheet.column_dimensions['A'].width = 40  # Adjust the width as needed
        
        # Access the row dimension of column A and set the height
        sheet.row_dimensions[1].height = 20 # Adjust the height as needed
        
        sheet.column_dimensions['B'].width = 20  # Adjust the width as needed
        sheet.column_dimensions['C'].width = 20  # Adjust the width as needed
        sheet.column_dimensions['D'].width = 20  # Adjust the width as needed
        
        # Add more column adjustments as needed
        
        # Save the workbook
        workbook.save(f'facture_{numero_facture_str}.xlsx')

    # df.to_excel(f'facture_{numero_facture_str}.xlsx', index=False)
    
    # Afficher le contenu du tableau Excel
    messagebox.showinfo("Tableau Excel", df.to_string(index=False))
    # Obtenez le chemin complet du dossier Téléchargements pour l'utilisateur actuel
    dossier_telechargements = os.path.join(os.path.expanduser("~"), "Downloads")
    # Déplacez le fichier vers le dossier Téléchargements
    fichier_destination = os.path.join(dossier_telechargements, f'facture_{numero_facture_str}.xlsx')
    shutil.move(f'facture_{numero_facture_str}.xlsx', fichier_destination)

    messagebox.showinfo("Facture Téléchargée", f"Facture {numero_facture_str} téléchargée avec succès.\nTelechargement : {fichier_destination}")
    # messagebox.showinfo("Facture Téléchargée", f"Facture {numero_facture_str} téléchargée avec succès.")





# Définir les widgets de la fenêtre de commande en tant que variables globales
site_combobox = None
periode_du_entry = None
periode_au_entry = None
quantite_entry = None

def creer_barre_menu_commande():
    menu_bar_commande = Menu(command_window)

    # Option "Aide" avec la fonctionnalité d'aide pour la fenêtre de commande
    def afficher_aide_commande():
        aide_texte = "Bienvenue dans l'Aide de l'application de commande de carburant.\n\n"
        aide_texte += "Pour passer une commande :\n"
        aide_texte += "1. Sélectionnez le site dans le menu déroulant.\n"
        aide_texte += "2. Entrez la période du et la période au au format DD-MM-YYYY.\n"
        aide_texte += "3. Entrez la quantité en litres.\n"
        aide_texte += "4. Cliquez sur le bouton 'Confirmer la Commande' pour enregistrer la commande.\n"
        aide_texte += "5. Le montant total TTC sera calculé et affiché.\n"
        messagebox.showinfo("À propos", aide_texte)

    menu_aide_commande = Menu(menu_bar_commande, tearoff=0)
    menu_aide_commande.add_command(label="Aide", command=afficher_aide_commande)
    menu_bar_commande.add_cascade(label="Aide", menu=menu_aide_commande)

    # Option "À propos" avec la fonctionnalité d'affichage des informations sur la fenêtre de commande
    def afficher_a_propos_commande():
        a_propos_texte = "Application de Commande de Carburant\n\n"
        a_propos_texte += "Développée par Kari Ihab\n"
        a_propos_texte += "Version 0.5\n"
        messagebox.showinfo("À propos", a_propos_texte)


    menu_a_propos_commande = Menu(menu_bar_commande, tearoff=0)
    menu_a_propos_commande.add_command(label="À propos", command=afficher_a_propos_commande)
    menu_bar_commande.add_cascade(label="À propos", menu=menu_a_propos_commande)

    # Option "Paramètres" avec la fonctionnalité d'affichage d'un message indiquant que les paramètres seront bientôt disponibles
    def afficher_parametres_commande():
        messagebox.showinfo("Paramètres", "Les paramètres de la fenêtre de commande seront bientôt disponibles.")

    menu_parametres_commande = Menu(menu_bar_commande, tearoff=0)
    menu_parametres_commande.add_command(label="Paramètres", command=afficher_parametres_commande)
    menu_bar_commande.add_cascade(label="Paramètres", menu=menu_parametres_commande)

    # Configurer la barre de menu pour la fenêtre de commande
    command_window.config(menu=menu_bar_commande)


def commande_action():
    # Masquer la fenêtre principale (application)
    root.withdraw()
    global command_window
    command_window = tk.Toplevel(root)
    command_window.title("Commande")

    # Définir un schéma de couleurs plus attrayant visuellement
    command_window.configure(bg="#E1F0FF")  # Fond bleu clair

    # Définir un schéma de couleurs plus attrayant visuellement
    command_window.configure(bg="#E1F0FF")  # Fond bleu clair

    # Créer la barre de menu pour la fenêtre de commande
    creer_barre_menu_commande()

    # Obtenir les dimensions de l'écran
    screen_width = command_window.winfo_screenwidth()
    screen_height = command_window.winfo_screenheight()

    # Définir la taille et la position de la fenêtre de commande
    window_width = 600
    window_height = 400
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2

    # Centrer la fenêtre de commande à l'écran
    command_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # Créer une Frame pour contenir les widgets de la fenêtre de commande
    command_frame = tk.Frame(command_window, bg="#E1F0FF")  # Fond bleu clair
    command_frame.pack(expand=True, padx=20, pady=20)

    # Liste des sites disponibles (à récupérer depuis la base de données)
    sites = obtenir_sites()

    # Variables pour stocker les valeurs des champs de saisie
    selected_site = tk.StringVar()
    periode_du_value = tk.StringVar()
    periode_au_value = tk.StringVar()
    quantite_value = tk.StringVar()

    site_label = tk.Label(command_frame, text="Sélectionnez le site :", font=("Helvetica", 14), bg="#E1F0FF")
    site_label.grid(row=0, column=0, padx=10, pady=10)
    site_combobox = ttk.Combobox(command_frame, values=sites, textvariable=selected_site, font=("Helvetica", 14), state="readonly")
    site_combobox.grid(row=0, column=1, padx=10, pady=10)

    periode_du_label = tk.Label(command_frame, text="Période du :", font=("Helvetica", 14), bg="#E1F0FF")  # Fond bleu clair
    periode_du_label.grid(row=1, column=0, padx=10, pady=10)

    periode_du_entry = DateEntry(command_frame, font=("Helvetica", 14), date_pattern="dd-mm-yyyy", textvariable=periode_du_value)
    periode_du_entry.grid(row=1, column=1, padx=10, pady=10)


     # Ajouter un label pour la période au
    periode_au_label = tk.Label(command_frame, text="Au :", font=("Helvetica", 14), bg="#E1F0FF")  # Fond bleu clair
    periode_au_label.grid(row=2, column=0, padx=10, pady=10)

    # Ajouter une saisie pour la période au
    periode_au_entry = DateEntry(command_frame, font=("Helvetica", 14), date_pattern="dd-mm-yyyy", textvariable=periode_au_value)
    periode_au_entry.grid(row=2, column=1, padx=10, pady=10)

    

    # Ajouter un label pour la quantité en litres
    quantite_label = tk.Label(command_frame, text="Quantité en litres :", font=("Helvetica", 14), bg="#E1F0FF")  # Fond bleu clair
    quantite_label.grid(row=3, column=0, padx=10, pady=10)

    # Ajouter une saisie pour la quantité en litres
    quantite_entry = tk.Entry(command_frame, font=("Helvetica", 14), textvariable=quantite_value)
    quantite_entry.grid(row=3, column=1, padx=10, pady=10)

    # def validate_date_format(date_str):
    #     try:
    #         datetime.strptime(date_str, "%d/%m/%Y")
    #         return True
    #     except ValueError:
    #         return False

    def validate_positive_number(input_str):
        try:
            number = float(input_str)
            if number > 0:
                return True
            else:
                return False
        except ValueError:
            return False
    
    def show_confirmation():
        site = selected_site.get()
        periode_du = periode_du_entry.get_date().strftime("%d-%m-%Y")
        periode_au = periode_au_entry.get_date().strftime("%d-%m-%Y")
        quantite = quantite_entry.get()

        # Récupérer le prix du litre à partir de la base de données
        prix_litre = obtenir_prix_litre(site)

        # Vérifications de validation des données
        if not site:
            messagebox.showwarning("Champ manquant", "Veuillez sélectionner un site.")
        elif not quantite or not validate_positive_number(quantite):
            messagebox.showwarning("Quantité invalide", "Veuillez entrer une quantité valide (nombre positif).")
        else:
            if prix_litre is not None:
                # Calculer le montant TTC
                montant_ttc = float(quantite) * prix_litre

                # Ajouter la commande à la liste
                commandes.append((site, periode_du, periode_au, quantite, prix_litre, montant_ttc))

                # Enregistrer la commande dans la table historique_commandes
                enregistrer_commande_historique(site, periode_du, periode_au, quantite, prix_litre, montant_ttc)

                # Afficher une boîte de dialogue pour confirmer la commande
                messagebox.showinfo("Confirmation", f"Commande passée avec succès pour le site {site}."
                                                    f"\nPériode du : {periode_du} au : {periode_au}"
                                                    f"\nQuantité en litres : {quantite}"
                                                    f"\nPrix TTC : {prix_litre}"
                                                    f"\nMontant TTC : {montant_ttc}")
            else:
                messagebox.showwarning("Prix non trouvé", f"Le prix du litre pour le site {site} n'est pas disponible.")


    def ajouter_commande():
        # Remettre à zéro les champs
        selected_site.set('')
        periode_du_entry.delete(0, tk.END)
        periode_au_entry.delete(0, tk.END)
        quantite_entry.delete(0, tk.END)

    
    # Ajouter un bouton pour ajouter une nouvelle commande
    add_command_button = tk.Button(command_frame, text="Ajouter une Commande", command=ajouter_commande, font=("Helvetica", 14), bg="#6C757D", fg="#F9F9F9")
    add_command_button.grid(row=5, column=0, columnspan=2, padx=10, pady=10)


    # Ajouter un bouton pour confirmer la commande
    confirm_button = tk.Button(command_frame, text="Confirmer la Commande", command=show_confirmation, font=("Helvetica", 14), bg="#6C757D", fg="#F9F9F9")
    confirm_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

    def retour_commande():
        # Fermer la fenêtre de commande lorsque le bouton "Retour" est cliqué
        command_window.destroy()
        # Réafficher la fenêtre principale (application)
        root.deiconify()

    # Ajouter un bouton pour retourner à la fenêtre principale (application)
    return_button = tk.Button(command_frame, text="Retour à l'Application", command=retour_commande, font=("Helvetica", 14), bg="#6C757D", fg="#F9F9F9")
    return_button.grid(row=6, column=0, columnspan=2, padx=10, pady=10)


def information_action():
    messagebox.showinfo("Information", "Rien pour le moment.")


def exit_action():
    if messagebox.askyesno("Quitter", "Êtes-vous sûr de vouloir quitter l'application ?"):
        root.destroy()


def telecharger_facture():
    enregistrer_facture(commandes)


def afficher_historique_commandes():
    if not commandes:
        messagebox.showinfo("Historique des Commandes", "Aucune commande enregistrée.")
    else:
        commande_str = ""
        for i, (site, periode_du, periode_au, quantite, prix_litre, montant_ttc) in enumerate(commandes, start=1):
            commande_str += f"\nCommande {i}:"
            commande_str += f"\nDesignation : {site}\nPériode du : {periode_du}\nAu : {periode_au}"
            commande_str += f"\nPrix TTC : {prix_litre}"
            commande_str += f"\nMontant TTC : {montant_ttc}\n"
        messagebox.showinfo("Historique des Commandes", commande_str)

def afficher_commandes_enregistrees():
    if not commandes:
        messagebox.showinfo("Commandes Enregistrées", "Aucune commande enregistrée.")
    else:
        commande_str = ""
        for i, (site, periode_du, periode_au, quantite, prix_litre, montant_ttc) in enumerate(commandes, start=1):
            commande_str += f"\nCommande {i}:"
            commande_str += f"\nDesignation : {site}\nPériode du : {periode_du}\nAu : {periode_au}"
            commande_str += f"\nPrix TTC : {prix_litre}"
            commande_str += f"\nMontant TTC : {montant_ttc}\n"
        messagebox.showinfo("Commandes Enregistrées", commande_str)


def afficher_aide():
    aide_texte = "Bienvenue dans l'Aide de l'application de commande de carburant.\n\n"
    aide_texte += "Pour passer une commande :\n"
    aide_texte += "1. Se dirigee vers la fenetre commande pour passer une commande.\n"
    aide_texte += "2. Voirs les informations du societe STG.\n"
    aide_texte += "3. Telecharger la facture en format csv.\n"
    aide_texte += "4. Quittez l'apllication."
    messagebox.showinfo("Aide", aide_texte)

def afficher_a_propos():
    a_propos_texte = "Application de Commande de Carburant\n\n"
    a_propos_texte += "Développée par Kari Ihab\n"
    a_propos_texte += "Version 0.5\n"
    messagebox.showinfo("À propos", a_propos_texte)

def afficher_parametres():
    messagebox.showinfo("Paramètres", "Les paramètres seront bientôt disponibles.")


global root
root = tk.Tk()
root.title("Application")


# Créer une Frame pour contenir les boutons
main_frame = tk.Frame(root, bg="#E1F0FF")  # Fond bleu clair
main_frame.pack(expand=True, padx=20, pady=20)


# Créer la barre de menu
menu_bar = Menu(root)
root.config(menu=menu_bar)

# Ajouter un menu "Aide" avec une option "Aide"
menu_aide = Menu(menu_bar, tearoff=0)
menu_aide.add_command(label="Aide", command=afficher_aide)
menu_bar.add_cascade(label="Aide", menu=menu_aide)

# Ajouter un menu "À propos" avec une option "À propos"
menu_a_propos = Menu(menu_bar, tearoff=0)
menu_a_propos.add_command(label="À propos", command=afficher_a_propos)
menu_bar.add_cascade(label="À propos", menu=menu_a_propos)

# Ajouter un menu "Paramètres" avec une option "Paramètres"
menu_parametres = Menu(menu_bar, tearoff=0)
menu_parametres.add_command(label="Paramètres", command=afficher_parametres)
menu_bar.add_cascade(label="Paramètres", menu=menu_parametres)


# Définir un schéma de couleurs plus attrayant visuellement
root.configure(bg="#E1F0FF")  # Fond bleu clair

# Obtenir les dimensions de l'écran
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Définir la taille et la position de la fenêtre
window_width = 800
window_height = 600
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2

# Centrer la fenêtre à l'écran
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# Créer une Frame pour contenir les boutons
main_frame = tk.Frame(root, bg="#E1F0FF")  # Fond bleu clair
main_frame.pack(expand=True, padx=20, pady=20)

# Ajouter un label de titre
title_label = tk.Label(main_frame, text="Societe de transport du gaz", font=("Helvetica", 24), bg="#E1F0FF")  # Fond bleu clair
title_label.pack(pady=20)

# Ajouter des boutons avec des emojis Unicode pour les icônes
commande_button = tk.Button(main_frame, text="Passer une Commande \n \U0001F4E6", command=commande_action, font=("Helvetica", 14), bg="#6C757D", fg="#F9F9F9")
commande_button.pack(pady=10, fill="x")

# Ajouter un bouton pour afficher l'historique des commandes
historique_button = tk.Button(main_frame, text="Historique des Commandes \n \U0001F4C1", command=afficher_historique_commandes, font=("Helvetica", 14), bg="#6C757D", fg="#F9F9F9")
historique_button.pack(pady=10, fill="x")

telecharger_facture_button = tk.Button(main_frame, text="Télécharger la Facture \n \U0001F4E5", command=telecharger_facture, font=("Helvetica", 14), bg="#6C757D", fg="#F9F9F9")
telecharger_facture_button.pack(pady=10, fill="x")

exit_button = tk.Button(main_frame, text="Quitter \n \U0001F6AA", command=exit_action, font=("Helvetica", 14), bg="#6C757D", fg="#F9F9F9")
exit_button.pack(pady=10, fill="x")

# Démarrer la boucle d'événements principale
root.mainloop()