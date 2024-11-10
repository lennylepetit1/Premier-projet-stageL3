# firststageproject
import pandas as pd
import tkinter as tk
from tkinter import simpledialog
import numpy as np
from tkinter import messagebox

class Questionnaire:
    def __init__(self, questions, propositions, categories, domain_names, facettes):
        self.questions = questions
        self.propositions = propositions
        self.categories = categories
        self.domain_names = domain_names
        self.facettes = facettes
        self.num_domains = len(domain_names)
        self.scores = [0] * self.num_domains
        self.question_counts = [0] * self.num_domains

        # Initialiser les scores des facettes
        self.facette_scores = {domain: [0] * len(facettes[domain]) for domain in domain_names}
        self.facette_counts = {domain: [0] * len(facettes[domain]) for domain in domain_names}

    def process_responses(self, responses):
        for question_num, response in enumerate(responses):
            domain = question_num % self.num_domains
            category = self.questions[question_num][1]  # Extraire la categorie de la question
            domain_name = self.questions[question_num][2]
            facette_index = self.questions[question_num][3]  # Extraire l'index de la facette

            if 1 <= response <= len(self.propositions):
                if category == "plus":
                    score = response
                else:  # "minus" category
                    score = 6 - response
                self.scores[domain] += score
                self.question_counts[domain] += 1

                # Ajouter le score a la facette correspondante
                self.facette_scores[domain_name][facette_index] += score
                self.facette_counts[domain_name][facette_index] += 1

    def display_scores(self):
        root = tk.Tk()
        root.title("Resultats du Questionnaire")
        canvas = tk.Canvas(root, width=800, height=800, bg="white")
        canvas.pack()

        for i, score in enumerate(self.scores):
            avg_score = score / self.question_counts[i]
            message = "Le score est eleve" if avg_score > 3 else "Le score est faible"
            bubble_text = f"{self.domain_names[i]}: {score} points\nMoyenne: {avg_score:.2f}\n{message}"
            x = 100 + (i % 2) * 300  # Position X
            y = 100 + (i // 2) * 200  # Position Y
            self.create_bubble(canvas, x, y, bubble_text, i)  # Utilisation de create_bubble

        root.mainloop()

    def create_bubble(self, canvas, x, y, text, domain_index):
        bubble = canvas.create_oval(x, y, x + 200, y + 200, fill="lightblue", tags=f"bubble_{domain_index}")
        text = canvas.create_text(x + 100, y + 100, text=text, font=("Helvetica", 12), width=180)
        canvas.tag_bind(f"bubble_{domain_index}", "<Button-1>", lambda e, idx=domain_index: self.display_facettes(canvas, idx))

    def display_facettes(self, canvas, domain_index):
        facettes_window = tk.Toplevel()
        facettes_window.title(f"Facettes de {self.domain_names[domain_index]}")
        facettes_canvas = tk.Canvas(facettes_window, width=800, height=800, bg="white")
        facettes_canvas.pack()

        center_x, center_y = 400, 400  # Centre de la fenetre

        # Dessiner la bulle pour le domaine principal
        domain_text = f"{self.domain_names[domain_index]}"
        facettes_canvas.create_oval(center_x - 100, center_y - 100, center_x + 100, center_y + 100, fill="lightblue")
        facettes_canvas.create_text(center_x, center_y, text=domain_text, font=("Helvetica", 16))

        # Calcul des positions pour les facettes
        num_facettes = len(self.facettes[self.domain_names[domain_index]])
        radius = 250
        facet_radius = 100
        for i, facette in enumerate(self.facettes[self.domain_names[domain_index]]):
            angle = 2 * np.pi * i / num_facettes
            x = center_x + radius * np.cos(angle) - facet_radius
            y = center_y + radius * np.sin(angle) - facet_radius

            # Calculer le score moyen de la facette
            if self.facette_counts[self.domain_names[domain_index]][i] > 0:
                avg_score = self.facette_scores[self.domain_names[domain_index]][i] / self.facette_counts[self.domain_names[domain_index]][i]
            else:
                avg_score = 0

            facette_text = f"{facette}\nScore: {avg_score:.2f}"
            facettes_canvas.create_oval(x, y, x + 2 * facet_radius, y + 2 * facet_radius, fill="lightgreen")
            facettes_canvas.create_text(x + facet_radius, y + facet_radius, text=facette_text, font=("Helvetica", 12))

        facettes_window.mainloop()

def read_responses_from_excel(file_path):
    # Creer une fenetre Tkinter pour demander le nom de la colonne
    root = tk.Tk()
    root.withdraw()  # Cacher la fenetre principale

    column_name = simpledialog.askstring("Nom de la colonne", "Entrez le nom de la colonne contenant les reponses:")
    root.destroy()  # Detruire la fenetre apres avoir obtenu le nom de la colonne

    if column_name:
        df = pd.read_excel(file_path)
        # Verifier si le nom de colonne est un nombre et le convertir en chaine de caracteres
        if column_name.isdigit():
            column_name = int(column_name)

        if column_name in df.columns:
            responses = df[column_name].tolist()
            return responses
        else:
            messagebox.showerror("Erreur", f"La colonne '{column_name}' n'existe pas dans le fichier Excel.")
            exit()
    else:
        messagebox.showerror("Erreur", "Aucun nom de colonne n'a ete entre.")
        exit()

questions = [
    ("Je m'inquiete a propos de choses", "plus", "Nevrosisme", 0),
    ("Je me fais des amis facilement", "plus", "Extraversion", 0),
    ("J'ai une imagination debordante", "plus", "Ouverture a l'experience", 0),
    ("Je fais confiance aux autres", "plus", "Agreabilite", 0),
    ("Je complete les taches avec succes", "plus", "Conscience", 0),
    ("Je me mets en colere facilement", "plus", "Nevrosisme", 1),
    ("J'adore les grandes fetes", "plus", "Extraversion", 1),
    ("Je crois en l'importance de l'art", "plus", "Ouverture a l'experience", 1),
    ("J'utilise les autres pour mes propres fins", "minus", "Agreabilite", 1),
    ("J'aime faire du rangement", "plus", "Conscience", 1),
    ("Je me sens souvent triste", "plus", "Nevrosisme", 2),
    ("J'aime prendre en charge", "plus", "Extraversion", 2),
    ("Je vis mes emotions intensement", "plus", "Ouverture a l'experience", 2),
    ("J'adore aider les autres", "plus", "Agreabilite", 2),
    ("Je tiens mes promesses", "plus", "Conscience", 2),
    ("J'eprouve de la difficulte a aborder les autres", "plus", "Nevrosisme", 3),
    ("Je suis toujours occupe", "plus", "Extraversion", 3),
    ("Je prefere la variete a  la routine", "plus", "Ouverture a l'experience", 3),
    ("J'adore une bonne bagarre", "minus", "Agreabilite", 3),
    ("Je travaille dur", "plus", "Conscience", 3),
    ("Je fais des exces", "plus", "Nevrosisme", 4),
    ("J'adore les sensations fortes", "plus", "Extraversion", 4),
    ("J'adore lire des documents stimulants", "plus", "Ouverture a l'experience", 4),
    ("Je crois etre meilleur que les autres", "minus", "Agreabilite", 4),
    ("Je suis toujours prepare", "plus", "Conscience", 4),
    ("Je panique facilement", "plus", "Nevrosisme", 5),
    ("Je respire la joie", "plus", "Extraversion", 5),
    ("J'ai tendance a promouvoir des valeurs sociales traditionnelles", "plus", "Ouverture a l'experience", 5),
    ("Je sympathise avec les sans-abri", "plus", "Agreabilite", 5),
    ("Je me lance dans des choses sans reflechir", "minus", "Conscience", 5),
    ("Je crains le pire", "plus", "Nevrosisme", 0),
    ("Je me sens confortable avec les gens", "plus", "Extraversion", 0),
    ("J'aime me perdre dans mes idees", "plus", "Ouverture a l'experience", 0),
    ("Je crois que les autres ont de bonnes intentions", "plus", "Agreabilite", 0),
    ("J'excelle dans ce que je fais", "plus", "Conscience", 0),
    ("Je suis irritable facilement", "plus", "Nevrosisme", 1),
    ("Je parle a beaucoup de personnes dans les fetes", "plus", "Extraversion", 1),
    ("Je vois une beaute dans les choses que d'autres pourraient ne pas remarquer", "plus", "Ouverture a l'experience", 1),
    ("Je triche pour avancer", "minus", "Agreabilite", 1),
    ("J'oublie souvent de remettre les choses a leur place", "minus", "Conscience", 1),
    ("Je n'ai pas beaucoup d'estime pour moi", "plus", "Nevrosisme", 2),
    ("J'essaie de diriger les autres", "plus", "Extraversion", 2),
    ("Je ressens les emotions des autres", "plus", "Ouverture a l'experience", 2),
    ("Je me preoccupe des autres", "plus", "Agreabilite", 2),
    ("Je dis la verite", "plus", "Conscience", 2),
    ("J'ai peur d'attirer l'attention", "plus", "Nevrosisme", 3),
    ("Je suis toujours en mouvement", "plus", "Extraversion", 3),
    ("Je prefere m'en tenir aux choses connues", "minus", "Ouverture a l'experience", 3),
    ("Je crie apres les gens", "minus", "Agreabilite", 3),
    ("Je fais plus que ce que l'on attend de moi", "plus", "Conscience", 3),
    ("Je fais rarement des exces", "minus", "Nevrosisme", 4),
    ("Je recherche l'aventure", "plus", "Extraversion", 4),
    ("J'evite les discussions philosophiques", "minus", "Ouverture a l'experience", 4),
    ("J'ai une tres grande estime de moi-meme", "minus", "Agreabilite", 4),
    ("Je realise mes objectifs", "plus", "Conscience", 4),
    ("Je me sens depasse par les evenements", "plus", "Nevrosisme", 5),
    ("J'ai beaucoup de plaisir", "plus", "Extraversion", 5),
    ("Je crois qu'il n'y a pas de bon ou de mauvais absolu", "plus", "Ouverture a l'experience", 5),
    ("J'ai de la sympathie pour les gens plus demunis que moi", "plus", "Agreabilite", 5),
    ("Je prends des decisions impulsives", "minus", "Conscience", 5),
    ("J'ai peur de beaucoup de choses", "plus", "Nevrosisme", 0),
    ("J'evite le contact avec les autres", "minus", "Extraversion", 0),
    ("J'adore revasser", "plus", "Ouverture a l'experience", 0),
    ("J'ai confiance en ce que les gens disent", "plus", "Agreabilite", 0),
    ("Je gere les taches facilement", "plus", "Conscience", 0),
    ("Je perds patience", "plus", "Nevrosisme", 1),
    ("Je prefere etre seul", "minus", "Extraversion", 1),
    ("Je n'aime pas la poesie", "minus", "Ouverture a l'experience", 1),
    ("Je profite des autres", "minus", "Agreabilite", 1),
    ("Je laisse ma chambre en desordre", "minus", "Conscience", 1),
    ("Il m'arrive souvent de broyer du noir", "plus", "Nevrosisme", 2),
    ("Je prends le controle des choses", "plus", "Extraversion", 2),
    ("Je suis rarement sensible a  mes reactions emotives", "minus", "Ouverture a l'experience", 2),
    ("Je suis indifferent aux sentiments des autres", "minus", "Agreabilite", 2),
    ("J'enfreins les regles", "minus", "Conscience", 2),
    ("Je me sens seulement a l'aise avec mes amis", "plus", "Nevrosisme", 3),
    ("Je fais beaucoup de choses pendant mon temps libre", "plus", "Extraversion", 3),
    ("Je n'aime pas les changements", "minus", "Ouverture a l'experience", 3),
    ("J'insulte les gens", "minus", "Agreabilite", 3),
    ("Je fais juste assez de travail pour m'en sortir", "minus", "Conscience", 3),
    ("Je resiste facilement a la tentation", "minus", "Nevrosisme", 4),
    ("J'agis de maniere effrenee", "plus", "Extraversion", 4),
    ("J'ai des difficultes a  comprendre les idees abstraites", "minus", "Ouverture a l'experience", 4),
    ("J'ai une tres bonne opinion de moi-meme", "minus", "Agreabilite", 4),
    ("Je perds mon temps", "minus", "Conscience", 4),
    ("Je me sens incapable de gerer les choses", "plus", "Nevrosisme", 5),
    ("J'adore la vie", "plus", "Extraversion", 5),
    ("J'ai tendance a promouvoir des valeurs sociales liberales", "minus", "Ouverture a l'experience", 5),
    ("Je ne suis pas interesse par les problemes des autres", "minus", "Agreabilite", 5),
    ("Je me precipite dans l'action", "minus", "Conscience", 5),
    ("Je deviens stresse facilement", "plus", "Nevrosisme", 0),
    ("Je garde les autres a distance", "minus", "Extraversion", 0),
    ("J'aime me perdre dans mes pensees", "plus", "Ouverture a l'experience", 0),
    ("Je ne fais pas confiance aux autres", "minus", "Agreabilite", 0),
    ("Je sais comment faire avancer les choses", "plus", "Conscience", 0),
    ("Je ne suis pas facilement agace", "minus", "Nevrosisme", 1),
    ("J'evite les foules", "minus", "Extraversion", 1),
    ("Je n'apprecie pas visiter des musees d'art", "minus", "Ouverture a l'experience", 1),
    ("J'entrave les plans des autres", "minus", "Agreabilite", 1),
    ("Je laisse trainer les choses", "minus", "Conscience", 1),
    ("Je suis a l'aise avec moi-meme", "minus", "Nevrosisme", 2),
    ("J'attends que les autres prennent les devants", "minus", "Extraversion", 2),
    ("Je ne comprends pas les gens emotifs", "minus", "Ouverture a l'experience", 2),
    ("Je ne prends pas de temps pour les autres", "minus", "Agreabilite", 2),
    ("Je ne tiens pas mes promesses", "minus", "Conscience", 2),
    ("Je ne suis pas derange par les situations sociales difficiles", "minus", "Nevrosisme", 3),
    ("J'aime prendre ca mollo", "minus", "Extraversion", 3),
    ("Je suis attache aux methodes conventionnelles", "minus", "Ouverture a l'experience", 3),
    ("Je m'en prends aux autres", "minus", "Agreabilite", 3),
    ("Je mets peu de temps et d'efforts dans mon travail", "minus", "Conscience", 3),
    ("Je suis capable de controler mes envies", "minus", "Nevrosisme", 4),
    ("J'agis de facon temeraire", "plus", "Extraversion", 4),
    ("Je ne suis pas interesse par les discussions theoriques", "minus", "Ouverture a l'experience", 4),
    ("Je me vante de mes vertus", "minus", "Agreabilite", 4),
    ("J'ai de la difficulte a commencer les taches", "minus", "Conscience", 4),
    ("Je demeure calme sous pression", "minus", "Nevrosisme", 5),
    ("Je regarde le bon cote de la vie", "plus", "Extraversion", 5),
    ("Je crois que nous devrions etre severes a propos des crimes", "minus", "Ouverture a l'experience", 5),
    ("Je tente de ne pas penser aux gens dans le besoin", "minus", "Agreabilite", 5),
    ("J'agis sans penser", "minus", "Conscience", 5)]

propositions = ["Pas du tout d'accord", "Plutot pas d'accord", "Ni d'accord ni pas d'accord", "Plutot d'accord","Tout a fait d'accord" ]

categories = ["plus", "minus"]

domain_names = ["Nevrosisme", "Extraversion", "Ouverture a l'experience", "Agreabilite", "Conscience"] 

facettes = {
    "Nevrosisme": ["Anxiete", "Colere", "Depression", "Conscience de soi", "Immoderation", "Vulnerabilite"],
    "Extraversion": ["Bienveillance", "Conformisme", "Assurance", "Activite", "Enthousiasme", "Gaiete"],
    "Ouverture a l'experience": ["Imagination", "Sens Artistique", "Affection", "Aventure", "Intellect", "Liberalisme"],
    "Agreabilite": ["Confiance", "Moralite", "Altruisme", "Cooperation", "Modestie", "Sympathie"],
    "Conscience": ["Auto_efficacite", "Organisation", "Respect", "Perseverance", "Autodiscipline", "Prudence"]
}

file_path = "C:/Users/lenny/Documents/Python Project/Excel pour le stage/Fichier excel du test.xlsx"  # Remplacez par le chemin correct de votre fichier Excel

responses = read_responses_from_excel(file_path)
questionnaire = Questionnaire(questions, propositions, categories, domain_names, facettes)
questionnaire.process_responses(responses)
questionnaire.display_scores()
