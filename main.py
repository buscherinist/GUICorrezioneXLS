#per compilare
#pyinstaller --onefile --noconsole main.py
import openpyxl
import tkinter as tk
from tkinter import messagebox  # Importa messagebox per la finestra di conferma
from tkinter import filedialog, scrolledtext, Toplevel

def center_window():
    # Aggiorna le dimensioni della finestra in base ai contenuti
    root.update_idletasks()  # Aggiorna il layout e le dimensioni

    # Ottieni le dimensioni della finestra
    width = root.winfo_width()
    height = root.winfo_height()

    # Ottieni le dimensioni dello schermo
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Calcola la posizione x e y per centrare la finestra
    x = (screen_width // 2) - (width // 2)

    # Ottieni l'altezza della barra delle applicazioni (Windows)
    taskbar_height = 30

    y = (screen_height // 2) - (height // 2) - taskbar_height

    # Imposta la geometria della finestra
    root.geometry(f'{width}x{height}+{x}+{y}')

# Funzione di conferma per uscire dall'applicazione
def exit_app():
    if messagebox.askokcancel("Exit", "Vuoi uscire?"):
        root.quit()


# Funzione per cambiare il colore del pulsante quando il mouse entra
def on_enter(e):
    e.widget['background'] = colore_sfondo_button_on
    e.widget['foreground'] = colore_testo_button_on

# Funzione per ripristinare il colore del pulsante quando il mouse esce
def on_leave(e):
    e.widget['background'] = colore_sfondo_button
    e.widget['foreground'] = colore_testo_button


def mostra_file():
    file_path = "correzione.txt"  # Nome del file nella stessa cartella di main.py
    try:
        with open(file_path, "r") as file:
            contenuto = file.read()

        # Creazione di una finestra secondaria
        finestra_secondaria = Toplevel(root)
        finestra_secondaria.title(f"Contenuto di {file_path}")

        finestra_testo = scrolledtext.ScrolledText(finestra_secondaria, width=50, height=20)
        finestra_testo.pack(expand=True, fill="both")
        finestra_testo.insert("1.0", contenuto)

        # Impedisce l'editing del contenuto
        finestra_testo.config(state="disabled")

    except FileNotFoundError:
        print(f"Il file '{file_path}' non è stato trovato.")

def carica_soluzioni(file_soluzioni):
    # Legge il file di testo con le soluzioni e i punti, organizzati per foglio
    soluzioni = {}
    with open(file_soluzioni, 'r') as file:
        linee = file.readlines()

    foglio_corrente = None
    i = 0
    while i < len(linee):
        line = linee[i].strip()

        # Se la linea è un nome di foglio
        if line.startswith("Foglio"):
            foglio_corrente = line
            soluzioni[foglio_corrente] = {}
            i += 1
            continue

        # Altrimenti, leggiamo cella, formula e punti
        if foglio_corrente:
            cella = line
            formula = linee[i + 1].strip()
            valore_atteso = linee[i + 2].strip()  # Il valore atteso
            punti = int(linee[i + 3].strip())
            soluzioni[foglio_corrente][cella] = (formula, valore_atteso,punti)
            i += 4

    return soluzioni

def controlla_formule(nome_file_excel, soluzioni):
    # Apre il file Excel e controlla le formule (data_only=False per ottenere la formula)
    workbook = openpyxl.load_workbook(nome_file_excel, data_only=False)  # Usa data_only=False per leggere le formule
    punteggio_totale = 0
    risultati = {nome_file_excel: {}}

    for foglio_nome, celle in soluzioni.items():
        foglio = workbook[foglio_nome]
        risultati[nome_file_excel][foglio_nome] = {}

        for cella, (formula_attesa, valore_atteso, punti) in celle.items():
            # Ottieni l'oggetto cella dal foglio
            cella_oggetto = foglio[cella]

            # Ottieni il valore calcolato della cella (risultato della formula)
            valore_cella = cella_oggetto.value
            print(f"Valore della cella (risultato formula): {valore_cella}")

            # Ottieni la formula (se presente)
            formula_cella = None
            if hasattr(cella_oggetto, 'formula'):  # Verifica se la cella ha un attributo 'formula'
                formula_cella = cella_oggetto.formula
                print(f"Formula della cella: {formula_cella}")
            else:
                print(f"La cella {cella} non contiene una formula.")

            # Confronta la formula e il valore
            if formula_cella == formula_attesa and valore_cella == valore_atteso:
                risultati[nome_file_excel][foglio_nome][cella] = f"Formula corretta: {formula_attesa} (+{punti} punti)"
                punteggio_totale += punti
            else:
                if formula_cella != formula_attesa:
                    risultati[nome_file_excel][foglio_nome][
                        cella] = f"Formula errata. Attesa: {formula_attesa}, Trovata: {formula_cella} (0 punti)"
                if valore_cella != valore_atteso:
                    risultati[nome_file_excel][foglio_nome][
                        cella] += f"\nValore errato. Atteso: {valore_atteso}, Trovato: {valore_cella} (0 punti)"

    return risultati, punteggio_totale

def calcola_punteggio_totale(file_elenco_excel, file_soluzioni):
    # Carica le soluzioni e inizializza il punteggio complessivo
    soluzioni = carica_soluzioni(file_soluzioni)
    risultati_globale = {}

    # Legge la lista dei file Excel
    with open(file_elenco_excel, 'r') as file:
        nomi_file_excel = [line.strip().lower() for line in file if line.strip()]

    # Calcola il punteggio per ciascun file Excel
    for nome_file_excel in nomi_file_excel:
        risultati, punteggio_totale = controlla_formule("./verifiche/"+nome_file_excel, soluzioni)
        risultati_globale.update(risultati)

    return risultati_globale, punteggio_totale

def correggi():
    file_elenco_excel = 'elencoalunni.txt'
    file_soluzioni = 'soluzioni.txt'
    risultati_globale, punteggio_totale = calcola_punteggio_totale(file_elenco_excel, file_soluzioni)
    with open("correzione.txt", "w") as file:
        file.write(f"Correzione\n")
    # Stampa i risultati e il punteggio totale complessivo
    for nome_file, fogli in risultati_globale.items():
        with open("correzione.txt", "a") as file:
            file.write(f"File: {nome_file}\n")
            print(f"File: {nome_file}")
            for foglio, celle in fogli.items():
                file.write(f"  Foglio: {foglio}\n")
                print(f"  Foglio: {foglio}")
                for cella, risultato in celle.items():
                    file.write(f"    Cella {cella}: {risultato}\n")
                    print(f"    Cella {cella}: {risultato}")
            file.write(f"\nPunteggio totale complessivo ottenuto: {punteggio_totale}\n\n\n")
            print(f"\nPunteggio totale complessivo ottenuto: {punteggio_totale}")
    mostra_file()

#main
#Definisce i colori
colore_sfondo_root = "#AAC2FB"
colore_sfondo_line = "white"
colore_sfondo_titolo = "#5F70E7"
colore_testo_titolo = "white"
colore_sfondo_label = "white"
colore_testo_label = "black"
colore_sfondo_button = "white"
colore_testo_button = "black"
colore_sfondo_button_on = "#FF3E3E"
colore_testo_button_on = "white"
colore_sfondo_entry = "white"
colore_testo_entry = "black"
# Imposta il font
font_titolo = ("Georgia", 18)  # Font Arial, dimensione 16
font = ("Georgia", 16)  # Font Arial, dimensione 16
font_developed = ("Arial", 10)  # Font Arial, dimensione 16

# Crea la finestra principale
root = tk.Tk()
root.title("Correttore Excel")
root.configure(bg=colore_sfondo_root)
larghezza_finestra=900
altezza_finestra=600
root.geometry(f"{larghezza_finestra}x{altezza_finestra}")
center_window()
# Sovrascrivi il comportamento del pulsante di chiusura
root.protocol("WM_DELETE_WINDOW", exit_app)

# Crea un frame
frame_center = tk.Frame(root, bg=colore_sfondo_root)
frame_center.pack(expand=True, fill="both")

# Prima riga: label e campo di input centrato
label_0 = tk.Label(frame_center, text="Creazione soluzione", bg=colore_sfondo_label, fg=colore_testo_label, font=font)
label_0.grid(row=0, column=0, padx=5, pady=2, sticky="w")

button_soluzione = tk.Button(frame_center, text="Crea soluzione", bg=colore_sfondo_button, fg=colore_testo_button, font=font, command=correggi)
button_soluzione.grid(row=0, column=1, pady=2, sticky="w")
# Bind degli eventi per il cambiamento di colore
button_soluzione.bind("<Enter>", on_enter)  # Quando il mouse entra nel pulsante
button_soluzione.bind("<Leave>", on_leave)  # Quando il mouse esce dal pulsante

# Prima riga: label e campo di input centrato
label_1 = tk.Label(frame_center, text="Correzione verifiche", bg=colore_sfondo_label, fg=colore_testo_label, font=font)
label_1.grid(row=1, column=0, padx=5, pady=2, sticky="w")

button_correggi = tk.Button(frame_center, text="Correggi", bg=colore_sfondo_button, fg=colore_testo_button, font=font, command=correggi)
button_correggi.grid(row=1, column=1, pady=2, sticky="w")
# Bind degli eventi per il cambiamento di colore
button_correggi.bind("<Enter>", on_enter)  # Quando il mouse entra nel pulsante
button_correggi.bind("<Leave>", on_leave)  # Quando il mouse esce dal pulsante

label_2 = tk.Label(frame_center, text="Ricorda: inserire le verifiche nella directory verifiche", bg=colore_sfondo_label, fg=colore_testo_label, font=font)
label_2.grid(row=1, column=2, padx=5, pady=2, sticky="w")

# Avvia il loop principale della finestra
root.mainloop()