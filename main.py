import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Text, Scrollbar, RIGHT, Y, END
from PIL import Image, ImageTk
from icalendar import Calendar, Event
from datetime import datetime
import os


# Funksjon for å vise loggvindu med tekst
def show_log(log_text):
    if log_text.strip() == "":
        return  # Ikke vis noe hvis det ikke er noen feilmeldinger
    log_window = Toplevel()
    log_window.title("Utelatte rader")

    scrollbar = Scrollbar(log_window)
    scrollbar.pack(side=RIGHT, fill=Y)

    text_widget = Text(log_window, wrap='word', yscrollcommand=scrollbar.set, width=100, height=20)
    text_widget.insert(END, log_text)
    text_widget.pack()

    scrollbar.config(command=text_widget.yview)


# Funksjon for å generere iCal-fil fra Excel-data
def generate_ical(file1):
    skipped_rows_log = ""

    try:
        df = pd.read_excel(file1)

        # Bruk norsk datoformat
        df['Dato'] = pd.to_datetime(df['Dato'], format='%d.%m.%Y', errors='coerce')

        # Lag ny kalender
        cal = Calendar()

        for index, row in df.iterrows():
            try:
                if pd.isna(row['Dato']) or pd.isna(row['Tid']):
                    skipped_rows_log += f"Rad {index + 2}: Mangler gyldig dato eller tid.\n"
                    continue

                tid = str(row['Tid'])
                start_time_str = tid.split('-')[0].strip()
                start_time = pd.to_datetime(start_time_str, format='%H:%M', errors='coerce').time()
                if pd.isna(start_time):
                    skipped_rows_log += f"Rad {index + 2}: Klarte ikke å tolke klokkeslett '{row['Tid']}'.\n"
                    continue

                event_start = datetime.combine(row['Dato'].date(), start_time)

                # Varighet
                if pd.notna(row['Timer']) and row['Timer'] > 0:
                    duration = pd.to_timedelta(row['Timer'], unit='h')
                else:
                    duration = pd.Timedelta(minutes=30)

                event_end = event_start + duration

                event = Event()
                event.add('summary', f"{row['Emne']} - {row['Tittel']}")
                event.add('dtstart', event_start)
                event.add('dtend', event_end)

                if pd.notna(row['Rom']):
                    event.add('location', row['Rom'])
                if pd.notna(row['Beskjed til studenter']):
                    event.add('description', f"Beskjed til studenter: {row['Beskjed til studenter']}")

                cal.add_component(event)

            except Exception as row_error:
                skipped_rows_log += f"Rad {index + 2}: Feil under behandling: {row_error}\n"

        # Lag filnavn basert på Excel-navn
        base_filename = os.path.splitext(os.path.basename(file1))[0]
        ical_filename = f"{base_filename}.ics"

        with open(ical_filename, 'wb') as f:
            f.write(cal.to_ical())

        messagebox.showinfo("Ferdig", f"iCal-fil lagret som:\n{ical_filename}")
        show_log(skipped_rows_log)

    except Exception as e:
        messagebox.showerror("Feil", f"Det oppstod en feil:\n{str(e)}")


# Filvelger
def browse_file1():
    file1_path.set(filedialog.askopenfilename(filetypes=[("Excel-filer", "*.xlsx")]))


# Start generering
def start_ical_generation():
    generate_ical(file1_path.get())


# GUI-oppsett
root = tk.Tk()
root.title("iCal Generator")
root.configure(bg='#b7c6cc')

# Prøv å laste inn logo
try:
    logo_img = Image.open("logo.png")
    logo_img = logo_img.resize((150, 150))
    logo = ImageTk.PhotoImage(logo_img)
    tk.Label(root, image=logo, bg='#b7c6cc').grid(row=0, column=0, columnspan=3, pady=10)
except Exception as e:
    messagebox.showwarning("Advarsel", f"Logo kunne ikke lastes: {str(e)}")

file1_path = tk.StringVar()

# GUI-elementer
tk.Label(root, text="Velg Excel-fil med timeplan:", bg='#b7c6cc').grid(row=1, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=file1_path, width=50, bg='#b7c6cc').grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Bla gjennom", command=browse_file1, bg='#b7c6cc').grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Lag iCal-fil", command=start_ical_generation, width=20, bg='#b7c6cc').grid(row=3, column=1,
                                                                                                  pady=20)

# Start GUI-loop
root.mainloop()
