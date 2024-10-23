import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from icalendar import Calendar, Event
from datetime import datetime


# Function to generate iCal file from Excel data
def generate_ical(file1):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file1)

        # Convert the 'Dato' column to datetime format, assuming day-month-year format
        df['Dato'] = pd.to_datetime(df['Dato'], format='mixed')

        # Create a new iCal Calendar
        cal = Calendar()

        # Iterate over each row in the DataFrame and add events to the calendar
        for index, row in df.iterrows():
            event = Event()
            event.add('summary', f"{row['Emne']} - {row['Fagfelt']}")
            event.add('location', row['Rom'])
            event.add('description', f"Beskjed til studenter: {row['Beskjed til studenter']}")

            # Parsing date and time for event start and end
            event_start = datetime.combine(row['Dato'].date(), pd.to_datetime(row['Tid']).time())
            event.add('dtstart', event_start)

            # Adding event duration
            duration = pd.to_timedelta(row['Timer'], unit='h')
            event_end = event_start + duration
            event.add('dtend', event_end)

            cal.add_component(event)

        # Save the calendar to a .ics file
        with open("calendar.ics", 'wb') as f:
            f.write(cal.to_ical())

        messagebox.showinfo("Success", "iCal file created successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")


# Function for browsing the file
def browse_file1():
    file1_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))


# Function to start iCal generation
def start_ical_generation():
    generate_ical(file1_path.get())


# Setting up tkinter GUI
root = tk.Tk()
root.title("iCal Generator")

# Set the background color for the main window
root.configure(bg='#b7c6cc')

# Adding logo to GUI
try:
    logo_img = Image.open("logo.png")
    logo_img = logo_img.resize((150, 150))
    logo = ImageTk.PhotoImage(logo_img)
    tk.Label(root, image=logo, bg='#b7c6cc').grid(row=0, column=0, columnspan=3, pady=10)
except Exception as e:
    messagebox.showwarning("Warning", f"Logo could not be loaded: {str(e)}")

# Variables to store file paths
file1_path = tk.StringVar()

# GUI setup
tk.Label(root, text="Choose Excel file with calendar:", bg='#b7c6cc').grid(row=1, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=file1_path, width=50, bg='#b7c6cc').grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_file1, bg='#b7c6cc').grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Generate iCal", command=start_ical_generation, width=20, bg='#b7c6cc').grid(row=3, column=1,
                                                                                                  pady=20)

# Run the GUI loop
root.mainloop()
