import tkinter as tk
from tkinter import ttk
from tkinter import PhotoImage
import subprocess
import os
import subprocess
import tkinter as tk
from tkinter import ttk

def run_scripts(script_files):
    for script in script_files:
        subprocess.run(['python', script], shell=True)

def open_folder():
    folder_path = "aller/rebo.xlsx"  # Change this to the desired folder path
    subprocess.run(['explorer', folder_path], shell=True)

def open_folder():
    script_directory = os.path.dirname(os.path.realpath(__file__))
    subprocess.Popen(['explorer', script_directory]) 
    

# Create the main window
root = tk.Tk()
root.title("TMPA Prices Scraper")

# Set up styles
style = ttk.Style()

# Configure the "TButton" style
style.configure("TButton", font=("Helvetica", 12), padding=10, background="#e74c3c", foreground="black")
style.map("TButton",
          foreground=[('active', 'white')],
          background=[('active', '#c0392b')])

# Configure the "TLabel" style
style.configure("TLabel", font=("Helvetica", 14), padding=10, foreground="#34495e")

# Configure the title style
style.configure("Title.TLabel", font=("Helvetica", 20, "bold"), padding=10, foreground="#3498db")

# Set up the title and picture
title_label = ttk.Label(root, text="TMPA Prices Scraper", style="Title.TLabel")
title_label.pack(pady=20)

# Insert the image path here
image_path = "head.jpg"
image = PhotoImage(file=image_path)
image_label = ttk.Label(root, image=image)
image_label.pack(pady=20)  # Added padding to center the image

# Create a frame to hold the buttons
button_frame = ttk.Frame(root)
button_frame.pack(pady=10)

# Create and configure the Scrap Aller and Aller1paxv button
scrap_aller_button = ttk.Button(button_frame, text="Scrap Aller", style="TButton", command=lambda: run_scripts(['aller/ftt.py', 'aller/fca.py', 'aller/fta.py', 'aller/aller1_reboot.py', 'aller/aller1_analysis.py', 'aller1paxv/fttv.py', 'aller1paxv/cav.py', 'aller1paxv/ftav.py', 'aller1paxv/aller1v_reboot.py', 'aller1paxv/aller1v_analysis.py', 'aller4paxv/FTTV.py', 'aller4paxv/CAV.py', 'aller4paxv/FTAV.py', 'aller4paxv/aller4Vreboot.py', 'aller4paxv/aller4V_analysis.py', 'comp.py']))
scrap_aller_button.pack(side="left", padx=10)

# Create and configure the Scrap Retour button
retour_button = ttk.Button(button_frame, text="Scrap Retour", style="TButton", command=lambda: run_scripts(['retour/ftt.py', 'retour/fca.py', 'retour/fta.py', 'retour/retour1_reboot.py', 'retour/retour1_analysis.py', 'retour1PaxV/fttv.py', 'retour1PaxV/cav.py', 'retour1PaxV/ftav.py', 'retour1PaxV/retou1v_reboot.py', 'retour1PaxV/retour1v_analysis.py', 'retour4PaxV/fttv.py', 'retour4PaxV/ftav.py', 'retour4PaxV/cav.py', 'retour4PaxV/retour4v_reboot.py', 'retour4PaxV/retour4v_analysis.py', 'comp.py']))
retour_button.pack(side="left", padx=10)

# Create and configure the Scrap Retour button
retour_button = ttk.Button(button_frame, text="Scrap Aller/Retour", style="TButton", command=lambda: run_scripts(['AllerR/fca.py', 'AllerR/fta.py', 'AllerR/ftt.py']))
retour_button.pack(side="left", padx=10)

# Create and configure the open folder button



folder_button = ttk.Button(root, text="Open Folder", command=open_folder)
folder_button.pack(pady=10)



# Start the GUI event loop
root.mainloop()
