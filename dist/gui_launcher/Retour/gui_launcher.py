import tkinter as tk
from tkinter import ttk
from tkinter import PhotoImage
import subprocess

def run_scripts():
    script_files = ['ftt.py', 'fca.py', 'fta.py', 'aller1_reboot.py','aller1_analysis.py']
    
    for script in script_files:
        subprocess.run(['python', script], shell=True)

def open_folder():
    folder_path = "out.xlsx"  # Change this to the desired folder path
    subprocess.run(['explorer', folder_path], shell=True)

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

# Create and configure the run button
run_button = ttk.Button(root, text="Scrap", style="TButton", command=run_scripts)
run_button.pack(pady=10)

# Create and configure the open folder button
folder_button = ttk.Button(root, text="Ouvrir fichier", style="TButton", command=open_folder)
folder_button.pack(pady=10)

# Start the GUI event loop
root.mainloop()
