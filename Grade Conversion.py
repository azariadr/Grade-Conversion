import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Window Setup
window = tk.Tk()
window.title("Grade Conversion")
window.resizable(False, False)
window.geometry("520x430")
window['bg'] = '#5885AF'

# TreeView Frame
frame = tk.LabelFrame(window, text="CSV Data")
frame.place(height=250, width=500, x=10, y=10)

# Treeview Widget
tv1 = ttk.Treeview(frame)
tv1.place(relheight=1, relwidth=1)

treescrolly = tk.Scrollbar(frame, orient="vertical", command=tv1.yview)
treescrollx = tk.Scrollbar(frame, orient="horizontal", command=tv1.xview)
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
treescrollx.pack(side="bottom", fill="x")
treescrolly.pack(side="right", fill="y")

# Style
style = ttk.Style()
style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
style.configure("Treeview", font=("Arial", 10))

# File Label
file_label = ttk.Label(text="No file selected")
file_label.place(x=10, y=265)

# Buttons
open_button = tk.Button(text="Open File", command=lambda: open_file())
open_button.place(y=300, x=50)

read_button = tk.Button(text="Read File", command=lambda: read_file())
read_button.place(y=300, x=150)

calculate_button = tk.Button(text="Calculate", command=lambda: calculate())
calculate_button.place(y=300, x=250)

# Save Label
save_label = ttk.Label(text="", font=("arial italic", 12),
                         background="#5885AF", foreground="white")
save_label.place(y=335, x=180)

# Progress Bar
progress = ttk.Progressbar(window, orient=tk.HORIZONTAL, length=400, mode='determinate')
progress.place(y=370, x=60)

# Status Bar
status_var = tk.StringVar()
status_var.set("Ready")
status_bar = ttk.Label(window, textvariable=status_var, relief=tk.SUNKEN, anchor='w')
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# Grade Mapping
grades = {
    86: "A",
    76: "AB",
    66: "B",
    61: "BC",
    56: "C",
    41: "D",
    0: "E",
}

def grade_conversion(value):
    for key, letter in grades.items():
        if value >= key:
            return letter

def clear_data():
    tv1.delete(*tv1.get_children())

def open_file():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("csv files", "*.csv"),
                                                    ("All files", '*.*')))
    file_label["text"] = filename
    status_var.set(f"File selected: {filename}")

def read_file():
    file_path = file_label["text"]
    try:
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", "No such file as {file_path}")
        return None

    # Check for required columns
    required_columns = [f"Homework {i}" for i in range(1, 4)] + \
                       [f"Quiz {i}" for i in range(1, 3)] + \
                       [f"Exam {i}" for i in range(1, 3)]

    if not all(col in df.columns for col in required_columns):
        messagebox.showerror("Error", f"File missing required columns:\n{required_columns}")
        return

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"

    for column in tv1["columns"]:
        tv1.heading(column, text=column)

    for row in df.to_numpy().tolist():
        tv1.insert("", "end", values=row)

    status_var.set("File loaded into table.")

def calculate():
    file_path = file_label["text"]
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        return

    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Unable to read file.\n{str(e)}")
        return

    progress['value'] = 10
    window.update_idletasks()

    # Grade Calculation
    for n in range(1, 4):
        df[f"Homework {n} Grade"] = df[f"Homework {n}"] / 100
    for n in range(1, 3):
        df[f"Quiz {n} Grade"] = df[f"Quiz {n}"] / 100
        df[f"Exam {n} Grade"] = df[f"Exam {n}"] / 100

    weight = pd.Series({
        "Homework 1 Grade": 0.2,
        "Homework 2 Grade": 0.2,
        "Homework 3 Grade": 0.2,
        "Quiz 1 Grade": 0.075,
        "Quiz 2 Grade": 0.075,
        "Exam 1 Grade": 0.125,
        "Exam 2 Grade": 0.125,
    })

    progress['value'] = 40
    window.update_idletasks()

    df["Final Grade"] = (df[weight.index] * weight).sum(axis=1)
    df["Ceiling Grade"] = np.ceil(df["Final Grade"] * 100)
    df["Letter Grade"] = df["Ceiling Grade"].map(grade_conversion)

    progress['value'] = 70
    window.update_idletasks()

    # Save file with dialog
    file_save = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             title="Save As")
    if file_save:
        df.to_excel(file_save, index=False)
        save_label["text"] = "File saved"
        status_var.set(f"Saved to {file_save}")

    progress['value'] = 100
    window.update_idletasks()

    # Plot Result
    plot_window = tk.Toplevel(window)
    plot_window.title("Grade Distribution")

    fig = Figure(figsize=(5, 4), dpi=100)
    ax = fig.add_subplot(111)
    grade_counts = df["Letter Grade"].value_counts().sort_index()
    grade_counts.plot(kind="bar", ax=ax)
    ax.set_title("Grade Distribution")
    ax.set_xlabel("Grade")
    ax.set_ylabel("Count")

    canvas = FigureCanvasTkAgg(fig, master=plot_window)
    canvas.draw()
    canvas.get_tk_widget().pack()

    status_var.set("Calculation completed.")

window.mainloop()
