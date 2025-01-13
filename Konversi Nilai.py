# import module
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# GUI tkinter
window = tk.Tk()
window.title("Konversi Nilai")
window.resizable(False, False)
window.geometry("500x400")
window['bg']='#5885AF'

# Frame untuk TreeView
frame = tk.LabelFrame(window, text="Data CSV")
frame.place(height=250, width=500)

# Buttons
buka_button = tk.Button(text = "Buka File", command = lambda: buka_file())
buka_button.place(y=300, x=125)

baca_button = tk.Button(text="Baca File", command=lambda: baca_file())
baca_button.place(y=300, x=225)

hitung_button = tk.Button(text="Hitung", command=lambda: hitung())
hitung_button.place(y=300, x=325)

# Label file
label_file = ttk.Label(text="Tidak ada file yang dipilih")
label_file.place(rely=0, relx=0)

label_simpan = ttk.Label(text="", font=("arial italic", 20),
                         background="#5885AF", foreground="white")
label_simpan.place(y=335, x=150)

# Treeview Widget
tv1 = ttk.Treeview(frame)
tv1.place(relheight=1, relwidth=1)

treescrolly = tk.Scrollbar(frame, orient="vertical", command=tv1.yview) # update sumbu y
treescrollx = tk.Scrollbar(frame, orient="horizontal", command=tv1.xview) # update sumbu x
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # scrollbars
treescrollx.pack(side="bottom", fill="x") #sumbu x scrollbar
treescrolly.pack(side="right", fill="y") # sumbu y scrollbar

# fungsi untuk membuka file
def buka_file():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("csv files", "*.csv"),
                                                    ("All files", '*.*')))
    
    label_file["text"] = filename
    return None

# fungsi untuk membaca file
def baca_file():
    file_path = label_file["text"]
    try:
        csv_filename = r"{}".format(file_path)
        if csv_filename[-4:] == ".csv":
            df = pd.read_csv(csv_filename)
        else:
            df = pd.read_excel(csv_filename)
    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", "No such file as {file_path}")
        return None

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # column heading = column name

    df_rows = df.to_numpy().tolist() # ubah dataframe menjadi list
    for row in df_rows:
        tv1.insert("", "end", values=row) # memasukkan list ke treeview
    return None

# fungsi untuk menghitung nilai
def hitung():
    baca_file()
    final_data = label_file["text"]
    final_data = pd.read_csv(final_data)

    # meghitung nilai tugas
    n_hw = 3
    for n in range(1, n_hw + 1):
        final_data[f"Nilai Tugas {n}"] = (final_data[f"Tugas {n}"] / 100)
    
    # menghitung nilai kuis
    n_quiz = 2
    for n in range(1, n_quiz + 1):
        final_data[f"Nilai Kuis {n}"] = (final_data[f"Kuis {n}"] / 100)
    
    # menghitung nilai ujian
    n_exams = 2
    for n in range(1, n_exams + 1):
        final_data[f"Nilai Evaluasi {n}"] = (final_data[f"Evaluasi {n}"] / 100)
    
    # menghitung bobot nilai
    bobot = pd.Series(
        {
            "Nilai Tugas 1": 0.2,
            "Nilai Tugas 2": 0.2,
            "Nilai Tugas 3": 0.2,
            "Nilai Kuis 1": 0.075,
            "Nilai Kuis 2": 0.075,
            "Nilai Evaluasi 1": 0.125,
            "Nilai Evaluasi 2": 0.125,
        }
    )
    
    final_data["Final Score"] = (final_data[bobot.index] * bobot).sum(axis=1)
    final_data["Ceiling Score"] = np.ceil(final_data["Final Score"] * 100)
    
    # konversi ke huruf
    letter_grades = final_data["Ceiling Score"].map(konversi_nilai)
    final_data["Final Grade"] = pd.Categorical(
    letter_grades, categories=grades.values(), ordered=True
          )
    
    # menyimpan hasil konversi
    final_data.to_excel('Nilai Akhir.xlsx', index = False)
    label_simpan["text"] = "File tersimpan"

    # bar plot
    root = tk.Tk()
    root.title("Plot")
    
    fig = Figure(figsize=(5, 5), dpi=(100))
    ax = fig.add_subplot(111)
    
    grade_counts = final_data["Final Grade"].value_counts().sort_index()
    
    grade_counts.plot(kind="bar",ax=ax)
    canvas = FigureCanvasTkAgg(fig, root)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
    
# konversi nilai 
grades = {
        86: "A",
        76: "AB",
        66: "B",
        61: "BC",
        56: "C",
        41: "D",
        0: "E",
    } 

def konversi_nilai(value):
    final_data = label_file["text"]
    for key, letter in grades.items():
            if value >= key:
                return letter

def clear_data():
    tv1.delete(*tv1.get_children())
    return None    

window.mainloop()