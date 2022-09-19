import tkinter as tk
from tkinter import filedialog, Text
import os

# setting up app layout and color scheme
root = tk.Tk()
root.geometry("400x400")
root.title("Company Directory Application")
root.iconbitmap("C:\\Path\\To\\File\LOGO.ico")

canvas = tk.Canvas(root, bg="#008B8B")  # dark cyan
canvas.place(relwidth=1, relheight=1)

framecolor = tk.Frame(root, bg="#E0FFFF")  # light cyan background frame
framecolor.place(relwidth=0.9, relheight=0.9, relx=0.05, rely=0.05)
frame = tk.Frame(root,  bg="#E0FFFF")  # used for button placement
frame.place(relx=0.08, rely=0.3)
pixelVirtual = tk.PhotoImage(width=1, height=1)  # used to set button unit
# setup project number entry
projlabel = tk.Label(text="NMBS Project Number:", bg="#E0FFFF")
projlabel.place(x=70, y=62)
projnum = tk.Entry(root, highlightthickness=2, justify="center")
projnum.config(highlightcolor="#008B8B", highlightbackground="#008B8B")
projnum.place(width=100, height=24, x=200, y=60)


def open_file(fpath):
    projtxt = projnum.get()
    if len(fpath) > 0:
        filename = filedialog.askopenfilenames(initialdir = f"P:\\{projtxt}\\{projtxt}{fpath}")
        for file in filename:
            os.startfile(file)
    else:
        filename = filedialog.askopenfilenames(initialdir = f"P:\\{projtxt}")
        for file in filename:
            os.startfile(file)
    return


def open_checkfile(fpath):
    projtxt = projnum.get()
    filename = filedialog.askopenfilenames(initialdir = f"P:\\{projtxt}\\{projtxt}_5_CHECKER\\{projtxt}{fpath}")
    for file in filename:
        os.startfile(file)
    return


def open_dwgfile(fpath):
    projtxt = projnum.get()
    filename = filedialog.askopenfilenames(initialdir = f"P:\\{projtxt}\\{fpath}")
    for file in filename:
        os.startfile(file)
    return


# setup directory buttons for folders
mfile = tk.Button(frame, text="Main File", image=pixelVirtual,
                  width=100, height=30, fg='white', bg="#008B8B",
                  compound="center", command=lambda: open_file(""))
Cspnd = tk.Button(frame, text="Correspondence", image=pixelVirtual,
                  width=100, height=30, fg='white', bg="#008B8B",
                  compound="center", command=lambda: open_file("_1_CORRESPONDENCE"))
OFA = tk.Button(frame, text="OFA", image=pixelVirtual,
                width=100, height=30, fg='white', bg="#008B8B",
                compound="center", command=lambda: open_file("_2_ OFA"))
BFA = tk.Button(frame, text="BFA", image=pixelVirtual,
                width=100, height=30, fg='white', bg="#008B8B",
                compound="center", command=lambda: open_file("_3_BFA"))
fUse = tk.Button(frame, text="Field Use", image=pixelVirtual,
                 width=100, height=30, fg='white', bg="#008B8B",
                 compound="center", command=lambda: open_file("_4_FIELD USE"))
checker = tk.Button(frame, text="Checker", image=pixelVirtual,
                    width=100, height=30, fg='white', bg="#008B8B",
                    compound="center", command=lambda: open_file("_5_CHECKER"))
appcheck = tk.Button(frame, text="Approval Check", image=pixelVirtual,
                     width=100, height=30, fg='white', bg="#008B8B",
                     compound="center", command=lambda: open_checkfile("_Approval Check"))
backcheck = tk.Button(frame, text="Back Check", image=pixelVirtual,
                      width=100, height=30, fg='white', bg="#008B8B",
                      compound="center", command=lambda: open_checkfile("_Back Check"))
bomcheck = tk.Button(frame, text="BOM Check", image=pixelVirtual,
                     width=100, height=30, fg='white', bg="#008B8B",
                     compound="center", command=lambda: open_checkfile("_Bom Check"))
finalcheck = tk.Button(frame, text="Final Check", image=pixelVirtual,
                       width=100, height=30, fg='white', bg="#008B8B",
                       compound="center", command=lambda: open_checkfile("_Final Check"))
Ctrctdwg = tk.Button(frame, text="Contract Drawings", image=pixelVirtual,
                     width=100, height=30, fg='white', bg="#008B8B",
                     compound="center", command=lambda: open_file("_7_CONTRACT DRAWINGS SHORTCUT"))
BOM = tk.Button(frame, text="BOM", image=pixelVirtual, width=100,
                height=30, fg='white', bg="#008B8B",
                compound="center", command=lambda: open_file("_8_BOM"))
cOrder = tk.Button(frame, text="Change Order", image=pixelVirtual,
                   width=100, height=30, fg='white', bg="#008B8B",
                   compound="center", command=lambda: open_file("_9_CHANGE ORDER"))
RFI = tk.Button(frame, text="RFI", image=pixelVirtual, width=100,
                height=30, fg='white', bg="#008B8B",
                compound="center", command=lambda: open_file("_10_RFI"))
caddwg = tk.Button(frame, text="CAD Drawings", image=pixelVirtual,
                   width=100, height=30, fg='white', bg="#008B8B",
                   compound="center", command=lambda: open_dwgfile("DWG"))

# setup button placement
mfile.grid(row=0, column=0, padx=2, pady=2)
Cspnd.grid(row=1, column=0, padx=2, pady=2)
OFA.grid(row=2, column=0, padx=2, pady=2)
BFA.grid(row=3, column=0, padx=2, pady=2)
fUse.grid(row=4, column=0, padx=2, pady=2)

checker.grid(row=0, column=1, padx=2, pady=2)
appcheck.grid(row=1, column=1, padx=2, pady=2)
backcheck.grid(row=2, column=1, padx=2, pady=2)
bomcheck.grid(row=3, column=1, padx=2, pady=2)
finalcheck.grid(row=4, column=1, padx=2, pady=2)

Ctrctdwg.grid(row=0, column=2, padx=2, pady=2)
BOM.grid(row=1, column=2, padx=2, pady=2)
cOrder.grid(row=2, column=2, padx=2, pady=2)
RFI.grid(row=3, column=2, padx=2, pady=2)
caddwg.grid(row=4, column=2, padx=2, pady=2)

# setting resizable to false
root.resizable(False, False)

root.mainloop()
