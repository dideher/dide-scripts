from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo, showwarning
import fitz
import os
import re


def atoi(text):
    return int(text) if text.isdigit() else text


def natural_keys(text):
    return [atoi(c) for c in re.split(r'(\d+)', text)]


class GUI:
    def __init__(self):
        self.window = Tk()

        self.window.title("Συνένωση αρχείων pdf")
        self.window.resizable(False, False)
        self.create_widgets()

    def get_input_dir(self):
        d_name = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο που περιέχει τα αρχεία")

        if d_name == "":
            return

        if d_name != self.input_dir.get():
            self.not_merged = True

        self.input_dir.set(d_name)
        self.btn_run.configure(state='normal')

    def run(self):
        if not self.not_merged:
            showwarning(title="Προσοχή ...", message="Έχετε εκτελέσει ήδη τη συνένωση των αρχείων αυτού του φακέλου.")
            return

        input_dir = self.input_dir.get()
        files_counter = 0

        pdf_out = fitz.open()

        if self.numbers_checked.get():
            files = list()

            for entry in os.scandir(input_dir):
                if entry.is_file() and entry.name[-4:] == ".pdf":
                    files.append(os.path.join(input_dir, entry.name))

            files.sort(key=natural_keys)

            for file in files:
                files_counter += 1

                with fitz.open(file) as mfile:
                    pdf_out.insertPDF(mfile)
        else:
            for entry in os.scandir(input_dir):
                if entry.is_file() and entry.name[-4:] == ".pdf":
                    files_counter += 1

                    file = os.path.join(input_dir, entry.name)

                    with fitz.open(file) as mfile:
                        pdf_out.insertPDF(mfile)

        if files_counter == 0:
            showwarning(title="Αποτυχία εκτέλεσης", message="Δεν υπήρχαν αρχεία για συνένωση.")
        else:
            output = os.path.join(input_dir, "output.pdf")
            pdf_out.save(output)
            pdf_out.close()

            self.not_merged = False
            showinfo(title="Ολοκλήρωση εκτέλεσης", message="Η συνένωση ολοκληρώθηκε.")

    def create_widgets(self):
        self.f_data = Frame(self.window)

        self.l_input_dir = Label(self.f_data, text="Φάκελος αρχείων:")
        self.l_input_dir.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.input_dir = StringVar()
        self.ntr_input_dir = Entry(self.f_data, width=128, state='disabled', textvariable=self.input_dir)
        self.ntr_input_dir.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btn_get_input_dir = Button(self.f_data, text="Επιλέξτε φάκελο...", command=self.get_input_dir)
        self.btn_get_input_dir.grid(column=2, row=0, padx=10, pady=10)

        self.numbers_checked = BooleanVar()
        self.ckb_numbers_checked = Checkbutton(self.f_data, text="Αριθμητική ταξινόμηση αρχείων",
                                               variable=self.numbers_checked)
        self.ckb_numbers_checked.grid(column=1, row=1, padx=10, pady=10)

        self.btn_run = Button(self.f_data, text="Εκτέλεση συνένωσης", command=self.run, state='disabled')
        self.btn_run.grid(column=1, row=10, padx=10, pady=10)

        self.f_data.pack()


gui = GUI()
gui.window.mainloop()
