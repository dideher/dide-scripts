from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo, showwarning
from PyPDF2 import PdfReader, PdfWriter
import os
import shutil
import re


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Συμπίεση αρχείων pdf")
        self.window.resizable(False, False)
        self.create_widgets()

    def get_input_dir_name(self):
        d_name = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο που περιέχει τα αρχεία")

        if d_name == "":
            return

        if d_name != self.input_dir_name.get():
            self.not_compressed = True

        self.input_dir_name.set(d_name)

        self.btn_run.configure(state='normal')

    def get_output_dir_name(self):
        d_name = filedialog.askdirectory(initialdir="./data/", title="Επιλέξτε τον φάκελο που θα αποθηκεύσετε τα αρχεία")

        if d_name == "":
            return

        if d_name == self.input_dir_name.get():
            showwarning(title="Προσοχή ...", message="Έχετε επιλέξει ως προορισμό των αρχείων την πηγή.")
            return

        self.output_dir_name.set(d_name)

        self.btn_run.configure(state='normal')

    def run(self):
        if not self.not_compressed:
            showwarning(title="Προσοχή ...", message="Έχετε εκτελέσει ήδη τη συμπίεση των αρχείων αυτού του φακέλου.")
            return

        if self.output_dir_name.get() == '':
            showwarning(title="Προσοχή ...", message="Δεν έχετε επιλέξει φάκελο για τα συμπιεσμένα αρχεία.")
            return

        input_dir = self.input_dir_name.get()
        output_dir = self.output_dir_name.get()
        files_counter = 0

        for entry in os.scandir(input_dir):
            if entry.is_file() and entry.name[-4:] == ".pdf":
                files_counter += 1

                input_file = os.path.join(input_dir, entry.name)
                output_file = os.path.join(output_dir, entry.name)
                input_file_size = os.stat(input_file).st_size

                if input_file_size > 1000000:
                    reader = PdfReader(input_file)
                    writer = PdfWriter()

                    for page in reader.pages:
                        page.compress_content_streams()  # This is CPU intensive!
                        writer.add_page(page)

                    with open(os.path.join(output_dir, entry.name), "wb") as f:
                        writer.write(f)

                    output_file_size = os.stat(output_file).st_size

                    print(f"{files_counter}: {entry.name}, {input_file_size} --> {output_file_size} !!!")
                else:
                    shutil.copy(input_file, output_file)

                    print(f"{files_counter}: {entry.name}, copied (smaller than 1M) !!!")

                self.not_compressed = False

        showinfo(title="Ολοκλήρωση εκτέλεσης", message="Η συνένωση ολοκληρώθηκε.")

    def create_widgets(self):
        self.f_data = Frame(self.window)

        self.l_input_dir_name = Label(self.f_data, text="Φάκελος ασυμπίεστων αρχείων:")
        self.l_input_dir_name.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.input_dir_name = StringVar()
        self.ntr_input_dir_name = Entry(self.f_data, width=128, state='disabled', textvariable=self.input_dir_name)
        self.ntr_input_dir_name.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btn_open_input_dir = Button(self.f_data, text="Επιλέξτε φάκελο...", command=self.get_input_dir_name)
        self.btn_open_input_dir.grid(column=2, row=0, padx=10, pady=10)

        self.l_output_dir_name = Label(self.f_data, text="Φάκελος συμπιεσμένων αρχείων:")
        self.l_output_dir_name.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.output_dir_name = StringVar()
        self.ntr_output_dir_name = Entry(self.f_data, width=128, state='disabled', textvariable=self.output_dir_name)
        self.ntr_output_dir_name.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btn_open_output_dir = Button(self.f_data, text="Επιλέξτε φάκελο...", command=self.get_output_dir_name)
        self.btn_open_output_dir.grid(column=2, row=1, padx=10, pady=10)

        self.btn_run = Button(self.f_data, text="Εκτέλεση συμπίεσης", command=self.run, state='disabled')
        self.btn_run.grid(column=1, row=10, padx=10, pady=10)

        self.f_data.pack()


gui = GUI()
gui.window.mainloop()
