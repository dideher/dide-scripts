import os.path
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo
from os import walk
from zipfile import ZipFile


class GUI:
    def __init__(self):
        self.window = Tk()

        self.window.title("Δημιουργία zip αρχείων ανά ΑΦΜ")
        self.window.resizable(False, False)
        self.create_widgets()

    def get_output_dir(self):
        d_name = filedialog.askdirectory(initialdir="./data/",
                                         title="Επιλέξτε τον φάκελο που θα αποθηκευτούν τα αρχεία zip")

        if d_name == "":
            return

        self.output_dir.set(d_name)
        self.btn_run.configure(state='normal')

    def get_input_dir(self):
        d_name = filedialog.askdirectory(initialdir="./data/",
                                         title="Επιλέξτε τον φάκελο με τα αρχεία εισόδου")

        if d_name == "":
            return

        self.input_dir.set(d_name)
        self.ntr_output_dir.configure(state='readonly')
        self.btn_get_output_dir.configure(state='normal')

    def verify_afm(self, value):
        afm = value

        if len(afm) != 9:
            return False

        if not afm.isdigit():
            return False

        chcknumbers = [0, 2, 4, 8, 16, 32, 64, 128, 256]
        lchcknumbers = len(chcknumbers) - 1
        sum = 0

        for i in range(9):
            sum += (int(afm[i]) * chcknumbers[lchcknumbers - i])

        ch_digit = int(afm[8])

        ypoloipo = sum % 11

        if ypoloipo == 10:
            ypoloipo = 0

        if ypoloipo == ch_digit:
            return True
        else:
            return False

    def get_afm(self, text):
        split_text = text.split('.')

        f_name = split_text[0]
        parts = f_name.split('_')

        for part in parts:
            if self.verify_afm(part):
                return part

        return None

    def create_afm_list(self):
        self.afms = set()
        self.errors = ''
        self.err_counter = 0

        for path, dirs, files in walk(self.input_dir.get()):
            for file in files:
                afm = self.get_afm(file)

                if afm is not None:
                    self.afms.add(afm)
                else:
                    self.err_counter += 1
                    self.errors += f'{self.err_counter}: {path}/{file}\n'

    def create_zips(self):
        input_dir = self.input_dir.get()

        for afm in self.afms:

            zip_obj = ZipFile(os.path.join(self.output_dir.get(), f'{afm}.zip'), 'w')

            for path, dirs, files in walk(input_dir):
                for file in files:
                    if afm in file:
                        file_path = os.path.join(path, file)
                        zip_path = file_path.replace(input_dir, '')
                        zip_obj.write(file_path, zip_path)

            zip_obj.close()

    def run(self):
        self.btn_run.configure(state='disabled')

        self.create_afm_list()
        self.create_zips()

        showinfo(title="Ολοκλήρωση εκτέλεσης", message="H δημιουργία των zip αρχείων ολοκληρώθηκε.")

        if self.errors != '':
            print(20 * '-', 'Λάθη', 20 * '-')
            print(self.errors)

        self.window.destroy()

    def create_widgets(self):
        self.f_data = Frame(self.window)

        self.l_input_dir = Label(self.f_data, text="Φάκελος με αρχεία εισόδου:")
        self.l_input_dir.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.input_dir = StringVar()
        self.ntr_input_dir = Entry(self.f_data, width=128, state='readonly', textvariable=self.input_dir)
        self.ntr_input_dir.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btn_get_input_dir = Button(self.f_data, text="Επιλέξτε φάκελο...", command=self.get_input_dir)
        self.btn_get_input_dir.grid(column=2, row=0, padx=10, pady=10)

        self.l_output_dir = Label(self.f_data, text="Φάκελος για αποθήκευση των αρχείων:")
        self.l_output_dir.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.output_dir = StringVar()
        self.ntr_output_dir = Entry(self.f_data, width=128, state='disabled', textvariable=self.output_dir)
        self.ntr_output_dir.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btn_get_output_dir = Button(self.f_data, text="Επιλέξτε φάκελο...", command=self.get_output_dir,
                                         state='disabled')
        self.btn_get_output_dir.grid(column=2, row=1, padx=10, pady=10)

        self.btn_run = Button(self.f_data, text="Εκτέλεση", command=self.run, state='disabled')
        self.btn_run.grid(column=1, row=10, padx=10, pady=10)

        self.f_data.pack()


gui = GUI()
gui.window.mainloop()
