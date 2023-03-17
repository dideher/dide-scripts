from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo
import fitz
import os


class GUI:
    def __init__(self):
        self.window = Tk()

        self.window.title("Διαχωρισμός αρχείου pdf ανά ΑΦΜ")
        self.window.resizable(False, False)
        self.create_widgets()

    def get_output_dir_name(self):
        d_name = filedialog.askdirectory(initialdir="./data/",
                                         title="Επιλέξτε τον φάκελο που θα αποθηκευτούν τα αρχεία")

        if d_name == "":
            return

        self.output_dir_name.set(d_name)
        self.btn_run.configure(state='normal')

    def get_data_filename(self):
        f_name = filedialog.askopenfilename(initialdir="./data/",
                                            title="Επιλέξτε το αρχείο pdf",
                                            filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))

        if f_name == "":
            return

        self.data_filename.set(f_name)
        self.ntr_output_dir_name.configure(state='readonly')
        self.btn_open_output_dir.configure(state='normal')

    def remove_spaces(self, text_list):
        while '' in text_list:
            text_list.remove('')

        return text_list

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

    def find_afm(self, text_list):
        afm_list = list()

        for item in text_list:
            if self.verify_afm(item) and item != '099709779':
                afm_list.append(item)

        if afm_list:
            return afm_list[0]
        else:
            return self.find_afm_special(text_list)

    def find_afm_special(self, text_list):
        afm_list = list()

        for i, item in enumerate(text_list):
            test_afm = ''
            if str(item).isdecimal():
                if i + 8 <= len(text_list):
                    for j in range(9):
                        test_afm += text_list[i + j]

                    if not test_afm.isdecimal():
                        continue
                else:
                    break
            else:
                continue

            if self.verify_afm(test_afm) and test_afm != '099709779':
                afm_list.append(test_afm)

        if afm_list:
            return afm_list[0]
        else:
            return self.last_afm

    def run(self):
        self.btn_run.configure(state='disabled')

        filename = self.data_filename.get()
        output_dir = self.output_dir_name.get()

        data = dict()

        pdf = fitz.open(filename)
        self.last_afm = '000000000'
        for page in range(pdf.page_count):
            page_obj = pdf[page]
            page_text = page_obj.get_text("text")
            text_list = self.remove_spaces(page_text.replace('\n', ' ').split(' '))

            afm = self.find_afm(text_list)
            self.last_afm = afm

            print(page, afm)
            if afm not in data:
                data[afm] = list()
            data[afm].append(page)

        for afm in data:
            pdf_out = fitz.open()

            for page in data[afm]:
                pdf_out.insert_pdf(pdf, from_page=page, to_page=page)

            output_pdf = os.path.join(output_dir, afm + ".pdf")
            pdf_out.save(output_pdf)
            pdf_out.close()

        pdf.close()

        showinfo(title="Ολοκλήρωση εκτέλεσης", message="Ο διαχωρισμός ολοκληρώθηκε.")

    def create_widgets(self):
        self.f_data = Frame(self.window)

        self.l_data = Label(self.f_data, text="Αρχείο:")
        self.l_data.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.data_filename = StringVar()
        self.ntr_data_filename = Entry(self.f_data, width=128, state='readonly', textvariable=self.data_filename)
        self.ntr_data_filename.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btn_open_data = Button(self.f_data, text="Επιλέξτε αρχείο...", command=self.get_data_filename)
        self.btn_open_data.grid(column=2, row=0, padx=10, pady=10)

        self.l_output_dir_name = Label(self.f_data, text="Φάκελος για αποθήκευση των αρχείων:")
        self.l_output_dir_name.grid(column=0, row=1, padx=10, pady=10, sticky=E)

        self.output_dir_name = StringVar()
        self.ntr_output_dir_name = Entry(self.f_data, width=128, state='disabled', textvariable=self.output_dir_name)
        self.ntr_output_dir_name.grid(column=1, row=1, padx=10, pady=10, sticky=W)

        self.btn_open_output_dir = Button(self.f_data, text="Επιλέξτε φάκελο...", command=self.get_output_dir_name,
                                          state='disabled')
        self.btn_open_output_dir.grid(column=2, row=1, padx=10, pady=10)

        self.btn_run = Button(self.f_data, text="Εκτέλεση διαχωρισμού", command=self.run, state='disabled')
        self.btn_run.grid(column=1, row=10, padx=10, pady=10)

        self.f_data.pack()


gui = GUI()
gui.window.mainloop()
