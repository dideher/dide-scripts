from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
from tkinter.messagebox import showinfo, showwarning
from PyPDF2 import PdfFileReader, PdfFileWriter
from sortedcontainers import SortedSet


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Εξαγωγή σελίδων αρχείου pdf")
        self.window.resizable(False, False)
        self.create_widgets()

    def getDataFilename(self):
        fName = filedialog.askopenfilename(initialdir="./data/",
                                           title="Επιλέξτε το αρχείο pdf",
                                           filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))

        if fName == "":
            return

        self.pdf = PdfFileReader(fName)
        self.ntrPagesNum.configure(state='readonly')
        self.pagesNum.set(str(self.pdf.getNumPages()))
        self.ntrPages.configure(state='normal')
        self.dataFilename.set(fName)
        self.btnRun.configure(state='normal')

    def checkValue(self, value):
        value = value.strip()
        pages_num = int(self.pagesNum.get())

        if value.isdigit():
            page = int(value) - 1
            if 0 <= page < pages_num:
                return page

        return -1

    def validate(self):
        pages_text = self.pages.get().strip()
        pages_list = pages_text.split(',')

        print(pages_list)

        self.pagesSet = SortedSet()
        for item in pages_list:
            # Check for number
            page = self.checkValue(item)
            if page != -1:
                self.pagesSet.add(page)

            # Check for range
            else:
                page_range = item.split('-')

                if len(page_range) != 2:
                    continue

                startPage = self.checkValue(page_range[0])
                endPage = self.checkValue(page_range[1])

                if startPage == -1 or endPage == -1 or startPage > endPage:
                    continue

                for i in range(startPage, endPage + 1):
                    self.pagesSet.add(i)

        print(self.pagesSet)

        allPagesSet = SortedSet()
        pages_num = int(self.pagesNum.get())

        for i in range(pages_num):
            allPagesSet.add(i)

        self.restOfPagesSet = allPagesSet.difference(self.pagesSet)

        print(self.restOfPagesSet)

        if len(self.pagesSet) == 0:
            showwarning(title='Μη ορισμός σελίδων', message='Πρέπει να καταχωρίσετε τουλάχιστον μια σελίδα.')
            return False

        return True

    def run(self):
        if not self.validate():
            return

        pdf_writer = PdfFileWriter()

        for page in self.pagesSet:
            pdf_writer.addPage(self.pdf.getPage(page))

        output = self.dataFilename.get().replace(".pdf", "_extraction.pdf")
        with open(output, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)

        rest_pdf_writer = PdfFileWriter()

        for page in self.restOfPagesSet:
            rest_pdf_writer.addPage(self.pdf.getPage(page))

        output = self.dataFilename.get().replace(".pdf", "_rest.pdf")
        with open(output, 'wb') as output_pdf:
            rest_pdf_writer.write(output_pdf)

        showinfo(title="Ολοκλήρωση εκτέλεσης", message="Η εξαγωγή ολοκληρώθηκε.")

    def create_widgets(self):
        self.fData = Frame(self.window)

        self.lData = Label(self.fData, text="Αρχείο:")
        self.lData.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.dataFilename = StringVar()
        self.ntrDataFilename = Entry(self.fData, width=128, state='readonly', textvariable=self.dataFilename)
        self.ntrDataFilename.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.btnOpenData = Button(self.fData, text="Επιλέξτε αρχείο...", command=self.getDataFilename)
        self.btnOpenData.grid(column=2, row=0, padx=10, pady=10)

        self.fPages = Frame(self.fData)
        self.fPages.grid(column=0, row=1, columnspan=3, padx=10, pady=10)

        self.lPagesNum = Label(self.fPages, text="Πλήθος σελίδων αρχείου:")
        self.lPagesNum.grid(column=0, row=0, padx=10, pady=10, sticky=E)

        self.pagesNum = StringVar()
        self.ntrPagesNum = Entry(self.fPages, width=20, state='disabled', textvariable=self.pagesNum, justify=CENTER)
        self.ntrPagesNum.grid(column=1, row=0, padx=10, pady=10, sticky=W)

        self.lPages = Label(self.fPages, text="Σελίδες για εξαγωγή:")
        self.lPages.grid(column=2, row=0, padx=10, pady=10, sticky=E)

        self.pages = StringVar()
        self.ntrPages = Entry(self.fPages, width=80, state='disabled', textvariable=self.pages)
        self.ntrPages.grid(column=3, row=0, padx=10, pady=10, sticky=W)

        self.btnRun = Button(self.fData, text="Εκτέλεση εξαγωγής", command=self.run, state='disabled')
        self.btnRun.grid(column=1, row=10, padx=10, pady=10)

        self.fData.pack()


gui = GUI()
gui.window.mainloop()
