# pip install python-dateutil
from tkinter import *
from tkinter.ttk import *
from tkinter.messagebox import showwarning
from datetime import date
from dateutil import relativedelta


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Υπολογισμός χρονικής διάρκειας")
        self.window.resizable(False, False)
        self.create_widgets()


    def setToday(self, id):
        self.daysList[id - 1].set(date.today().day)
        self.monthsList[id - 1].set(date.today().month)
        self.yearsList[id - 1].set(date.today().year)


    def calcDuration(self):
        self.result.set('')
        try:
            y1 = int(self.year1.get())
            y2 = int(self.year2.get())
            m1 = int(self.month1.get())
            m2 = int(self.month2.get())
            d1 = int(self.day1.get())
            d2 = int(self.day2.get())
        except:
            showwarning("Προσοχή ...", "Όλα τα πεδία πρέπει να είναι συμπληρωμένα με αριθμούς.")
        else:
            try:
                date1 = date(y1, m1, d1)
                date2 = date(y2, m2, d2)
            except:
                showwarning("Προσοχή ...", "1) Η τιμή των πεδίων 'Ημέρα' πρέπει να είναι μεταξύ 1 και 31.\n" +
                                           "2) Η τιμή των πεδίων 'Ημέρα' σχετίζεται με την τιμή των πεδίων 'Μήνας' (π.χ. ο Φεβρουάριος σίγουρα δεν μπορεί να έχει 30 ή 31 ημέρες).\n" +
                                           "3) Η τιμή των πεδίων 'Μήνας' πρέπει να είναι μεταξύ 1 και 12.\n" +
                                           "4) Η τιμή των πεδίων 'Έτος' πρέπει να είναι μεγαλύτερη του 1.")
            else:
                diff = relativedelta.relativedelta(date2, date1)

                years = diff.years
                months = diff.months
                days = diff.days

                self.result.set('{} Έτη {} Μήνες {} Ημέρες'.format(years, months, days))


    def create_widgets(self):
        self.fDateCalc = Frame(self.window)

        Label(self.fDateCalc, text="\nΗμέρα:").grid(column=1, row=0, padx=10, pady=1)
        Label(self.fDateCalc, text="\nΜήνας:").grid(column=2, row=0, padx=10, pady=1)
        Label(self.fDateCalc, text="\nΈτος:").grid(column=3, row=0, padx=10, pady=1)

        Label(self.fDateCalc, text="Από:").grid(column=0, row=1, padx=10, pady=1, sticky=E)
        Label(self.fDateCalc, text="Μέχρι:").grid(column=0, row=2, padx=10, pady=10, sticky=E)
        Label(self.fDateCalc, text="Αποτέλεσμα:").grid(column=0, row=4, padx=10, pady=10, sticky=E)

        self.day1 = StringVar()
        self.ntrDay1 = Entry(self.fDateCalc, width=10, textvariable=self.day1, justify=CENTER)
        self.ntrDay1.grid(column=1, row=1, padx=10, pady=1)

        self.month1 = StringVar()
        self.ntrMonth1 = Entry(self.fDateCalc, width=10, textvariable=self.month1, justify=CENTER)
        self.ntrMonth1.grid(column=2, row=1, padx=10, pady=1)

        self.year1 = StringVar()
        self.ntrYear1 = Entry(self.fDateCalc, width=20, textvariable=self.year1, justify=CENTER)
        self.ntrYear1.grid(column=3, row=1, padx=10, pady=1)

        self.btnSetToday1 = Button(self.fDateCalc, text="Σήμερα", command=lambda: self.setToday(1))
        self.btnSetToday1.grid(column=4, row=1, padx=10, pady=1)

        self.day2 = StringVar()
        self.ntrDay2 = Entry(self.fDateCalc, width=10, textvariable=self.day2, justify=CENTER)
        self.ntrDay2.grid(column=1, row=2, padx=10, pady=10)

        self.month2 = StringVar()
        self.ntrMonth2 = Entry(self.fDateCalc, width=10, textvariable=self.month2, justify=CENTER)
        self.ntrMonth2.grid(column=2, row=2, padx=10, pady=10)

        self.year2 = StringVar()
        self.ntrYear2 = Entry(self.fDateCalc, width=20, textvariable=self.year2, justify=CENTER)
        self.ntrYear2.grid(column=3, row=2, padx=10, pady=10)

        self.btnSetToday2 = Button(self.fDateCalc, text="Σήμερα", command=lambda: self.setToday(2))
        self.btnSetToday2.grid(column=4, row=2, padx=10, pady=10)

        self.btnCount = Button(self.fDateCalc, text="Yπολογισμός", command=self.calcDuration)
        self.btnCount.grid(column=0, row=3, columnspan=5, padx=10, pady=10)

        self.result = StringVar()
        self.ntrResult = Entry(self.fDateCalc, width=60, textvariable=self.result, justify=CENTER, state='readonly')
        self.ntrResult.grid(column=1, row=4, columnspan=4, padx=10, pady=10, sticky=EW)

        self.daysList = [self.day1, self.day2]
        self.monthsList = [self.month1, self.month2]
        self.yearsList = [self.year1, self.year2]

        self.fDateCalc.pack()


gui = GUI()
gui.window.mainloop()
