# pip install python-dateutil
from tkinter import *
from tkinter.ttk import *
from tkinter.messagebox import showwarning
from datetime import date
from dateutil import relativedelta


class GUI():
    def __init__(self):
        self.window = Tk()

        self.window.title("Υπολογισμός ημερομηνίας")
        self.window.resizable(False, False)
        self.create_widgets()


    def setToday(self):
        self.day.set(date.today().day)
        self.month.set(date.today().month)
        self.year.set(date.today().year)


    def calcDate(self):
        self.result.set('')
        try:
            y1 = int(self.year.get())
            m1 = int(self.month.get())
            d1 = int(self.day.get())
            y2 = int(self.years.get())
            m2 = int(self.months.get())
            d2 = int(self.days.get())

            if y2 < 0 or m2 < 0 or d2 < 0:
                showwarning("Προσοχή ...", "Όλα τα πεδία πρέπει να είναι συμπληρωμένα με θετικούς αριθμούς.")
                return
        except:
            showwarning("Προσοχή ...", "Όλα τα πεδία πρέπει να είναι συμπληρωμένα με αριθμούς.")
        else:
            try:
                date1 = date(y1, m1, d1)
            except:
                showwarning("Προσοχή ...", "1) Η τιμή των πεδίων 'Ημέρα' πρέπει να είναι μεταξύ 1 και 31.\n" +
                                           "2) Η τιμή των πεδίων 'Ημέρα' σχετίζεται με την τιμή των πεδίων 'Μήνας' (π.χ. ο Φεβρουάριος σίγουρα δεν μπορεί να έχει 30 ή 31 ημέρες).\n" +
                                           "3) Η τιμή των πεδίων 'Μήνας' πρέπει να είναι μεταξύ 1 και 12.\n" +
                                           "4) Η τιμή των πεδίων 'Έτος' πρέπει να είναι μεγαλύτερη του 1.")
            else:
                if self.cbAddSubtract.current() == -1:
                    showwarning("Προσοχή ...", "Πρέπει να επιλέξετε Πρόσθεση ή Αφαίρεση.")
                    return

                if self.cbAddSubtract.current() == 0:
                    resultDate = date1 + relativedelta.relativedelta(years=y2, months=m2, days=d2)
                else:
                    resultDate = date1 - relativedelta.relativedelta(years=y2, months=m2, days=d2)

                self.result.set("{}-{}-{}".format(resultDate.day, resultDate.month, resultDate.year))


    def create_widgets(self):
        self.fDateCalc = Frame(self.window)

        Label(self.fDateCalc, text="\nΗμέρα:").grid(column=1, row=0, padx=10, pady=1)
        Label(self.fDateCalc, text="\nΜήνας:").grid(column=2, row=0, padx=10, pady=1)
        Label(self.fDateCalc, text="\nΈτος:").grid(column=3, row=0, padx=10, pady=1)

        Label(self.fDateCalc, text="Από:").grid(column=0, row=1, padx=10, pady=1, sticky=E)
        Label(self.fDateCalc, text="\nΠρόσθεση/Αφαίρεση:").grid(column=0, row=2, padx=10, pady=1)
        Label(self.fDateCalc, text="\nΈτη:").grid(column=1, row=2, padx=10, pady=1)
        Label(self.fDateCalc, text="\nΜήνες:").grid(column=2, row=2, padx=10, pady=1)
        Label(self.fDateCalc, text="\nΗμέρες:").grid(column=3, row=2, padx=10, pady=1)
        Label(self.fDateCalc, text="Αποτέλεσμα:").grid(column=0, row=5, padx=10, pady=10, sticky=E)

        self.day = StringVar()
        self.ntrDay = Entry(self.fDateCalc, width=10, textvariable=self.day, justify=CENTER)
        self.ntrDay.grid(column=1, row=1, padx=10, pady=1)

        self.month = StringVar()
        self.ntrMonth = Entry(self.fDateCalc, width=10, textvariable=self.month, justify=CENTER)
        self.ntrMonth.grid(column=2, row=1, padx=10, pady=1)

        self.year = StringVar()
        self.ntrYear = Entry(self.fDateCalc, width=10, textvariable=self.year, justify=CENTER)
        self.ntrYear.grid(column=3, row=1, padx=10, pady=1)

        self.btnSetToday = Button(self.fDateCalc, text="Σήμερα", command=self.setToday)
        self.btnSetToday.grid(column=4, row=1, padx=10, pady=1)

        self.add_subtract = StringVar()
        self.cbAddSubtract = Combobox(self.fDateCalc, width=10, textvariable=self.add_subtract, state='readonly')
        self.cbAddSubtract.grid(column=0, row=3, padx=10, pady=10, sticky='EW')
        self.cbAddSubtract['values'] = ["Πρόσθεση (+)", "Αφαίρεση (-)"]

        self.years = StringVar()
        self.ntrYears = Entry(self.fDateCalc, width=10, textvariable=self.years, justify=CENTER)
        self.ntrYears.grid(column=1, row=3, padx=10, pady=10)

        self.months = StringVar()
        self.ntrMonths = Entry(self.fDateCalc, width=10, textvariable=self.months, justify=CENTER)
        self.ntrMonths.grid(column=2, row=3, padx=10, pady=10)

        self.days = StringVar()
        self.ntrDays = Entry(self.fDateCalc, width=10, textvariable=self.days, justify=CENTER)
        self.ntrDays.grid(column=3, row=3, padx=10, pady=10)

        self.btnCount = Button(self.fDateCalc, text="Yπολογισμός", command=self.calcDate)
        self.btnCount.grid(column=0, row=4, columnspan=5, padx=10, pady=10)

        self.result = StringVar()
        self.ntrResult = Entry(self.fDateCalc, width=40, textvariable=self.result, justify=CENTER, state='readonly')
        self.ntrResult.grid(column=1, row=5, columnspan=4, padx=10, pady=10, sticky=EW)

        self.fDateCalc.pack()


gui = GUI()
gui.window.mainloop()
