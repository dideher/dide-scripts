# dide-scripts
# Λίστα προγραμμάτων
## Προσοχή:
Για τη σωστή εκτέλεση των προγραμμάτων που αναπτύχθηκαν από τη ΔΔΕ προτείνουμε:

- την αντιγραφή τους τοπικά στον υπολογιστή σας
- σε περίπτωση χρήσης Antivirus (π.χ. Avast), να προσθέσετε το πρόγραμμα στις «εφαρμογές που εξαιρούνται από έλεγχο» ή την απενεργοποίηση του Antivirus για το χρονικό διάστημα που θα εκτελείτε το πρόγραμμα.
## Σχετικά με Excel:

|**A/A**|**Όνομα εφαρμογής**|**Περιγραφή**|
| :-: | - | - |
||ConvertCsv2Xlsx|Μετατροπή αρχείου csv σε xlsx, σε περίπτωση που το άμεσο άνοιγμα του αρχείου csv δεν εμφανίζει σωστά τα δεδομένα.|
||ConvertXls2Xlsx|Μαζική μετατροπή αρχείων xls (παλαιού τύπου) σε xlsx (νέου τύπου).|
||CreateTeachersAssignmentTable|Δημιουργία πίνακα τοποθετήσεων (αναπληρωτών) σε word από αρχείο xlsx.|
||JoinXlsx|<p>Σύνδεση εγγραφών δύο ξεχωριστών αρχείων excel μέσω κοινών πεδίων.</p><p>Σε περίπτωση απουσίας πεδίου με μοναδικές τιμές π.χ. ΑΦΜ μπορείτε να χρησιμοποιήσετε μέχρι 4 πεδία για να δημιουργήσετε τη μοναδικότητα της εγγραφής π.χ. ΕΠΙΘΕΤΟ-ΟΝΟΜΑ-ΠΑΤΡΩΝΥΜΟ-ΕΙΔΙΚΟΤΗΤΑ.</p>|
||MergeXlsxBooks-generic|<p>Συνένωση βιβλίων excel με κοινά πεδία για τη δημιουργία ενός βιβλίου.</p><p>Για την επιτυχή συνένωση, τα αρχεία πρέπει να εμφανίζουν συγκεκριμένα χαρακτηριστικά:</p><p>- Η πρώτη γραμμή περιέχει τις κεφαλίδες των δεδομένων</p><p>- Δεν υπάρχουν συγχωνευμένα κελιά.</p>|
||ΜergeXlsxBooks-enorkoi|<p>Συνένωση βιβλίων excel με κοινά πεδία για τη δημιουργία ενός βιβλίου.</p><p>Τροποποιημένη έκδοση για επεξεργασία συγκεκριμένων αρχείων.</p>|
||MergeXlsxBooks-myschool|<p>Συνένωση βιβλίων excel με κοινά πεδία για τη δημιουργία ενός βιβλίου.</p><p>Τροποποιημένη έκδοση για επεξεργασία συγκεκριμένων αρχείων.</p>|
||MergeXlsxSheets|Συνένωση των φύλλων ενός βιβλίου excel με κοινά πεδία για τη δημιουργία ενός φύλλου.|
||SplitXls\_x-Common-Different|<p>Σύγκριση δύο αρχείων excel (xls, xlsx) και δημιουργία τριών αρχείων για:</p><p>- Εγγραφές που υπάρχουν μόνο στο πρώτο αρχείο</p><p>- Εγγραφές που υπάρχουν μόνο στο δεύτερο αρχείο</p><p>- Εγγραφές που είναι κοινές και στα δύο αρχεία.</p>|
||SplitXlsxFilter|<p>Αντικαθιστά την ακόλουθη χρονοβόρα εργασία σε ένα αρχείο xlsx: να εισάγω φίλτρο, να επιλέξω τιμή στο φίλτρο του πεδίου, να αντιγράψω, να δημιουργήσω νέο βιβλίο excel, να επικολλήσω, να αποθηκεύσω το νέο αρχείο.</p><p>Παράδειγμα:</p><p>Στο αρχείο xlsx με την κατανομή των μαθητών σε σχολεία, εμφανίζεται σε κάποια στήλη το σχολείο που τοποθετούνται οι μαθητές. Το πρόγραμμα θα δημιουργήσει ξεχωριστό αρχείο για κάθε σχολείο με τους μαθητές που ανήκουν σε αυτό.</p>|
||SplitXlsxMultipleColumnsFilter|Φιλτράρισμα με τιμές μέχρι τεσσάρων πεδίων ενός αρχείου xlsx.|
||Text2UpperCaseInXlsx|Μετατροπή του περιεχομένου ενός αρχείου xlsx σε κεφαλαία και καθαρισμός των τιμών από περιττά κενά.|
||TransferBooks2Sheets|Μεταφορά ξεχωριστών αρχείων xlsx σε φύλλα ενός αρχείου xlsx.|
||TransferSheets2Books|Δημιουργία ξεχωριστών αρχείων xlsx από τα φύλλα ενός αρχείου xlsx.|
||TransformVacanciesXlsx-Aggregate|Δημιουργία συγκεντρωτικού πίνακα κενών-πλεονασμάτων (για ΠΥΣΔΕ).|
||TransformVacanciesXlsx-Separately|Δημιουργία επιμέρους πινάκων για κενά-πλεονασμάτα (για ΠΥΣΔΕ).|

## Σχετικά με τη Διαύγεια:

|**A/A**|**Όνομα εφαρμογής**|**Περιγραφή**|
| :-: | - | - |
||ContractsUpload|Μαζική μεταφόρτωση αρχείων (π.χ. συμβάσεων) στη Διαύγεια.|
||CorrectCopiesUpload|Μαζική μεταφόρτωση διορθωμένων αρχείων (π.χ. συμβάσεων) στη Διαύγεια.|

## Σχετικά με Pdf:

|**A/A**|**Όνομα εφαρμογής**|**Περιγραφή**|
| :-: | - | - |
||ConvertDocxs2Pdfs|Μαζική μετατροπή αρχείων word σε pdf.|
||ExtractPagesFromPdf|Εξαγωγή κομματιών από ένα αρχείο pdf. Μπορούμε να προσδιορίσουμε συγκεκριμένες σελίδες αλλά και εύρος σελίδων, π.χ. 1-4, 7, 9-11.|
||MergePdfFiles|<p>Συνένωση αρχείων pdf σε ένα pdf.</p><p>Η συνένωση μπορεί να γίνει είτε:</p><p>- με αλφαβητική ταξινόμηση (το «10» προηγείται του «1»)</p><p>- με αριθμητική ταξινόμηση.</p>|
||SplitPdf2Pages|Διαχωρισμός των σελίδων ενός αρχείου pdf σε ξεχωριστά αρχεία pdf.|

## Σχετικά με Υπολογισμό χρόνου:

|**A/A**|**Όνομα εφαρμογής**|**Περιγραφή**|
| :-: | - | - |
||DateCalculator|Υπολογισμός ημερομηνίας με πρόσθεση/αφαίρεση ετών/ μηνών/ημερών σε αρχική ημερομηνία.|
||DurationCalculator|Υπολογισμός χρονικής διάρκειας (έτη-μήνες-ημέρες) μεταξύ δύο ημερομηνιών.|
||WorkExperienceCalculator|Υπολογισμός χρόνου προϋπηρεσίας από επιμέρους προϋπηρεσίες λαμβάνοντας υπόψη αν είναι κανονικές (30 ημέρες/μήνα) ή ωρομίσθιες (25 ημέρες/μήνα) προϋπηρεσίες. |
## Σχετικά με την κατανομή μαθητών σε σχολεία:

|**A/A**|**Όνομα εφαρμογής**|**Περιγραφή**|
| :-: | - | - |
||CalcDistancesSchools2Schools|<p>Υπολογισμός απόστασης μεταξύ σχολικών μονάδων, π.χ. Δημοτικών Σχολείων σε σχέση με Γυμνάσια.</p><p>Η εφαρμογή διευκολύνει τον εντοπισμό ύποπτων μετακινήσεων.</p>|
||CalcDistancesStudents2Schools|<p>Υπολογισμός απόστασης από την οικεία των μαθητών σε σχέση με σχολικές μονάδες.</p><p>Η εφαρμογή διευκολύνει την εύρεση εναλλακτικού σχολείου σε περίπτωση υπεραριθμίας κατά την κατανομή των μαθητών.</p>|
||CheckE-eggrafes|Σύγκριση του αρχείου κατανομής με την τελική αναφορά κατανομής του E-eggrafes, για εύρεση ενδεχόμενων λαθών κατά τη διαδικασία επιλογής και ανάθεσης αιτήσεων σε σχολικές μονάδες.|
||ConvertKML2Xlsx|Μετατροπή αρχείου kml (αρχείο συντεταγμένες xlsx), για χρήση σε άλλες εφαρμογές αυτής της κατηγορίας.|
||CountStudents|Καταμέτρηση μαθητών ανά σχολική μονάδα μετά από τη διαδικασία της κατανομής.|
||SchoolAvailability|Εκτίμηση δυνατότητας υποδοχής μαθητών ανά σχολική μονάδα βάσει των μαθητών που έχουν δηλωθεί ως Γ’ τάξη στο Myschool για το τρέχων σχολικό έτος.|
||SplitStudents2UrbanRural|Διαχωρισμός του αρχείου xlsx με τους μαθητές για κατανομή σε σχέση με το αν η οικία τους είναι εντός πόλης ή σε χωριό.|
||Students2Schools-addresses|<p>Εύρεση κατανομής μαθητών βάσει της διεύθυνσης κατοικίας τους και σύμφωνα με τα όρια της χωροταξικής κατανομής.</p><p>Στην εφαρμογή εισάγονται:</p><p>- τα όρια των σχολικών μονάδων (αρχείο kml)</p><p>- οι μαθητές με τις διευθύνσεις τους (αρχείο xlsx)</p><p>- τα κλειδιά για τους χάρτες (Google maps, Bing maps, Here maps)</p>|
||Students2Schools-points|<p>Εύρεση κατανομής μαθητών βάσει του στίγματος της κατοικίας τους και σύμφωνα με τα όρια της χωροταξικής κατανομής.</p><p>Στην εφαρμογή εισάγονται:</p><p>- τα όρια των σχολικών μονάδων (αρχείο kml)</p><p>- οι μαθητές με το στίγμα της κατοικίας τους (αρχείο xlsx)</p><p>- τα κλειδιά για τους χάρτες (Google maps, Bing maps, Here maps)</p>|
||VerifyAddresses|Η εύρεση της κατανομής για κάθε μαθητή πραγματοποιείται με τέσσερις διαφορετικούς χάρτες. Σε περίπτωση που δεν υπάρχει απόλυτη συμφωνία, χρησιμοποιούμε τη συγκεκριμένη εφαρμογή για να συγκρίνουμε τη διεύθυνση του μαθητή με τις διευθύνσεις που αναγνώρισαν οι τέσσερις χάρτες.|


## Για αποστολή μαζικής αλληλογραφίας:

|**A/A**|**Όνομα εφαρμογής**|**Περιγραφή**|
| :-: | - | - |
||Mail2Schools-Exams|<p>Μαζική αποστολή email σε σχολικές μονάδες με αυτόματη επισύναψη αρχείων μέσω ενός Gmail λογαριασμού της Διεύθυνσης.</p><p>Η εφαρμογή χρειάζεται ένα βιβλίο διευθύνσεων που να περιέχει εγγραφές με το όνομα, τον κωδικό και το mail του σχολείου.</p><p>Ο εντοπισμός των συνημμένων αρχείων γίνεται βάσει του κωδικού του σχολείου. Τα συνημμένα αρχεία είναι δυνατόν να βρίσκονται και σε υποφακέλους σε σχέση με τον φάκελο που εισάγουμε στην εφαρμογή ως «Φάκελος με αρχεία προς αποστολή».</p><p>Η εφαρμογή υποστηρίζει και επισύναψη κοινού αρχείου (διαβιβαστικού).</p>|
||Mail2Schools-Names|<p>Αντίστοιχη εφαρμογή με την Mail2Schools-Exams.</p><p>Ο εντοπισμός των συνημμένων αρχείων γίνεται βάσει του ονόματος του σχολείου.</p>|
||schMail2Schools-Exams|<p>Αντίστοιχη εφαρμογή με την Mail2Schools-Exams.</p><p>Η αποστολή των mail γίνεται μέσω ενός λογαριασμού της Διεύθυνσης που ανήκει στο ΠΣΔ.</p>|
||schMail2Schools-Names|<p>Αντίστοιχη εφαρμογή με την Mail2Schools-Names.</p><p>Η αποστολή των mail γίνεται μέσω ενός λογαριασμού της Διεύθυνσης που ανήκει στο ΠΣΔ.</p>|
||schMail2Teachers|<p>Αντίστοιχη εφαρμογή με την schMail2Schools-Names.</p><p>Η εφαρμογή χρειάζεται ένα βιβλίο διευθύνσεων που να περιέχει εγγραφές με το επίθετο, το όνομα, το πατρώνυμο, την ειδικότητα και το mail του εκπαιδευτικού.</p><p>Ο εντοπισμός των συνημμένων αρχείων γίνεται βάσει του συνενωμένου κειμένου που προκύπτει από το «ΕΠΙΘΕΤΟ ΟΝΟΜΑ ΠΑΤΡΩΝΥΜΟ ΕΙΔΙΚΟΤΗΤΑ».</p>|

## Σχετικά με το Myschool:

|**A/A**|**Όνομα εφαρμογής**|**Περιγραφή**|
| :-: | - | - |
||CheckMyschoolStatistics|Έλεγχος του συγκεντρωτικού σε σχέση με τους επιμέρους πίνακες του Myschool για την Ειδική Αγωγή.|


## Σχετικά με τη μισθοδοσία:

|**A/A**|**Όνομα εφαρμογής**|**Περιγραφή**|
| :-: | - | - |
||CreateZipsForAFMs|<p>Δημιουργία ξεχωριστού αρχείου zip για κάθε ΑΦΜ εκπαιδευτικού από επιμέρους αρχεία που βρίσκονται στον φάκελο που δίνεται ως είσοδος στην εφαρμογή (αλλά και στους υποφακέλους αυτού).</p><p>Σημείωση:</p><p>Στο όνομα των επιμέρους αρχείων πρέπει να περιέχεται το ΑΦΜ του εκπαιδευτικού. </p>|
||SplitPdfWithAFM|<p>Η εφαρμογή αναγνωρίζει τα ΑΦΜ των εκπαιδευτικών -ένα σε κάθε σελίδα- του αρχείου pdf και δημιουργεί ξεχωριστά αρχεία pdf για κάθε ΑΦΜ.</p><p>Αν το ΑΜΦ ενός εκπαιδευτικού εμφανίζεται σε περισσότερες από μία σελίδες, το αρχείο που θα δημιουργηθεί για αυτόν τον εκπαιδευτικό θα περιέχει όλες αυτές τις σελίδες.</p>|
||Rename2AFM|<p>Η εφαρμογή μετονομάζει μαζικά αρχεία που περιέχουν στο όνομα τους ΑΦΜ σε «καθαρό» ΑΦΜ.</p><p>Σημείωση:</p><p>Η εφαρμογή σχετίζεται με την εφαρμογή Rename-AFM-Name.</p>|
||Rename-AFM-Name|<p>Η εφαρμογή χρησιμοποιεί ένα αρχείο xlsx με την αντιστοίχιση των ΑΦΜ σε Ονοματεπώνυμο για να πραγματοποιήσει τους παρακάτω τρόπους μαζικής μετονομασίας αρχείων:</p><p>- ΑΦΜ --> Ονοματεπώνυμο',</p><p>- Ονοματεπώνυμο --> ΑΦΜ',</p><p>- ΑΦΜ --> [ΑΦΜ] Ονοματεπώνυμο',</p><p>- ΑΦΜ --> Ονοματεπώνυμο [ΑΦΜ]',</p><p>- Ονοματεπώνυμο --> [ΑΦΜ] Ονοματεπώνυμο',</p><p>- Ονοματεπώνυμο --> Ονοματεπώνυμο [ΑΦΜ]</p>|
||WorkHoursCertificates|<p>Η εφαρμογή διευκολύνει τη δημιουργία βεβαιώσεων μείωσης διδακτικού ωραρίου για εκπαιδευτικούς (τακτικού προϋπολογισμού και ΕΣΠΑ).</p><p>Υπάρχει η δυνατότητα εισαγωγής λίστας εκπαιδευτικών προς επεξεργασία καθώς και η καταχώριση κάθε βεβαίωσης σε βάση της εφαρμογής.</p>|

