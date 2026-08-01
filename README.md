# Evidenta Turisti

Aplicatie web pentru evidenta turistilor, excursiilor si inscrierilor.
Ruleaza local, pe unul din cele 2 PC-uri, si e accesibila din browser
de pe amandoua (din reteaua locala).

## Ce fisiere ai primit

```
tourist_app/
├── app.py              -> aplicatia (rutele: pagini, formulare, salvare)
├── models.py            -> structura bazei de date (Turist, Excursie, Inscriere)
├── import_excel.py      -> aduce datele din vechiul baza_de_date.xls
├── requirements.txt     -> lista de librarii Python necesare
├── static/style.css     -> aspectul vizual
└── templates/           -> paginile HTML (una pe ecran)
```

Nu trebuie sa atingi `models.py` sau `app.py` ca sa folosesti aplicatia -
sunt "motorul". Le atingi doar daca vrei sa adaugi un camp nou sau o
functionalitate noua.

Baza de date in sine (fisierul cu toate datele) se creeaza automat prima
data cand pornesti aplicatia, intr-un fisier nou `instance/evidenta.db`.
Acela e fisierul important de care faci copie de siguranta (backup) din
cand in cand - pur si simplu copiaza-l pe un stick/cloud.

## Pasul 1 - instaleaza Python

Daca nu ai deja Python pe calculator: descarca de pe
https://www.python.org/downloads/ (varianta 3.11 sau 3.12) si la
instalare bifeaza **"Add Python to PATH"**.

## Pasul 2 - deschide proiectul in VS Code

1. Instaleaza [Visual Studio Code](https://code.visualstudio.com/) daca
   nu il ai deja.
2. In VS Code: `File > Open Folder...` si alege folderul `tourist_app`.
3. Instaleaza extensia oficiala **Python** (Microsoft) din tab-ul
   Extensions (iconita cu patrate in stanga) - se cauta "Python".

## Pasul 3 - instaleaza librariile necesare

In VS Code deschide un terminal nou: meniul `Terminal > New Terminal`.
Acolo scrii, pe rand, si apesi Enter dupa fiecare linie:

```bash
python -m venv venv
```

Aceasta creeaza un mediu Python izolat, ca sa nu incurci alte programe
Python de pe calculator. Apoi il activezi:

- **Windows:** `venv\Scripts\activate`
- **Mac/Linux:** `source venv/bin/activate`

Ar trebui sa vezi `(venv)` la inceputul liniei din terminal. Apoi
instalezi librariile:

```bash
cd tourist_app


pip install -r requirements.txt
python app.py
```

## Pasul 4 - (optional) importa datele vechi din Excel

Daca vrei sa aduci automat datele din `baza_de_date.xls` (turistii si
excursiile deja existente):

```bash
python import_excel.py "C:\Users\CiDr2\others\00_IMPORTANT\aplicatie\baza1.xls"
python import_excel.py "C:\Users\CiDr2\Downloads\Tabel_turisti_complet111.doc"
```

Scriptul va afisa la final o lista cu inscrierile care nu au putut fi
asociate automat (de obicei pentru ca numele turistului scris in foaia
`Pers.Exc` nu se potriveste exact, litera cu litera, cu numele din foaia
`Pers` - virgule, spatii sau diacritice diferite). Pe acelea le adaugi
manual din aplicatie, dureaza cateva minute - restul (turisti, excursii,
si inscrierile care s-au potrivit) sunt deja acolo.

Poti rula scriptul o singura data. Daca vrei sa iei totul de la capat,
sterge fisierul `instance/evidenta.db` si ruleaza-l din nou.

## Pasul 5 - porneste aplicatia

```bash
python app.py
```

In terminal va aparea ceva de genul:

```
Running on http://127.0.0.1:5000
Running on http://192.168.1.15:5000
```

- Pe **acelasi calculator**: deschizi in browser `http://localhost:5000`
- Pe **celalalt PC din retea**: deschizi in browser adresa de tipul
  `http://192.168.1.15:5000` (a doua adresa afisata in terminal, cu
  cifrele reale de la tine - nu neaparat astea). Ambele PC-uri trebuie
  sa fie in aceeasi retea (acelasi WiFi/router).

Lasi acest terminal deschis cat timp vrei ca aplicatia sa fie
accesibila. O inchizi cu `Ctrl+C` cand termini lucrul pe ziua respectiva.

Daca al doilea PC nu reuseste sa se conecteze, cel mai probabil
Windows Firewall blocheaza conexiunea - la prima rulare iti va aparea
un mesaj "Windows Defender Firewall has blocked some features", apasa
**"Allow access"**.

## Ce poate face aplicatia acum

- **Turisti**: adaugare, editare, cautare dupa nume, stergere, poza CI
  atasata (imagine sau PDF). La editare, aplicatia tine minte automat
  schimbarile de nume/adresa/CI intr-un istoric. Fiecare turist are si o
  sectiune "Extra" pentru notite private, vizibile doar acolo.
- **Excursii**: adaugare, editare, cautare, sortare/filtrare dupa an,
  stergere, cu toate campurile din cerinte (perioada, obiective, cazare,
  transport, ghid, pret/persoana). Export automat intr-un document Word
  cu lista pentru autocar. Si aici exista o sectiune "Extra" pentru
  notite private.
- **Inscrieri**: din pagina unei excursii, alegi un turist dintr-o lista
  (poti scrie ca sa filtrezi) si completezi suma achitata, discount,
  gratuitate, penalizare.
- **Excursie cadou**: la fiecare turist inscris intr-o excursie exista un
  patratel bifabil "excursie cadou" - il bifezi cand acea excursie a fost
  oferita gratuit (dupa 10 excursii platite). Nu se numara la calculul
  urmatorului premiu.
- **Costuri interne** (nu apar in Word): fiecare excursie are o sectiune
  "Costuri" (se deschide la click) cu hoteluri, obiective si cheltuieli
  extra, plus o lista de achitatii pentru urmarirea platilor turistilor.
- **Atentionari**: sectiunea "Curand" de pe pagina principala arata cele
  mai apropiate zile de nastere/onomastici ale turistilor si cine
  urmeaza sa primeasca o excursie cadou; fiecare excursie isi arata
  propriile atentionari (zile speciale in perioada ei).
- **Rapoarte de baza**: pagina fiecarui turist arata cate excursii a
  facut si cate premii a acumulat (regula: 1 premiu la fiecare 10
  excursii platite); pagina fiecarei excursii arata lista completa a
  turistilor inscrisi si sumele.

## Import din mai multe surse

`import_excel.py` accepta acum si fisiere Word (`.docx`) cu tabele in
interior, nu doar Excel - poti amesteca tipurile in aceeasi comanda:

```bash
python import_excel.py fisier1.xls fisier2.docx fisier3.xlsx
```

## Ce nu e inclus inca (le putem adauga oricand)

- Notificari automate pe email/SMS (acum atentionarile se vad doar in
  aplicatie, cand o deschizi)

## Daca ceva nu merge

- **"python nu este recunoscut ca o comanda"** - Python nu e instalat

Cipi a fost aici
  corect sau nu a fost adaugat la PATH; reinstaleaza-l si bifeaza
  optiunea de la Pasul 1.
- **Eroare la `pip install`** - verifica sa ai `(venv)` la inceputul
  liniei din terminal inainte sa rulezi comanda.
- **Pagina nu se incarca deloc** - verifica ca terminalul cu
  `python app.py` e inca deschis si nu a dat eroare.
