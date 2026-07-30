"""
Alertele "eventuale" cerute in document:
  - aniversari de zi de nastere in perioada unei excursii (sau in jurul ei)
  - onomastici (nume de sfant) in perioada unei excursii
  - turisti care au ajuns la un numar de excursii ce da dreptul la premiu

Calendarul de onomastici de mai jos acopera cele mai raspandite prenume
romanesti si datele lor consacrate. Nu e o lista exhaustiva a calendarului
ortodox (acela are un sfant aproape in fiecare zi) - e gandita sa prinda
majoritatea prenumelor intalnite in practica. Daca observi un prenume
frecvent care lipseste, se adauga usor mai jos.
"""
import re
import unicodedata
from datetime import date, timedelta


def fara_diacritice(text):
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def normalizeaza(text):
    return fara_diacritice(text).upper().strip()


# prenume (fara diacritice, majuscule) -> (luna, zi)
ONOMASTICI = {
    "ION": (1, 7), "IOAN": (1, 7), "IOANA": (1, 7), "IONEL": (1, 7), "IONUT": (1, 7),
    "VASILE": (1, 1), "VASILICA": (1, 1),
    "ANTON": (1, 17), "ANTONIA": (1, 17), "ANTONIU": (1, 17),
    "GRIGORE": (1, 25),
    "EFTIMIE": (1, 20),
    "TATIANA": (1, 12),
    "MACARIE": (1, 19),
    "TEODOSIE": (1, 11),
    "CONSTANTIN": (5, 21), "CONSTANTA": (5, 21), "COSTEL": (5, 21), "COSTICA": (5, 21),
    "ELENA": (5, 21), "ELENA-CONSTANTA": (5, 21), "LENUTA": (5, 21),
    "ANDREI": (11, 30), "ANDREEA": (11, 30),
    "GHEORGHE": (4, 23), "GEORGE": (4, 23), "GEORGETA": (4, 23), "GIGI": (4, 23),
    "MARIA": (8, 15), "MARIANA": (8, 15), "MARIOARA": (8, 15), "MARIN": (8, 15),
    "MARINELA": (8, 15), "MARIUS": (8, 15),
    "DUMITRU": (10, 26), "DUMITRA": (10, 26), "MITRU": (10, 26), "MITICA": (10, 26),
    "NICOLAE": (12, 6), "NICOLETA": (12, 6), "NICU": (12, 6), "NICUSOR": (12, 6),
    "MIHAI": (11, 8), "MIHAELA": (11, 8), "MIHAIL": (11, 8), "MICHAELA": (11, 8),
    "GAVRIL": (11, 8), "GABRIEL": (11, 8), "GABRIELA": (11, 8),
    "RAFAEL": (11, 8),
    "PETRU": (6, 29), "PETRE": (6, 29), "PETRICA": (6, 29),
    "PAVEL": (6, 29), "PAUL": (6, 29), "PAULA": (6, 29),
    "ANA": (7, 26), "ANUTA": (7, 26), "ANISOARA": (7, 26),
    "ILIE": (7, 20), "ELIA": (7, 20),
    "MARINA": (7, 17),
    "CHIRIL": (7, 6),
    "CRISTINA": (7, 24), "CRISTIAN": (7, 24), "CRISTI": (7, 24),
    "IACOB": (10, 23),
    "LUCA": (10, 18),
    "SIMION": (5, 3), "SIMONA": (5, 3),
    "FILIP": (11, 14), "FILOFTEIA": (12, 7),
    "TOMA": (10, 6),
    "STEFAN": (12, 27), "STEFANIA": (12, 27), "STELIAN": (12, 27), "STELUTA": (12, 27),
    "VICTOR": (11, 11), "VICTORIA": (11, 11),
    "SERGIU": (10, 7),
    "VALENTIN": (7, 6), "VALENTINA": (7, 6),
    "IULIA": (5, 18), "IULIAN": (5, 18), "IULIANA": (5, 18),
    "SANDA": (9, 1), "ALEXANDRA": (8, 30), "ALEXANDRU": (8, 30), "SANDU": (8, 30),
    "ADRIAN": (9, 8), "ADRIANA": (9, 8),
    "DANIEL": (12, 17), "DANIELA": (12, 17), "DAN": (12, 17), "DANA": (12, 17),
    "EMIL": (5, 27), "EMILIA": (5, 27), "EMILIAN": (5, 27),
    "FLORIN": (5, 18), "FLOREA": (5, 18), "FLORICA": (5, 18), "FLORENTINA": (5, 18),
    "IOACHIM": (9, 9),
    "LUMINITA": (11, 21),
    "RADU": (10, 19), "RADA": (10, 19),
    "ROXANA": (9, 14),
    "SORIN": (10, 26), "SORINA": (10, 26),
    "TUDOR": (10, 26), "TUDORA": (10, 26),
    "VERONICA": (7, 30),
    "VIOLETA": (11, 21),
    "VIOREL": (7, 20), "VIORICA": (7, 20),
    "IRINA": (5, 5),
    "LAURA": (8, 10), "LAURENTIU": (8, 10),
    "MONICA": (5, 4),
    "OANA": (1, 7),
    "RALUCA": (10, 14), "PARASCHIVA": (10, 14),
    "ROBERT": (11, 8),
    "SILVIA": (1, 3),
    "TEODORA": (1, 11), "TEODOR": (1, 11),
    "VALERIU": (4, 22), "VALERIA": (4, 22),
    "ANGELA": (11, 8), "ANGELICA": (11, 8),
    "CATALIN": (8, 7), "CATALINA": (11, 25),
    "CLAUDIA": (8, 18), "CLAUDIU": (8, 18),
    "CORNEL": (9, 13), "CORNELIA": (9, 13), "CORNELIU": (9, 13),
    "DOINA": (10, 1),
    "ECATERINA": (11, 25),
    "AURICA": (10, 4), "AUREL": (10, 4), "AURELIA": (10, 4), "AURORA": (10, 4),
    "LILIANA": (5, 18),
    "MARCEL": (10, 7),
    "NARCISA": (10, 29),
    "OCTAVIAN": (8, 30),
    "OLGA": (7, 11),
    "RODICA": (9, 8),
    "TIBERIU": (11, 16),
}


def extrage_prenume_tokens(nume_complet):
    """
    Datele existente au formatul "Nume Prenume" (numele de familie primul).
    Extragem toate cuvintele afara de primul si le curatam de cratime.
    """
    if not nume_complet:
        return []
    cuvinte = re.split(r"[\s\-]+", nume_complet.strip())
    return [normalizeaza(c) for c in cuvinte[1:] if c]


def gaseste_onomastica(nume_complet):
    """Returneaza (luna, zi) daca vreun cuvant din nume e in calendar, altfel None."""
    for token in extrage_prenume_tokens(nume_complet):
        if token in ONOMASTICI:
            return ONOMASTICI[token]
    return None


def data_in_perioada(luna, zi, data_inceput, data_sfarsit, marja_zile=3):
    """Verifica daca data (luna, zi) din anul excursiei cade in perioada
    excursiei, extinsa cu 'marja_zile' inainte si dupa."""
    if not data_inceput:
        return False
    data_sfarsit = data_sfarsit or data_inceput
    an = data_inceput.year
    try:
        data_referinta = date(an, luna, zi)
    except ValueError:
        return False  # 29 februarie intr-un an nebisect etc.

    inceput = data_inceput - timedelta(days=marja_zile)
    sfarsit = data_sfarsit + timedelta(days=marja_zile)
    return inceput <= data_referinta <= sfarsit


def atentionari_excursie(excursie, marja_zile=3):
    """
    Pentru o excursie, returneaza o lista plata de dict-uri:
    {turist, data, tip} - cate un rand pentru fiecare zi speciala (aniversare
    sau onomastica) a fiecarui turist inscris, care cade in perioada
    excursiei (+/- marja_zile). Sortata cronologic.
    """
    rezultate = []
    for inscriere in excursie.inscrieri:
        t = inscriere.turist

        if t.data_nasterii and data_in_perioada(
            t.data_nasterii.month, t.data_nasterii.day,
            excursie.data_inceput, excursie.data_sfarsit, marja_zile
        ):
            data_txt = date(excursie.data_inceput.year, t.data_nasterii.month, t.data_nasterii.day)
            rezultate.append({"turist": t, "data": data_txt, "tip": "zi de nastere"})

        onomastica = gaseste_onomastica(t.nume)
        if onomastica and excursie.data_inceput and data_in_perioada(
            onomastica[0], onomastica[1], excursie.data_inceput, excursie.data_sfarsit, marja_zile
        ):
            data_txt = date(excursie.data_inceput.year, onomastica[0], onomastica[1])
            rezultate.append({"turist": t, "data": data_txt, "tip": "onomastica"})

    rezultate.sort(key=lambda r: r["data"])
    return rezultate


def turisti_cu_premiu_de_acordat(turisti):
    return [t for t in turisti if t.premii_de_acordat > 0]


LUNI_RO = [
    "ianuarie", "februarie", "martie", "aprilie", "mai", "iunie",
    "iulie", "august", "septembrie", "octombrie", "noiembrie", "decembrie",
]


def formateaza_zi_luna(luna, zi):
    return f"{zi} {LUNI_RO[luna - 1]}"


def urmatoarea_aparitie(luna, zi, azi=None):
    """Data urmatoare (incepand de azi) cand pica o zi recurenta anuala
    (luna, zi) - anul acesta daca n-a trecut inca, altfel anul viitor."""
    azi = azi or date.today()
    try:
        d = date(azi.year, luna, zi)
    except ValueError:
        return None  # 29 februarie
    if d < azi:
        try:
            d = date(azi.year + 1, luna, zi)
        except ValueError:
            return None
    return d


def zile_speciale_apropiate(turisti, in_zile=3650):
    """
    Pentru sectiunea "Curand" de pe pagina principala: toate zilele de
    nastere si onomasticile care urmeaza, pentru toti turistii, indiferent
    de excursii - sortate de la cea mai apropiata la cea mai indepartata.
    """
    azi = date.today()
    rezultate = []
    for t in turisti:
        if t.data_nasterii:
            d = urmatoarea_aparitie(t.data_nasterii.month, t.data_nasterii.day, azi)
            if d:
                rezultate.append({"turist": t, "data": d, "tip": "zi de nastere", "zile_pana_la": (d - azi).days})

        onomastica = gaseste_onomastica(t.nume)
        if onomastica:
            d = urmatoarea_aparitie(onomastica[0], onomastica[1], azi)
            if d:
                rezultate.append({"turist": t, "data": d, "tip": "onomastica", "zile_pana_la": (d - azi).days})

    rezultate = [r for r in rezultate if r["zile_pana_la"] <= in_zile]
    rezultate.sort(key=lambda r: r["zile_pana_la"])
    return rezultate
