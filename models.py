"""
Modelele bazei de date pentru aplicatia de evidenta turisti.

Trei tabele principale, exact ca in documentul de cerinte:
  - Turist:    date personale + CI ale turistului
  - Excursie:  pachetele de excursii organizate
  - Inscriere: tabelul de legatura turist <-> excursie (cine a fost la ce
               excursie, cat a platit, ce discount a avut)
"""
from datetime import date
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Turist(db.Model):
    __tablename__ = "turisti"

    id = db.Column(db.Integer, primary_key=True)
    nume = db.Column(db.String(150), nullable=False)
    cnp = db.Column(db.String(13), unique=True, nullable=True)
    serie_ci = db.Column(db.String(20))
    emisa_de = db.Column(db.String(100))          # ex: "SPCEP Sector 5"
    adresa = db.Column(db.String(250))
    localitate = db.Column(db.String(100))
    judet = db.Column(db.String(100))
    data_nasterii = db.Column(db.Date, nullable=True)
    telefon = db.Column(db.String(30))
    email = db.Column(db.String(120))
    observatii = db.Column(db.Text)

    # numele fisierului cu poza CI, salvat in instance/uploads/
    poza_ci = db.Column(db.String(300))

    # cate premii i s-au acordat efectiv pana acum (apasand butonul din
    # aplicatie) - separat de cate a "castigat" matematic, ca sa stim
    # exact ce mai trebuie oferit
    premii_acordate = db.Column(db.Integer, default=0)

    # ajustari manuale (corectii introduse de operator cand numararea
    # automata nu porneste de la valoarea reala - ex: turist vechi care a
    # facut excursii inainte de a fi introdus in aplicatie). Aceste valori
    # se ADUNA peste ce numara automat aplicatia din inscrieri, iar
    # numararea continua normal de aici incolo la fiecare excursie noua.
    ajustare_excursii = db.Column(db.Integer, default=0)   # excursii platite in plus
    ajustare_cadouri = db.Column(db.Integer, default=0)    # excursii cadou deja primite in plus

    # istoric simplu de modificari (schimbare adresa/CI/nume) - text liber,
    # fiecare linie noua se adauga la cele vechi, cu data
    istoric_modificari = db.Column(db.Text)

    # notite private, vizibile DOAR pe fisa acestui turist - nu apar in
    # export Word, nici in liste/rapoarte, nici la alti turisti
    notite_extra = db.Column(db.Text)

    inscrieri = db.relationship(
        "Inscriere", back_populates="turist", cascade="all, delete-orphan"
    )

    @property
    def numar_excursii(self):
        """Total excursii efectuate (platite + primite cadou), incluzand
        ajustarile manuale."""
        return len(self.inscrieri) + (self.ajustare_excursii or 0) + (self.ajustare_cadouri or 0)

    @property
    def numar_excursii_platite(self):
        """Excursiile care CONTEAZA pentru premiu - fara cele deja primite cadou.
        Se adauga ajustarea manuala (excursii facute inainte de introducerea
        in aplicatie), ca numararea sa porneasca de la valoarea corecta."""
        automat = len([i for i in self.inscrieri if not i.excursie_cadou])
        return automat + (self.ajustare_excursii or 0)

    @property
    def numar_premii(self):
        """Regula simpla: un premiu la fiecare 10 excursii PLATITE (nu se
        numara si excursia cadou primita anterior, ca sa nu se acorde
        premii pe baza altor premii)."""
        return self.numar_excursii_platite // 10

    @property
    def numar_cadouri_acordate(self):
        """Cate excursii-cadou au fost deja bifate la persoana asta, plus
        premiile marcate manual prin sistemul vechi (pastrat pentru
        compatibilitate cu ce era deja marcat inainte de aceasta versiune),
        plus ajustarea manuala de cadouri deja primite."""
        return (
            len([i for i in self.inscrieri if i.excursie_cadou])
            + (self.premii_acordate or 0)
            + (self.ajustare_cadouri or 0)
        )

    @property
    def premii_de_acordat(self):
        """Premii castigate dar neacordate inca (asta genereaza atentionarea)."""
        return max(0, self.numar_premii - self.numar_cadouri_acordate)

    @property
    def varsta(self):
        if not self.data_nasterii:
            return None
        today = date.today()
        return today.year - self.data_nasterii.year - (
            (today.month, today.day) < (self.data_nasterii.month, self.data_nasterii.day)
        )


class Excursie(db.Model):
    __tablename__ = "excursii"

    id = db.Column(db.Integer, primary_key=True)
    cod = db.Column(db.String(30), unique=True, nullable=True)   # ex: delta0824
    nume = db.Column(db.String(200), nullable=False)
    data_inceput = db.Column(db.Date, nullable=True)
    data_sfarsit = db.Column(db.Date, nullable=True)
    nr_zile = db.Column(db.Integer)
    obiective = db.Column(db.Text)
    cazare = db.Column(db.String(200))
    transport = db.Column(db.String(150))
    ghid = db.Column(db.String(100))
    punct_contact = db.Column(db.String(200))
    observatii = db.Column(db.Text)

    # pretul standard al excursiei (per persoana) - folosit ca sa stim cand
    # cineva a achitat integral, in sectiunea de Costuri
    pret = db.Column(db.Float)

    # notite private, vizibile DOAR pe fisa acestei excursii - nu apar in
    # export Word, nici in liste/rapoarte, nici la alte excursii
    notite_extra = db.Column(db.Text)

    inscrieri = db.relationship(
        "Inscriere", back_populates="excursie", cascade="all, delete-orphan"
    )

    @property
    def numar_participanti(self):
        return len(self.inscrieri)


class Inscriere(db.Model):
    __tablename__ = "inscrieri"

    id = db.Column(db.Integer, primary_key=True)
    turist_id = db.Column(db.Integer, db.ForeignKey("turisti.id"), nullable=False)
    excursie_id = db.Column(db.Integer, db.ForeignKey("excursii.id"), nullable=False)

    suma_achitata = db.Column(db.Float, default=0)
    discount = db.Column(db.Float, default=0)         # suma sau procent, la alegere
    gratuitate = db.Column(db.Boolean, default=False)
    penalizare = db.Column(db.Float, default=0)
    observatii = db.Column(db.String(250))

    # bifa "excursie cadou" (patratelul negru/alb) - NU apare in Word,
    # doar in aplicatie/baza de date
    excursie_cadou = db.Column(db.Boolean, default=False)

    # marcat manual ca achitat integral, indiferent de suma introdusa
    achitat_manual = db.Column(db.Boolean, default=False)

    turist = db.relationship("Turist", back_populates="inscrieri")
    excursie = db.relationship("Excursie", back_populates="inscrieri")

    __table_args__ = (
        db.UniqueConstraint("turist_id", "excursie_id", name="uix_turist_excursie"),
    )

    @property
    def este_achitat(self):
        if self.achitat_manual:
            return True
        if self.excursie and self.excursie.pret and self.suma_achitata >= self.excursie.pret:
            return True
        return False


class CostExcursie(db.Model):
    """Un cost individual asociat unei excursii: hotel, obiectiv sau alta
    cheltuiala 'extra'. Nu apare niciodata in Word-ul exportat."""
    __tablename__ = "costuri_excursie"

    id = db.Column(db.Integer, primary_key=True)
    excursie_id = db.Column(db.Integer, db.ForeignKey("excursii.id"), nullable=False)

    tip = db.Column(db.String(20), nullable=False)  # 'hotel' | 'obiectiv' | 'extra'
    nume = db.Column(db.String(200))
    pret = db.Column(db.Float)

    # doar pentru hotel: cum se calculeaza pretul + date de contact
    unitate_pret = db.Column(db.String(30))   # 'camera/noapte' sau 'persoana/noapte'
    contact = db.Column(db.Text)

    # pentru obiectiv/extra (si optional hotel): notite libere
    notite = db.Column(db.Text)

    excursie = db.relationship("Excursie", backref=db.backref("costuri", cascade="all, delete-orphan"))


