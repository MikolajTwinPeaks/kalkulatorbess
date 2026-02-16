# Logika działania aplikacji – Kalkulator Ofertowy Energii

## Spis treści

1. [Jak działa aplikacja (ogólnie)](#1-jak-działa-aplikacja-ogólnie)
2. [Strona 1: Dane klienta](#2-strona-1-dane-klienta)
3. [Strona 2: Analiza & Rekomendacje](#3-strona-2-analiza--rekomendacje)
   - [Kalkulator kontraktu ee (FIX/RDN/MIX)](#31-kalkulator-kontraktu-ee)
   - [Kalkulator PV (fotowoltaika)](#32-kalkulator-pv)
   - [Kalkulator BESS (magazyn energii)](#33-kalkulator-bess)
   - [Kalkulator DSR (redukcja popytu)](#34-kalkulator-dsr)
   - [Kalkulator KMB (moc bierna)](#35-kalkulator-kmb)
4. [Strona 3: Finansowanie](#4-strona-3-finansowanie)
5. [Strona 4: Generuj ofertę](#5-strona-4-generuj-ofertę)
6. [Skąd się biorą liczby (założenia)](#6-skąd-się-biorą-liczby)

---

## 1. Jak działa aplikacja (ogólnie)

Aplikacja pomaga handlowcowi szybko przygotować ofertę energetyczną dla klienta biznesowego (zakład produkcyjny, firma).

Przepływ jest prosty:

```
Handlowiec wpisuje dane klienta
        ↓
Aplikacja automatycznie oblicza rekomendacje
        ↓
Handlowiec przegląda wyniki i opcje finansowania
        ↓
Pobiera gotowe pliki: ofertę XLSX, raport DOCX, formularz klienta
```

Aplikacja składa się z 4 stron (nawigacja w menu bocznym po lewej).

---

## 2. Strona 1: Dane klienta

Tu handlowiec wpisuje informacje o kliencie zebrane na spotkaniu lub z faktur.

### Jakie dane zbieramy i po co:

| Dane | Po co są potrzebne |
|------|--------------------|
| **Nazwa firmy, NIP, branża** | Identyfikacja klienta w ofercie |
| **Dni pracy, godziny pracy** | Określenie profilu zużycia (1 zmiana / 2 zmiany / 24h). Wpływa na to, jaki produkt energetyczny rekomendujemy |
| **Roczne zużycie ee (kWh)** | Główna liczba do wszystkich obliczeń – ile energii klient zużywa rocznie |
| **Moc umowna (kW)** | Ile mocy klient ma zakontraktowane. Potrzebne do BESS (peak shaving) i DSR |
| **Grupa taryfowa** | Określa jak klient jest rozliczany za dystrybucję |
| **Średni rachunek ee (PLN/mies.)** | Do obliczenia % redukcji kosztów |
| **Cena ee z umowy (PLN/kWh)** | Obecna cena – porównujemy ją z cenami rynkowymi żeby wyliczyć oszczędność |
| **Opłata dystrybucyjna (PLN/kWh)** | Potrzebna do obliczenia wartości autokonsumpcji PV i KMB |
| **Stawka opłaty mocowej (PLN/MWh)** | Do obliczenia oszczędności z peak shaving (BESS) |
| **Kategoria mocowa (K1-K4)** | Mówi nam jak nierówny jest profil zużycia klienta. K4 = bardzo nierówny = duży potencjał oszczędności |
| **Data końca umowy** | Kiedy można podpisać nową umowę |
| **Typ umowy (FIX/RDN/MIX)** | Co klient ma teraz – porównujemy z naszą rekomendacją |
| **Dane o PV** | Czy klient ma panele, ile produkują. Wpływa na sizing BESS i rekomendację dodatkowego PV |
| **cos(φ)** | Współczynnik mocy – jeśli < 0.95, klient płaci kary za moc bierną i potrzebuje KMB |
| **Powierzchnia dachu (m²)** | Ile miejsca jest na dodatkowe panele PV |

### Przycisk "Załaduj dane DEMO"

Wypełnia formularz przykładowymi danymi typowego zakładu produkcyjnego (800 MWh/rok, 350 kW mocy, PV 200 kWp). Służy do szybkiego przetestowania aplikacji.

---

## 3. Strona 2: Analiza & Rekomendacje

Po wpisaniu danych aplikacja automatycznie uruchamia 5 kalkulatorów. Każdy działa niezależnie.

---

### 3.1. Kalkulator kontraktu ee

**Co robi:** Rekomenduje najlepszy produkt energetyczny – FIX, RDN lub MIX.

**Jak działa:**

1. Bierze 3 ceny rynkowe na 2026 rok:
   - **FIX** = 0,58 PLN/kWh (stała cena, bezpieczna)
   - **RDN** = 0,50 PLN/kWh (cena zmienna, średnia z giełdy TGE)
   - **MIX** = 0,54 PLN/kWh (50% FIX + 50% RDN)

2. Porównuje obecną cenę klienta z każdą z tych cen:
   ```
   Oszczędność = (obecna cena - nowa cena) × roczne zużycie
   ```

3. Wybiera produkt na podstawie profilu pracy klienta:
   - **Praca 24h / 3 zmiany** → rekomenduje **RDN** (bo klient może kupować tanią energię w nocy)
   - **Ma PV** → rekomenduje **MIX** (stabilność + korzystanie z tanich godzin gdy PV produkuje)
   - **Standardowy profil** i obecna cena > FIX × 1.1 → rekomenduje **FIX** (bezpieczeństwo)
   - **Standardowy profil** i cena bliska rynkowej → rekomenduje **MIX**

**Przykład:** Klient płaci 0,68 PLN/kWh, zużywa 800 000 kWh/rok. Rekomendacja MIX (0,54 PLN/kWh):
```
Oszczędność = (0,68 - 0,54) × 800 000 = 112 000 PLN/rok
```

---

### 3.2. Kalkulator PV

**Co robi:** Sprawdza czy klient potrzebuje nowej/dodatkowej instalacji fotowoltaicznej i dobiera optymalną moc.

**Jak działa:**

1. **Sprawdza czy jest sens** – jeśli wolna powierzchnia dachu < 20 m², zwraca "brak rekomendacji"

2. **Oblicza ile PV potrzeba:**
   - Cel: pokrycie 70% rocznego zużycia ee
   - Odejmuje istniejącą produkcję PV (jeśli klient już ma panele)
   - 1 kWp paneli produkuje ~1 050 kWh/rok w Polsce
   - Sprawdza ile zmieści się na dachu (1 kWp = ~5,5 m²)
   - Bierze mniejszą z dwóch wartości (potrzeba vs miejsce na dachu)
   - Zaokrągla do 5 kWp (tak się kupuje panele w praktyce)

3. **Oblicza CAPEX:**
   - < 50 kWp: 3 800 PLN za kWp
   - 50-200 kWp: 3 200 PLN za kWp
   - > 200 kWp: 2 800 PLN za kWp

4. **Oblicza autokonsumpcję** (ile % energii z PV klient zużyje sam, a ile odda do sieci):
   - Praca 24h/3 zmiany: 50% (zakład pracuje też w dzień gdy PV produkuje)
   - Praca 6 dni: 40%
   - Praca Pn-Pt: 35%

5. **Oblicza roczną oszczędność:**
   ```
   Oszczędność = produkcja PV × (autokonsumpcja × cena pełna + reszta × cena net-billing)
   ```
   Gdzie:
   - Cena pełna = cena energii + opłata dystrybucyjna (bo nie kupujemy z sieci)
   - Cena net-billing = 50% ceny energii (tyle dostajemy za oddanie do sieci)

6. **Okres zwrotu = CAPEX / oszczędność roczna**

**Przykład:** Klient ma PV 200 kWp (produkuje 210 000 kWh), zużywa 800 000 kWh, dach 800 m²:
```
Potrzeba: 800 000 × 70% - 210 000 = 350 000 kWh → 333 kWp
Miejsce na dachu: 800 / 5.5 = 145 kWp
Rekomendacja: 145 kWp (ogranicza dach)
CAPEX: 145 × 3 200 = 464 000 PLN
```

---

### 3.3. Kalkulator BESS (magazyn energii)

**Co robi:** Dobiera pojemność magazynu energii i oblicza oszczędności z 3 źródeł.

**Jak działa sizing (dobór pojemności):**

Bierze maksimum z 3 celów:

1. **Nadwyżka PV** – ile energii z PV dziennie nie jest zużywane przez zakład:
   ```
   Nadwyżka dzienna = (produkcja PV × (1 - autokonsumpcja)) / 365
   ```

2. **Peak shaving** – wygładzenie szczytów zużycia:
   ```
   Peak target = moc umowna × 30% × 3 godziny
   ```

3. **Arbitraż** – kupowanie taniej energii na giełdzie i zużywanie jej w szczycie:
   ```
   Arbitraż target = dzienne zużycie × 20%
   ```

Bierze największą z tych 3 wartości (min. 50 kWh), dodaje 25% na degradację baterii i zaokrągla do 50 kWh.

Moc magazynu = pojemność / 2 (tzn. magazyn 500 kWh ma moc 250 kW – może się rozładować w 2 godziny).

**Jak oblicza CAPEX:**
```
CAPEX = pojemność × 2 000 PLN/kWh + 30 000 PLN (system EMS) + 10% (instalacja)
```

**3 strumienie oszczędności:**

1. **Autokonsumpcja PV** – zamiast oddawać nadwyżki do sieci po niskiej cenie net-billing, magazynujemy je i zużywamy wieczorem po pełnej cenie:
   ```
   Oszczędność = nadwyżka zmagazynowana × (cena pełna - cena net-billing)
   ```

2. **Arbitraż cenowy** – kupujemy tanią energię z giełdy TGE (np. w południe za ~0,20 PLN) i zużywamy w szczycie (np. wieczorem za ~0,50 PLN):
   ```
   Oszczędność = pojemność × sprawność 90% × 50% × spread 0,30 PLN × 300 dni
   ```

3. **Peak shaving** – wygładzamy profil zużycia, dzięki czemu spadamy z kategorii mocowej (np. z K3 na K1) i płacimy niższą opłatę mocową:
   ```
   Kategorie: K1=17%, K2=40%, K3=70%, K4=100% stawki bazowej
   Przejście z K3 (70%) na K1 (17%) = oszczędność 53% opłaty mocowej
   ```

**Okres zwrotu = CAPEX / suma 3 strumieni**

---

### 3.4. Kalkulator DSR

**Co robi:** Sprawdza czy klient może zarabiać na redukcji zużycia energii na wezwanie operatora sieci (PSE).

**Jak działa:**

1. Oblicza potencjał redukcji = 15% mocy umownej
2. Jeśli potencjał < 50 kW → "nieopłacalny" (za mało)
3. Jeśli >= 50 kW:
   ```
   Przychód roczny = potencjał (kW) × 300 PLN/kW/rok
   Koszt wdrożenia = 15 000 PLN + potencjał × 50 PLN
   ```

**Przykład:** Moc umowna 350 kW:
```
Potencjał = 350 × 15% = 52.5 → zaokrąglone do 50 kW
Przychód = 50 × 300 = 15 000 PLN/rok
Koszt = 15 000 + 50 × 50 = 17 500 PLN
```

---

### 3.5. Kalkulator KMB

**Co robi:** Sprawdza czy klient potrzebuje kompensacji mocy biernej i ile może zaoszczędzić.

**Kiedy się uruchamia:** Tylko jeśli klient NIE ma jeszcze KMB i cos(φ) < 0.95.

**Jak działa:**

1. Oblicza ile mocy biernej trzeba skompensować (trygonometria):
   ```
   Q = moc umowna × (tg(arccos(obecny cosφ)) - tg(arccos(0.95)))
   ```
   To po prostu oblicza różnicę między obecną mocą bierną a docelową (przy cos(φ) = 0.95).

2. CAPEX = Q × 120 PLN za kvar

3. Oszczędność = 10% rocznego kosztu dystrybucji (bo opłaty za moc bierną to zwykle ~10% rachunku dystrybucyjnego)

**Przykład:** Moc umowna 350 kW, cos(φ) = 0.85:
```
Q ≈ 100 kvar
CAPEX = 100 × 120 = 12 000 PLN
Oszczędność = 800 000 kWh × 0.27 PLN/kWh × 10% = 21 600 PLN/rok
Zwrot = 12 000 / 21 600 = 0.6 roku (!)
```

---

### Podsumowanie na stronie Analiza

Na górze strony wyświetlamy 4 kluczowe liczby:

```
Łączny CAPEX = suma CAPEX wszystkich produktów (BESS + PV + DSR + KMB)
Oszczędność roczna = suma oszczędności ze wszystkich produktów
Okres zwrotu = łączny CAPEX / łączna oszczędność
Redukcja kosztów = oszczędność / (rachunek ee + rachunek gaz) × 12 miesięcy
```

---

## 4. Strona 3: Finansowanie

**Co robi:** Pokazuje 6 opcji sfinansowania łącznego CAPEX-u.

### 6 opcji finansowania:

| # | Model | Jak działa |
|---|-------|------------|
| 1 | **Zakup za gotówkę** | Klient płaci 100% z góry. Amortyzacja 10%/rok daje korzyść podatkową (CIT 19%). Najszybszy zwrot. |
| 2 | **Leasing operacyjny (7 lat)** | 0% wkładu. Rata liczona jak annuitet przy RRSO 6.5%. Cała rata idzie w koszty podatkowe (KUP). Nie obciąża bilansu. |
| 3 | **Leasing finansowy (10 lat)** | 5% wkładu. Rata przy 6.5% na 120 miesięcy. W koszty idą amortyzacja + odsetki. |
| 4 | **Kredyt ekologiczny BGK** | 50% kosztów pokrywa premia ekologiczna (dotacja). Kredyt tylko na pozostałe 50% przy 7% na 10 lat. |
| 5 | **ESCO** | 0 PLN z góry. Firma ESCO finansuje inwestycję, spłata z gwarantowanych oszczędności przez 15 lat. Koszt ~150% CAPEX. |
| 6 | **On-site PPA** | 0 PLN z góry. Developer buduje PV na terenie klienta, klient kupuje energię po stałej cenie 10-15 lat. Dotyczy tylko PV. |

### Jak liczone są raty (leasing/kredyt):

Wzór annuitetowy (równe raty):
```
rata = CAPEX × (r × (1+r)^n) / ((1+r)^n - 1)
```
Gdzie r = oprocentowanie miesięczne, n = liczba miesięcy.

### Tabela porównawcza

Pokazuje skróconą tabelę (model, wkład własny, rata, koszt całkowity) + wykres kolumnowy kosztów.

---

## 5. Strona 4: Generuj ofertę

**Co robi:** Pozwala pobrać 3 gotowe dokumenty.

### 1. Oferta XLSX (7 arkuszy)

Pełna, sformatowana oferta w Excelu:
- Podsumowanie z danymi klienta i tabelą produktów
- Szczegóły kontraktu ee z porównaniem FIX/RDN/MIX
- Szczegóły BESS (system, CAPEX, 3 strumienie oszczędności, strategia)
- Dodatkowe rekomendacje (PV, DSR, KMB)
- Tabela opcji finansowania z porównaniem
- Projekcja finansowa na 10 lat (z uwzględnieniem 3% wzrostu cen ee i 2% degradacji BESS rocznie)
- Lista korzyści (finansowe, operacyjne, ESG)

### 2. Raport DOCX

Raport analityczny (~20 stron) o PV + BESS w Polsce:
- Regulacje prawne (taryfy dynamiczne, koncesje, net-billing, dotacje)
- Analiza technologiczna (producenci BESS, LFP vs NMC, oprogramowanie EMS)
- Analiza rynkowa (ceny TGE, ujemne ceny, arbitraż, peak shaving)
- Business case i rekomendacje

Ten raport jest generyczny (nie zależy od danych klienta) – służy jako materiał edukacyjny/sprzedażowy.

### 3. Formularz klienta XLSX

Profesjonalny formularz do zbierania danych od klienta (6 arkuszy):
- Dane klienta (firma, profil, PPE, PPG)
- Checklist dokumentów do zebrania (faktury, umowy)
- Informacje techniczne (PV, infrastruktura, oczekiwania)
- Analiza umowy (obecny kontrakt ee i gaz)
- Szablon na dane z 12 faktur
- Workflow analizy (krok po kroku od zebrania danych do podpisania umowy)

---

## 6. Skąd się biorą liczby (założenia)

### Ceny energii (2026)
| Parametr | Wartość | Źródło |
|----------|---------|--------|
| Cena FIX | 0,58 PLN/kWh | Kontrakty terminowe TGE |
| Cena RDN (średnia) | 0,50 PLN/kWh | Średnia RDN z optymalizacją |
| Cena MIX | 0,54 PLN/kWh | 50% FIX + 50% RDN |
| Opłata mocowa | 219,40 PLN/MWh | Taryfa URE na 2026 |
| Spread arbitrażowy | 0,30 PLN/kWh | Konserwatywny spread intraday |

### Koszty inwestycyjne
| Parametr | Wartość |
|----------|---------|
| BESS (bateria) | 2 000 PLN/kWh |
| BESS (EMS) | 30 000 PLN |
| BESS (instalacja) | +10% kosztu baterii |
| PV (< 50 kWp) | 3 800 PLN/kWp |
| PV (50-200 kWp) | 3 200 PLN/kWp |
| PV (> 200 kWp) | 2 800 PLN/kWp |
| KMB | 120 PLN/kvar |
| DSR (wdrożenie) | 15 000 PLN + 50 PLN/kW |

### Parametry techniczne
| Parametr | Wartość |
|----------|---------|
| Produkcja PV w Polsce | 1 050 kWh/rok na 1 kWp |
| Sprawność BESS (RTE) | 90% |
| Degradacja BESS | 2%/rok |
| Wzrost cen ee | 3%/rok |
| Net-billing | 50% ceny energii |
| Autokonsumpcja PV (bez BESS) | 35-50% zależnie od profilu |
| DSR (przychód) | 300 PLN/kW/rok |
| KMB (oszczędność) | ~10% kosztów dystrybucji |

### Ograniczenia i uproszczenia

- Ceny rynkowe są stałe (nie aktualizują się automatycznie)
- Autokonsumpcja PV jest szacowana na podstawie profilu pracy, nie z danych 15-minutowych
- Arbitraż BESS zakłada 300 efektywnych dni handlowych w roku
- Peak shaving zakłada przejście o 2 kategorie mocowe (np. K4→K2)
- Obliczenia nie uwzględniają inflacji ani zmiany regulacji
- Raport DOCX jest generyczny i nie jest personalizowany danymi klienta

---

## Struktura plików

```
BESS_PV_Kalkulator/
├── app.py                         ← główna aplikacja Streamlit (interfejs)
├── kalkulator_oferta.py           ← silnik obliczeniowy + generator oferty XLSX
├── kalkulator_bess.py             ← osobny kalkulator BESS (stand-alone)
├── generuj_raport.py              ← generator raportu DOCX
├── generuj_formularz_klienta.py   ← generator formularza klienta XLSX
├── logo.svg                       ← logo SunHelp (białe)
└── requirements.txt               ← zależności Python
```

### Kto co robi:

- **app.py** – wyświetla interfejs (formularze, wykresy, przyciski). Nie liczy nic sam – wywołuje funkcje z kalkulator_oferta.py
- **kalkulator_oferta.py** – cała logika obliczeniowa (5 kalkulatorów + finansowanie) + generowanie oferty XLSX
- **kalkulator_bess.py** – osobny, niezależny kalkulator BESS (do użycia bez pełnej oferty, z CLI)
- **generuj_raport.py** – tworzy raport DOCX (statyczny, bez danych klienta)
- **generuj_formularz_klienta.py** – tworzy formularz intake XLSX (statyczny)
