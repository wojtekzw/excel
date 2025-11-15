# Podstawy technologii informacyjnej i aplikacje biurowe - Excel


## Wprowadzenie
Niniejszy zestaw ćwiczeń ma na celu zapoznanie studentów z podstawowymi i średniozaawansowanymi funkcjami programu Microsoft Excel w wersji desktopowej na przykładzie rozliczania projektów informatycznych w firmie programistycznej. Studenci będą pracować z rzeczywistymi danymi dotyczącymi pracowników, projektów, zadań oraz terminów ich realizacji.

## Dane wejściowe
Do wykonania ćwiczeń potrzebne będą następujące pliki CSV (w formacie amerykańskim, gdzie separatorem jest przecinek, a separatorem dziesiętnym kropka):

1. [pracownicy.csv](pracownicy.csv) - zawiera informacje o pracownikach i ich stawkach godzinowych
2. [zadania.csv](zadania.csv) - zawiera informacje o zadaniach, liczbie godzin i przypisanych pracownikach
3. [projekty.csv](projekty.csv) - zawiera informacje o projektach, klientach i stawkach
4. [daty_realizacji.csv](daty_realizacji.csv) - zawiera informacje o terminach realizacji projektów
5. [projekt-spotkanie-1.xlsx](projekt-spotkanie-1.xlsx) - przykładowy plik Excel z danymi do analizy

## Ćwiczenie 1: Importowanie danych
**Cel:** Nauczenie się importowania danych z plików CSV do Excela oraz formatowania ich jako tabele.

**Zadania:**

1. Zaimportuj dane z pliku `pracownicy.csv` do nowego arkusza o nazwie "Pracownicy":
   - Kliknij zakładkę "Dane" na wstążce
   - Wybierz "Z tekstu/CSV"
   - Przejdź do lokalizacji pliku `pracownicy.csv` i wybierz go
   - W oknie podglądu upewnij się, że separator to przecinek, a separator dziesiętny to kropka
   - Kliknij "Załaduj" (lub "Załaduj do" aby wybrać miejsce docelowe)
2. Zastosuj odpowiednie formaty liczbowe dla kolumny z kosztami godzinowymi:
   - Zaznacz kolumnę z kosztami
   - Kliknij prawym przyciskiem myszy i wybierz "Formatuj komórki"
   - Wybierz kategorię "Walutowy" z dwoma miejscami po przecinku
   - Ustaw symbol waluty "zł" i kliknij "OK"
3. Powtórz kroki 1-2 dla pozostałych plików CSV, tworząc odpowiednio nazwane arkusze.

**Wskazówki:**

- Podczas importu możesz od razu zmienić typy danych dla poszczególnych kolumn
- Jeśli napotkasz problemy z kodowaniem znaków, spróbuj wybrać inne kodowanie w oknie importu (np. UTF-8)

## Ćwiczenie 2: Podstawowe formuły i wyliczenia
**Cel:** Nauczenie się podstawowych formuł Excela do wykonywania prostych obliczeń biznesowych.

**Zadania:**

1. W arkuszu "Pracownicy" dodaj nową kolumnę "Miesięczne wynagrodzenie", w której obliczysz miesięczne wynagrodzenie każdego pracownika, zakładając że miesiąc ma 160 godzin pracy. Zaokrąglij wyniki do pełnych 10 zł:
   - W komórce F2 (zakładając, że kolumna E zawiera koszt godziny) wpisz formułę: `=ZAOKR(E2*160;-1)`
   - Przeciągnij formułę w dół do wszystkich pracowników
2. Oblicz średnie, minimalne i maksymalne miesięczne wynagrodzenie dla każdego stanowiska:
   - Utwórz tabelę przestawną (na karcie "Wstawianie" wybierz "Tabela przestawna")
   - Jako wiersze wybierz pole "stanowisko"
   - Jako wartości wybierz "Miesięczne wynagrodzenie" (trzykrotnie)
   - Zmień obliczenia wartości na Średnia, Min i Maks (trzeba wybrać Wartości i pod prawym klawiszem myszy jest menu)
3. Dodaj kolumnę z połączonym imieniem i nazwiskiem w arkuszu "Pracownicy" w kolumnie C
4. W arkuszu "Zadania" dodaj kolumnę "Koszt zadania", która będzie iloczynem liczby godzin i stawki godzinowej pracownika:
   - Użyj funkcji WYSZUKAJ.PIONOWO do pobierania stawki pracownika z arkusza "Pracownicy"
   - Przykładowa formuła: `=C2*WYSZUKAJ.PIONOWO(D2;Pracownicy!C:E;3;FAŁSZ)`
5. Oblicz sumę kosztów zadań dla każdego projektu:
   - Utwórz tabelę przestawną z polem "projekt" jako wiersze i sumą "Koszt zadania" jako wartości

**Wskazówki:**

- Funkcja ZAOKR pozwala zaokrąglać do wybranej liczby cyfr (użyj -1 dla dziesiątek)
- W wersji desktopowej Excela możesz używać polskich nazw funkcji
- Pamiętaj o poprawnym formatowaniu zakresów w funkcji WYSZUKAJ.PIONOWO
- Możesz używać bezwzględnych odniesień ($) do blokowania wierszy lub kolumn, np. `$A$1:$D$15`

## Ćwiczenie 3: Tworzenie wykresów
**Cel:** Nauka tworzenia i formatowania różnych typów wykresów do wizualizacji danych biznesowych.

**Zadania:**

1. Utwórz wykres kołowy pokazujący procentowy udział kosztów pracy różnych grup pracowników:
   - Przygotuj dane wejściowe (suma kosztów według stanowiska, moesz dodać WYSZUKAJ.PIONOWO aby w arkuszu Zadania pojawiły się te stanowiska, a później dodaj tabelę przestawną)
   - Zaznacz dane i przejdź do karty "Wstawianie" > "Wykres kołowy"
   - Dostosuj tytuł, legendę i etykiety danych (kliknij prawym przyciskiem myszy na wykres)
   - Dodaj etykiety procentowe (kliknij prawym przyciskiem myszy na wykres > "Dodaj etykiety danych")
2. Stwórz wykres kolumnowy pokazujący koszty realizacji poszczególnych projektów:
   - Zaznacz dane dotyczące kosztów projektów
   - Na karcie "Wstawianie" wybierz "Wykres kolumnowy"
   - Dostosuj formatowanie wykresu używając narzędzi na karcie "Projektowanie wykresu"
3. Utwórz wykres słupkowy pokazujący miesięczne zarobki pracowników:
   - Zaznacz dane z imionami pracowników i ich miesięcznymi zarobkami
   - Wybierz "Wstawianie" > "Wykres słupkowy"
   - Posortuj dane od najwyższych do najniższych (kliknij prawym przyciskiem myszy > "Sortuj")
4. Stwórz wykres liniowy pokazujący rozkład godzin pracy w poszczególnych projektach:
   - Przygotuj dane o sumie godzin w poszczególnych projektach
   - Wybierz "Wstawianie" > "Wykres liniowy"
   - Dodaj linie trendu (kliknij prawym przyciskiem myszy na linię > "Dodaj linię trendu")

**Wskazówki:**

- W wersji desktopowej Excel oferuje rozszerzone opcje formatowania wykresów
- Możesz korzystać z wielu elementów wykresu dostępnych na karcie "Projektowanie wykresu"
- Skorzystaj z opcji "Zmień typ wykresu" aby eksperymentować z różnymi wariantami
- Użyj karty "Format" aby dostosować szczegółowo wygląd elementów wykresu

## Ćwiczenie 4: Analiza rentowności projektów
**Cel:** Nauczenie się wykonywania złożonej analizy biznesowej przy użyciu zaawansowanych funkcji Excela.

**Zadania:**

1. Utwórz nowy arkusz o nazwie "Rentowność projektów".
2. Dla każdego projektu oblicz:
   - Całkowite koszty realizacji: `=SUMA.JEŻELI(Zadania!A:A;A2;Zadania!E:E)`
   - Przychód z projektu (z arkusza "Projekty"): `=WYSZUKAJ.PIONOWO(A2;Projekty!A:E;5;FAŁSZ)`
   - Marżę absolutną: `=przychód - koszty`
   - Marżę procentową: `=(przychód - koszty)/przychód`
3. Stwórz tabelę podsumowującą rentowność wszystkich projektów.
4. Zastosuj formatowanie warunkowe:
   - Zaznacz kolumnę z marżą procentową
   - Na karcie "Narzędzia główne" wybierz "Formatowanie warunkowe" > "Paski danych"
   - Dodaj kolejne reguły formatowania dla wyróżnienia najwyższej i najniższej marży

**Wskazówki:**

- Funkcja SUMA.JEŻELI jest idealna do sumowania wartości spełniających określone kryteria
- W wersji desktopowej Excela masz dostęp do bardziej zaawansowanych opcji formatowania warunkowego
- Możesz używać opcji "Zarządzaj regułami" aby edytować lub usuwać reguły formatowania warunkowego
- Skorzystaj z funkcji JEŻELI do dodania komentarzy opartych na wartościach marży

## Ćwiczenie 5: Analiza czasowa projektów
**Cel:** Nauka pracy z datami w Excelu oraz tworzenia analizy czasowej projektów.

**Zadania:**

1. W arkuszu "Terminy realizacji" oblicz dla każdego projektu:
   - Planowany czas trwania projektu w dniach: `=DNI.ROBOCZE(planowana_data_zakonczenia;data_rozpoczecia)`
   - Faktyczny czas trwania projektu: `=JEŻELI(faktyczna_data_zakonczenia<>"";DNI.ROBOCZE(faktyczna_data_zakonczenia;data_rozpoczecia);"W trakcie")`
   - Odchylenie od planu: `=JEŻELI(faktyczna_data_zakonczenia<>"";DNI(faktyczna_data_zakonczenia;planowana_data_zakonczenia);"Nie dotyczy")`
2. Oblicz średnie dzienne tempo realizacji zadań (dzieląc godziny przez dni trwania, uzyj dni roboczych).
3. Uwzględnij dni wolne od pracy w Polsce w 2024 roku przy obliczaniu dni roboczych, wykorzystując drugi parametr funkcji DNI.ROBOCZE.NIESTAND:
   - 1 stycznia 2024 (poniedziałek) - Nowy Rok
   - 6 stycznia 2024 (sobota) - Święto Trzech Króli
   - 1 kwietnia 2024 (poniedziałek) - Poniedziałek Wielkanocny
   - 1 maja 2024 (środa) - Święto Pracy
   - 3 maja 2024 (piątek) - Święto Konstytucji 3 Maja
   - 30 maja 2024 (czwartek) - Boże Ciało

**Wskazówki:**

- W wersji desktopowej Excel oferuje więcej funkcji do pracy z datami (np. DNI.ROBOCZE.NIESTAND)
- Możesz skorzystać z opcji formatowania liczb jako dat w różnych formatach
- Skorzystaj z warunkowego formatowania komórek aby wyróżnić projekty opóźnione
- Aby uwzględnić święta w funkcji DNI.ROBOCZE.NIESTAND, utwórz zakres komórek z datami świąt i podaj go jako 4 parametr funkcji, np.: `=DNI.ROBOCZE(data_rozpoczecia;data_zakonczenia;weekend;święta)`

## Ćwiczenie 6: Optymalizacja alokacji zasobów
**Cel:** Nauka wykorzystania Excela do optymalizacji zasobów i planowania.

**Zadania:**

1. Utwórz nowy arkusz "Alokacja zasobów".
2. Przygotuj tabelę przestawną pokazującą liczbę godzin przepracowanych przez pracowników:
   - Na karcie "Wstawianie" wybierz "Tabela przestawna"
   - Jako wiersze wybierz pole "pracownik"
   - Jako kolumny wybierz pole "projekt"
   - Jako wartości wybierz sumę "liczba_godzin"
3. Oblicz stopień wykorzystania czasu pracy (zakładając 160h/miesiąc):
   - Dodaj kolumnę "Suma godzin" i "Wykorzystanie (%)"
   - Wykorzystanie = Suma godzin / 160 * 100%
4. Zidentyfikuj pracowników przeciążonych i niedociążonych:
   - Użyj formatowania warunkowego: "Narzędzia główne" > "Formatowanie warunkowe" > "Reguły wyróżniania komórek"
   - Stwórz reguły dla wartości > 100% i < 80%

**Wskazówki:**

- Wersja desktopowa Excela oferuje pełną funkcjonalność tabel przestawnych

## Ćwiczenie 7: Dashboardy i raportowanie
**Cel:** Nauczenie się tworzenia interaktywnych dashboardów i raportów w Excelu.

**Zadania:**

1. Utwórz nowy arkusz "Dashboard":
   - Dodaj pasek tytułowy i menu nawigacyjne (hiperłącza do innych arkuszy)
   - Podziel obszar na sekcje (np. używając kształtów lub obramowań)
2. Umieść na dashboardzie:
   - Wskaźniki KPI używające formuł odwołujących się do danych z innych arkuszy
   - Wykresy z innych arkuszy (użyj obiektów połączonych)
   - Skopiuj najważniejsze tabele podsumowujące

**Wskazówki:**

- Możesz tworzyć połączone obiekty, które automatycznie aktualizują się po zmianie danych źródłowych
- Korzystaj z nazwanych zakresów dla czytelności formuł (Formuły > Menedżer nazw)

## Ćwiczenie 8: Finalna integracja i prezentacja
**Cel:** Integracja wszystkich wcześniejszych ćwiczeń i przygotowanie profesjonalnej prezentacji biznesowej.

**Zadania:**

1. Stwórz menu główne:
   - Użyj hiperłączy lub przycisków z przypisanymi makrami
   - Dodaj logo, datę automatyczną i informacje o autorze
2. Przygotuj raport dla zarządu:
   - Utwórz arkusz z kompleksowym podsumowaniem wszystkich analiz
   - Dodaj najważniejsze wykresy i wskaźniki
   - Zastosuj profesjonalne formatowanie (style, kolorystyka firmowa)
3. Zabezpiecz arkusze:
   - Na karcie "Recenzja" wybierz "Chroń arkusz/skoroszyt"
   - Ustaw hasło i zaznacz elementy dostępne dla użytkowników
4. Przygotuj wersję dystrybucyjną:
   - Usuń dane poufne lub zastąp je danymi przykładowymi
   - Sprawdź i usuń komentarze oraz dane ukryte
   - Sprawdź kompatybilność z innymi wersjami Excela

**Wskazówki:**

- Korzystaj z pełnych możliwości formatowania w wersji desktopowej
- Używaj własnych motywów i stylów dla spójnego wyglądu całego skoroszytu
- Testuj działanie zabezpieczeń przed dystrybucją
- Możesz utworzyć szablony (.xltx) do wykorzystania w przyszłych projektach

## Kryteria oceny
Każde ćwiczenie będzie oceniane pod kątem:
1. Poprawności wykonania zadań (50%)
2. Estetyki i czytelności rozwiązania (25%)
3. Kreatywności i zaawansowania zastosowanych technik (25%)


