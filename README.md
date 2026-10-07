# NPS Event Viewer

Graficzny interfejs (WPF) do przeglądania i filtrowania zdarzeń autoryzacji **Network Policy Server (NPS)** na serwerach Windows.

## Funkcjonalność

Aplikacja oferuje trzy główne sekcje:

### 📋 Zakładka 1: Dziennik Security
- Przeglądanie zdarzeń autoryzacji z dziennika Security (ID: 6272–6280)
- Filtrowanie po:
  - **Wyniku** (udzielono dostępu, odmowa, odrzucono, kwarantanna, dostęp warunkowy, pełny dostęp, zablokowanie/odblokowanie konta)
  - **Switch/Kliencie RADIUS** (urządzenie wysyłające żądanie)
  - **Zasadzie sieciowej** (policy)
  - **Typie uwierzytelniania** (EAP, PEAP, MSCHAPv2, etc.)
  - **Użytkowniku / MAC** (różne formaty)
  - **Komputerze**
  - **Pełnotekstowy wyszukiwaniu**
- Obsługa zdalnych serwerów
- Dostosowywalne zakresy czasowe (ostatnia godzina, 24h, 7 dni, 30 dni, własny zakres)
- Eksport wyników do CSV
- Automatyczne odświeżanie co 60 sekund

### 📁 Zakładka 2: Pliki logów RADIUS (.log)
- Analiza plików logów NPS (Accounting → Log File)
- Obsługiwane formaty:
  - **DTS-compliant (XML)** — zalecany
  - **IAS (best effort)**
- Źródła danych:
  - Folder lokalny
  - Pojedyncze pliki
  - Ścieżka UNC (np. `\\NPS01\c$\Windows\System32\LogFiles`)
- Automatyczne uzupełnianie danych z Access-Request (User-Name, MAC, port) dla odpowiedzi
- Opcja wyświetlania Access-Request / Challenge i zdarzeń Accounting
- Takie same filtry jak w zakładce 1
- Domyślna lokalizacja: `%SystemRoot%\System32\LogFiles`

### 🖥️ Zakładka 3: Zdarzenia systemowe NPS
- Przeglądanie zdarzeń systemowych z dziennika **System** (opcjonalnie **Application**)
- Obejmuje:
  - Niezgodny shared secret / Message-Authenticator (ID 18, 14)
  - Nieznany klient RADIUS (ID 13)
  - Problemy z kontrolerem domeny (ID 4402)
  - Inne problemy infrastruktury NPS
- Filtrowanie po:
  - **Poziomie** (krytyczne, błędy, ostrzeżenia, informacje)
  - **ID zdarzenia**
  - **Źródle** (NPS, IAS, Microsoft-Windows-NPS)
  - **Kliencie RADIUS (IP)**
  - **Zawartości wiadomości**
- Wskazówki diagnostyczne dla typowych problemów

### 🕒 Historia urządzenia / użytkownika / komputera (oś czasu)
Osobne okno z pełnym „łańcuchem zdarzeń” jednego **MAC**, **użytkownika** albo **komputera** – jak oś czasu w konsolach antywirusowych. Przydatne, gdy urządzenie przechodzi np. z Wi-Fi na kabel i z powrotem, a gdzieś po drodze dostaje odmowę.
- Otwieranie: **dwuklik na wierszu** tabeli, menu kontekstowe (*Historia tego urządzenia / użytkownika / komputera*) albo przycisk **Historia urządzenia / użytkownika...** w nagłówku (można też wpisać MAC w dowolnym formacie). Można mieć otwartych kilka okien naraz.
- **Zakres czasu**: od ręki z danych wczytanych w zakładkach albo wczytanie w tle ostatniej godziny, 24 h, 7 / 30 / 90 dni lub własnego zakresu – z dziennika Security (dla MAC filtr wykonuje sam dziennik, więc działa szybko także na dużym logu) i/lub z plików `.log` (filtr już podczas czytania plików). Dopasowanie i budowa osi czasu też odbywają się w tle, więc okna nie zamarzają nawet przy dziesiątkach tysięcy zdarzeń.
- **Zapis MAC**: w dzienniku Security szukane są zapisy `AA-BB-CC-DD-EE-FF`, `AA:BB:…`, `AABBCCDDEEFF`, `aabb.ccdd.eeff` (Cisco), `aabbcc-ddeeff` (HP), `aabb-ccdd-eeff` (Huawei / H3C), `aa.bb.cc.dd.ee.ff` – wielkimi i małymi literami. Inny zapis: zaznacz *Dopasowanie częściowe*. Adres IP w Calling-Station-Id (VPN) nie jest traktowany jak MAC.
- **Access-Request / Challenge** z plików `.log` (kilkanaście na jedno logowanie EAP) są wczytywane tylko po zaznaczeniu tej opcji – dzięki temu nie zajmują limitu *MAKS. ZDARZEŃ*; gdy limit zostanie osiągnięty, okno o tym informuje.
- **Oś czasu**: nagłówki dni (z liczbą zdarzeń i odmów), kolorowe kropki wyniku, plakietka **LAN / Wi-Fi / VPN**, switch / port / SSID, zasada, metoda uwierzytelnienia, adres IP i czas sesji z accountingu.
- **Zaznaczone zdarzenia przełomowe**: zmiana sieci (LAN ↔ Wi-Fi, inny switch / port / SSID), inny użytkownik na urządzeniu albo inne urządzenie użytkownika, przerwy dłuższe niż godzina. Porównywane są tylko dane znane po obu stronach (np. wpis accounting bez typu portu albo nazwa portu `Gi1/0/12` zamiast numeru nie oznaczają zmiany), a ponowne połączenie VPN z innym numerem tunelu to nie zmiana miejsca. Powtórzenia (np. reautoryzacje co kilka minut) są zgrupowane jako „×N od–do”.
- **Pasek aktywności** – słupki udzielono / odmowa w czasie; kliknięcie słupka zawęża widok do tego odcinka. Menu: *tylko ten dzień*, *godzina przed i po*, *od / do tego momentu*.
- **Podsumowanie**: pierwsze / ostatnie zdarzenie, liczba odmów i ostatnia odmowa, zmiany miejsca w sieci, udział LAN / Wi-Fi; powiązani użytkownicy, urządzenia (MAC) i komputery – kliknięcie otwiera ich historię – oraz switche, SSID i zasady.
- Szczegóły zaznaczonego zdarzenia jak w zakładkach, **Kopiuj oś czasu** (tekst do zgłoszenia) i **Eksport CSV**. Zdarzenia widoczne jednocześnie w dzienniku Security i w pliku `.log` są pokazywane raz („Security + log”).

## Wymagania

- **Windows Server** 2012 R2 lub nowszy
- **PowerShell 5.1**
- Uruchomienie **jako Administrator** (lub konto w grupie *Event Log Readers*)
- Do zdalnych serwerów: reguła firewalla **"Remote Event Log Management"**

## Instalacja

1. Pobierz plik `Logi_NPS.ps1`
2. Kliknij prawym przyciskiem → *Uruchom w PowerShell*
   - lub otwórz PowerShell jako Administrator i wykonaj: `& 'C:\ścieżka\Logi_NPS.ps1'`

## Uwagi

- **Wczytywanie w tle**: UI nie blokuje się podczas pobierania danych (także w oknie historii)
- **Test dymny**: `tests/Smoke-NpsViewer.ps1` (uruchamiany w GitHub Actions na Windows PowerShell 5.1) sprawdza logikę osi czasu (`tests/HistoryLogic.Tests.ps1`), otwiera okna na danych testowych, zamyka okno w trakcie wczytywania i mierzy czas przy 20000 zdarzeń
- **Dane nie odblokowywane**: Pliki `.log` są otwierane bez blokowania — NPS może dalej pisać
- **Audyt**: Jeśli brak zdarzeń, sprawdź audyt:
  ```powershell
  auditpol /get /subcategory:"Network Policy Server"
  ```
- **Kodowanie**: Zapisz skrypt jako UTF-8 z BOM (dla polskich znaków w PowerShell 5.1)
- **Kolory kodów odpowiedzi**: 
  - 🟢 Zielony = powodzenie
  - 🔴 Czerwony = błąd
  - 🟡 Żółty = ostrzeżenie

## Wskazówki diagnostyczne

Aplikacja zawiera automatyczne wskazówki dla najczęstszych problemów:

- **ID 13**: Klient RADIUS nie jest zarejestrowany
- **ID 14**: Brak Message-Authenticator
- **ID 18**: Niezgodny shared secret
- **ID 4402**: Brak dostępu do kontrolera domeny
- **Kod 8**: Konto nie istnieje
- **Kod 16**: Złe poświadczenia
- **Kod 34/36**: Konto wyłączone/zablokowane
- **Kod 48**: Żądanie nie pasuje do żadnej zasady

## Rozwój

Projekt jest open-source. Zapraszamy do zgłaszania problemów i propozycji usprawnień.

---

**Autor**: borubarczyk  
**Licencja**: Sprawdź plik LICENSE
