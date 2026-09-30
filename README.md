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

- **Wczytywanie w tle**: UI nie blokuje się podczas pobierania danych
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
