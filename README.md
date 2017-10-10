
# Ceny_streamsoft

Aplikacja stworzona napisana w pythonie 3.

## Wprowadzenie

Służy do przetworzenia plików xls i txt na csv.
W arkusz kalkulacyjnym mamy możliwość wydrukowania fiszek cenowych wybranych towarów i i wybranej formie.

## Opis działania aplikacji

Opis działania aplikacji:
program wczytuje zawartość pliku towary.xls (baza towarowa aktualna), zawartość pliku towarys.csv(towary posortowane alfabetycznie malejąco po ostatniej modyfikacji), zawartość pliku cennik.xls (plik wygenerowany z programu sprzedażowego - wybrane towary z modułu magazynowego), zawartość pliku dokument.txt (plik wygenerowany z programu sprzedażowego - wybrane towary z dokumentu sprzedażowego), zawartość pliku nazwy00.txt (1-wiersz: początkowe nazwy towarów lub 2-wiersz: nazwy grup towarowych które mają być pominięte) .
Po przetworzeniu tych plików zapisuje w pliku dane_towar.txt.
Natomiast w pliku towaryss.csv umieszcza te towary posortowane alfabetycznie malejąco (arkusz kalkulacyjny pobierze i zmieni nazwę).

## uruchomienie

program uruchamiamy w lini poleceń (lub w pliku wsadowym):
python ceny_stream.py 

## Licencja



