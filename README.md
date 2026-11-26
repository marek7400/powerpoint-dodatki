# PowerPoint (2019 działa) DODATKI

v2.0 w przygotowaniu...

<img width="1920" height="199" alt="Makra PowerPoint 2019" src="https://github.com/user-attachments/assets/654c95b6-ff69-42e0-8b38-1a563efdd2c4" />

---

2 pliki wrzuca się je do (nazawaużytkownika-podaje się swoją)

C:\Users\nazawaużytkownika\AppData\Roaming\Microsoft\AddIns

potem PowerPoint-Plik-Opcje-Dodatki-Zarządzaj z menu wybiera się "Dodatki programu PowerPoint)-klikamy przycisk "Przejdź"

w oknie "Dodatki" zaznaczamy checkbox przy "makra" i klikamy przycisk "Zamknij"

Restartujemy PowerPoint. Powinno się pojawić dodatkowe menu o nazwie "Moje Makra"

<img width="1086" height="151" alt="1menu" src="https://github.com/user-attachments/assets/47e5b606-3f84-4bca-829d-84bb14420674" />

a w nim:

<img width="690" height="159" alt="2makra" src="https://github.com/user-attachments/assets/8e835751-b94f-43c9-94d8-0c481594bef2" />

1. ZapiszSzerokość - po zaznaczeniu elementu na slajdzie, zapisuje jego szerokość w pamięci

2. ZastosujSzerokość - po zaznaczeniu jednego (klik) lub wielu obiektów (SHIFT+klik) i wybraniu - zmienia ich szerokość do takiej jaka była zapamiętana

3. ZapiszWysokość - po zaznaczeniu elementu na slajdzie, zapisuje jego wysokość w pamięci

4. ZastosujWysokość - po zaznaczeniu jednego (klik) lub wielu obiektów (SHIFT+klik) i wybraniu - zmienia ich wysokość do takiej jaka była zapamiętana

przydatne do wyrównywania wymiarów dwóch lub więcej elementów hurtem


5. Wstaw plik z pamięcią (tworząc slajdy ZAWSZE warto je wstawiać korzystając z tego) - zapamiętuje wtedy dla każdego wstawianego elementu z dysku, jego dane

a jak potem na slajdzie klikniemy na taki element i klikniemy 


6. Odczytaj dane pliku to nam pokaże:

<img width="561" height="298" alt="6" src="https://github.com/user-attachments/assets/46d82adb-d5fb-4346-926c-01a1069bdb2a" />

Prawdziwa nazwa pliku (domyślnie po wstawieniu PowerPoint  JĄ NISZCZY), rozmiar pliku i z jakiego katalogu był wstawiany

Pomaga to czasem w szukaniu czego użyto do tworzenia slajdu.

<br>
7.  Eksportuj Slajdy (można eksportować do pliku pojedyncze slajdy, z zakresu od-do, wybrane lub wszystkie-zapisywany jest osobno każdy slajd)

Wybór co chcemy z prezentacji zapisać i jak (domyślnie PowerPoint nie ma takich opcji).

<img width="607" height="329" alt="7" src="https://github.com/user-attachments/assets/1784d203-db83-4bce-b2f5-f06bd6a0af98" />

po wyborze wybieramy folder zapisu i format zapisu (można wybrać kilka na raz) i jakość

<img width="617" height="295" alt="7b" src="https://github.com/user-attachments/assets/55c2871d-8d57-4407-b0e4-e3a54ff03a4e" />

W wypadku wybrania opcji 4 Wypakuj elementy slajdu do folderu -

zostanie stworzony katalog/katalogi o numerze wybranego slajdu/ów+nazwa prezentacji i do niego zostaną wypakowane WSZYSTKIE elementy ze slajdu (nawet te których nie widać, bo na slajdzie są zakryte innymi elementami)

przykład wypakowanego slajdu nr 1:

<img width="767" height="607" alt="8" src="https://github.com/user-attachments/assets/e97ed67b-d857-4f71-816c-9e02afbb5770" />

Wypakowuje wideo, audio, animacje gif, obrazki, tekst, kształty jako png.

*** Innym prostym sposobem dostania się do plików użytych w prezentacji jest zmiana rozszerzenia pliku z *.pptx na *.zip, i wypakowania plików z podkatalogu MEDIA ***
---

Przydatne sztuczki przy tworzeniu własnej zakładki ze skrótami w menu PowerPoint (Wstążka)

(do dopisania):

Office RibbonX Editor - edytor wstążki github.com/fernandreu/office-ribbonx-editor/releases/latest 

UWAGA: jeśli używasz <command idMso="Paste" onAction="PrzechwycWklejanie" />

trzeba użyć 
Public Sub PrzechwycWklejanie()

a nie
Sub PrzechwycWklejanie(control As IRibbonControl, ByRef cancelDefault As Boolean)
(com się kurna tego naszukał...)

Własne ikony - Office 2010 Add-In: Icons Gallery https://www.microsoft.com/en-us/download/details.aspx?id=21103&page=1 
ustawiamy w idMso=""

olevba - jak wyciągnąć makro  VBA z pliku *.ppam


pip install oletools

olevba "sciezka\do\pliku.ppam" > kod_vba.txt lub olevba -c "l:\b\ZapisywanieMediow.ppam" > kod_vba.txt lub olevba -c "sciezka\do\vbaProject.bin" > kod_vba.txt (bin wyciągnąć ze środka archiwum ppam--->zip)

Jak wyciągnąłeś kod z pliku *.ppam i go wkleisz w inny *pptx lub *.pptm- to na liście makr go nie zobaczysz-póki nie usuniesz z kodu linii (control As IRibbonControl) z nagłówków.
