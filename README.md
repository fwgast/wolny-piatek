> [!TIP]
> Plik bardziej czytelny na githubie: https://github.com/fwgast/wolny-piatek

# Wolny Piatek

apka sluzy do wysylania maili na bazie pliku z excela, dziala tylko na windowsie

## Przedmowa, przemyslenia autora

### errata
witam, jesli ktos to czyta to wspolxzuje w chuj, bo to znaczy ze cos nie dziala, a w kodzie jest tak nasrane ze glowa mala, ale na swoje usprawiedliwienie powiem ze Windows (i szerzej Microsoft) to jebane gowno i tak prosta rzecz jak wysylanie maili nie moze zostac zrobiona przez pythona, tylko przez tego zjebanego power automate'a bo:
1. microsoft blokuje konta ktore wysylaja maile przez skrypt pythonowy i trzeba przesylac potem skan dowodu xdd
2. maile trafiaja do spamu
3. Authenticator nie ulatwia sprawy
4. App password na outlooku dziala tragicznie albo wcale

\+ apka nie jest w 100% user-friendly bo jakiekolwiek przydatne funkcje sa Premium :upside_down_face: np. triggerowanie flow


## Wymagania
- \>50 iq (wymog konieczny)
- Plik excel w formacie:  (kolory i wielkosc liter nie maja znaczenia, wazne tylko zeby gdy dostarczono dokument to bylo Y/y/T/t a gdy nie dostarczono N/n)
![alt text](https://github.com/fwgast/wolny-piatek/blob/main/additonal/excel2.png)
\
update: mea culpa, przez miscommunication wyszło ze po lewej jest jeszcze jedna kolumna z imieniem i nazwiskiem, w najnowszej wersji apka dziala niezalezie od tego czy w excelu jest ta kolumna czy nie

- Link z Power Automate: (w przeplywie -> info, tam gdzies powinien byc link podobny do tego):

```
ms-powerautomate:/console/flow/run?environmentid=Default-XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXX&workflowid=XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXX&source=Other
```
update: link jest nie potrzebny bo to funkcja premium :clown_face:
- zainstalowany python

## Krotki opis dzialania (TL;DR)
* apka pobiera dane z excela 
* przetwarza je
* tworzy plik csv, ktory bedzie uzywal power automate 
* otwiera power automate
* nawiguje po GUI
* odpala flow na power automate
* power automate wysyla maile bazujac na pliku csv (plik csv uzywa niestandardowego delimitera `, bo inaczej stopki sie pierdola)

## Debugowanie / uzywanie / "kompilacja"

> [!NOTE] 
> Manuale dla endusera/ow sa tez w osobnym pdfie
1. pobrac/sklonowac repo
2. w cmd: ( w odpowiednim folderze)
```powershell
python -m venv .venv
```
```powershell
.\.venv\Scripts\activate.bat
```
^ tu moze wyjebac sie na plecki^ in case: (jako admin)
*   ```powershell
    Set-ExecutionPolicy Unrestricted -Force
    ```
3. dalej w cmd:
```powershell
pip install -r requirements.txt
```
Od teraz kod powinien sie uruchamiac :))

### python script do .exe
zeby zrobic z tego apke w exe:
trzeba uzyc pyinstallera

```powershell
pyinstaller --icon="PATH\TO\additional\icon.ico" --onefile --add-data "PATH\TO\additional\dance.gif;." --add-data "PATH\TO\additional\icon.ico;." --windowed '.\Wolny Piatek v3.py'
```

### Struktura plikow
- generalnie prawie caly kod ktory cos robi jest w pliku 'Wolny Piatek v3.py', wiem ze nie powinno tak byc ale to jest moj kod i chuj mnnie obchodza konsekwencje
- niezbedny jest tez plik NIE_RUSZAC.py
- zeby kod zadziałał musi byc jeszcze plik config.py w ktorym znajduja sie zmienne: MAIL_BODY_SENSITIVE_DATA, PA_LINK, PATH_GETTER_PATH, SUBJECT, SENDER, FOOTER, ktorych nie ma na githubie z przyczyn oczywistych (dane osobowe, numer telefonu itd), jesli nie wiesz co ma byc w tych zmiennych to pytania kierowac do sami wiecie kogo, nie do mnie
- przed plikami warto dodawac: poniewaz czasem sie pierdola polskie znaki ale to wina raczej Outlooka niz pythona
*   ```python
    # -*- coding: utf-8 -*-
    ```
- w folderze additional:
  - plik dance.gif do loading screena  
  - plik icon.ico, no chyba wiadomo po co

update: zmienna PA_LINK jest niepotrzebna :slightly_smiling_face: , ale zostawiam moze kiedys to bedzie za free


### dlaczego tak a nie inaczej?
otoz moi drodzy: 
- robienie tego w 100% w power automate jest fizycznie niewykonalne przez tak krotki okres jaki pracowalem w tej firmie
- robienie tego w 100% w pythonie mija sie z celem jak wykazano wczesniej
- robienie tego na SharePoincie:
  1. to obraza dla intelektu i godnosci czlowieka
  2. sam proces approvali itd trwałby za dlugo
  3. z uwagi na to ze tylko jedna osoba (na dzien pisania tego) uzywałaby tej apki, nikt by tego nie klepnał
  4. nawet jezeli proces approvali skonczyłby sie pozytywnie, to:
        1. znając etos pracy działu IT w tej firmie, nikomu nie chciałoby sie jej robic 
        2. nawet jezeli ktos by sie podjał, prawdopodonie aplikacje robiłby rafal ;) (nie działałoby)

- dlaczego NIE_RUSZAC.py jest osobnym plikiem? po zrobieniu skryptu do execa, windows nie wpuszcza glob do windowsApps'ow bo to jakis folder chroniony czy cos, zmiana
- dlaczego uzywac CSV a nie JSON, YAML, XML? na dzien dzisiejszy power automate w wersji za free wspiera tylko csv :upside_down_face:
- dlaczego kod nie jest podzielony na osobne pliki? bo mi sie nie chciało, pozdrawiam
- dlaczego nie uzywalem dekoratorow? jak wyzej
- dlaczego importy sa niezoptymalizowane? jak wyzej
- dlaczego attachmenty sa tak dziwnie przekazywane i zapisywane? kwestia legacy, na poczatku dzialalo to na JSON'ie.
- dlaczego w takim razie nie zostało to zmienione? jak dwa wyzej
- dlaczego zmienna jest przekazywana do power automate przez zmienna srodowiskowa? sprobuj inaczej, nie majac premium :upside_down_face:
- dlaczego nie uzywam jakis triggerow do uruchamiania flow? bo sie nie da :slightly_smiling_face:


## Adnotacje koncowe, uwagi
* Szczegoly flow w folderze power_automate
* flow musi byc pierwsze od góry
* jesli nie działa nawigowanie po GUI w power automate to pewnie trzeba zmienic czas, bo prawdopodobnie nie da się wnioskowac tego dynamicznie
* jesli kiedys zmieni sie GUI power automate to pewnie sie wszystko wywali
* Jesli cos instalowales/as to na koniec:
```powershell
pip freeze > requirements.txt
```
* Plik CSV uzywa niestandardowego delimitera "`" bo inaczej stopki sie psują, trzeba to tez zmienic na power automate'cie
* Jesli po kliknieciu Send Mails przez usera, wystepuje lag, to normalne bo windows dlugo ustawia zmienne srodowiskowe
* trzeba uzupełnic config_template.py i zmienic nazwe na config.py
* nalezy zmienic lokalizacje pliku NIE_RUSZAC.py na wskazana w configu
* jesli kiedykolwiek trigerowanie flow przez link byłoby za darmo, nalezy:
1. odkomentowac tę linie:
```python
#webbrowser.open(PA_LINK)           #only on premium -_-
```
2. zakomentowac nastepne linie
3. odrobine zmienic manual

## Raportowanie bledow, zglaszanie reklamacji itd.
ze wszystkimi zazaleniami prosze pisac na:
```mail
chuj_mnie_to@fairwind.com
```

buziaki, \
<em>stasiek<em>
