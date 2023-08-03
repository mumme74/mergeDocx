# Merge docx
This is a simple application to mers docx files to a single docx.
It was requested by a collegue.

It has also been a learning experience for me as I have never used tkinter before.
And my python skills are not that great.

### intended audience
People who now swedish as all user facing strings are in swedish.

## Svenska
Detta är en simpel applikation för att slå samman docx filer till en gemensam docx.
Det var efterfrågat av en kolla till mig.

Det har också var en läro process för mig då jag aldrig tidigare jobba med tkinter
Jag har också fått lära mig lite mer om python, då det inte är mitt bästa språk.

## Användning
Det finns både ett GUI interface och ett CLI

### python
Denna applikation använder python.
Steg för att få igång den:
1 Installera python om det inte redan finns på datorn.
1 Installera beroenden
1 starta applikationen

### Steg 1 python
För att kunna använda detta måste du installera python först.
I riktiga operativsystem finns detta installerat från början, i Windows kan man gå till marketplace för att ladda ned det.
[länk till python i windows store](https://apps.microsoft.com/store/detail/python-311/9NRWMJP3717K)

I macos och linux bör det vara installerat på förhand.

### Steg2 Installera beroenden
Kör scriptet *install_prerequisites.py* så kommer det hämta beroenden.
Ifall det strular kan du prova i kommandotolken (denna mapp som working directory)[^1]
```
pip install -r requirements.txt
```

### Steg3 starta programmet
Dubbelklicka på filen *mergeDocxGui.py* så öppnas programmet.
Om allt funkat ser du nu:
![Färsk start](/images/Screenshot1.png)


### Använda programmet
Först väljer du en mapp som innehåller flera docx filer.
Klicka på *Välj mapp* och bläddra fram till din mapp.
![Välja mapp](/images/Screenshot2.png)
Det kan se något annorlunda ut på din dator, jag skriver detta just nu på en Ubuntu maskin.

När du gjort detta visar programmet vilka docx filer som finns i mappen:
![Filer i mapp](/images/Screenshot3.png)

Sedan är det bara att trycka på *Slå ihop filer*.
![Sammanslagna filer](/images/Screenshot4.png)

### Sammanslagen fil
Som du ser finns nu en sammanslagen fil i mappen, som heter merged0.docx
![Merge fil i filhanteraren](/images/Screenshot5.png)

Filerna är sorterade i den ordning som operativsystemet har som standard.
Vanligen senast sparad först.


[^1]:Du kan behöva ladda ned pip (python package manager), beror på hur python är installerat på datorn.

### Intallera pip
Pip behövs för att kunna installera några beroenden (externa bibliotek som behövs).
Du kan readn ha pip installerar, dDetta steg är kanske inte nödvändigt
[länk för att kontrollera pip](https://packaging.python.org/en/latest/tutorials/installing-packages/)

### Kommando tolken (Power users)
Om det strular är det bra för felsökningen skull att använda terminal fönstret.
Öppna terminal fönstret och *cd* dig till detta programs mapp.

Cli programmet är en annan fil *mergeDocx.py*.
Det tar 1 argument som är mappen till dina docx filer.
```
exempel:
 hemmapp - docxmapp - Doc1.docx
         |          - Doc2.docx
         |          - Doc3.docx
         |          - etc...
         |
         - Hämtade filer - mergeDocx - mergeDocx.py
                                     - mergeDocxGui.py
                                     - etc ...

för att komma till hemma mapp:
  cd ~

för att komma till progmmet:
  cd ~/Hämtade\ filer/mergeDocx/

starta programmet och referera till docxmapp
  python3 mergeDocx.py ../../docxmapp/
 alt. på vissa datorer
  py mergeDocx.py ../../docxmapp/
```

Resultatet kan se ut så här:
![Terminal](/images/Screenshot6.png)

### Bug rapportering
Om du hittar en bugg, raportera gärna på github.com/mumme74/mergeDocx
