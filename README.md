# MSWord_Keystroke-Counter
Mit Hilfe eines Makros für Microsoft-Word können die Tastaturanschläge pro Zeile bestimmt werden.


# Voraussetzungen
- **Software**
    - Microsoft Word (getestet ab Version 2010)
- **Dokumentenaufbau (siehe KeystrokeCounter_Template.docx)**
    - Das Dokument muss mindestens zwei Tabellen enthalten. Die zweite Tabelle sollte eine Zeile und zwei Spalten haben.
    - Der zu untersuchende Text muss in der zweiten Tabelle in der ersten Spalte/Zelle eingetragen werden. In die zweite Spalte/Zelle werden die Anschlagszahlen eingetragen.


# Installation
- **Quellcode herunterladen**    
- **Makro hinzufügen**
    - Empfohlen: Global (für alle Dokumente): Entwicklertools ⇒ Visual Basic ⇒ Normal ⇒ Module ⇒ Rechts Klick ⇒ Datei importieren ⇒ KeystrokeCounter.bas auswählen
    - Lokal (für aktuelles Dokument): Entwicklertools ⇒ Visual Basic ⇒ Project (*\<Dokumentname\>*) ⇒ Module ⇒ Rechts Klick auf Module ⇒ Datei importieren ⇒ KeystrokeCounter.bas auswählen  
      ⚠️ *Word-Dokument muss anschließend als \*.docm (Dokument mit Makros) gespeichert werden.* ⚠️        
      
- **Makro zu Menüband hinzufügen**
    Wordoptionen ⇒ Menüband anpassen ⇒ Befehle auswählen: Makros ⇒ gewünschte Registerkarte und Gruppe auswählen.
    (Ggf. Beschriftung und Icon anpassen) z. B.  
    ![grafik](https://github.com/mexterng/MSWord_Keystroke-Counter/assets/16732689/6be5b0ea-0c61-4581-8a52-b4b8acab78e2)
   

# Verwendung
- **countKeystrokes()**: Zählt die Tastaturanschläge pro Zeile und trägt diese in die zweite Spalte der zweiten Tabelle ein.


# Mögliche Veränderungen
- Verhalten von Zeilenumbrüchen: In Sub countKeystrokes() kann das Verhalten von Zeilenumbrüchen eingestellt werden. Mit queryIgnoreLineBreak kann das Popup-Fenster aktiviert (True) oder deaktiviert (False) werden. Das Standardverhalten (wenn keine Nutzerabfrage stattfindet) kann mit der Variable ignoreLineBreak eingestellt werden, ob Zeilenumbrüche ignoriert werden sollen (True) oder nicht (False)
- Anpassung der doppelten Anschlagszahlen: In der Funktion isDoubleKeystroke() sind alle Zeichen mit doppelten Tasttauranschlag einzutragen.
