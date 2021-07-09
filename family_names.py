import pandas as pd
import numpy as np
from selenium import webdriver
import time
import os



class family_names_webscraping:

    def __init__(self, namens_liste, driver_path):
        """
        

        Parameters
        ----------
        namens_liste : str-list
            Liste mit den Namen, die gesucht werden sollen.
        driver_path : str
            Pfad, wo dein Google Chrome Webdriver gespeichert ist.

        Returns
        -------
        None.

        """
        self.namens_liste = namens_liste
        
        # Die Startseite des Namen-Dictionary´s
        self.website = "https://www.oxfordreference.com"
        
        # Erstellt einen Driver mit Hilfe des Google Chrome Webdrivers.
        self.driver = webdriver.Chrome(driver_path)
        
        # Liste, in denen der Name und seine Beschreibung gespeichert werden,
        # falls diese im online Namen-Dictionary gefunden wurden.
        self.ergebnisse = []
        
        # Liste, in denen der Name gespeichert wird, falls dieser nicht gefunden
        # wurde.
        self.nicht_gefunden = []
        
        # Zählindex, der im Auge behält, bei dem wievielten Namen man gerade ist.
        self.i = 0





    def auf_website_gehen(self):
        # Ruft die Website oxfordreference.com auf.
        self.driver.get(self.website)        
        
                   
        
        
        
    def im_dokumnet_suchen(self, name):
        """
        

        Parameters
        ----------
        name : str
            Der Name, der gesucht werden soll.

        Returns
        -------
        None.

        """
        # Klickt auf das Suchenfeld auf der linken Seite und gibt dort den Namen
        # ein, der gesucht werden soll. Anschließend wird Enter gedrückt.
        inputElement = self.driver.find_element_by_id("searchWithinDocumentForm")
        inputElement = inputElement.find_element_by_id("q_within")
        inputElement.send_keys(name)
        inputElement.submit() 
        
        
        

        
    def klick_auf_erstes_ergebnis(self):
        # Man befindet sich auf der Seite mit den Suchergebnissen. Diese Funktion
        # klickt nun auf der erste Ergebnis.
        results_table = self.driver.find_element_by_id("searchContent")
        results_table = results_table.find_element_by_id("abstract_link1")
        results_table.click()





    def ergebnisse_auslesen(self, name):
        """
        

        Parameters
        ----------
        name : str
            Name, der gerade gesucht wird.

        Returns
        -------
        None.

        """
        # Man befindet sich auf der Seite des gesuchten Namens. Dort steht eine
        # Beschreibung des Namens und seine Herkunft. Diese Funktion ließt die
        # Beschreibung aus und speichert sie im self.ergebnisse String.
        content = self.driver.find_element_by_id("contentRoot")
        content2 = content.text
        text = content2
        self.ergebnisse.append([name, text])
        
        
        
        
        
    def write_to_excel(self):
        # Nach dem alle Namen aus der self.namens_liste gefunden wurden,
        # schreibt diese Funktion self.ergebnisse, self.nicht_gefunden und den
        # letzten Index bis wohin des Programm gekommen ist in ein Excel Sheet.
        # Hat man schon vorher Namen rausgesucht, werden diese nicht über-
        # schrieben, sondern unter die schon bestehende Tabelle gesetzt.
        ergebnisse_alt = pd.read_excel("output.xlsx", sheet_name="Ergebnisse")
        nicht_gefunden_alt = pd.read_excel("output.xlsx", sheet_name="Nicht Gefunden")
        letzter_index_alt = pd.read_excel("output.xlsx", sheet_name="Weiteres").values[0][0]
        
        ergebnisse = pd.DataFrame(self.ergebnisse, columns=["Name", "Description"])
        nicht_gefunden = pd.DataFrame(self.nicht_gefunden, columns=["Name"])
        letzter_index = pd.DataFrame([self.i + letzter_index_alt], columns=["Letzter Index"])
        
        ergebnisse = ergebnisse_alt.append(ergebnisse, ignore_index=True)
        nicht_gefunden = nicht_gefunden_alt.append(nicht_gefunden, ignore_index=True)
        
        with pd.ExcelWriter('output.xlsx') as writer:  
            ergebnisse.to_excel(writer, sheet_name='Ergebnisse', index=False)
            nicht_gefunden.to_excel(writer, sheet_name='Nicht Gefunden', index=False)
            letzter_index.to_excel(writer, sheet_name='Weiteres', index=False)
            
            
            
            
            
    def namens_liste_auswahl(self):
        # Wenn man mit dem letzten Programmdurchlauf z.B. die ersten 20 Namen
        # bearbeitet hat, würde die Namensliste so gekürzt, dass das Programm
        # nun bei dem 21. Namen anfangen würde.
        letzter_index = pd.read_excel("output.xlsx", sheet_name="Weiteres").values[0][0]
        self.namens_liste = self.namens_liste[letzter_index:]
        
    
    
    
    
    def leeres_output_sheet(self):
        # Falls noch kein Excel Sheet mit dem Namen "output.xlsx" vorhanden ist,
        # wird ein leeres erstellt, mit den jeweiligen Überschriften der Tabellen.
        if "output.xlsx" not in os.listdir():
            ergebnisse = pd.DataFrame(np.array([np.nan, np.nan]).reshape(1,2), columns=["Name", "Description"])
            nicht_gefunden = nicht_gefunden = pd.DataFrame(np.array(np.nan).reshape(1,1), columns=["Name"])
            letzter_index = pd.DataFrame([0], columns=["Letzter Index"])
            with pd.ExcelWriter('output.xlsx') as writer:  
                ergebnisse.to_excel(writer, sheet_name='Ergebnisse', index=False)
                nicht_gefunden.to_excel(writer, sheet_name='Nicht Gefunden', index=False)
                letzter_index.to_excel(writer, sheet_name='Weiteres', index=False)
            
        
        
        
        
    def main(self):
        self.leeres_output_sheet()
        self.namens_liste_auswahl()
        self.auf_website_gehen()
        
        # Pausen sind wichtig, damit das Programm keinen Fehler wirft, weil die
        # Seite nicht schnell genug geladen hat, oder deswegen ein Name nicht
        # gefunden werden konnte. Bei langsamen Internet sollte die Pause erhöht
        # werden.
        time.sleep(1)
        
        for name in self.namens_liste:
            self.im_dokumnet_suchen(name)
            time.sleep(1)
            
            # Wenn ein Name nicht gefunden werden sollte, springt man in den
            # except-Block, außerdem wird der Name zu self.nicht_gefunden 
            # hinzugefügt und man springt zurück zur Startseite des online 
            # Namen-Dictionary´s.
            try:
                self.klick_auf_erstes_ergebnis()
                time.sleep(1)
                self.ergebnisse_auslesen(name)
            except:
                self.nicht_gefunden.append(name)
                self.auf_website_gehen()
                time.sleep(1)
                
            # Printet aus, bei dem wievielten Namen man gerade ist, wieviel
            # Prozent man schon geschafft hat und wie der Name lautet.
            print(str(self.i) + " (" + str(int(((self.i+1)/len(self.namens_liste))*100)) + "%) " + str(name))
            
            # Setzte self.i ums eins hoch.
            self.i += 1










class country_filter:

    def __init__(self, ergebnisse, nicht_gefunden, country_adj):
        """
        

        Parameters
        ----------
        ergebnisse : DataFrame
            DataFrame mit den Namen die gefunden wurden + Beschreibung: "Name", "Description".
        nicht_gefunden : DataFrame
            DataFrame mit den Namen, die nicht gefunden wurden: "Name".
        country_adj : DataFrame
            DataFrame mit den Länder Adjektiven: "Länder Adjektive".

        Returns
        -------
        None.

        """
        self.ergebnisse = ergebnisse.values
        self.nicht_gefunden = nicht_gefunden.drop_duplicates()
        
        # Wird zu einer Menge umgebaut.
        self.country_adj = set(country_adj["Länder Adjektive"])
        
        
        
        
        
    def string_to_word_array(self, string):
        """
        

        Parameters
        ----------
        string : str
            Satz, der in seine Wörter zerlegt werden soll.

        Returns
        -------
        str-array
            Array bestehend aus den Wörtern des Strings.

        """
        # Zunächst werden die Satzzeichen gelöscht. Dann Wird eine Liste aus
        # den Wörtern des Strings erstellt.
        satzzeichen = [":", ",", ".", ";", "!", "?"]

        for sz in satzzeichen:
            string = string.replace(sz, "")
            
        return string.split()
        



        
    def set_to_string(self, s):
        """
        

        Parameters
        ----------
        s : set
            Menge mit den Gefundenen Herkunftsländern.

        Returns
        -------
        string : str
            String welcher die Herkunftsländer enthält.

        """
        # Die Wörter der Liste werden zu einem String kombiniert.
        string = ""
        for i, e in enumerate(s):
            string += e
            if i+1 != len(s):
                string += ", "
                
        return string
    
            
            
      
    
    def filter_countries(self):
        """
        

        Returns
        -------
        gefunden : DataFrame
            DataFrame mit den Namen, die gefunden wurden: "Name", "Herkunftsländer", "Beschreibung".

        """
        # Geht alle Wörter des Textes durch. Wenn ein Länder Adjektiv gefunden
        # wird, wird dieses in die Liste "gefunden" gepackt. Zum Schluss wird
        # daraus ein DataFrame gemacht und Duplikate gelöscht.
        gefunden = []
        for zeile in self.ergebnisse:
            lander = []
            satz = self.string_to_word_array(zeile[1])
            for w in satz:
                if w in self.country_adj:
                    lander.append(w)
            lander = self.set_to_string(set(lander))
            gefunden.append([zeile[0], lander, zeile[1]])
        gefunden = pd.DataFrame(gefunden, columns=["Name", "Herkunftsländer", "Beschreibung"]).drop_duplicates()
        
        return gefunden
    
    
    
    
    
    def alle_namen_df(self, gefunden):
        """
        

        Parameters
        ----------
        gefunden : DataFrame
            DataFrame mit den Namen, die gefunden wurden: "Name", "Herkunftsländer", "Beschreibung".

        Returns
        -------
        alle_namen : DataFrame
            DataFrame mit den Namen die gefunden wurden + die Namen, die nicht 
            gefunden wurden: "Name", "Herkunftsländer", "Beschreibung".

        """
        # Es wird ein DataFrame erstellt, der sich an der Liste mit den Namen
        # orientiert. Das DataFrame enthält also alle Namen und in den Spalten
        # daneben wird ergänzt, was das Herkunftsland ist und wie die 
        # Beschreibung des Namens lautet.
        names = pd.read_excel("names.xls", sheet_name="Sheet1")
        names = pd.DataFrame(np.ones((len(names), 1)), index=names["name"])
        
        gefunden.set_index("Name", inplace=True)
        alle_namen = names.join(gefunden).drop(0, axis=1)
        
        return alle_namen
    
    
    
    
    
    def write_to_excel(self, alle_namen, gefunden):
        """
        

        Parameters
        ----------
        alle_namen : DataFrame
            DataFrame mit den Namen die gefunden wurden + die Namen, die nicht 
            gefunden wurden: "Name", "Herkunftsländer", "Beschreibung".
        gefunden : DataFrame
            DataFrame mit den Namen, die gefunden wurden: "Name", "Herkunftsländer", "Beschreibung".

        Returns
        -------
        None.

        """
        # Die drei DataFrames werden in ein Excel Sheet geschrieben.
        with pd.ExcelWriter('FINAL.xlsx') as writer:
            alle_namen.to_excel(writer, sheet_name='Namen', index=True)
            gefunden.to_excel(writer, sheet_name='Gefunden', index=True)
            self.nicht_gefunden.to_excel(writer, sheet_name='Nicht Gefunden', index=False)
            
            
            
            
            
    def main(self):
        gefunden = self.filter_countries()
        alle_namen = self.alle_namen_df(gefunden)   
        self.write_to_excel(alle_namen, gefunden)










if __name__ == '__main__':
  
    ##############
    ### Part 1 ###
    ##############
    
    # Sucht die Namen aus der Namensliste im online Namens Dictionary und lädt
    # die Beschreibung herunter.
    
    # Wird aus der Excel Datei gelesen, die du mir geschickt hast. Diese sollte
    # daher im selben Verzeichnis liegen.
    namens_liste = list(pd.read_excel("names.xls", "Sheet1")["name"])
    
    # Pfad zu dem Ort, wo du den Google Chrome Webdriver gespeichert hast.
    driver_path = "**************************"
    
    # Erstellt das Objekt.
    x1 = family_names_webscraping(namens_liste, driver_path)
    
    # Main Programm der Klasse. Erstellt das Excel Sheet "output.xlsx".
    x1.main()
    
    # Wenn manuel abgebrochen wird, muss diese Zeile noch ausgeführt Werden.
    x1.write_to_excel()




    
    ##############
    ### Part 2 ###
    ##############

    # Filtere die Herkunftsländer der Namen aus der Beschreibung heraus.
    
    # "ergebnisse", "nicht_gefunden" und "country_adj" werden in Part 1 
    # erstellt und aus dem dort erstellten Excel Sheet "output.xlsx" gelesen.
    ergebnisse = pd.read_excel("output.xlsx", sheet_name="Ergebnisse")
    nicht_gefunden = pd.read_excel("output.xlsx", sheet_name="Nicht Gefunden")
    country_adj = pd.read_excel("country_adjectives.xlsx", sheet_name="Blatt 1")

    # Erstellt das Objekt
    x2 = country_filter(ergebnisse, nicht_gefunden, country_adj)
    
    # Main Programm der zweiten Klasse. Erstellt das Excel Sheet "FINAL.xlsx".
    x2.main()