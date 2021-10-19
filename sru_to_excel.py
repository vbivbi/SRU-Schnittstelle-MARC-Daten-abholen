#Für das Vorgehen werden einige Bibliotheken benötigt, die in Python importiert werden müssen 

#urllib.request --> braucht es um die "SRU-URLs" gut lesbar abzubilden.
import urllib.request
#pandas as pd --> pandas ist eine Bibliothek, die hilft Daten zu verwalten. Hier wird sie z.B. gebraucht um Excel auszulesen. 
import pandas as pd
#time --> time wird hier benötigt, um die Arbeitsschritte zwischenzeitlich zu pausieren, damit die SRU-Schnittstelle nicht zu viele Anfragen auf einmal bearbeiten muss. 
import time
#bs4 BeautifulSoup --> Beautfulsoup hilft, die MARC-XML in CSV umzuwandeln. 
from bs4 import BeautifulSoup
#csv --> Die csv-Bibliothek hilft, die csv-Datei zu erstellen. 
import csv


#Als ersters wird die benötigte Excel-Datei mit den MMS-IDs ausgelesen. Wichtig ist, dass das Excelfile im gleichen Ordner wie das Jupyter-Notebook oder das py-file gespeichert ist.
df = pd.read_excel ('stw_oktober_2021.xlsx')

#In einem weiteren Arbeitsschritt wird die benötigte Excelspalte ausgelesen. (In Anführungszeichen in der eckigen Klammer.)
#Die IDs werden ausgelesen und mit dem Befehlt "tolist" in einer Liste gespeichert.
id_list = df['MMS-ID'].tolist()

#Mit den IDs kann nun für jede Aufnahme ein Link auf die SRU-Schnittstelle erstellt werden. Der Link ist immer gleich aufgebaut, die MMS-ID soll am Schluss eingefügt werden.
sru_anfang = 'https://slsp-network.alma.exlibrisgroup.com/view/sru/41SLSP_UBS?version=1.2&operation=searchRetrieve&recordSchema=marcxml&query=alma.mms_id=='

#Ein csv "f" wird im Schreibmodus geöffnet. Achtung, dass encoding UTF-8 macht danach allenfalls in Excel Probleme und muss manuell auf UTF-16 gesetzt werden.
f = csv.writer(open("mein_auszug.csv", "w", encoding="UTF-8", newline=''))

#Hier werden die gewünschten Bezeichnungen der Spalten definiert. In Anführungszeichen und mit Komma abgetrennt.
f.writerow(["001", "650 a", "650 0", "AVA d", "710 0", "520 a", "245 a", "245 p", "856 u", "990 f", "008", "264 c"])



#Hier werden Funktionen erstellt, um die benötigsten Infos aus den MARC-Feldern zu holen.
#Es gibt unglaubliche viele Arten von MARC-Feldern (mit und ohne Unterfelder, wiederholbar nicht wiederholbar etc.)
#Hier wurden für zwei Arten Funktionen erstellt:
#1. Unwiederholbare Felder ohne Unterfeld (z.B. 001 mit MMS-ID)
#2. Felder, die mehrmals vorkommen können. Die Unterfelder sind aber pro Feld einmalig. (z.B. Schlagwörter in 6XX-Felder)
#Mit diesem Script sind also nicht alle Arten von Feldern abgedeckt. Mit einer weiteren Funktion kann dies aber hier ergänzt werden.


def unwiederholbare_ohne_unterfeld(feld_tag):
    #Bei beiden Feldern wurde try- und except eingefügt, weil ansonsten häufig aufgrund eines fehlerhaften Datensatzes das ganze Script abgebrochen wird. Aber daher ist auch empfehlenswert zu prüfen, ob wirklich alle Datensätze in mein CSV übernommen wurden.
    try:
        unwiederholbares_feld = soup.find(tag=feld_tag).get_text()
        return(unwiederholbares_feld)
    except:
        pass
    
def unterfeldurchgehen(feld_tag, unterfeld_code):
    try:
        feldname_list = [unterfeld.find(code=unterfeld_code).get_text() 
                         for unterfeld in soup.find_all(tag=feld_tag)]
        #wenn es mehere MARC-Feldern mit den gleichen Unterfeldern gibt, werden die Unterfelder mit einem | abgetrennt
        #weil es bei der Verarbeitung von csv in Excel immer Probleme gibt, wurden hier ; zu / umgewandelt. Dies ist aber fakultativ.
        beautiful_string = "| ".join(feldname_list).replace(";", "/")
        return beautiful_string
    except:
        pass


#Hier beginnt die eigentliche Programmschleife. Die einzelnen IDs aus der Excelliste werden laufend zum SRU-String zusammengefügt.
#Mit Hilfe der Bibliothek "urllib" werden die einzelnen Links ausgelesen.
for cur_id in id_list:
    sru_strings= sru_anfang + str(cur_id)
    
    sru_data = urllib.request.urlopen(sru_strings).read()
   
 #Damit die SRU-Schnittselle nicht überfodert wird, werden die Abfragen nur im 2-Sekunden-Takt ausgeführt.   
    time.sleep(2)
#Mit Hilfe der Bibliothek "BeautifulSoup" werden die erhaltenen Daten als XML erkannt und können ausgelesen werden.
    soup = BeautifulSoup(sru_data, 'xml')

    
 #Die letzte Zeile kann nach Bedarf angepasst werden. Hier werden nämlich die MARC-Felder definiert, die in der CSV-Datei auglesen werden sollen.
#Dafür werden die oben definierten Funktionen ausgeführt. Wichtig ist, dass die ausgewählten MARC-Felder mit den oben definierten Spaltennamen übereinstimmten.
    f.writerow([ unwiederholbare_ohne_unterfeld("001"), unterfeldurchgehen("650", "a"), unterfeldurchgehen("650", "0"), unterfeldurchgehen("AVA", "d"), 
                unterfeldurchgehen("710", "0"), unterfeldurchgehen("520", "a"), unterfeldurchgehen("245", "a"),  unterfeldurchgehen("245", "p"), unterfeldurchgehen("856", "u"), 
                unterfeldurchgehen("990", "f"),  unwiederholbare_ohne_unterfeld("008") , unterfeldurchgehen("264", "c") ])