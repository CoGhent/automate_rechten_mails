import pandas as pd
import os

#script om op basis van een achternaam te zoeken naar de padnamen van bestanden op een specifieke plek

#set zoekfolder, broncsv en doelcsv (mogen ook hetzelfde zijn)

folder = r""
namenlijst = "adressenlijst.csv"
doelcsv = "adressenlijst.csv"

#kies extensies waarop gezocht moet worden
extensies = (".pdf", ".docx")
# open de namenlijst en maak dataframe met achternamen om op te zoeken
df_rechten = pd.read_csv(namenlijst)
achternamen = df_rechten["lastname"]

df_bestanden = pd.DataFrame(
    columns=["filepath"]
)

# nog te verbeteren: script maakt nu eerst een dataframe met alle bestanden in de folder
for root, dirs, files in os.walk(folder):
    # select file name
    for file in files:
        # check the extension of files
        if file.endswith(extensies):
            # print whole path of files
           naam = os.path.join(root, file)
           df_bestanden.loc[df_bestanden.shape[0]] = naam

print("Bestandspaden ophalen klaar")
df_bestanden.to_csv("bestandsnamen.csv")

#zoekt in het gemaakte dataframe van alle bestanden naar de pdf met achternaam en plaatst die in objectenlijst
#+ zoekt naar docx met achternaam en plaatst die in "toestemming"kolom
#haalt daarna de padnaam waar mails moeten worden opgeslagen op door de map waarin bijlagen staan toe te voegen aan kolom padnaam
# TO DO: verbeter script want wat met achternamen die meerdere keren voorkomen?


df_bestandsnamen = pd.read_csv("bestandsnamen.csv")

for i in range(len(df_rechten)):
    achternaam = df_rechten['lastname'].values[i]
    df_padnaam = df_bestandsnamen.loc[df_bestandsnamen["filepath"].str.contains(achternaam)]
    #maakt een dataframe met enkel de padnamen die achternaam bevatten
    pdf = df_padnaam.loc[df_padnaam["filepath"].str.contains(".pdf")]
    objectenlijst = str(pdf.iloc[0][1])
    docx = df_padnaam.loc[df_padnaam["filepath"].str.contains(".docx")]
    toestemming = str(docx.iloc[0][1])
    #haalt de waarde op van padnaam met .pdf in, door eerst dataframe te maken met enkel die pdf en dan waarde op te halen als string
    df_rechten["objectenlijst"].iloc[i] = objectenlijst
    df_rechten["toestemming"].iloc[i] = toestemming
    folder = os.path.dirname(toestemming)
    df_rechten["pathname"].iloc[i] = folder
    print(achternaam + " Bijlagen done")
    # TO DO: efficiënter maken van script, moet korter kunnen
df_rechten.to_csv(doelcsv)




