from pathlib import Path
import pandas as pd 

def excel(input_paths, output_path):
    # change this
    # dit is de periode 
    PERIOD = "2023-02"

    # hier wordt de huidige directory verkregen met path.cwd()
    ROOT_DIR = Path.cwd()
    # hier wordt een nieuwe padvariabele gemaakt door ROOT_DIR te combineren met PERIOD
    PERIOD_DIR = ROOT_DIR / PERIOD
    # hier worden twee paden gecombineert namelijk PERIOD_DIR en de submap data. hiermee wordt een padvariabele gemaakt DATA_DIR gemaakt die wijst naar de "data"submap binnen de PERIOD_DIR
    DATA_DIR = PERIOD_DIR / "data"

    # exact goederenstromen 
    # load dataframe
    # hier wordt pandas gebruikt om een excel bestand te lezen. DATA_DIR is een object dat wijst naar een map. In dit geval wordt verwezen naar de map data. De periode staat in de string zodat de computer weet welk excel bestandje hij moet gebruiken. Skiprows betekent dat hij de eerste 13 rijen van het excel bestand negeert. Skipfooter betekent dat hij de laatste rij van het excel bestand negeert. 
    exact_goederenstromen_df = pd.read_excel(DATA_DIR / f"Exact_ICV_{PERIOD}.xlsx", skiprows=13, skipfooter=1)

    # rename columns
    # hier krijgt de variabele exact_goederenstromen_df een nieuwe waarde. Er zijn kolommen die een nieuwe waarde krijgen. Bedrag (Debet / Credit) wordt Bedrag debet en Unnamed: 7 wordt Bedrag credit. 
    exact_goederenstromen_df = exact_goederenstromen_df.rename(columns={
        "Bedrag (Debet / Credit)": "Bedrag debet",
        "Unnamed: 7": "Bedrag credit",
    })

    # filter by grootboekrekening
    # deze regel voegt een nieuwe kolom toe met de naam Grootboekrekening_code aan de dataframe exact_goederenstromen_df. Deze nieuwe kolom wordt ingevuld door de grootboekrekening kolom op te splitsten op basis van het teken "-" hier wordt alleen het eerste deel van groetboekrekening behouden ofwel het deel voor de -. 
    exact_goederenstromen_df["Grootboekrekening_code"] = exact_goederenstromen_df["Grootboekrekening"].str.split(" - ", expand=True)[0]
    # nadat de grootboekrekening_code kolom is gevuld met waarden uit de grootboekrekening kolom wordt deze omgezet naar het gegevenstype integer
    exact_goederenstromen_df["Grootboekrekening_code"] = exact_goederenstromen_df["Grootboekrekening_code"].astype(int)
    # hier wordt de data van de dataframe gefilterd zodat alleen de rijen met de waardes 3000, 5511, 7110, 7350 behouden blijven. 
    exact_goederenstromen_df = exact_goederenstromen_df.loc[exact_goederenstromen_df["Grootboekrekening_code"].isin({3000, 5511, 7110, 7350}), :]
    # hier wordt de kolom groetboekrekening_code verwijdert. 
    exact_goederenstromen_df = exact_goederenstromen_df.drop(columns=["Grootboekrekening_code"])

    # drop empty amounts
    # hier wordt in de dataframe exact_goederenstromen_df de rijen verwijdert waarbij de kolom Bedrag debet leeg is ofwel NaN is. 
    exact_goederenstromen_df = exact_goederenstromen_df.dropna(subset=["Bedrag debet"])

    # extract relatie code
    # deze regel voegt een nieuwe kolom toe met de naam Relatie_code aan de dataframe exact_goederenstromen_df. Deze nieuwe kolom wordt ingevuld door de relatie kolom op te splitsten op basis van het teken "-" hier wordt alleen het eerste deel van de relatie behouden ofwel het deel voor de -. 
    exact_goederenstromen_df["Relatie code"] = exact_goederenstromen_df["Relatie"].str.split(" - ", expand=True)[0]
    # nadat de relatie code kolom is gevuld met waarden uit de relatie kolom wordt deze omgezet naar het gegevenstype integer
    exact_goederenstromen_df["Relatie code"] = exact_goederenstromen_df["Relatie code"].astype(int)
    # dit is een controle regel
    exact_goederenstromen_df

    # exact overzicht boekingen 
    # hier wordt pandas gebruikt om een excel bestand te lezen. DATA_DIR is een object dat wijst naar een map. In dit geval wordt verwezen naar de map data. De periode staat in de string zodat de computer weet welk excel bestandje hij moet gebruiken. Skiprows betekent dat hij de eerste 8 rijen van het excel bestand negeert. 
    exact_boekingen_df = pd.read_excel(DATA_DIR / f"Exact_overzicht_boekingen_{PERIOD}.xlsx", skiprows=8)
    # deze regel verwijdert kolomen axis = 1 uit de dataframe waar alle waarden ontbreken (NaN) in die kolommen. How=all geeft aan dat een kolom alleen wordt verwijdert als alle waarden in die kolom ontbreken. Hierdoor worden onnodige kolommen verwijdert. 
    exact_boekingen_df = exact_boekingen_df.dropna(axis=1, how="all")
    # deze regel geeft de kolom uw ref. de nieuwe naam factuurnummer
    exact_boekingen_df = exact_boekingen_df.rename(columns={"Uw ref.": "Factuurnummer"})
    # dit is een controle regel
    exact_boekingen_df

    # exact overzicht relaties 
    # hier wordt pandas gebruikt om een excel bestand te lezen. DATA_DIR is een object dat wijst naar een map. In dit geval wordt verwezen naar de map data. De periode staat in de string zodat de computer weet welk excel bestandje hij moet gebruiken. Skiprows betekent dat hij de eerste 11 rijen van het excel bestand negeert. 
    exact_relaties_df = pd.read_excel(DATA_DIR / f"Exact_overzicht_relaties_{PERIOD}.xlsx", skiprows=11)
    # deze regel verwijdert kolomen axis = 1 uit de dataframe waar alle waarden ontbreken (NaN) in die kolommen. How=all geeft aan dat een kolom alleen wordt verwijdert als alle waarden in die kolom ontbreken. Hierdoor worden onnodige kolommen verwijdert. 
    exact_relaties_df = exact_relaties_df.dropna(axis=1, how="all")
    # deze regel geeft de kolom code de nieuwe naam relatie code en de kolom naam wordt relatie naam. 
    exact_relaties_df = exact_relaties_df.rename(columns={"Code": "Relatie code", "Naam": "Relatie naam"})
    # dit is een controle regel
    exact_relaties_df

    # landcodes
    # hier wordt pandas gebruikt om een tabel te lezen die zijn ingebed in een html pagina. de [0] wordt gebruikt om het eerste uit de dataframes die op deze webpagina staan te selecteren. 
    iso_3166_df = pd.read_html("https://nl.wikipedia.org/wiki/ISO_3166-1")[0]
    # hier worden de kolomen van de tabel van hierboven genaamd 3-letterig, nummer en ISO 3166-2 codes verwijdert
    iso_3166_df = iso_3166_df.drop(columns=["3-letterig", "Nummer", "ISO 3166-2 codes"])
    # hier wordt de kolom van de tabel van hierboven genaamd 2-letterig hernoemt naar land code 
    iso_3166_df = iso_3166_df.rename(columns={"2-letterig": "Land code"})
    # dit is een controle regel
    iso_3166_df

    # factuurnummer en land van herkomst toevoegen
    # Hier wordt een join operatie tussen de dataframe exact_goederenstromen_df en exact_boekingen_df op de kolom bkst.nr het resultaat wordt opgeslagen in exact_goederenstromen_df 
    exact_goederenstromen_df = exact_goederenstromen_df.join(exact_boekingen_df.set_index("Bkst.nr."), on="Bkst.nr.", rsuffix="_boeking")
    # Hier wordt een join operatie tussen de dataframe exact_goederenstromen_df en exact_relaties_df op de kolom relaties code het resultaat wordt opgeslagen in exact_goederenstromen_df 
    exact_goederenstromen_df = exact_goederenstromen_df.join(exact_relaties_df.set_index("Relatie code"), on="Relatie code", rsuffix="_relatie")
    # Hier wordt een join operatie tussen de dataframe exact_goederenstromen_df en iso_3166_df op de kolom land het resultaat wordt opgeslagen in exact_goederenstromen_df 
    exact_goederenstromen_df = exact_goederenstromen_df.join(iso_3166_df.set_index("Land"), on="Land")

    # deze regel selecteert een subset van kolommen uit de dataframe deze kolommen worden toegewezen aan de dataframe exact_goederenstromen_df. alleen de opgegeven kolomen blijven bestaan de rest van de kolommen wordt verwijdert. 
    exact_goederenstromen_df = exact_goederenstromen_df[["Boekjaar / Periode", "Datum", "Bkst.nr.", "Factuurnummer", "Relatie code", "Relatie naam", "Land", "Land code", "Omschrijving", "Bedrag debet"]]
    # de reset_index wordt gebruikt hierdoor wordt de huidige index verwijdert door drop=true te gebruiken, vervolgens wordt een nieuwe index toegevoegd. 
    exact_goederenstromen_df = exact_goederenstromen_df.reset_index(drop=True)
    # dit is een controle regel
    exact_goederenstromen_df

    # vendit overzicht ingekochte producten 
    # hier wordt pandas gebruikt om een excel bestand te lezen. DATA_DIR is een object dat wijst naar een map. In dit geval wordt verwezen naar de map data. De periode staat in de string zodat de computer weet welk excel bestandje hij moet gebruiken.
    vendit_ingekocht_df = pd.read_excel(DATA_DIR / f"Vendit_overzicht_ingekochte_producten_{PERIOD}.xlsx")
    # dit is een controle regel
    vendit_ingekocht_df

    # vendit producten export 
    # hier wordt pandas gebruikt om een excel bestand te lezen. DATA_DIR is een object dat wijst naar een map. In dit geval wordt verwezen naar de map data. De periode staat in de string zodat de computer weet welk excel bestandje hij moet gebruiken.
    vendit_producten_df = pd.read_excel(DATA_DIR / f"Vendit_producten_export_{PERIOD}.xlsx")

    # remove duplicates'
    # hier worden diplicaten verwijdert aan de hand van de kolommen levarancier en productnummer
    vendit_producten_df = vendit_producten_df.drop_duplicates(subset=["Leverancier", "Productnummer"])

    # extract HS code
    # uit de kolom product subomschrijving hier wordt met str.extract overeenkomsten gezocht met de opgegeven code deze code gaat dus binnen de kolom product subomschrijving naar HS codes gezocht. 
    vendit_producten_df["HS Code"] = vendit_producten_df["Product Subomschrijving"].str.extract(r"HS Code:\s?(\d+)")

    # we only want the first 8 digits
    # om de HS code correct te extraheren blijven alleen de eerste 8 tekens bestaan. 
    vendit_producten_df["HS Code"] = vendit_producten_df["HS Code"].str.slice(0, 8)
    # Met fillchar=0 wordt als de HS code minder dan 8 tekens lang is de code met nullen aangevuld. 
    vendit_producten_df["HS Code"] = vendit_producten_df["HS Code"].str.pad(8, fillchar="0")

    # dit is een controle regel 
    vendit_producten_df

    # zet hs code voor producten
    # dit is een for loop die door de unieke elementen van product_group loopt en deze sorteert.  
    for product_group in vendit_producten_df["Groep Omschrijving"].sort_values().unique():
        # Hier wordt een nieuwe dataframe gemaakt en worden alleen de rijen geselecteert waarin de kolom Groep Omschrijving gelijk is aan de huidige waarde van product group 
        vendit_product_group_df = vendit_producten_df.loc[vendit_producten_df["Groep Omschrijving"] == product_group, :]
        # hier wordt gecontroleert of er in de dataframe zowel rijen zijn met ontbrekende waarden NaN als ontbrekende waarden in de kolom HS code 
        if vendit_product_group_df["HS Code"].isna().any() and vendit_product_group_df["HS Code"].notna().any():
            # find the most common hs code for this product group
            # hier wordt de meest voorkomende HS code bepaalt door .mode met dropna = True worden eventuele ontbrekende waarden genegeerd. Dit wordt toegewezen aan product_group_hs_code. 
            product_group_hs_code = vendit_product_group_df["HS Code"].mode(dropna=True).values[0]
            # hier wordt de oorspronkelijke vendit_producten_df bijgewerkt voor de rijen waarin groep omschrijving overeenkomt met de waarde van product_groep en waarin HS code ontbreekt (NaN) deze rijen krijgen de waarde product_group_hs_code, dit betekent dat deze worden ingevuld door de meest voorkomende waarde voor die specifieke productgroep. 
            vendit_producten_df.loc[(vendit_producten_df["Groep Omschrijving"] == product_group) & (vendit_producten_df["HS Code"].isna()), "HS Code"] = product_group_hs_code

    # dit is een loop die loopt over indexen van de rijen waarin HS code ontbreekt binnen de dataframe vendit_producten_df. 
    for no_hs_code_idx in vendit_producten_df.loc[vendit_producten_df["HS Code"].isna(), :].index:
        # hier wordt de groep omschrijving van de huidige rij van de huidige rij waarin de HS code ontbreekt opgeslagen in de variabele product_group. 
        product_group = vendit_producten_df.loc[no_hs_code_idx, "Groep Omschrijving"]
        # dit begint een while loop die blijft draaien zolang de HS code ontbreekt. 
        while pd.isna(vendit_producten_df.loc[no_hs_code_idx, "HS Code"]):
            # als het volgende karakter niet voorkomt in de product_group dan breakt je uit de while loop.
            if " / " not in product_group:
                break
            # hier wordt het eerste deel van de splitsing van product_group. Het tweede deel wordt verwijdert. 
            product_group = product_group.rsplit(" / ", 1)[0]
            # hier wordt een nieuwe dataframe vendit_groep_df gemaakt door alleen de rijen te selecteren waarvan de Groeps Omschrijving begint met product_group 
            vendit_product_group_df = vendit_producten_df.loc[vendit_producten_df["Groep Omschrijving"].str.startswith(product_group), :]
            # deze voorwaarde controleert of er in het vendit_product_groep_df ten minste een niet-ontbrekende HS-code is. 
            if vendit_product_group_df["HS Code"].notna().any():
                # hier wordt de HS code van de huidige rij (waarin oorspronkelijk HS code ontbrak) ingesteld op de modus
                vendit_producten_df.loc[no_hs_code_idx, "HS Code"] = vendit_product_group_df["HS Code"].mode(dropna=True).values[0]

    # verwijder lege hs codes 
    # drop empty HS codes
    # hier worden de rijen die een lege HS Code hebben verwijdert. 
    vendit_producten_df = vendit_producten_df.dropna(subset=["HS Code"])

    # dit is een controle regel
    vendit_producten_df

    # voeg gewicht toe van PIM
    # hier wordt pandas gebruikt om een csv bestand te lezen. DATA_DIR is een object dat wijst naar een map. In dit geval wordt verwezen naar de map data. De periode staat in de string zodat de computer weet welk excel bestandje hij moet gebruiken. hier worden de namen van de kolomen toegevoegd. 
    pim_df = pd.read_csv(DATA_DIR / f"PIM_producten_gewicht_export_{PERIOD}.csv", names=["Merk", "Productnummer", "Product naam", "Gewicht"])
    # hier worden duplicaten verwijdert aan de hand van de kolommen levarancier en productnummer
    pim_df = pim_df.drop_duplicates(subset=["Merk", "Productnummer"])

    # drop any product weights that don't contain a number
    # deze regel selecteert alleen de rijen in het dataframe waarin de gewicht kolom ten minste een cijfer bevat.
    pim_df = pim_df.loc[pim_df["Gewicht"].str.contains(r"\d"), :]

    # clean up product weights
    # verwijdert voorloop en achterloop spaties
    pim_df["Gewicht"] = pim_df["Gewicht"].str.strip()
    # verwijdert het teken Â zonder onderscheidt te maken tussen hoofdletters en kleine letters
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace("Â", "", case=False)
    # verwijdert het teken 
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace("±", "")
    # verwijdert het woord Ca ongeacht hoofdletters of kleine letters 
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace("Ca.", "", case=False)
    # verwijdert het woord gewicht
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace("Gewicht:", "")
    # verwijdert overmatige spaties en vervangt ze door enkele spaties
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace(r"\s+", " ")
    # vervangt punten door komma's
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace(".", ",", regex=False)
    # verwijdert opnieuw voorloop en achterloop spaties 
    pim_df["Gewicht"] = pim_df["Gewicht"].str.strip()

    # hier worden getallen met een optionele komma geextraheert gevolgd door optionele spaties end de eenheden gram, gr of g
    pim_df["Gewicht (gr)"] = pim_df["Gewicht"].str.extract(r"(\d+,?\d*)\s?(?:gram|gr|g)")
    # hier worden getallen met een optionele komma geextraheert gevolgd door optionele spaties end de eenheid kg
    pim_df["Gewicht (kg)"] = pim_df["Gewicht"].str.extract(r"(\d+,?\d*)\s?(?:kg)")
    # hier wordt een update gemaakt met waarden van de extracties met een optionele komma aan het einde van de kolom gewicht
    pim_df["Gewicht (kg)"].update(pim_df["Gewicht"].str.extract(r"(\d+,?\d*)$", expand=False))

    # drop original "Gewicht" column
    # hier wordt de gewicht kolom verwijdert omdat deze niet meer nodig is. 
    pim_df = pim_df.drop(columns=["Gewicht"])

    # convert to numeric
    # hier worden de eventuele komma's vervangen door punten en wordt de uiteindelijke kolomlijst (gr) omgezet in numerieke waarden
    pim_df["Gewicht (gr)"] = pd.to_numeric(pim_df["Gewicht (gr)"].str.replace(",", "."))
    # hier worden de eveuntuele komma's vervangen door punten en wordt de uiteindelijke kolomlijst (kg) omgezet in numerieke waarden
    pim_df["Gewicht (kg)"] = pd.to_numeric(pim_df["Gewicht (kg)"].str.replace(",", "."))
    # hier wordt de gewicht kolom geupdatet met de waarde van de gewicht(gr) kolom gedeeld door 1000 waardoor deze kg worden. 
    pim_df["Gewicht (kg)"].update(pim_df["Gewicht (gr)"] / 1000)

    # drop any products that don't have a weight
    # hier worden rijen verwijdert waarin de gewicht(kg) kolom ontbreekt. 
    pim_df = pim_df.dropna(subset=["Gewicht (kg)"])
    # drop grams column
    # hier wordt de kolom gewicht (gr) verwijdert 
    pim_df = pim_df.drop(columns=["Gewicht (gr)"])

    # controle regel
    pim_df

    # hier wordt de vendit_producten_df dataframe samengevoegd met de pim_df dataframe op basis van de kolommen merk en productnummer hierdoor wordt er in vendit_producten_df gewicht toegevoegd.
    vendit_producten_df = vendit_producten_df.join(
        pim_df.set_index(["Merk", "Productnummer"])["Gewicht (kg)"],
        on=["Merk", "Productnummer"]
    )
    # controle regel
    vendit_producten_df

    # hier wordt de dataframe geëxporteerd naar een excel bestand die met behulp van variabele PERIOD gevonden wordt. 
    vendit_producten_df.to_excel(PERIOD_DIR / f"Vendit_producten_export_{PERIOD}_met_hs_codes_en_gewicht.xlsx", index=False)

    # voeg HS code toe aan vendit ingekochte producten 
    # hier wordt de vendit_ingekocht_df dataframe samengevoegd met de vendit_ingekocht_df dataframe op basis van de kolommen leverancier en productnummer hierdoor wordt er in vendit_producten_df HS code en gewicht toegevoegd. 
    vendit_ingekocht_df = vendit_ingekocht_df.join(
        vendit_producten_df.set_index(["Leverancier", "Productnummer"])[["HS Code", "Gewicht (kg)"]],
        on=["Leverancier", "Poduct Nr."]
    )
    # controle regel
    vendit_ingekocht_df

    # voeg pakbonnummer toe aan exact dataframe om overeen te komen met vendit 
    # in deze lus wordt een overeenkomst gezocht tussen de kolomen factuurnummer in exact_goederenstroom_df en pakbonnummer in vendit_ingekocht_df. als er een overeenkomst wordt gevonden wordt het pakbonnumer toegevoegd aan exact_goederen_stroom om de pakbonnummers van beide dataframes aan elkaar te matchen. 
    for i, factuur_nummer in enumerate(exact_goederenstromen_df["Factuurnummer"]):
        for j, pakbon_nummer in enumerate(vendit_ingekocht_df["Pakbonnummer"]):
            if str(pakbon_nummer).endswith(str(factuur_nummer)):
                exact_goederenstromen_df.loc[i, "Pakbonnummer (Vendit)"] = pakbon_nummer
    # controle regel
    exact_goederenstromen_df

    # hier wordt de dataframe geëxporteerd naar een excel bestand die met behulp van variabele PERIOD gevonden wordt. 
    exact_goederenstromen_df.to_excel(PERIOD_DIR / f"Exact_goederenstromen_geconsolideerd_{PERIOD}.xlsx", index=False)

    # CBS rapport opmaken 
    # hier wordt een nieuwe dataframe cbs_df gemaakt door rijen te selecteren uit vendit_ingekocht_df waarvan de pakbonnummer voorkomt in de lijst pakbonnumer(vendit) uit exact_goederenstroom_df.
    cbs_df = vendit_ingekocht_df.loc[vendit_ingekocht_df["Pakbonnummer"].isin(exact_goederenstromen_df["Pakbonnummer (Vendit)"]), :]
    # hier wordt de land code van cbs_df ingevuld op basis van overeenkomstige pakbonnummers in exact_goederenstromen_df. De Land code wordt opgehaald uit exact_goederenstroom_df op basis van overeenkomstige pakbonnummers en toegevoegd aan cbs_df. 
    cbs_df["Land code"] = cbs_df["Pakbonnummer"].apply(lambda pakbon_nummer: exact_goederenstromen_df.loc[exact_goederenstromen_df["Pakbonnummer (Vendit)"] == pakbon_nummer, "Land code"].values[0])
    # controle regel
    cbs_df

    # het cbs_df dataframe wordt geëxporteerd naar een excel bestand met een naam die gebaseerd wordt op variabele PERIOD.
    cbs_df.to_excel(PERIOD_DIR / f"CBS_export_{PERIOD}.xlsx", index=False)