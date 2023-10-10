from pathlib import Path
import pandas as pd

def excel(input_paths, output_path):
    print("ik ben in begin.py")
    PERIOD = "2023-04"

    ROOT_DIR = Path.cwd()

    PERIOD_DIR = ROOT_DIR / PERIOD

    DATA_DIR = PERIOD_DIR / "data"
    print("stap 1: het inlezen van exact_goederenstromen_df")
    exact_goederenstromen_df = pd.read_excel(input_paths[0], skiprows=13, skipfooter=1)
    exact_goederenstromen_df = exact_goederenstromen_df.rename(columns={
        "Bedrag (Debet / Credit)": "Bedrag debet",
        "Unnamed: 7": "Bedrag credit",
    })

    exact_goederenstromen_df["Grootboekrekening_code"] = exact_goederenstromen_df["Grootboekrekening"].str.split(" - ", expand=True)[0]
    exact_goederenstromen_df["Grootboekrekening_code"] = exact_goederenstromen_df["Grootboekrekening_code"].astype(int)
    exact_goederenstromen_df = exact_goederenstromen_df.loc[exact_goederenstromen_df["Grootboekrekening_code"].isin({3000, 5511, 7110, 7350}), :]
    exact_goederenstromen_df = exact_goederenstromen_df.drop(columns=["Grootboekrekening_code"])

    exact_goederenstromen_df = exact_goederenstromen_df.dropna(subset=["Bedrag debet"])

    exact_goederenstromen_df["Relatie code"] = exact_goederenstromen_df["Relatie"].str.split(" - ", expand=True)[0]
    exact_goederenstromen_df["Relatie code"] = exact_goederenstromen_df["Relatie code"].astype(int)
    exact_goederenstromen_df

    print("stap 2: het inlezen van exact_boekingen_df")
    exact_boekingen_df = pd.read_excel(input_paths[1], skiprows=8)
    exact_boekingen_df = exact_boekingen_df.dropna(axis=1, how="all")
    exact_boekingen_df = exact_boekingen_df.rename(columns={"Uw ref.": "Factuurnummer"})
    exact_boekingen_df

    print("stap 3: het inlezen van exact_relaties_df")
    exact_relaties_df = pd.read_excel(input_paths[2], skiprows=11)
    exact_relaties_df = exact_relaties_df.dropna(axis=1, how="all")
    exact_relaties_df = exact_relaties_df.rename(columns={"Code": "Relatie code", "Naam": "Relatie naam"})
    exact_relaties_df

    print("stap 4: het inlezen van tabel uit wikipedia pagina")
    iso_3166_df = pd.read_html("https://nl.wikipedia.org/wiki/ISO_3166-1")[0]
    iso_3166_df = iso_3166_df.drop(columns=["3-letterig", "Nummer", "ISO 3166-2-codes"])
    iso_3166_df = iso_3166_df.rename(columns={"2-letterig": "Land code"})
    iso_3166_df

    print("stap 5: voeg exact_goederenstromen_df samen met exact_boeking_df, exact_relaties_df, iso_3166_df")
    exact_goederenstromen_df = exact_goederenstromen_df.join(exact_boekingen_df.set_index("Bkst.nr."), on="Bkst.nr.", rsuffix="_boeking")
    exact_goederenstromen_df = exact_goederenstromen_df.join(exact_relaties_df.set_index("Relatie code"), on="Relatie code", rsuffix="_relatie")
    exact_goederenstromen_df = exact_goederenstromen_df.join(iso_3166_df.set_index("Land"), on="Land")

    exact_goederenstromen_df = exact_goederenstromen_df[["Boekjaar / Periode", "Datum", "Bkst.nr.", "Factuurnummer", "Relatie code", "Relatie naam", "Land", "Land code", "Omschrijving", "Bedrag debet"]]
    exact_goederenstromen_df = exact_goederenstromen_df.reset_index(drop=True)
    exact_goederenstromen_df

    print("stap 6: het inlezen van vendit_ingekocht_df")
    vendit_ingekocht_df = pd.read_excel(input_paths[4])
    vendit_ingekocht_df

    print("stap 7: het inlezen van vendit_producten_df")
    vendit_producten_df = pd.read_excel(input_paths[5])
    vendit_producten_df = vendit_producten_df.drop_duplicates(subset=["Leverancier", "Productnummer"])
    vendit_producten_df["HS Code"] = vendit_producten_df["Product Subomschrijving"].str.extract(r"HS Code:\s?(\d+)")
    vendit_producten_df["HS Code"] = vendit_producten_df["HS Code"].str.slice(0, 8)
    vendit_producten_df["HS Code"] = vendit_producten_df["HS Code"].str.pad(8, fillchar="0")

    vendit_producten_df

    print("stap 8: zet hs codes voor producten in dataframe")
    for product_group in vendit_producten_df["Groep Omschrijving"].sort_values().unique():
        vendit_product_group_df = vendit_producten_df.loc[vendit_producten_df["Groep Omschrijving"] == product_group, :]
        if vendit_product_group_df["HS Code"].isna().any() and vendit_product_group_df["HS Code"].notna().any():
            product_group_hs_code = vendit_product_group_df["HS Code"].mode(dropna=True).values[0]
            vendit_producten_df.loc[(vendit_producten_df["Groep Omschrijving"] == product_group) & (vendit_producten_df["HS Code"].isna()), "HS Code"] = product_group_hs_code

    print("stap 9: verwijdert rijen waar de kolom hs codes leeg is")
    for no_hs_code_idx in vendit_producten_df.loc[vendit_producten_df["HS Code"].isna(), :].index:
        product_group = vendit_producten_df.loc[no_hs_code_idx, "Groep Omschrijving"]
        while pd.isna(vendit_producten_df.loc[no_hs_code_idx, "HS Code"]):
            if " / " not in product_group:
                break
            product_group = product_group.rsplit(" / ", 1)[0]
            vendit_product_group_df = vendit_producten_df.loc[vendit_producten_df["Groep Omschrijving"].str.startswith(product_group), :]
            if vendit_product_group_df["HS Code"].notna().any():
                vendit_producten_df.loc[no_hs_code_idx, "HS Code"] = vendit_product_group_df["HS Code"].mode(dropna=True).values[0]

    vendit_producten_df = vendit_producten_df.dropna(subset=["HS Code"])

    vendit_producten_df

    print("stap 10: het inlezen, opschonen en omrekenen van pim_df")
    pim_df = pd.read_csv(input_paths[3], names=["Merk", "Productnummer", "Product naam", "Gewicht"])
    pim_df = pim_df.drop_duplicates(subset=["Merk", "Productnummer"])

    pim_df = pim_df.loc[pim_df["Gewicht"].str.contains(r"\d"), :]

    pim_df["Gewicht"] = pim_df["Gewicht"].str.strip()
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace("Â", "", case=False)
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace("±", "")
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace("Ca.", "", case=False)
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace("Gewicht:", "")
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace(r"\s+", " ")
    pim_df["Gewicht"] = pim_df["Gewicht"].str.replace(".", ",", regex=False)
    pim_df["Gewicht"] = pim_df["Gewicht"].str.strip()

    pim_df["Gewicht (gr)"] = pim_df["Gewicht"].str.extract(r"(\d+,?\d*)\s?(?:gram|gr|g)")
    pim_df["Gewicht (kg)"] = pim_df["Gewicht"].str.extract(r"(\d+,?\d*)\s?(?:kg)")
    pim_df["Gewicht (kg)"].update(pim_df["Gewicht"].str.extract(r"(\d+,?\d*)$", expand=False))

    pim_df = pim_df.drop(columns=["Gewicht"])

    pim_df["Gewicht (gr)"] = pd.to_numeric(pim_df["Gewicht (gr)"].str.replace(",", "."))
    pim_df["Gewicht (kg)"] = pd.to_numeric(pim_df["Gewicht (kg)"].str.replace(",", "."))
    pim_df["Gewicht (kg)"].update(pim_df["Gewicht (gr)"] / 1000)

    pim_df = pim_df.dropna(subset=["Gewicht (kg)"])

    pim_df = pim_df.drop(columns=["Gewicht (gr)"])

    pim_df
    
    print("stap 11: samenvoegen van vendit_producten_df en pim_df")
    vendit_producten_df = vendit_producten_df.join(
        pim_df.set_index(["Merk", "Productnummer"])["Gewicht (kg)"],
        on=["Merk", "Productnummer"]
    )

    vendit_producten_df

    print("stap 12: hier wordt de vendit_producten_df geëxporteerd naar het output_path")
    vendit_producten_df.to_excel(output_path, index=False)
    # vendit_producten_df.to_excel(PERIOD_DIR / f"Vendit_producten_export_{PERIOD}_met_hs_codes_en_gewicht.xlsx", index=False)

    print("stap 13: hier wordt de vendit_ingekocht_df samengevoegd met vendit_producten_df")
    vendit_ingekocht_df = vendit_ingekocht_df.join(
        vendit_producten_df.set_index(["Leverancier", "Productnummer"])[["HS Code", "Gewicht (kg)"]],
        on=["Leverancier", "Poduct Nr."]
    )

    vendit_ingekocht_df
    
    print("stap 14: matchen dataframes")
    for i, factuur_nummer in enumerate(exact_goederenstromen_df["Factuurnummer"]):
        for j, pakbon_nummer in enumerate(vendit_ingekocht_df["Pakbonnummer"]):
            if str(pakbon_nummer).endswith(str(factuur_nummer)):
                exact_goederenstromen_df.loc[i, "Pakbonnummer (Vendit)"] = pakbon_nummer

    exact_goederenstromen_df

    print("stap 15: exporteren excel bestand")
    exact_goederenstromen_df.to_excel(output_path, index=False)
    # exact_goederenstromen_df.to_excel(PERIOD_DIR / f"Exact_goederenstromen_geconsolideerd_{PERIOD}.xlsx", index=False)

    print("stap 16: cbs rapport opmaken")
    cbs_df = vendit_ingekocht_df.loc[vendit_ingekocht_df["Pakbonnummer"].isin(exact_goederenstromen_df["Pakbonnummer (Vendit)"]), :]
    cbs_df["Land code"] = cbs_df["Pakbonnummer"].apply(lambda pakbon_nummer: exact_goederenstromen_df.loc[exact_goederenstromen_df["Pakbonnummer (Vendit)"] == pakbon_nummer, "Land code"].values[0])

    cbs_df

    cbs_df.to_excel(output_path, index=False)