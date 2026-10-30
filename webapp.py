import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from openpyxl.styles import Font

def splits_adres(volledig_adres):
    match = re.match(r"^(.+)\s+(\d+)\s*(.*)$", volledig_adres)
    if match:
        return match.group(1).strip(), match.group(2).strip(), match.group(3).strip()
    return volledig_adres, "", ""

st.title("YES Energie - Factuurverwerker")
st.write("Upload een PDF-factuur om de PerfectView importbestanden te genereren.")

uploaded_file = st.file_uploader("Sleep de YES-factuur hierheen", type="pdf")

if uploaded_file is not None:
    # --- BASIS STRUCTUUR ---
    factuur_data = {
        # --- DEEL 1: Klantgegevens PV ---
        "Relatiecode": "", "Relatiebeheerder": "", "Status": "", "Klant": "", 
        "Bezoekadres straat": "", "Bezoekadres huisnummer": "", "Bezoekadres huisnummer toevoeging": "", "Bezoekadres postcode": "", "Bezoekadres plaats": "",
        "Postadres straat": "", "Postadres huisnummer": "", "Postadres huisnummer toevoeging": "", "Postadres postcode": "", "Postadres plaats": "",
        "Factuuradres straat": "", "Factuuradres huisnummer": "", "Factuuradres huisnummer toevoeging": "", "Factuuradres postcode": "", "Factuuradres plaats": "",
        "Aangemaakt op (relatie)": "", "Algemeen e-mailadres": "", "Algemeen telefoonnummer": "", "Financieel e-mailadres": "", 
        "Iban code": "", "Nummer KvK": "",
        
        # --- DEEL 2: Verbruiksgegevens PV ---
        "Activiteitencode ": "", "Organisatienaam": "", "Klantnummer ": "", "Leverancier ": "", "Ingangsdatum": "", "Einddatum": "", 
        "Aansluitadres straat": "", "Aansluitadres huisnummer": "", "Aansluitadres huisnummer toevoeging": "", "Aansluitadres postcode": "", "Aansluitadres plaats": "", 
        "E-mailadres": "", "Contactpersooncode ": "", "Roepnaam": "", "Achternaam": "", "Voorletter(s)": "", 
        "Elektra dal Levering 1": "", "Elektra enkel Levering 1": "", "Elektra piek Levering 1": "", "TL Elektra piek  ": "", "TL Elektra dal": "", "Verbruik gas 1": "", 
        "Stroomtarief dal ": "", "Stroomtarief piek ": "", "Gastarief ": "", "Stroomtarief teruglevering piek ": "", "Stroomtarief teruglevering dal ": "", 
        "Contractnummer ": "", "Type contract ": "", "EAN code electra": "", "EAN code gas": "", "Profielcode elektra ": "", "Profielcode gas ": ""
    }

    # --- 1. PDF TEKST INLADEN ---
    volledige_tekst = ""
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            tekst = page.extract_text()
            if tekst:
                volledige_tekst += tekst + "\n"
    
    regels = volledige_tekst.split('\n')
    
    # --- 2. LOGICA UITVOEREN ---
    for i, regel in enumerate(regels):
        # 1. NAAM ZOEKEN
        if "contractantnaam" in regel.lower():
            if not factuur_data["Klant"]: 
                rechter_deel = re.split(r"contractantnaam", regel, flags=re.IGNORECASE)[-1]
                schoon_naam = re.sub(r"[\:\|\-\–]", "", rechter_deel).strip()
                if schoon_naam:
                    factuur_data["Klant"] = schoon_naam
                    factuur_data["Organisatienaam"] = schoon_naam

        # 2. DATUMS ZOEKEN
        if "Startdatum factuurperiode" in regel:
            match = re.search(r"(\d{4}-\d{2}-\d{2})", regel)
            if match: factuur_data["Ingangsdatum"] = match.group(1)
            
        if "Einddatum factuurperiode" in regel:
            match = re.search(r"(\d{4}-\d{2}-\d{2})", regel)
            if match: factuur_data["Einddatum"] = match.group(1)

        # 3. FACTUURADRES
        if "Voor het adres:" in regel or "Voorhetadres:" in regel:
            if not factuur_data["Factuuradres postcode"]:
                schoon_adres = regel.replace("Voorhetadres:", "").replace("Voor het adres:", "").strip()
                match = re.search(r"(.*?)([0-9]{4}\s?[A-Za-z]{2})\s*(.*)", schoon_adres)
                if match:
                    straat, huisnr, toev = splits_adres(match.group(1).strip())
                    factuur_data["Factuuradres straat"] = straat
                    factuur_data["Factuuradres huisnummer"] = huisnr
                    factuur_data["Factuuradres huisnummer toevoeging"] = toev
                    factuur_data["Factuuradres postcode"] = match.group(2).strip()
                    factuur_data["Factuuradres plaats"] = match.group(3).strip()

        # 4. AANSLUITADRES
        if "AANSLUITADRES" in regel:
            if not factuur_data["Bezoekadres postcode"]:
                try:
                    ruwe_straat = regels[i+1].split(":")[-1].split("|")[-1].replace("Straatnaam", "").strip()
                    ruw_huisnr = regels[i+2].split(":")[-1].split("|")[-1].replace("Huisnummer + toevoeging", "").strip()
                    straat, huisnr, toev = splits_adres(f"{ruwe_straat} {ruw_huisnr}".strip())
                    postcode = regels[i+3].split(":")[-1].split("|")[-1].replace("Postcode", "").strip()
                    plaats = regels[i+4].split(":")[-1].split("|")[-1].replace("Plaats", "").strip()

                    for prefix in ["Bezoekadres", "Postadres", "Aansluitadres"]:
                        factuur_data[f"{prefix} straat"] = straat
                        factuur_data[f"{prefix} huisnummer"] = huisnr
                        factuur_data[f"{prefix} huisnummer toevoeging"] = toev
                        factuur_data[f"{prefix} postcode"] = postcode
                        factuur_data[f"{prefix} plaats"] = plaats
                except IndexError:
                    pass

        # 5. IBAN & KLANTNUMMER
        if "IBAN" in regel:
            match = re.search(r"(NL[0-9]{2}[A-Z]{4}[0-9]{10})", regel)
            if match: factuur_data["Iban code"] = match.group(1)

        if "Klantnummer:" in regel:
            if not factuur_data["Klantnummer "]:
                factuur_data["Klantnummer "] = regel.split("Klantnummer:")[-1].strip()

        # 6. EAN CODES
        if "EAN-code elektriciteit" in regel:
            match = re.search(r"(\d{18})", regel)
            if match: factuur_data["EAN code electra"] = match.group(1)
        if "EAN-code gas" in regel:
            match = re.search(r"(\d{18})", regel)
            if match: factuur_data["EAN code gas"] = match.group(1)

        # 7. VERBRUIK & TERUGLEVERING
        if "Totaal verbruik dal" in regel:
            match = re.search(r"(\d+)", regel.split("|")[-1])
            if match: factuur_data["Elektra dal Levering 1"] = match.group(1)
        if "Totaal verbruik piek" in regel:
            match = re.search(r"(\d+)", regel.split("|")[-1])
            if match: factuur_data["Elektra piek Levering 1"] = match.group(1)
        if "Gasverbruik" in regel:
            match = re.search(r"(\d+)", regel.split("|")[-1])
            if match: factuur_data["Verbruik gas 1"] = match.group(1)
            
        if "Stroom teruglevering daltarief" in regel and "kWh" in regel:
            match = re.search(r"(\d+)\s*kWh", regel)
            if match: factuur_data["TL Elektra dal"] = match.group(1)
        if "Stroom teruglevering normaaltarief" in regel and "kWh" in regel:
            match = re.search(r"(\d+)\s*kWh", regel)
            if match: factuur_data["TL Elektra piek  "] = match.group(1)

        # 8. CONTRACT INFO
        if "Naam leverancier" in regel:
            naam = regel.split("|")[-1].strip()
            if naam: factuur_data["Leverancier "] = naam
        if "Contracttype" in regel:
            type_c = regel.split("|")[-1].replace(":", "").strip()
            if type_c: factuur_data["Type contract "] = type_c
        if "Tarief gas (excl. btw)" in regel:
            match = re.search(r"([\d\.,]+)", regel.split("|")[-1])
            if match: factuur_data["Gastarief "] = match.group(1)

    # --- NA-VERWERKING: NAAM SPLITSEN ---
    volledige_naam = factuur_data["Klant"]
    if volledige_naam:
        delen = volledige_naam.split(" ", 1)
        if len(delen) == 2:
            if "." in delen[0]:
                factuur_data["Voorletter(s)"] = delen[0]
                factuur_data["Achternaam"] = delen[1]
            else:
                factuur_data["Roepnaam"] = delen[0]
                factuur_data["Achternaam"] = delen[1]
        else:
            factuur_data["Achternaam"] = volledige_naam

    # --- VEEG-SCRIPTJE VOOR LEGE VELDEN ---
    for veld, waarde in factuur_data.items():
        if waarde == "":
            factuur_data[veld] = "Handmatig invullen"

    # --- DATAFRAMES AANMAKEN ---
    klant_kolommen = [
        "Relatiecode", "Relatiebeheerder", "Status", "Klant", 
        "Bezoekadres straat", "Bezoekadres huisnummer", "Bezoekadres huisnummer toevoeging", "Bezoekadres postcode", "Bezoekadres plaats",
        "Postadres straat", "Postadres huisnummer", "Postadres huisnummer toevoeging", "Postadres postcode", "Postadres plaats",
        "Factuuradres straat", "Factuuradres huisnummer", "Factuuradres huisnummer toevoeging", "Factuuradres postcode", "Factuuradres plaats",
        "Aangemaakt op (relatie)", "Algemeen e-mailadres", "Algemeen telefoonnummer", "Financieel e-mailadres", 
        "Iban code", "Nummer KvK"
    ]

    verbruik_kolommen = [
        "Activiteitencode ", "Organisatienaam", "Klantnummer ", "Leverancier ", "Ingangsdatum", "Einddatum", 
        "Aansluitadres straat", "Aansluitadres huisnummer", "Aansluitadres huisnummer toevoeging", "Aansluitadres postcode", "Aansluitadres plaats", 
        "E-mailadres", "Contactpersooncode ", "Roepnaam", "Achternaam", "Voorletter(s)", 
        "Elektra dal Levering 1", "Elektra enkel Levering 1", "Elektra piek Levering 1", "TL Elektra piek  ", "TL Elektra dal", "Verbruik gas 1", 
        "Stroomtarief dal ", "Stroomtarief piek ", "Gastarief ", "Stroomtarief teruglevering piek ", "Stroomtarief teruglevering dal ", 
        "Contractnummer ", "Type contract ", "EAN code electra", "EAN code gas", "Profielcode elektra ", "Profielcode gas "
    ]

    klant_data_dict = {k: factuur_data[k] for k in klant_kolommen}
    verbruik_data_dict = {k: factuur_data[k] for k in verbruik_kolommen}

    df_klant = pd.DataFrame([klant_data_dict])
    df_verbruik = pd.DataFrame([verbruik_data_dict])

    # --- 3. EXCEL MAKEN EN DOWNLOADBAAR MAKEN ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_klant.to_excel(writer, index=False, sheet_name='Klantgegevens PV')
        df_verbruik.to_excel(writer, index=False, sheet_name='Verbruiksgegevens PV')
        
        rode_tekst = Font(color="FF0000", bold=True)
        for sheet_naam in ['Klantgegevens PV', 'Verbruiksgegevens PV']:
            worksheet = writer.sheets[sheet_naam]
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value == "Handmatig invullen":
                        cell.font = rode_tekst
                        
    output.seek(0)
    st.success("Factuur succesvol uitgelezen en de data is gekoppeld!")
    
    st.download_button(
        label="📥 Download PerfectView Importbestand (.xlsx)",
        data=output,
        file_name="Perfectview_Upload_Gesplitst.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )