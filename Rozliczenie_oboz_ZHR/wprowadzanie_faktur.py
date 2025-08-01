import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Wczytanie danych z pliku
plik = "rozliczenie_excel.xlsx"
df = pd.read_excel(plik)

# Konfiguracja przeglądarki
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=chrome_options)

# Otwarcie strony logowania
driver.get("https://ziher.zhr.pl/hal-owlkp/users/sign_in")
input("🔐 Zaloguj się do ZiHeR i naciśnij ENTER...")

# Mapy ID pól formularza
mapa_wydatkow = {
    "w": "entry_items_attributes_0_amount",  # Wyposażenie
    "m": "entry_items_attributes_1_amount",  # Materiały
    "j": "entry_items_attributes_2_amount",  # Wyżywienie
    "u": "entry_items_attributes_3_amount",  # Usługi
    "t": "entry_items_attributes_4_amount",  # Transport
    "d": "entry_items_attributes_5_amount",  # Delegacje
    "z": "entry_items_attributes_6_amount",  # Zakwaterowanie
    "wyn": "entry_items_attributes_7_amount" # Wynagrodzenia
}

mapa_wplywow = {
    "skladka":    "entry_items_attributes_0_amount", # Składka programowa
    "bon":        "entry_items_attributes_1_amount", # Bon turystyczny
    "obsluga":    "entry_items_attributes_2_amount", # Koszty obsługi HAL 3% (ze znakiem minus)
    "konto":      "entry_items_attributes_3_amount", # Wpłaty/wypłaty z konta
    "ko":         "entry_items_attributes_4_amount", # Dotacja KO
    "dotacja":    "entry_items_attributes_5_amount", # Dotacja inne
    "darowizna":  "entry_items_attributes_6_amount", # Darowizna
    "wlasne":     "entry_items_attributes_7_amount", # Środki własne
    "poobozowe":  "entry_items_attributes_8_amount", # Środki poobozowe (ze znakiem minus)
    "rohis":      "entry_items_attributes_9_amount" # ROHiS
}

# Grupowanie po numerze dokumentu
df["Numer dokumentu"] = df["Numer dokumentu"].astype(str)
grupy = df.groupby("Numer dokumentu", sort=False)

for dokument, grupa in grupy:
    if pd.isna(dokument) or dokument.strip() == "":
        print(f"🔚 Pominięto pusty numer dokumentu")
        continue

    pierwszy = grupa.iloc[0]

    try:
        data = (
            pierwszy["Data"].strftime("%Y-%m-%d")
            if isinstance(pierwszy["Data"], datetime)
            else pd.to_datetime(pierwszy["Data"]).strftime("%Y-%m-%d")
        )
        opis = str(pierwszy.get("Opis", "")).strip()
        ksiazka = str(pierwszy.get("Ksiazka", "")).strip().lower()
        rodzaj = str(pierwszy.get("Rodzaj", "")).strip().lower()

        # Wybór odpowiedniego formularza
        if ksiazka == "b" and rodzaj == "e":
            driver.get("https://ziher.zhr.pl/hal-owlkp/entries/new?is_expense=true&journal_id=727")
        elif ksiazka == "f" and rodzaj == "e":
            driver.get("https://ziher.zhr.pl/hal-owlkp/entries/new?is_expense=true&journal_id=726")
        elif ksiazka == "b" and rodzaj == "r":
            driver.get("https://ziher.zhr.pl/hal-owlkp/entries/new?is_expense=false&journal_id=727")
        elif ksiazka == "f" and rodzaj == "r":
            driver.get("https://ziher.zhr.pl/hal-owlkp/entries/new?is_expense=false&journal_id=726")
        else:
            print(f"⚠️ Nieznane wartości 'Ksiazka' lub 'Rodzaj' dla dokumentu {dokument}")
            continue

        # Oczekiwanie na załadowanie formularza
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "entry_date")))

        # Wypełnienie pól podstawowych
        pole_data = driver.find_element(By.ID, "entry_date")
        pole_data.clear()
        pole_data.send_keys(data)
        pole_data.send_keys(Keys.ENTER)

        driver.find_element(By.ID, "entry_document_number").send_keys(dokument)
        driver.find_element(By.ID, "entry_name").send_keys(opis)

        if rodzaj == "e":
            # Wydatki – wiele kategorii w jednej fakturze
            for _, row in grupa.iterrows():
                kwota = str(row.get("Kwota", "")).replace(" zł", "").replace(",", ".").strip()
                kategoria = str(row.get("Kategoria_wydatek", "")).strip().lower()

                if not kwota or not kategoria:
                    continue

                pole_id = mapa_wydatkow.get(kategoria)
                if pole_id:
                    driver.find_element(By.ID, pole_id).send_keys(kwota)
                else:
                    print(f"⚠️ Nieznana kategoria wydatku '{kategoria}' w dokumencie {dokument}")

        elif rodzaj == "r":
            # Wpływy – zakładamy jeden wpis
            kwota = str(pierwszy.get("Kwota", "")).replace(" zł", "").replace(",", ".").strip()
            kategoria = str(pierwszy.get("Kategoria_wplyw", "")).strip().lower()

            pole_id = mapa_wplywow.get(kategoria)
            if pole_id:
                driver.find_element(By.ID, pole_id).send_keys(kwota)
            else:
                print(f"⚠️ Nieznana kategoria wpływu '{kategoria}' w dokumencie {dokument}")

        # Zapisanie formularza
        driver.find_element(By.NAME, "commit").click()
        print(f"✅ Zapisano dokument: {dokument}")
        time.sleep(1.5)

    except Exception as e:
        print(f"❌ Błąd przy dokumencie {dokument}: {e}")
        continue

print("🎉 Wszystkie dokumenty zostały przetworzone.")
