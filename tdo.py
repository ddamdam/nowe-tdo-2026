import streamlit as st
import pandas as pd
import io
import re
import os
from typing import Dict, Optional, List, Tuple

# Używamy PyMuPDF (fitz) dla maksymalnej dokładności
import fitz

# --- Logika Ekstrakcji Danych (klasa OCR) ---

class ApartmentPlanOCR:
    def __init__(self):
        # Pełna, zaktualizowana lista kolumn
        self.columns = [
            "Nazwa pliku", "NR budynku", "NR Lokalu", "Nazwa robocza typu", "Typ",
            "Etap realizacyjny", "Strony świata (okna)", "Liczba kondygnacji",
            "Poddasze Uzytkowe/Nieuzytkowe",
            "Glowne wejscie do lokalu - kondygnacja",
            "Numer lokalu z dokumentacji budowlanej",
            "Powierzchnia z dokumentacji budowlanej dla GW (m2) z garazami",
            "Powierzchnia z dokumentacji budowlanej dla GW (m2) + pow. schodow",
            "Powierzchnia uzytkowa parter (m2)",
            "Powierzchnia uzytkowa pietro (m2)",
            "Powierzchnia pod sciankami i schodow Parter (m2)",
            "Powierzchnia pod sciankami i schodow Pietro (m2)",
            "Laczna powierzchnia lokalu poddasze uzytkowe (m2)",
            "Powierzchnia lacznie (m2)",
            "Powierzchnia ogrodek (m2)",
            "Ilosc pokoi",
            "Garderoba (0 brak, 1 jest)",
            "Kuchnia (0 brak, 1 jest)",
            "Aneks kuchenny (0 brak, 1 jest)",
            "Okno w kuchni lub aneksie (0 brak, 1 jest)",
            "Schowek (0 brak, 1 jest)",
            "Spiżarnia (0 brak, 1 jest)",
            "Pom. Gosp. (0 brak, 1 jest)",
            "Lokal połaczony z garażem (0 brak, 1 jest)",
            "Druga łazienka z prysznicem (0 brak, 1 jest)",
            "Gabinet (0 brak, 1 jest)",
            "Balkon (m2)",
            "Loggia (m2)",
            "Taras (m2)",
            "Kominek (0 brak, 1 jest)",
            "Oznaczenie miejsca postojowego na terenie/w garażu",
            "Oznaczenie dodatkowego miejsca postojowego",
            "Oznaczenie terenu do wyłącznego korzystania",
            "Wysokosc parteru (m)",
            "Wysokosc pietra (m)",
            "Wysokosc piwnicy (m2)"
        ]

    def extract_text_from_pdf(self, pdf_file_obj: io.BytesIO) -> Tuple[str, int]:
        """Używa metody "blocks" z PyMuPDF dla maksymalnej dokładności ekstrakcji tekstu."""
        full_text = ""
        page_count = 0
        try:
            pdf_bytes = pdf_file_obj.getvalue()
            with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
                page_count = doc.page_count
                for i, page in enumerate(doc):
                    blocks = page.get_text("blocks", sort=True)
                    page_text = " ".join([block[4].strip().replace('\n', ' ') for block in blocks if block[4].strip()])
                    full_text += f"\n--- STRONA {i + 1} ---\n" + page_text
            return full_text, page_count
        except Exception as e:
            st.error(f"Błąd podczas odczytu pliku PDF za pomocą PyMuPDF: {e}")
            return "", 0

    def extract_height_in_cm(self, text: str) -> str:
        """Wydobywa wysokość pomieszczenia (np. H=280) z tekstu strony i zwraca w CM."""
        match = re.search(r'H\s*=\s*(\d+)', text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_info_from_filename(self, filename: str) -> Tuple[Optional[str], Optional[str]]:
        """Analizuje nazwę pliku, aby wyciągnąć numer budynku i lokalu."""
        try:
            base_name = os.path.splitext(filename)[0]
            parts = base_name.split('_')
            if len(parts) > 1:
                apartment_code = parts[-1]
                building_num_match = re.search(r'(\d+)', apartment_code)
                building_num = building_num_match.group(1) if building_num_match else None
                return building_num, apartment_code
        except Exception:
            return None, None
        return None, None

    def extract_apartment_code_from_text(self, text: str) -> Optional[str]:
        """Wyodrębnia kod lokalu Z TEKSTU, jeśli nie uda się z nazwy pliku."""
        patterns = [r"Oznaczenie lokalu[:\s]+([\w\d/.-]+)"]
        matches = re.findall(patterns[0], text, re.IGNORECASE)
        for match in matches:
            if not re.match(r'\d{2}-\d{3}', match):
                return match.strip()
        return None

    def extract_building_number_from_text(self, text: str, apartment_code: Optional[str]) -> Optional[str]:
        """Wyodrębnia numer budynku Z TEKSTU."""
        match = re.search(r'(\d{2})-\d{3}\s+Warszawa', text)
        if match:
            return match.group(1)
        if apartment_code:
            match = re.search(r'^(\d+)', apartment_code)
            if match:
                return match.group(1)
        return None

    # --- POPRAWIONA FUNKCJA ---
    def _extract_value(self, text: str, patterns: List[str], default_value: str = "") -> str:
        """
        Znajduje wszystkie dopasowania i zwraca OSTATNIE z nich.
        Ulepszona wersja ignoruje teraz tekstowe prefiksy (np. "ok.") przed liczbą.
        """
        normalized_text = re.sub(r'\s+', ' ', text).replace(',', '.')
        for pattern in patterns:
            # Wzorzec dopasowuje pattern, potem ignoruje opcjonalne znaki i słowa (np. ':', 'ok.'), a na końcu wyciąga liczbę
            full_pattern = pattern + r'[:\s.a-zA-Z]*?(\d+\.?\d*)'
            matches = re.findall(full_pattern, normalized_text, re.IGNORECASE)
            if matches:
                return matches[-1].strip()
        return default_value

    def extract_surface_area(self, text: str, keywords: List[str]) -> str:
        """Generyczna funkcja do ekstrakcji powierzchni na podstawie słów kluczowych."""
        return self._extract_value(text, keywords)

    def extract_floor_areas(self, text: str) -> Dict[str, str]:
        """
        Poprawiona logika do ekstrakcji powierzchni z podziałem na strony (parter/piętro).
        """
        areas = {
            "parter_uzytkowa": "", "pietro_uzytkowa": "",
            "parter_scianki": "", "pietro_scianki": "",
            "laczna": ""
        }
        
        patterns_uzytkowa = [r"POWIERZCHNIA UŻYTKOWA"]
        patterns_scianki = [r"POWIERZCHNIA POD ŚCIANKAMI(?: I SCHODÓW)?"]
        # --- POPRAWIONY WZORZEC ---
        patterns_laczna = [r"Razem Łączna Powierzchnia Lokalu", r"Łączna Powierzchnia Lokalu"]

        sections = re.split(r'--- STRONA \d+ ---', text)

        if len(sections) > 1:
            page1_text = sections[1]
            areas["parter_uzytkowa"] = self._extract_value(page1_text, patterns_uzytkowa)
            areas["parter_scianki"] = self._extract_value(page1_text, patterns_scianki)

        if len(sections) > 2:
            page2_text = sections[2]
            areas["pietro_uzytkowa"] = self._extract_value(page2_text, patterns_uzytkowa)
            areas["pietro_scianki"] = self._extract_value(page2_text, patterns_scianki)
            
        # Szukamy łącznej powierzchni w całym tekście, używając poprawionych wzorców
        areas["laczna"] = self._extract_value(text, patterns_laczna)
        return areas


    def process_pdf(self, pdf_file_obj: io.BytesIO, filename: str) -> Tuple[Dict[str, str], str]:
        building_num, apartment_code = self.extract_info_from_filename(filename)
        text, page_count = self.extract_text_from_pdf(pdf_file_obj)
        if not text:
            return {col: "" for col in self.columns}, ""
        if not apartment_code:
            apartment_code = self.extract_apartment_code_from_text(text)
        if not building_num:
            building_num = self.extract_building_number_from_text(text, apartment_code)

        floor_areas = self.extract_floor_areas(text)
        
        # Ekstrakcja powierzchni z użyciem poprawionej funkcji _extract_value
        garden_area = self.extract_surface_area(text, [r"Powierzchnia ogrodu", r"ogród"])
        balkon_area = self.extract_surface_area(text, [r"balkon"])
        loggia_area = self.extract_surface_area(text, [r"loggia"])
        taras_area = self.extract_surface_area(text, [r"taras"])

        parter_height_cm = ""
        pietro_height_cm = ""
        sections = re.split(r'--- STRONA \d+ ---', text)
        if len(sections) > 1:
            parter_height_cm = self.extract_height_in_cm(sections[1])
        if len(sections) > 2:
            pietro_height_cm = self.extract_height_in_cm(sections[2])
            
        parter_height_m = f"{float(parter_height_cm) / 100:.2f}" if parter_height_cm.isdigit() else ""
        pietro_height_m = f"{float(pietro_height_cm) / 100:.2f}" if pietro_height_cm.isdigit() else ""

        result = {
            "Nazwa pliku": filename, "NR budynku": building_num or "", "NR Lokalu": apartment_code or "",
            "Nazwa robocza typu": "", "Typ": "", "Etap realizacyjny": "", "Strony świata (okna)": "",
            "Liczba kondygnacji": str(page_count) if page_count > 0 else "",
            "Poddasze Uzytkowe/Nieuzytkowe": "",
            "Glowne wejscie do lokalu - kondygnacja": "Parter" if page_count > 0 else "",
            "Numer lokalu z dokumentacji budowlanej": apartment_code or "",
            "Powierzchnia z dokumentacji budowlanej dla GW (m2) z garazami": "",
            "Powierzchnia z dokumentacji budowlanej dla GW (m2) + pow. schodow": "",
            "Powierzchnia uzytkowa parter (m2)": floor_areas.get("parter_uzytkowa", ""),
            "Powierzchnia uzytkowa pietro (m2)": floor_areas.get("pietro_uzytkowa", ""),
            "Powierzchnia pod sciankami i schodow Parter (m2)": floor_areas.get("parter_scianki", ""),
            "Powierzchnia pod sciankami i schodow Pietro (m2)": floor_areas.get("pietro_scianki", ""),
            "Laczna powierzchnia lokalu poddasze uzytkowe (m2)": "",
            "Powierzchnia lacznie (m2)": floor_areas.get("laczna", ""),
            "Powierzchnia ogrodek (m2)": garden_area,
            "Ilosc pokoi": "", "Garderoba (0 brak, 1 jest)": "", "Kuchnia (0 brak, 1 jest)": "",
            "Aneks kuchenny (0 brak, 1 jest)": "", "Okno w kuchni lub aneksie (0 brak, 1 jest)": "",
            "Schowek (0 brak, 1 jest)": "", "Spiżarnia (0 brak, 1 jest)": "", "Pom. Gosp. (0 brak, 1 jest)": "",
            "Lokal połaczony z garażem (0 brak, 1 jest)": "", "Druga łazienka z prysznicem (0 brak, 1 jest)": "",
            "Gabinet (0 brak, 1 jest)": "",
            "Balkon (m2)": balkon_area, "Loggia (m2)": loggia_area, "Taras (m2)": taras_area,
            "Kominek (0 brak, 1 jest)": "", "Oznaczenie miejsca postojowego na terenie/w garażu": "",
            "Oznaczenie dodatkowego miejsca postojowego": "", "Oznaczenie terenu do wyłącznego korzystania": "",
            "Wysokosc parteru (m)": parter_height_m,
            "Wysokosc pietra (m)": pietro_height_m,
            "Wysokosc piwnicy (m2)": ""
        }
        return result, text

# --- Interfejs Użytkownika Streamlit ---

st.set_page_config(layout="wide")
st.title("🤖 OCR do ekstrakcji danych v4.7") # Zwiększona wersja dla przejrzystości
st.subheader("Silnik ekstrakcji: PyMuPDF (metoda 'blocks', logika 'ostatniego dopasowania')")
st.write("Wgraj jeden lub więcej plików PDF. Numery lokali i budynków będą odczytywane głównie z nazw plików.")

ocr = ApartmentPlanOCR()

uploaded_files = st.file_uploader(
    "Wybierz pliki PDF (np. 'Projekt_Budynku_5A.pdf')",
    type="pdf",
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 Przetwórz pliki"):
        all_results = []
        progress_bar = st.progress(0, text="Oczekiwanie...")

        with st.spinner("Analizuję dokumenty..."):
            for i, uploaded_file in enumerate(uploaded_files):
                progress_bar.progress((i + 1) / len(uploaded_files), text=f"Przetwarzanie: {uploaded_file.name}")

                result, full_text = ocr.process_pdf(uploaded_file, uploaded_file.name)
                all_results.append(result)

                if full_text:
                    with st.expander(f"Pokaż/Ukryj pełny tekst odczytany z '{uploaded_file.name}'"):
                        st.text_area("Surowy tekst OCR:", full_text, height=300, key=f"text_area_{i}")
                else:
                    st.warning(f"Nie udało się odczytać tekstu z pliku '{uploaded_file.name}'.")

        if all_results:
            st.markdown("---")
            st.header("📊 Zbiorcze wyniki")
            st.success(f"Zakończono! Przetworzono {len(all_results)} plików.")
            df_results = pd.DataFrame(all_results, columns=ocr.columns)

            st.dataframe(df_results)

            st.session_state['df_results'] = df_results
        else:
            st.warning("Nie udało się wydobyć danych z żadnego z plików.")

if 'df_results' in st.session_state:
    df_to_download = st.session_state['df_results']

    @st.cache_data
    def to_excel(df: pd.DataFrame):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Wyniki_OCR')
        return output.getvalue()

    excel_data = to_excel(df_to_download)

    st.download_button(
        label="📥 Pobierz wyniki jako plik Excel",
        data=excel_data,
        file_name="wyniki_ocr.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheet.sheet"
    )