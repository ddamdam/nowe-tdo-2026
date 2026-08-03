import streamlit as st
import pandas as pd
import io
import re
import os
from pathlib import Path
from typing import Dict, Optional, List, Tuple

import pymupdf as fitz
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Wiersz 1 = sekcja (kolor), wiersz 2 = kto uzupełnia, wiersz 3 = nagłówki, dane od wiersza 4.
SECTION_LABELS = {
    "lokalizacja": "LOKALIZACJA",
    "dokumentacja": "DOKUMENTACJA BUDOWLANA",
    "rozporzadzenie": "BILANS POWIERZCHNI ZGODNIE Z ROZPORZĄDZENIEM",
    "ustawa": "BILANS POWIERZCHNI ZGODNIE Z USTAWĄ",
    "zalaczniki": "DANE Z ZAŁĄCZNIKÓW TECHNICZNYCH",
    "marketing": "MARKETING",
}

SECTION_COLORS = {
    "LOKALIZACJA": "92D050",  # zielony
    "DOKUMENTACJA BUDOWLANA": "FDB913",  # żółty ciemny / złoty
    "BILANS POWIERZCHNI ZGODNIE Z ROZPORZĄDZENIEM": "FCE4D6",  # morelowy jasny
    "BILANS POWIERZCHNI ZGODNIE Z USTAWĄ": "BDD7EE",  # niebieski
    "DANE Z ZAŁĄCZNIKÓW TECHNICZNYCH": "D9E6BC",  # zielony jasny (jak w próbce)
    "MARKETING": "F2F2F2",  # szary
}

EXCEL_COLUMNS: List[Dict] = [
    # --- LOKALIZACJA ---
    {"idx": 1, "source": "ocr", "section": "lokalizacja", "header": "NR budynku"},
    {"idx": 2, "source": "ocr", "section": "lokalizacja", "header": "NR Lokalu"},
    {"idx": 3, "source": "ocr", "section": "lokalizacja", "header": "Nazwa robocza typu"},
    {"idx": 4, "source": "architekt", "section": "lokalizacja", "header": "Typ"},
    {"idx": 5, "source": "ocr", "section": "lokalizacja", "header": "Etap"},
    {"idx": 6, "source": "architekt", "section": "lokalizacja", "header": "Strony świata (okna)"},
    {"idx": 7, "source": "ocr", "section": "lokalizacja", "header": "Liczba kondygnacji"},
    {"idx": 8, "source": "ocr", "section": "lokalizacja", "header": "Poddasze Użytkowe/Nieużytkowe"},
    {"idx": 9, "source": "ocr", "section": "lokalizacja", "header": "Główne wejscie do lokalu - kondygnacja"},
    # --- DOKUMENTACJA BUDOWLANA ---
    {"idx": 10, "source": "ocr", "section": "dokumentacja", "header": "Numer lokalu z dokumentacji budowlanej"},
    {"idx": 11, "source": "architekt", "section": "dokumentacja", "header": "Powierzchnia z dokumentacji budowlanej dla GW (m2) bez garaży\n*UWAGA: powierzchnia do zweryfikowania po otrzymaniu proj. budowlanego"},
    # --- BILANS ZGODNIE Z ROZPORZĄDZENIEM ---
    {"idx": 12, "source": "architekt", "section": "rozporzadzenie", "header": "Powierzchnia użytkowa adm. parter (m2)"},
    {"idx": 13, "source": "architekt", "section": "rozporzadzenie", "header": "Powierzchnia użytkowa adm. piętro (m2)"},
    {"idx": 14, "source": "architekt", "section": "rozporzadzenie", "header": "Łączna powierzchnia lokalu piętro 2 użytkowe (m2)"},
    {"idx": 15, "source": "ocr", "section": "rozporzadzenie", "header": "Powierzchnia użytkowa adm. suma (m2)"},
    {"idx": 16, "source": "architekt", "section": "rozporzadzenie", "header": "Powierzchnia łącznie  (m2)"},
    # --- BILANS ZGODNIE Z USTAWĄ ---
    {"idx": 17, "source": "ocr", "section": "ustawa", "header": "Powierzchnia użytkowa parter (m2)"},
    {"idx": 18, "source": "ocr", "section": "ustawa", "header": "Powierzchnia użytkowa piętro (m2)"},
    {"idx": 19, "source": "ocr", "section": "ustawa", "header": "Powierzchnia użytkowa piętro 2 (m2)"},
    {"idx": 20, "source": "ocr", "section": "ustawa", "header": "Powierzchnia użytkowa suma (m2)"},
    # --- DANE Z ZAŁĄCZNIKÓW TECHNICZNYCH ---
    {"idx": 21, "source": "ocr", "section": "zalaczniki", "header": "Powierzchnia ogródek (m2)"},
    {"idx": 22, "source": "ocr", "section": "zalaczniki", "header": "Ilość pokoi"},
    {"idx": 23, "source": "ocr", "section": "zalaczniki", "header": "Garderoba(0 brak, 1 jest)"},
    {"idx": 24, "source": "ocr", "section": "zalaczniki", "header": "Kuchnia (0 brak, 1 jest)"},
    {"idx": 25, "source": "ocr", "section": "zalaczniki", "header": "Aneks kuchenny (0 brak, 1 jest)"},
    {"idx": 26, "source": "ocr", "section": "zalaczniki", "header": "Okno w kuchni lub aneksie (0 brak, 1 jest)"},
    {"idx": 27, "source": "ocr", "section": "zalaczniki", "header": "Schowek (0 brak, 1 jest)"},
    {"idx": 28, "source": "ocr", "section": "zalaczniki", "header": "Spiżarnia (0 brak, 1 jest)"},
    {"idx": 29, "source": "ocr", "section": "zalaczniki", "header": "Pom. Gosp. (0 brak, 1 jest)"},
    {"idx": 30, "source": "ocr", "section": "zalaczniki", "header": "Lokal połączony z garażem (0 brak, 1 jest)"},
    {"idx": 31, "source": "ocr", "section": "zalaczniki", "header": "Druga łazienka z prysznicem (0 brak, 1 jest)"},
    {"idx": 32, "source": "ocr", "section": "zalaczniki", "header": "Gabinet (0 brak, 1 jest)"},
    {"idx": 33, "source": "ocr", "section": "zalaczniki", "header": "Balkon (m2)"},
    {"idx": 34, "source": "ocr", "section": "zalaczniki", "header": "Loggia (m2)"},
    {"idx": 35, "source": "ocr", "section": "zalaczniki", "header": "Taras (m2)"},
    {"idx": 36, "source": "marketing", "section": "zalaczniki", "header": "Oznaczenie miejsca postojowego na terenie/w garażu"},
    {"idx": 37, "source": "marketing", "section": "zalaczniki", "header": "Oznaczenie dodatkowego miejsca postojowego"},
    {"idx": 38, "source": "marketing", "section": "zalaczniki", "header": "Oznaczenie terenu do wyłącznego korzystania"},
    {"idx": 39, "source": "ocr", "section": "zalaczniki", "header": "Wysokość parteru (m)"},
    {"idx": 40, "source": "ocr", "section": "zalaczniki", "header": "Wysokość piętra (m)"},
    {"idx": 41, "source": "ocr", "section": "zalaczniki", "header": "Wysokość piwnicy (m2)"},
    # --- MARKETING (na końcu) ---
    {"idx": 42, "source": "marketing", "section": "marketing", "header": "OSIEDLE"},
    {"idx": 43, "source": "marketing", "section": "marketing", "header": "Miasto/Miejscowość"},
    {"idx": 44, "source": "marketing", "section": "marketing", "header": "Dzielnica"},
    {"idx": 45, "source": "marketing", "section": "marketing", "header": "OPIS LOKALU"},
    {"idx": 46, "source": "marketing", "section": "marketing", "header": "URL"},
    {"idx": 47, "source": "marketing", "section": "marketing", "header": "GARAŻ 1"},
    {"idx": 48, "source": "marketing", "section": "marketing", "header": "CENA GARAŻU 1"},
    {"idx": 49, "source": "marketing", "section": "marketing", "header": "GARAŻ 2"},
    {"idx": 50, "source": "marketing", "section": "marketing", "header": "CENA GARAŻU 2"},
    {"idx": 51, "source": "marketing", "section": "marketing", "header": "MIEJSCE POSTOJOWE 1"},
    {"idx": 52, "source": "marketing", "section": "marketing", "header": "CENA MP 1"},
    {"idx": 53, "source": "marketing", "section": "marketing", "header": "MIEJSCE POSTOJOWE 2"},
    {"idx": 54, "source": "marketing", "section": "marketing", "header": "CENA MP 2"},
    {"idx": 55, "source": "marketing", "section": "marketing", "header": "NAROŻNY"},
    {"idx": 56, "source": "marketing", "section": "marketing", "header": "CENA"},
    {"idx": 57, "source": "marketing", "section": "marketing", "header": "CENA M2"},
    {"idx": 58, "source": "marketing", "section": "marketing", "header": "FEATURES"},
]

# Fonty Lato w części załączników (np. Arkady Walendów) mają zepsute ToUnicode —
# cyfry wychodzą jako litery IPA/cyrylica. Mapowanie odtworzone z wartości w PDF.
_BROKEN_DIGIT_MAP = str.maketrans(
    {
        "\u018f": "0",  # Ə
        "\u0191": "2",  # Ƒ
        "\u0192": "3",  # ƒ
        "\u0193": "4",  # Ɠ
        "\u0194": "5",  # Ɣ
        "\u0475": "6",  # ѵ
        "\u0195": "7",  # ƕ
        "\u0476": "8",  # Ѷ
        "\u0196": "9",  # Ɩ
        "\u0137": ",",  # ķ
    }
)


TEMPLATE_BASENAME = "ostateczny - kopia.xlsx"


def source_label(source: str) -> str:
    """Tekst w komórce wiersza „kto uzupełnia” nad danymi."""
    labels = {
        "ocr": "OCR",
        "architekt": "Architekt",
        "marketing": "marketing",
    }
    return labels.get(source, "")


def column_display_names() -> List[str]:
    return [
        f"{get_column_letter(c['idx'])} | {c['header'].replace(chr(10), ' ')}"
        for c in EXCEL_COLUMNS
    ]


def section_label(section_key: str) -> str:
    return SECTION_LABELS.get(section_key, SECTION_LABELS["marketing"])


def section_row_values() -> List[str]:
    return [section_label(col["section"]) for col in EXCEL_COLUMNS]


def _section_groups() -> List[Tuple[str, List[int]]]:
    groups: List[Tuple[str, List[int]]] = []
    for col_def in EXCEL_COLUMNS:
        label = section_label(col_def["section"])
        if groups and groups[-1][0] == label:
            groups[-1][1].append(col_def["idx"])
        else:
            groups.append([label, [col_def["idx"]]])
    return [(label, indices) for label, indices in groups]


def _apply_section_fill(ws, row: int, col_idx: int, section_name: str) -> None:
    color = SECTION_COLORS.get(section_name, "F2F2F2")
    cell = ws.cell(row, col_idx)
    cell.fill = PatternFill("solid", fgColor=color)
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _column_width_for_header(header: str) -> float:
    """Szersze kolumny dla długich nagłówków; tekst i tak się zawija."""
    plain = header.replace("\n", " ").strip()
    length = len(plain)
    if length >= 60:
        return 28
    if length >= 35:
        return 22
    if length >= 20:
        return 16
    if length >= 12:
        return 12
    return 10


def apply_sheet_layout(ws, data_row_count: int = 0) -> None:
    """Szerokości kolumn + zawijanie tekstu w nagłówkach i danych."""
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000")
    header_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_def in EXCEL_COLUMNS:
        col_idx = col_def["idx"]
        ws.column_dimensions[get_column_letter(col_idx)].width = _column_width_for_header(
            col_def["header"]
        )
        for row_idx in (1, 2, 3):
            cell = ws.cell(row_idx, col_idx)
            cell.alignment = header_align
            cell.border = header_border

    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 60

    for row_idx in range(4, 4 + max(data_row_count, 0)):
        ws.row_dimensions[row_idx].height = 18
        for col_def in EXCEL_COLUMNS:
            cell = ws.cell(row_idx, col_def["idx"])
            cell.alignment = data_align


def write_column_metadata(ws) -> None:
    """Wiersz 1 = sekcja, wiersz 2 = kto uzupełnia, wiersz 3 = nagłówki."""
    max_idx = max(col_def["idx"] for col_def in EXCEL_COLUMNS)
    used_indices = {col_def["idx"] for col_def in EXCEL_COLUMNS}

    for merge in list(ws.merged_cells.ranges):
        if merge.min_row == 1 and merge.max_row == 1:
            ws.unmerge_cells(str(merge))

    for col_idx in range(1, max_idx + 1):
        for row_idx in (1, 2, 3):
            ws.cell(row_idx, col_idx, None)
            ws.cell(row_idx, col_idx).fill = PatternFill()
            ws.cell(row_idx, col_idx).font = Font()

    for col_def in EXCEL_COLUMNS:
        section_name = section_label(col_def["section"])
        _apply_section_fill(ws, 1, col_def["idx"], section_name)
        ws.cell(2, col_def["idx"], source_label(col_def["source"]) or None)
        ws.cell(3, col_def["idx"], col_def["header"])

    for section_name, indices in _section_groups():
        start_col, end_col = indices[0], indices[-1]
        ws.cell(1, start_col, section_name)
        if end_col > start_col:
            ws.merge_cells(
                start_row=1, start_column=start_col, end_row=1, end_column=end_col
            )
        for col_idx in range(start_col, end_col + 1):
            if col_idx in used_indices:
                _apply_section_fill(ws, 1, col_idx, section_name)

    apply_sheet_layout(ws)


def _script_dir() -> Path:
    return Path(__file__).resolve().parent


def create_excel_template(path: Path) -> None:
    """Tworzy szablon (wiersz 2 = źródło, wiersz 3 = nagłówki, dane od wiersza 4)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Arkusz1"
    write_column_metadata(ws)
    wb.save(path)


def ensure_excel_template() -> str:
    """Szablon obok skryptu — tworzy plik, jeśli go nie ma (Streamlit Cloud)."""
    script_dir = _script_dir()
    for name in (TEMPLATE_BASENAME, "ostateczny — kopia.xlsx"):
        path = script_dir / name
        if path.is_file():
            return str(path)
    matches = sorted(script_dir.glob("ostateczny*kopia*.xlsx"))
    if matches:
        return str(matches[0])
    path = script_dir / TEMPLATE_BASENAME
    create_excel_template(path)
    return str(path)


def find_excel_template() -> str:
    return ensure_excel_template()


def build_full_row(result: Dict[str, str]) -> List[str]:
    """Pełny wiersz danych: OCR wypełnione, marketing/architekt puste."""
    row = []
    for col_def in EXCEL_COLUMNS:
        if col_def["source"] == "ocr":
            row.append(result.get(col_def["header"], ""))
        else:
            row.append("")
    return row


def source_row_values() -> List[str]:
    return [source_label(c["source"]) for c in EXCEL_COLUMNS]


def _column_section_map() -> Dict[str, str]:
    return {
        name: section_label(col["section"])
        for name, col in zip(column_display_names(), EXCEL_COLUMNS)
    }


def results_to_dataframe(results: List[Dict[str, str]]) -> pd.DataFrame:
    columns = column_display_names()
    rows = (
        [section_row_values()]
        + [source_row_values()]
        + [build_full_row(result) for result in results]
    )
    return pd.DataFrame(rows, columns=columns)


def style_results_dataframe(df: pd.DataFrame):
    section_by_column = _column_section_map()

    def color_columns(col: pd.Series) -> List[str]:
        color = SECTION_COLORS.get(section_by_column.get(col.name, ""), "F2F2F2")
        return [
            f"background-color: #{color}; white-space: normal; word-wrap: break-word"
        ] * len(col)

    styled = df.style.apply(color_columns, axis=0)
    try:
        styled = styled.set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        ("white-space", "normal"),
                        ("word-wrap", "break-word"),
                        ("vertical-align", "top"),
                    ],
                }
            ]
        ).set_properties(**{"white-space": "normal", "word-wrap": "break-word"})
    except Exception:
        pass
    return styled


def export_results_to_excel(results: List[Dict[str, str]]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Arkusz1"
    write_column_metadata(ws)

    for row_idx, result in enumerate(results, start=4):
        for col_def, value in zip(EXCEL_COLUMNS, build_full_row(result)):
            ws.cell(row_idx, col_def["idx"], value if value != "" else None)

    apply_sheet_layout(ws, data_row_count=len(results))

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


class ApartmentPlanOCR:
    def __init__(self):
        self.columns = [c["header"] for c in EXCEL_COLUMNS]

    def empty_row(self) -> Dict[str, str]:
        return {header: "" for header in self.columns}

    def _safe_float_sum(self, *args: str) -> str:
        total = 0.0
        for val in args:
            if isinstance(val, str) and val.strip():
                try:
                    total += float(val.replace(",", ".").strip())
                except (ValueError, TypeError):
                    continue
        return f"{total:.2f}" if total > 0 else ""

    @staticmethod
    def _decode_broken_digits(text: str) -> str:
        """Naprawia cyfry z błędnym ToUnicode (Lato w Arkady Walendów itd.)."""
        if not text:
            return text
        decoded = text.translate(_BROKEN_DIGIT_MAP)
        decoded = decoded.replace("ohĺ", "ok.").replace("lѹ", "m²")
        # Indeksy pomieszczeń typu 0ĺ1 → 0.1 (ĺ nie mapujemy globalnie na kropkę)
        decoded = re.sub(r"(\d)ĺ(\d)", r"\1.\2", decoded)
        return decoded

    @staticmethod
    def _html_to_text(html: str) -> str:
        """Prosty strip tagów — w HTML często są poprawne cyfry z tabeli RAZEM."""
        text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
        text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
        text = re.sub(r"<[^>]+>", "\n", text)
        text = re.sub(r"&nbsp;", " ", text)
        text = re.sub(r"&amp;", "&", text)
        text = re.sub(r"[ \t]+", " ", text)
        return text

    def extract_text_from_pdf(self, pdf_file_obj: io.BytesIO) -> Tuple[str, int]:
        full_text = ""
        page_count = 0
        try:
            pdf_bytes = pdf_file_obj.getvalue()
            with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
                page_count = doc.page_count
                for i, page in enumerate(doc):
                    page_text = self._decode_broken_digits(
                        page.get_text("text", sort=True)
                    )
                    # HTML: nazwy pomieszczeń i RAZEM bywają czytelniejsze niż surowy text
                    try:
                        html_text = self._decode_broken_digits(
                            self._html_to_text(page.get_text("html"))
                        )
                    except Exception:
                        html_text = ""
                    combined = page_text
                    if html_text and html_text.strip() != page_text.strip():
                        combined = page_text + "\n" + html_text
                    full_text += f"\n--- STRONA {i + 1} ---\n" + combined
            return full_text, page_count
        except Exception as e:
            st.error(f"Błąd podczas odczytu pliku PDF za pomocą PyMuPDF: {e}")
            return "", 0

    def extract_exposure_from_pdf(self, pdf_file_obj: io.BytesIO) -> str:
        """Ekspozycja okien na podstawie pozycji strzałki północy na rzucie."""
        try:
            pdf_bytes = pdf_file_obj.getvalue()
            with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
                if doc.page_count == 0:
                    return ""
                page = doc[0]
                width, height = page.rect.width, page.rect.height
                arrows = []
                for x0, y0, x1, y1, word, *_ in page.get_text("words"):
                    letter = word.strip().upper()
                    if not re.fullmatch(r"[NSWE]", letter):
                        continue
                    letter_h = y1 - y0
                    if letter_h < 8 or y0 > height * 0.7:
                        continue
                    arrows.append((letter_h, letter, x0, y0))
                if arrows:
                    arrows.sort(reverse=True)
                    _, letter, x0, y0 = arrows[0]
                    # Strzałka N/S po lewej/prawej → okna na elewacjach W-E
                    if x0 < width * 0.35 or x0 > width * 0.65:
                        return "E, W"
                    # Strzałka u góry/dołu → okna na elewacjach N-S
                    if y0 < height * 0.35 or y0 > height * 0.55:
                        return "N, S"
                    return letter

                search_rect = fitz.Rect(width * 0.75, 0, width, height * 0.15)
                text = page.get_text("text", clip=search_rect, sort=True)
                matches = re.findall(r"\b[NSWE]{1,4}\b", text.upper())
                if matches:
                    letters = sorted(set("".join(matches)))
                    return ", ".join(letters)
        except Exception:
            pass
        return ""

    def count_rooms_and_garderoba(self, text: str) -> Tuple[int, str, int]:
        """
        Zlicza pokoje z etykiet planu / tabeli.
        Garderoba > 5 m2 = pokój; flaga Garderoba=1 tylko gdy jest garderoba ≤ 5 m2.
        Preferuje wiersze tabeli „0.1 NAZWA powierzchnia”; przy duplikatach bierze największą powierzchnię.
        """
        rooms_count = 0
        bathrooms_count = 0
        has_small_garderoba = False
        seen_areas: set = set()

        def take_area(raw: str, min_area: float = 0.0) -> Optional[float]:
            try:
                area = float(raw.replace(",", "."))
            except ValueError:
                return None
            if area <= min_area or area > 80:
                return None
            key = f"{area:.2f}"
            if key in seen_areas:
                return None
            seen_areas.add(key)
            return area

        def classify_name(name: str, area: float) -> None:
            nonlocal rooms_count, bathrooms_count, has_small_garderoba
            name_l = name.lower()
            if re.search(r"łazienk|lazienk", name_l):
                bathrooms_count += 1
            elif re.search(r"garderob", name_l):
                if area > 5.0:
                    rooms_count += 1
                else:
                    has_small_garderoba = True
            elif re.search(r"salon|sypialn|gabine|pok[oó]j", name_l):
                # Śmieci z wymiarów planu (np. 0.4 / 1.39 m²) nie liczą się jako pokój
                if area >= 4.0:
                    rooms_count += 1

        table_hits = list(
            re.finditer(
                r"(\d+\.\d+)\s+([A-ZĄĆĘŁŃÓŚŹŻ][A-ZĄĆĘŁŃÓŚŹŻa-ząćęłńóśźż ]{1,40}?)\s+(\d+[,.]\d+)",
                text,
            )
        )
        if table_hits:
            best_by_idx: Dict[str, Tuple[str, float]] = {}
            for match in table_hits:
                idx = match.group(1)
                # Tylko indeksy pomieszczeń typu 0.1–3.9 (odrzuca np. 547.5 SALON 0.4)
                if not re.fullmatch(r"[0-3]\.\d{1,2}", idx):
                    continue
                name = re.sub(r"\s+", " ", match.group(2).strip())
                if len(name) < 3 or not re.search(
                    r"salon|sypialn|gabine|pok[oó]j|garderob|łazienk|lazienk|hol|kuchn|aneks|wc|gara",
                    name,
                    re.IGNORECASE,
                ):
                    continue
                try:
                    area = float(match.group(3).replace(",", "."))
                except ValueError:
                    continue
                if area <= 0 or area > 80:
                    continue
                prev = best_by_idx.get(idx)
                if prev is None or area > prev[1]:
                    best_by_idx[idx] = (name, area)

            for name, area in best_by_idx.values():
                key = f"{area:.2f}"
                if key in seen_areas:
                    continue
                seen_areas.add(key)
                classify_name(name, area)
            return rooms_count, ("1" if has_small_garderoba else "0"), bathrooms_count

        label_specs = [
            ("room", r"\bsalon(?:\s+z\s+aneks[a-ząćęłńóśźż]*)?"),
            ("room", r"\bsypialn[a-ząćęłńóśźż]*"),
            ("room", r"\bgabine[a-ząćęłńóśźż]*"),
            ("room", r"\bpok[oó]j[a-ząćęłńóśźż]*"),
            ("garderoba", r"\bgarderob[a-ząćęłńóśźż]*"),
            ("bath", r"\b(?:ł|l)azienk[a-ząćęłńóśźż]*"),
        ]

        for kind, pat in label_specs:
            for match in re.finditer(pat, text, re.IGNORECASE):
                window = text[match.end() : match.end() + 160]
                area_match = re.search(r"(\d+[,.]\d+)\s*m\s*[2²]", window, re.IGNORECASE)
                if not area_match:
                    area_match = re.search(r"^\s*(\d+[,.]\d+)", window)
                if not area_match:
                    continue
                min_area = 0.0 if kind in ("bath", "garderoba") else 4.0
                area = take_area(area_match.group(1), min_area=min_area)
                if area is None:
                    continue
                if kind == "bath":
                    bathrooms_count += 1
                elif kind == "garderoba":
                    if area > 5.0:
                        rooms_count += 1
                    else:
                        has_small_garderoba = True
                else:
                    rooms_count += 1

        return rooms_count, ("1" if has_small_garderoba else "0"), bathrooms_count

    def extract_height_in_cm(self, text: str) -> str:
        match = re.search(r"H\s*=\s*(\d+)", text, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    def extract_info_from_filename(self, filename: str) -> Tuple[Optional[str], Optional[str]]:
        try:
            base_name = os.path.splitext(filename)[0]
            normalized = base_name.replace("\u2013", "-").replace("\u2014", "-").strip()
            apt_matches = re.findall(r"\b(\d+\s*[A-Za-z]{1,2})\b", normalized)
            apartment_code: Optional[str] = None
            if apt_matches:
                apartment_code = re.sub(r"\s+", "", apt_matches[-1])
            else:
                tail = normalized.split("_")[-1].split("-")[-1]
                m = re.search(r"(\d+\s*[A-Za-z]{1,2})", tail)
                if m:
                    apartment_code = re.sub(r"\s+", "", m.group(1))

            building_num: Optional[str] = None
            if apartment_code:
                m2 = re.match(r"^(\d+)", apartment_code)
                if m2:
                    building_num = m2.group(1)

            if building_num or apartment_code:
                return building_num, apartment_code
        except Exception:
            return None, None
        return None, None

    def extract_apartment_code_from_text(self, text: str) -> Optional[str]:
        """Kod lokalu z bloku 'Oznaczenie lokalu:' (np. 1A), nie z adresu biura."""
        for block in re.split(r"Oznaczenie lokalu:", text, flags=re.IGNORECASE)[1:]:
            match = re.search(r"(?<![\w/-])(\d+[A-Z])\s*(?:\n|ok\.)", block, re.IGNORECASE)
            if match:
                return match.group(1).upper()
        matches = re.findall(r"Oznaczenie lokalu[:\s]+([\w\d/.-]+)", text, re.IGNORECASE)
        for match in matches:
            if re.fullmatch(r"\d+[A-Za-z]", match.strip()):
                return match.strip().upper()
        return None

    def extract_typ_name_from_filename(self, filename: str) -> str:
        match = re.search(r"typ_([^.]+?)(?:\.pdf)?$", filename, re.IGNORECASE)
        return match.group(1) if match else ""

    def extract_etap_from_text(self, text: str) -> str:
        match = re.search(
            r"Etap\s*(?:realizacyjny)?\s*[:\s-]*(\d+)", text, re.IGNORECASE
        )
        return match.group(1) if match else ""

    def extract_etap_from_filename(self, filename: str) -> str:
        match = re.search(r"ETAP[_\s-]*(\d+)", filename, re.IGNORECASE)
        return match.group(1) if match else ""

    def extract_kitchen_window(self, has_kuchnia: str, has_aneks: str) -> str:
        """Osobna kuchnia ma okno (1); aneks kuchenny w próbkach Osiedle Fikus — 0."""
        if has_kuchnia == "1":
            return "1"
        if has_aneks == "1":
            return "0"
        return "0"

    def extract_garage_connected(self, text: str) -> str:
        if re.search(r"\d+\.\d+\s+GARA(?:Ż|Z)\b", text, re.IGNORECASE):
            return "1"
        if re.search(
            r"lokal\s+połączon\w*\s+z\s+gara(?:ż|z)|"
            r"(?:drzwi|wejśc\w*).*?do\s+gara(?:ż|z)|"
            r"gara(?:ż|z).*?(?:drzwi|wejśc\w*)\s+do\s+lokal",
            text,
            re.IGNORECASE,
        ):
            return "1"
        return "0"

    def extract_building_number_from_text(self, text: str, apartment_code: Optional[str]) -> Optional[str]:
        if apartment_code:
            match = re.search(r"^(\d+)", apartment_code)
            if match:
                return match.group(1)
        return None

    def _normalize_number(self, value: str) -> str:
        return value.replace(",", ".").strip()

    def _extract_value(self, text: str, patterns: List[str], default_value: str = "") -> str:
        normalized_text = re.sub(r"\s+", " ", text).replace(",", ".")
        for pattern in patterns:
            full_pattern = pattern + r"[\s:*\-–—a-zA-ZąćęłńóśźżĄĆĘŁŃÓŚŹŻ°]*?(\d+\.?\d*)"
            matches = re.findall(full_pattern, normalized_text, re.IGNORECASE)
            if matches:
                return matches[-1].strip()
        return default_value

    def _extract_area_after_label(self, text: str, label_pattern: str) -> str:
        match = re.search(
            label_pattern + r"\s*([\d,.]+)\s*m²",
            text,
            re.IGNORECASE,
        )
        return self._normalize_number(match.group(1)) if match else ""

    def _extract_razem_from_section(self, section_text: str) -> str:
        match = re.search(r"RAZEM\s+([\d,.]+)", section_text, re.IGNORECASE)
        return self._normalize_number(match.group(1)) if match else ""

    def _to_float(self, value: str) -> Optional[float]:
        if not value or not str(value).strip():
            return None
        try:
            return float(str(value).replace(",", ".").strip())
        except ValueError:
            return None

    def _fmt_area(self, value: float) -> str:
        return f"{value:.2f}"

    def _closest(self, target: float, candidates: List[float], tol: float = 0.06) -> Optional[float]:
        best = None
        best_diff = tol + 1
        for value in candidates:
            diff = abs(value - target)
            if diff <= tol and diff < best_diff:
                best = value
                best_diff = diff
        return best

    def _extract_summary_cluster_areas(self, text: str) -> Tuple[Dict[str, str], List[float]]:
        """
        Walendów: etykiety i liczby w ramce podsumowania są w innej kolejności.
        Zbieramy liczby m² z okolicy etykiet i klasyfikujemy względem RAZEM / sumy.
        """
        out = {"parter": "", "pietro": "", "laczna": "", "admin": "", "garden": ""}
        garden = re.search(r"ok\.\s*([\d,.]+)\s*m²", text, re.IGNORECASE)
        if garden:
            out["garden"] = self._normalize_number(garden.group(1))

        nums: List[float] = []
        # Pokoje zwykle < ~40 m²; wartości ramki podsumowania są większe
        for raw in re.findall(r"(\d+[,.]\d+)\s*m²", text, re.IGNORECASE):
            val = self._to_float(raw)
            if val is not None and 40 <= val <= 400:
                nums.append(val)

        unique: List[float] = []
        for val in nums:
            if all(abs(val - u) > 0.001 for u in unique):
                unique.append(val)

        return out, unique

    def extract_surface_area(self, text: str, keywords: List[str]) -> str:
        return self._extract_value(text, keywords)

    def extract_floor_areas(self, text: str) -> Dict[str, str]:
        areas = {
            "parter_uzytkowa": "",
            "pietro_uzytkowa": "",
            "pietro2_uzytkowa": "",
            "parter_admin": "",
            "pietro_admin": "",
            "pietro2_admin": "",
            "admin_laczna": "",
            "laczna": "",
            "garden": "",
        }
        sections = re.split(r"--- STRONA \d+ ---", text)
        page_keys = [
            (1, "parter_uzytkowa"),
            (2, "pietro_uzytkowa"),
            (3, "pietro2_uzytkowa"),
        ]

        for page_idx, uzyt_key in page_keys:
            if len(sections) <= page_idx:
                continue
            areas[uzyt_key] = self._extract_razem_from_section(sections[page_idx])

        if len(sections) > 1:
            page1 = sections[1]
            if not areas["parter_uzytkowa"]:
                areas["parter_uzytkowa"] = self._extract_area_after_label(page1, r"Parter")
            if not areas["pietro_uzytkowa"]:
                areas["pietro_uzytkowa"] = self._extract_area_after_label(page1, r"Piętro")
            areas["laczna"] = self._extract_area_after_label(
                page1, r"Powierzchnia Użytkowa \*"
            ) or self._extract_value(
                text,
                [r"Razem Łączna Powierzchnia Lokalu", r"Łączna Powierzchnia Lokalu"],
            )
            areas["admin_laczna"] = self._extract_area_after_label(
                page1, r"Powierzchnia Użytkowa Administracyjna \*\*"
            )

        summary, candidates = self._extract_summary_cluster_areas(text)
        if summary.get("garden"):
            areas["garden"] = summary["garden"]

        parter_f = self._to_float(areas["parter_uzytkowa"])
        pietro_f = self._to_float(areas["pietro_uzytkowa"])

        if candidates:
            # Szukaj pary kondygnacji, która dokładnie sumuje się do „użytkowej *”
            pair = None
            total_match = None
            for total in sorted(candidates, reverse=True):
                for a in candidates:
                    if a >= total - 0.01:
                        continue
                    b = total - a
                    b_hit = self._closest(b, candidates, tol=0.05)
                    if b_hit is None or abs(a - b_hit) < 0.001:
                        continue
                    if abs(a + b_hit - total) <= 0.05:
                        pair = sorted([a, b_hit])
                        total_match = total
                        break
                if pair:
                    break

            if pair and parter_f is not None and pietro_f is not None:
                # Przypisz po bliskości do RAZEM / kolejności kondygnacji
                if abs(pair[0] - parter_f) + abs(pair[1] - pietro_f) <= abs(
                    pair[0] - pietro_f
                ) + abs(pair[1] - parter_f):
                    areas["parter_uzytkowa"] = self._fmt_area(pair[0])
                    areas["pietro_uzytkowa"] = self._fmt_area(pair[1])
                else:
                    areas["parter_uzytkowa"] = self._fmt_area(pair[1])
                    areas["pietro_uzytkowa"] = self._fmt_area(pair[0])
                parter_f = self._to_float(areas["parter_uzytkowa"])
                pietro_f = self._to_float(areas["pietro_uzytkowa"])
            else:
                if parter_f is not None:
                    better = self._closest(parter_f, candidates, tol=0.06)
                    if better is not None:
                        areas["parter_uzytkowa"] = self._fmt_area(better)
                        parter_f = better
                if pietro_f is not None:
                    better = self._closest(pietro_f, candidates, tol=0.06)
                    if better is not None:
                        areas["pietro_uzytkowa"] = self._fmt_area(better)
                        pietro_f = better

            if total_match is not None:
                areas["laczna"] = self._fmt_area(total_match)
            elif parter_f is not None and pietro_f is not None:
                target = parter_f + pietro_f
                total = self._closest(target, candidates, tol=0.08)
                if total is not None:
                    areas["laczna"] = self._fmt_area(total)

            used = {
                self._to_float(areas["parter_uzytkowa"]),
                self._to_float(areas["pietro_uzytkowa"]),
                self._to_float(areas["laczna"]),
            }
            used.discard(None)
            admin_candidates = [
                c for c in candidates if all(abs(c - u) > 0.05 for u in used)
            ]
            admin_label = self._extract_area_after_label(
                text, r"Powierzchnia Użytkowa Administracyjna \*\*"
            )
            admin_label_f = self._to_float(admin_label)
            if admin_label_f is not None and any(
                abs(admin_label_f - c) < 0.05 for c in admin_candidates
            ):
                areas["admin_laczna"] = self._fmt_area(admin_label_f)
            elif admin_candidates:
                direct = [
                    self._to_float(x)
                    for x in re.findall(
                        r"Administracyjna \*\*\s*([\d,.]+)\s*m²", text, re.IGNORECASE
                    )
                ]
                direct = [x for x in direct if x is not None]
                chosen = None
                for d in direct:
                    if any(abs(d - c) < 0.05 for c in admin_candidates):
                        chosen = d
                        break
                if chosen is None:
                    chosen = admin_candidates[-1]
                areas["admin_laczna"] = self._fmt_area(chosen)

        if not areas["laczna"]:
            areas["laczna"] = self._extract_value(
                text,
                [r"Razem Łączna Powierzchnia Lokalu", r"Łączna Powierzchnia Lokalu"],
            )

        return areas

    def extract_main_entrance(self, text: str, page_count: int) -> str:
        sections = re.split(r"--- STRONA \d+ ---", text)
        if len(sections) > 1:
            page1 = sections[1]
            if re.search(r"\bparter\b", page1, re.IGNORECASE):
                return "parter"
            if re.search(r"\bpi[eę]tro\b", page1, re.IGNORECASE):
                return "piętro"
        return "parter" if page_count > 0 else ""

    def extract_outdoor_area_m2(self, text: str, keyword: str) -> str:
        value = self.extract_surface_area(text, [keyword])
        if value:
            return value
        return "0" if re.search(rf"\b{keyword}\w*\b", text, re.IGNORECASE) else "0"

    def extract_garden_area(self, text: str) -> str:
        match = re.search(
            r"(?:Powierzchnia ogrodu\s*)?(?:ok\.\s*)([\d,.]+)\s*m²",
            text,
            re.IGNORECASE,
        )
        if match:
            return self._normalize_number(match.group(1))
        match = re.search(
            r"Powierzchnia ogrodu\s*(?:ok\.\s*)?([\d,.]+)\s*m²",
            text,
            re.IGNORECASE,
        )
        if match:
            return self._normalize_number(match.group(1))
        return self.extract_surface_area(text, [r"Powierzchnia ogrodu", r"ogród"])

    def extract_wc(self, text: str) -> str:
        if re.search(r"\d+\.\d+\s+WC\b", text, re.IGNORECASE):
            return "1"
        if re.search(r"\bWC\b", text, re.IGNORECASE):
            return "1"
        return "0"

    def process_pdf(
        self, pdf_file_obj: io.BytesIO, filename: str, default_etap: str = ""
    ) -> Tuple[Dict[str, str], str]:
        text, page_count = self.extract_text_from_pdf(pdf_file_obj)
        if not text:
            return self.empty_row(), ""

        _, file_apartment = self.extract_info_from_filename(filename)
        apartment_code = file_apartment or self.extract_apartment_code_from_text(text)
        building_num = self.extract_building_number_from_text(text, apartment_code)
        if not building_num:
            building_num, _ = self.extract_info_from_filename(filename)

        typ_name = self.extract_typ_name_from_filename(filename)
        etap = (
            self.extract_etap_from_text(text)
            or self.extract_etap_from_filename(filename)
            or default_etap
        )
        garage_connected = self.extract_garage_connected(text)

        floor_areas = self.extract_floor_areas(text)
        parter_uzytkowa = floor_areas.get("parter_uzytkowa", "")
        pietro_uzytkowa = floor_areas.get("pietro_uzytkowa", "")
        pietro2_uzytkowa = floor_areas.get("pietro2_uzytkowa", "")
        laczna_uzytkowa = (
            floor_areas.get("laczna", "")
            or self._safe_float_sum(parter_uzytkowa, pietro_uzytkowa, pietro2_uzytkowa)
        )
        admin_laczna = floor_areas.get("admin_laczna", "")

        garden_area = floor_areas.get("garden") or self.extract_garden_area(text)
        main_entrance = self.extract_main_entrance(text, page_count)

        sections = re.split(r"--- STRONA \d+ ---", text)
        parter_height_cm = self.extract_height_in_cm(sections[1]) if len(sections) > 1 else ""
        pietro_height_cm = self.extract_height_in_cm(sections[2]) if len(sections) > 2 else ""

        def cm_to_m(cm: str) -> str:
            if not cm.isdigit():
                return ""
            value = float(cm) / 100.0
            text_m = f"{value:.2f}".rstrip("0").rstrip(".")
            return text_m

        parter_height_m = cm_to_m(parter_height_cm)
        pietro_height_m = cm_to_m(pietro_height_cm)

        def present(patterns: List[str]) -> str:
            return "1" if any(re.search(p, text, re.IGNORECASE) for p in patterns) else "0"

        if bool(re.search(r"\bsalon\s+z\s+(kuchni|aneks)\w*", text, re.IGNORECASE)):
            has_kuchnia, has_aneks = "0", "1"
        else:
            has_aneks = present([r"\baneks\w*\s+kuchenn\w*"])
            has_kuchnia = "0" if has_aneks == "1" else present([r"\bkuchni\w*"])

        rooms_count, has_garderoba, bathrooms_count = self.count_rooms_and_garderoba(text)
        if rooms_count == 0:
            rooms_count = len(
                re.findall(
                    r"\d+\.\d+\s+(?:salon|sypialn\w*|gabine\w*|pok[oó]j\w*)",
                    text,
                    re.IGNORECASE,
                )
            )
        if bathrooms_count == 0:
            bathrooms_count = len(
                re.findall(r"\d+\.\d+\s+(?:ł|l)azienk\w*", text, re.IGNORECASE)
            )

        result = self.empty_row()
        kitchen_window = self.extract_kitchen_window(has_kuchnia, has_aneks)
        ocr_values = {
            "NR budynku": building_num or "",
            "NR Lokalu": apartment_code or "",
            "Nazwa robocza typu": typ_name,
            "Etap": etap,
            "Liczba kondygnacji": str(page_count) if page_count > 0 else "",
            "Poddasze Użytkowe/Nieużytkowe": "U" if page_count >= 3 else "N",
            "Główne wejscie do lokalu - kondygnacja": main_entrance,
            "Numer lokalu z dokumentacji budowlanej": apartment_code or "",
            "Powierzchnia użytkowa adm. suma (m2)": admin_laczna,
            "Powierzchnia użytkowa parter (m2)": parter_uzytkowa,
            "Powierzchnia użytkowa piętro (m2)": pietro_uzytkowa,
            "Powierzchnia użytkowa piętro 2 (m2)": (
                pietro2_uzytkowa if page_count >= 3 else "0.00" if page_count > 0 else ""
            ),
            "Powierzchnia użytkowa suma (m2)": laczna_uzytkowa,
            "Powierzchnia ogródek (m2)": garden_area,
            "Ilość pokoi": str(rooms_count) if rooms_count > 0 else "",
            "Garderoba(0 brak, 1 jest)": has_garderoba,
            "Kuchnia (0 brak, 1 jest)": has_kuchnia,
            "Aneks kuchenny (0 brak, 1 jest)": has_aneks,
            "Okno w kuchni lub aneksie (0 brak, 1 jest)": kitchen_window,
            "Schowek (0 brak, 1 jest)": present([r"\bschow\w*"]),
            "Spiżarnia (0 brak, 1 jest)": present([r"\bspi(?:ż|z)arni\w*"]),
            "Pom. Gosp. (0 brak, 1 jest)": present(
                [r"\bpom\.?\s*gosp\.?\b", r"\bpomieszczenie\s+gospodarcze\b"]
            ),
            "Lokal połączony z garażem (0 brak, 1 jest)": garage_connected,
            "Druga łazienka z prysznicem (0 brak, 1 jest)": (
                "1" if bathrooms_count >= 2 else "0"
            ),
            "Gabinet (0 brak, 1 jest)": present([r"\bgabine\w*"]),
            "Balkon (m2)": self.extract_outdoor_area_m2(text, "balkon"),
            "Loggia (m2)": self.extract_outdoor_area_m2(text, "loggi"),
            "Taras (m2)": self.extract_outdoor_area_m2(text, "taras"),
            "Wysokość parteru (m)": parter_height_m,
            "Wysokość piętra (m)": pietro_height_m,
            "Wysokość piwnicy (m2)": "0",
        }
        result.update(ocr_values)
        return result, text


# --- Interfejs Użytkownika Streamlit ---

st.set_page_config(layout="wide")
st.title("OCR do ekstrakcji danych — format Excel ostateczny")
st.subheader("PyMuPDF + eksport zgodny z szablonem (tylko kolumny OCR)")
st.write(
    "Wgraj pliki PDF. Program wypełni kolumny oznaczone jako **ocr** "
    "(w tym numer lokalu, nazwę typu, etap, okno w kuchni i połączenie z garażem). "
    "Kolumny **marketing** i pozostałe **architekt** uzupełniacie ręcznie."
)

default_etap_value = st.text_input(
    "Etap — domyślny, gdy brak w PDF / nazwie pliku",
    value="1A",
    help="Np. 1, 10, 1A. Dla Pod Dębem często 1A.",
)
default_etap = default_etap_value.strip()

ocr = ApartmentPlanOCR()

uploaded_files = st.file_uploader(
    "Wybierz pliki PDF (np. 'OP6_ZT-1A.pdf')",
    type="pdf",
    accept_multiple_files=True,
)

if uploaded_files:
    if st.button("Przetwórz pliki"):
        all_results = []
        progress_bar = st.progress(0, text="Oczekiwanie...")

        with st.spinner("Analizuję dokumenty..."):
            for i, uploaded_file in enumerate(uploaded_files):
                progress_bar.progress(
                    (i + 1) / len(uploaded_files), text=f"Przetwarzanie: {uploaded_file.name}"
                )
                result, full_text = ocr.process_pdf(
                    uploaded_file, uploaded_file.name, default_etap=default_etap
                )
                all_results.append(result)

                if full_text:
                    with st.expander(f"Pokaż/Ukryj pełny tekst odczytany z '{uploaded_file.name}'"):
                        st.text_area("Surowy tekst OCR:", full_text, height=300, key=f"text_area_{i}")
                else:
                    st.warning(f"Nie udało się odczytać tekstu z pliku '{uploaded_file.name}'.")

        if all_results:
            st.markdown("---")
            st.header("Zbiorcze wyniki (wszystkie kolumny)")
            st.success(f"Zakończono! Przetworzono {len(all_results)} plików.")
            st.caption(
                "Wiersz 1 = **sekcja** (kolory), wiersz 2 = **kto uzupełnia** (OCR / marketing / Architekt). "
                "Ten sam układ jest w eksporcie Excel."
            )
            df_results = results_to_dataframe(all_results)
            column_config = {
                name: st.column_config.Column(
                    width="large" if len(col_def["header"]) >= 35 else (
                        "medium" if len(col_def["header"]) >= 18 else "small"
                    )
                )
                for name, col_def in zip(column_display_names(), EXCEL_COLUMNS)
            }
            st.dataframe(
                style_results_dataframe(df_results),
                use_container_width=True,
                column_config=column_config,
            )
            st.session_state["df_results"] = all_results
        else:
            st.warning("Nie udało się wydobyć danych z żadnego z plików.")

if "df_results" in st.session_state:
    excel_data = export_results_to_excel(st.session_state["df_results"])

    st.download_button(
        label="Pobierz wyniki jako plik Excel (format szablonu)",
        data=excel_data,
        file_name="wyniki_ocr.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheet.sheet",
    )
