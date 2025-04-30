from flask import Flask, request, send_file, jsonify
from google.cloud import vision
from PIL import Image
from io import BytesIO
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import os
import re

app = Flask(__name__)

def clean_number(val):
    val = re.sub(r'[^\d,\.\-]', '', val).replace(',', '.').replace(' ', '')
    return val if re.match(r'^\d+(\.\d+)?$', val) else ''

def ocr_google_vision(image_bytes: bytes) -> str:
    client = vision.ImageAnnotatorClient()
    image = vision.Image(content=image_bytes)
    response = client.text_detection(image=image)
    texts = response.text_annotations
    if not texts:
        return ""
    return texts[0].description

def parse_ocr_text(ocr_text):
    produkty = []
    ostatni_produkt = None
    linia_oczekujaca = None
    tymczasowy_rabat = 0
    czekamy_na_rabat = False

    koncowe_frazy = ["PTU", "Kwota", "Suma", "Razem", "Płatność", "RAZEM", "nr:", "Data", "Godzina"]
    linie = [l.strip() for l in ocr_text.split('\n') if l.strip()]

    for line in linie:
        if linia_oczekujaca:
            match = re.match(r"^([\d, ]+)\s+([\d, ]+)$", line)
            if match:
                cena1 = clean_number(match.group(1))
                cena2 = clean_number(match.group(2))
                if cena1 == cena2:
                    try:
                        cena = float(cena1)
                        if ostatni_produkt:
                            produkty.append({**ostatni_produkt, "rabat": -tymczasowy_rabat})
                        ostatni_produkt = {
                            "nazwa": linia_oczekujaca,
                            "ilosc": 1,
                            "cena_jedn": cena,
                            "cena_laczna": cena,
                            "rabat": 0,
                            "wlasciciel": "wspolne"
                        }
                        linia_oczekujaca = None
                        tymczasowy_rabat = 0
                        czekamy_na_rabat = True
                        continue
                    except:
                        pass

        match = re.match(r"^(.*?)(\d[\d, ]*)\s*\*\s*([\d, ]+)\s+([\d, ]+)\w?$", line)
        if match:
            nazwa = match.group(1).strip() or linia_oczekujaca or ""
            nazwa = re.sub(r"\s+\d{1,3}$", "", nazwa).strip()
            raw_ilosc = clean_number(match.group(2)) or "1"
            liczby = re.findall(r"\d{1,3},\s?\d{2}", line)
            liczby = [clean_number(x) for x in liczby if clean_number(x)]
            raw_cena_laczna = liczby[-1] if liczby else "0"
            raw_cena_jedn = liczby[-2] if len(liczby) >= 2 else "0"
            try:
                ilosc = float(raw_ilosc)
                cena_jedn = float(raw_cena_jedn)
                cena_laczna = float(raw_cena_laczna)
                if ostatni_produkt:
                    produkty.append({**ostatni_produkt, "rabat": -tymczasowy_rabat})
                ostatni_produkt = {
                    "nazwa": nazwa,
                    "ilosc": ilosc,
                    "cena_jedn": cena_jedn,
                    "cena_laczna": cena_laczna,
                    "rabat": 0,
                    "wlasciciel": "wspolne"
                }
                linia_oczekujaca = None
                tymczasowy_rabat = 0
                czekamy_na_rabat = True
                continue
            except:
                pass

        if "*" in line:
            liczby = [clean_number(x) for x in re.findall(r"[\d,\.\s]{1,10}", line) if clean_number(x)]
            if len(liczby) >= 3:
                try:
                    ilosc = float(liczby[0])
                    cena_jedn = float(liczby[1])
                    cena_laczna = float(liczby[2])
                    nazwa = linia_oczekujaca or line.split("*")[0].strip()
                    if ostatni_produkt:
                        produkty.append({**ostatni_produkt, "rabat": -tymczasowy_rabat})
                    ostatni_produkt = {
                        "nazwa": nazwa,
                        "ilosc": ilosc,
                        "cena_jedn": cena_jedn,
                        "cena_laczna": cena_laczna,
                        "rabat": 0,
                        "wlasciciel": "wspolne"
                    }
                    linia_oczekujaca = None
                    tymczasowy_rabat = 0
                    czekamy_na_rabat = True
                    continue
                except:
                    pass

        rabat_match = re.search(r"-([\d,\s\w]+)", line)
        if rabat_match and czekamy_na_rabat and "*" not in line:
            raw = re.sub(r"[^\d,.-]", "", rabat_match.group(1))
            raw = clean_number(raw)
            try:
                rabat = float(raw)
                tymczasowy_rabat += rabat
            except:
                pass
            continue

        if not re.search(r"-\s*[\d,]+", line) and any(line.startswith(f) for f in koncowe_frazy):
            czekamy_na_rabat = False

        if not re.search(r"\d\s*\*\s*[\d, ]+", line) and not re.search(r"-\s*[\d,]+", line):
            linia_oczekujaca = line.strip()

        if re.search(r"\d\s*\*\s*[\d, ]+", line) and ostatni_produkt:
            produkty.append({**ostatni_produkt, "rabat": -tymczasowy_rabat})
            ostatni_produkt = None
            linia_oczekujaca = None
            tymczasowy_rabat = 0
            czekamy_na_rabat = False

    if ostatni_produkt:
        produkty.append({**ostatni_produkt, "rabat": -tymczasowy_rabat})

    return produkty

@app.route("/upload", methods=["POST"])
def upload():
    if 'file' not in request.files:
        return jsonify({"error": "Brak pliku"}), 400

    file = request.files['file']
    image_bytes = file.read()
    ocr_text = ocr_google_vision(image_bytes)
    produkty = parse_ocr_text(ocr_text)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Paragon"

    headers = ["Nazwa", "Ilosc", "Cena jednostkowa", "Cena laczna", "Rabat", "Do zaplaty", "Wlasciciel"]
    ws.append(headers)

    for prod in produkty:
        do_zaplaty = prod['cena_laczna'] - prod['rabat']
        ws.append([
            prod['nazwa'], prod['ilosc'], prod['cena_jedn'], prod['cena_laczna'],
            prod['rabat'], do_zaplaty, prod['wlasciciel']
        ])

    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.freeze_panes = "A2"

    last_row = len(produkty) + 1
    ws[f"E{last_row + 2}"].value = "Suma moje:"
    ws[f"F{last_row + 2}"].value = f"=SUMIF(G2:G{last_row},\"ja\",F2:F{last_row})"
    ws[f"E{last_row + 3}"].value = "Suma ona:"
    ws[f"F{last_row + 3}"].value = f"=SUMIF(G2:G{last_row},\"ona\",F2:F{last_row})"
    ws[f"E{last_row + 4}"].value = "Suma wspolne:"
    ws[f"F{last_row + 4}"].value = f"=SUMIF(G2:G{last_row},\"wspolne\",F2:F{last_row})"
    ws[f"E{last_row + 6}"].value = "Ona ma mi oddac:"
    ws[f"F{last_row + 6}"].value = (
        f"=SUMIF(G2:G{last_row},\"ona\",F2:F{last_row}) + SUMIF(G2:G{last_row},\"wspolne\",F2:F{last_row})/2"
    )

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="paragon.xlsx"
    )

if __name__ == "__main__":
    app.run(debug=True)
