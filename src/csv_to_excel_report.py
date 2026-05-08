from pathlib import Path
import csv
import logging
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_FILE = BASE_DIR / "input" / "sample_data.csv"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"

REQUIRED_COLUMNS = ["data", "cliente", "categoria", "descrizione", "importo", "stato"]


def setup_logging():
    LOG_DIR.mkdir(exist_ok=True)

    log_file = LOG_DIR / "app.log"

    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        encoding="utf-8"
    )


def valida_file_csv(file_path):
    if not file_path.exists():
        raise FileNotFoundError(f"File CSV non trovato: {file_path}")

    if file_path.stat().st_size == 0:
        raise ValueError("Il file CSV è vuoto.")

    with open(file_path, mode="r", encoding="utf-8") as file:
        reader = csv.DictReader(file)

        if reader.fieldnames is None:
            raise ValueError("Il file CSV non contiene intestazioni.")

        colonne_mancanti = [
            colonna for colonna in REQUIRED_COLUMNS
            if colonna not in reader.fieldnames
        ]

        if colonne_mancanti:
            raise ValueError(f"Colonne mancanti nel CSV: {', '.join(colonne_mancanti)}")


def leggi_transazioni(file_path):
    transazioni = []

    with open(file_path, mode="r", encoding="utf-8") as file:
        reader = csv.DictReader(file)

        for numero_riga, row in enumerate(reader, start=2):
            if not any(row.values()):
                logging.warning(f"Riga vuota ignorata: {numero_riga}")
                continue

            try:
                row["importo"] = float(row["importo"])
            except ValueError:
                raise ValueError(
                    f"Importo non valido alla riga {numero_riga}: {row.get('importo')}"
                )

            if not row["cliente"]:
                raise ValueError(f"Cliente mancante alla riga {numero_riga}")

            if not row["stato"]:
                raise ValueError(f"Stato mancante alla riga {numero_riga}")

            transazioni.append(row)

    if not transazioni:
        raise ValueError("Nessuna transazione valida trovata nel CSV.")

    return transazioni


def formatta_excel(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        for cell in row:
            cell.number_format = '€ #,##0.00'

    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column

        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[get_column_letter(column)].width = max_length + 3


def genera_report_excel(transazioni):
    OUTPUT_DIR.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = OUTPUT_DIR / f"report_transazioni_{timestamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Report Transazioni"

    headers = ["Data", "Cliente", "Categoria", "Descrizione", "Importo", "Stato"]
    ws.append(headers)

    for transazione in transazioni:
        ws.append([
            transazione["data"],
            transazione["cliente"],
            transazione["categoria"],
            transazione["descrizione"],
            transazione["importo"],
            transazione["stato"],
        ])

    ultima_riga = ws.max_row + 2
    totale = sum(t["importo"] for t in transazioni)
    totale_pagato = sum(t["importo"] for t in transazioni if t["stato"] == "Pagato")
    totale_attesa = sum(t["importo"] for t in transazioni if t["stato"] == "In attesa")

    ws[f"D{ultima_riga}"] = "Totale"
    ws[f"E{ultima_riga}"] = totale

    ws[f"D{ultima_riga + 1}"] = "Totale pagato"
    ws[f"E{ultima_riga + 1}"] = totale_pagato

    ws[f"D{ultima_riga + 2}"] = "Totale in attesa"
    ws[f"E{ultima_riga + 2}"] = totale_attesa

    for row in range(ultima_riga, ultima_riga + 3):
        ws[f"D{row}"].font = Font(bold=True)
        ws[f"E{row}"].font = Font(bold=True)
        ws[f"E{row}"].number_format = '€ #,##0.00'

    formatta_excel(ws)

    wb.save(output_file)
    return output_file


def main():
    setup_logging()

    try:
        logging.info("Avvio generazione report")
        print("Avvio generazione report...")

        valida_file_csv(INPUT_FILE)
        transazioni = leggi_transazioni(INPUT_FILE)
        report_path = genera_report_excel(transazioni)

        logging.info(f"Report generato correttamente: {report_path}")
        print(f"Report generato correttamente: {report_path}")

    except Exception as errore:
        logging.error(f"Errore durante la generazione report: {errore}")
        print(f"Errore: {errore}")


if __name__ == "__main__":
    main()