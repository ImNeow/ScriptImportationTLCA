import argparse
import csv
from datetime import datetime
import openpyxl

if __name__ == "__main__":
    """Ajout des Arguments"""
    parser = argparse.ArgumentParser(description="Script pour transférer des résultats d'un fichier csv vers xlsx")
    parser.add_argument('moodle', type=str, help='Fichier Moodle')
    parser.add_argument('input_tlca', type=str, help='Fichier TLCA input')
    parser.add_argument('output_tlca', type=str, help='Fichier TLCA output')
    parser.add_argument('comp', nargs="+", type=str, help='Compétences à valider')
    args = parser.parse_args()

    """Récupération des données dans le fichier Moodle"""
    fmoodle = open(args.moodle, 'r')
    moodle_result = csv.reader(fmoodle, delimiter=';')

    """Ecriture dans le fichier tlca"""
    wb = openpyxl.load_workbook(filename=args.input_tlca)
    sheet = wb.active

    for row in moodle_result:
        for i in range(1, sheet.max_row):
            # Si les noms sont égaux
            if row[0] == sheet['B' + str(i)].value and row[1] == sheet['C' + str(i)].value:

                # Si FINISHED
                if row[2] == 'Finished':
                    sheet['E' + str(i)].value = 'x'

                    # DATE
                    date = datetime.strptime(row[4][:-3], "%d %B %Y %H:%M")
                    sheet['F' + str(i)].value = date.strftime("%y-%m-%d %H:%M")

                    # COMMENT
                    if sheet['G' + str(i)].value is not None:
                        if float(row[6]) * 10 > float(sheet['G' + str(i)].value[:-1]):
                            sheet['G' + str(i)].value = str(round(float(row[6]) * 10)) + '%'
                    else:
                        sheet['G' + str(i)].value = str(round(float(row[6]) * 10)) + '%'

                    # COMP
                    if float(row[6]) >= 7.50:
                        if "DEV-201" in args.comp:
                            sheet['H' + str(i)].value = "x"
                        if "DEV-203" in args.comp:
                            sheet['I' + str(i)].value = "x"

    """Sauvegarde du fichier"""
    if args.output_tlca:
        wb.save(args.output_tlca)
    else:
        wb.save(args.input_tlca)
    """Fermeture du fichier"""
    fmoodle.close()
