import tkinter as tk
from tkinter import filedialog

import difflib

import pandas as pd
import numpy as np
import os
import math

from lxml import etree

folder_path = None

namespaces = {'cfdi': "http://www.sat.gob.mx/cfd/4"}


def choose_folder():
    global folder_path
    folder_path = filedialog.askdirectory()
    directory_label.config(text=f"Selected directory: {folder_path}")


class parsedXmlObject:
    clave: int
    desc: str
    cantidad: float
    precio: float
    total: float


def run():
    if folder_path == None:
        return

    xml_files = []
    excel_file = ''

    for file in os.listdir(folder_path):
        if file.endswith('.xlsx'):
            excel_file = file
        if file.endswith('.xml'):
            xml_files.append(file)

    if len(excel_file) == 0:
        print('Error, no excel file')
        return

    if len(xml_files) == 0:
        print('Error, no xml')
        return

    excel_df = pd.read_excel(os.path.join(folder_path, excel_file), 0)
    empty = excel_df.isnull()

    comparison_df = pd.DataFrame(columns=['e_clave', 'e_desc', 'e_qty', 'e_precio',
                                 'e_total', 'f_clave', 'f_desc', 'f_qty', 'f_precio', 'f_total', 'observaciones'])

    row_index = 17

    # Loop until you reach the end of the DataFrame or an empty row
    while row_index < len(excel_df) and not empty.iloc[row_index].iloc[0]:
        toAdd = excel_df.iloc[row_index]
        row = {
            'e_clave': toAdd.iloc[0],
            'e_desc': toAdd.iloc[1],
            'e_qty': toAdd.iloc[2],
            'e_precio': toAdd.iloc[3],
            'e_total': toAdd.iloc[4],
            'f_clave': None,
            'f_desc': None,
            'f_qty': None,
            'f_precio': None,
            'f_total': None,
            'observaciones': '',
        }

        comparison_df = comparison_df._append(row, ignore_index=True)

        row_index += 1

    # cargar XML

    for p in xml_files:
        doc: etree._ElementTree = etree.parse(os.path.join(folder_path, p))

        root = doc.getroot()

        conceptos = root.xpath('cfdi:Conceptos', namespaces=namespaces)[0]

        for concepto in conceptos.xpath('cfdi:Concepto', namespaces=namespaces):
            attr: dict[str, any] = concepto.attrib

            desc: str = attr.get('Descripcion')

            lookalikes = difflib.get_close_matches(
                desc.lower(), comparison_df['e_desc'].str.lower(), cutoff=0.7)

            if len(lookalikes) == 0:
                # empty add new row
                row = {
                    'e_clave': None,
                    'e_desc': '',
                    'e_qty': 0,
                    'e_precio': 0,
                    'e_total': 0,
                    'f_clave': attr.get('ClaveProdServ'),
                    'f_desc': attr.get('Descripcion'),
                    'f_qty': attr.get('Cantidad'),
                    'f_precio': attr.get('ValorUnitario'),
                    'f_total': attr.get('Importe'),
                    'observaciones': f'No en excel, Archivo: {p}',
                }

                comparison_df = comparison_df._append(row, ignore_index=True)
            else:
                # not empty
                for matcheo in lookalikes:
                    makeExit = False

                    for i, leRow in comparison_df[comparison_df["e_desc"].str.lower() == matcheo].iterrows():
                        # leRow = comparison_df[comparison_df["e_desc"].str.lower() == matcheo].head(
                        #     1)
                        # i = leRow.index[0]

                        # check if not already filled
                        if comparison_df.at[i, 'f_clave'] == None and -0.01 < float(leRow['e_qty']) - float(attr.get('Cantidad')) < 0.01:
                            comparison_df.at[i, 'f_clave'] = attr.get(
                                'ClaveProdServ')
                            comparison_df.at[i, 'f_desc'] = attr.get(
                                'Descripcion')
                            comparison_df.at[i, 'f_qty'] = attr.get('Cantidad')
                            comparison_df.at[i, 'f_precio'] = attr.get(
                                'ValorUnitario')
                            comparison_df.at[i, 'f_total'] = attr.get(
                                'Importe')
                            comparison_df.at[i,
                                             'observaciones'] = f'Archivo: {p}'

                            makeExit = True

                            break

                        pass

                    if makeExit:
                        break

                    pass

                pass

                if not makeExit:
                    row = {
                        'e_clave': None,
                        'e_desc': '',
                        'e_qty': 0,
                        'e_precio': 0,
                        'e_total': 0,
                        'f_clave': attr.get('ClaveProdServ'),
                        'f_desc': attr.get('Descripcion'),
                        'f_qty': attr.get('Cantidad'),
                        'f_precio': attr.get('ValorUnitario'),
                        'f_total': attr.get('Importe'),
                        'observaciones': f'Repetida o no tiene match. Archivo: {p}',
                    }

                    comparison_df = comparison_df._append(
                        row, ignore_index=True)

    comparison_df.to_csv(os.path.join(folder_path, 'datos procesados.csv'))


root = tk.Tk()
root.title("Folder Chooser")

choose_button = tk.Button(root, text="Choose Folder", command=choose_folder)
choose_button.pack(pady=10)

directory_label = tk.Label(root, text="")
directory_label.pack(pady=10)

start_button = tk.Button(
    root, text="Start", command=run)
start_button.pack(pady=10)

root.mainloop()
