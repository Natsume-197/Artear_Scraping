# Importing the required libraries
# THIS CODE REQUIRES xlsxwriter LIBRARY
import os
import re
import sys
import tkinter
import requests
import threading
import pandas as pd
from tkinter import ttk
from bs4 import BeautifulSoup
from tkinter.messagebox import showinfo

channel_options = {
    "trece": "ARTEAR EL TRECE",
    "tn": "ARTEAR TN TODO NOTICIAS",
    "resistencia": "ARTEAR EL NUEVE RESISTENCIA",
    "magazine": "ARTEAR CIUDAD MAGAZINE",
    "volver": "ARTEAR VOLVER",
    "satelital": "ARTEAR EL TRECE SATELITAL",
    "canal12": "ARTEAR CANAL 12 CORDOBA ARGENTINA",
    "canal10-mardelplata": "ARTEAR CANAL 10 MAR DEL PLATA",
    "canal9": "ARTEAR CANAL 9 PARANA",
    "airevalle": "ARTEAR CANAL 10 GENERAL ROCA",
    "canal10-tucuman": "ARTEAR CANAL 10 TUCUMAN",
    "quiero": "ARTEAR QUIERO"
}


def formaturl(url):
    if not re.match("(?:http|ftp|https)://", url):
        return "http://{}".format(url)
    return url


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(
        os.path.abspath('__file__')))
    return os.path.join(base_path, relative_path)

class Found(Exception):
    pass

def set_name_file(URL_BASE):
    name_file = "DESCONOCIDO"
    for value in channel_options:
        try:
            if value in URL_BASE:
                name_file = channel_options[value]
                raise Found

        except Found:
            return name_file

    return name_file


class application:

    def gui_loading(self):

        self.main_window = tkinter.Tk()
        self.main_window.title("Gestor de descargas Artear")
        self.main_window.resizable(False, False)

        self.main_label = tkinter.Label(
            text="Inserte aquí el link de la grilla de contenido Artear a descargar:"
        )
        self.main_label.place(x=10, y=10, height=25)

        self.input_url = tkinter.Entry(self.main_window)
        self.input_url.place(x=15, y=40, width=400, height=25)

        try:
            self.download_button = ttk.Button(
                text="Descargar",
                command=self.start_submit_thread)
        except:
            showinfo(
                message="Dirección no válida, verifique el link insertado.",
                title="Error",
            )
            sys.exit(1)

        self.download_button.place(x=422, y=39, height=25)

        self.status_label = tkinter.Label(
            text="Estado: No hay ninguna descarga activa")
        self.status_label.place(x=10, y=70, height=25)

        self.progressbar = ttk.Progressbar(
            mode="determinate", length=100, orient='horizontal')
        self.progressbar.place(x=15, y=95, width=470)

        self.main_window.geometry("505x130")
        self.main_window.mainloop()

    def start_submit_thread(self):
        global submit_thread
        self.submit_thread = threading.Thread(
            target=lambda: self.process_file(self.input_url.get()))
        self.submit_thread.daemon = True
        self.status_label['text'] = 'Estado: Descargando grilla...'
        self.submit_thread.start()

    def process_file(self, URL_BASE):
        name_file = set_name_file(URL_BASE)

        import time
        for i in range(50):
            time.sleep(0.01)
            self.progressbar['value'] = i

        try:
            URL_BASE = formaturl(URL_BASE)
            page = requests.get(URL_BASE)
        except:
            self.status_label['text'] = 'Estado: No hay ninguna descarga activa.'
            self.progressbar['value'] = 0
            showinfo(
                message="No fue posible acceder a la página, verifique el link ingresado.",
                title="Error",
            )

            sys.exit(1)

        soup = BeautifulSoup(page.text, "html.parser")

        main_grid = soup.find("frame", {"name": "Grillas"})

        if (URL_BASE.__contains__('mkt')):
            pass
        else:
            try:
                URL_BASE_2 = "http://www.canal13.artear.com.ar/" + \
                    main_grid["src"]
                page2 = requests.get(URL_BASE_2)
                soup = BeautifulSoup(page2.text, "html.parser")
            except:
                self.status_label['text'] = 'Estado: No hay ninguna descarga activa.'
                self.progressbar['value'] = 0
                showinfo(
                    message="No se encontró información a descargar. Verifique la URL/contenido.",
                    title="Error",
                )
                sys.exit(1)

        table = ''

        try:
            table = soup.find_all("table")[1]
            list_dates = []
            for row in table.find_all("td", {"class": "header"}):
                list_dates.append(row.text.strip().split(" ")[-1])
        except:
            self.status_label['text'] = 'Estado: No hay ninguna descarga activa.'
            self.progressbar['value'] = 0
            showinfo(
                message="No se encontró información a descargar. Verifique la URL/contenido.",
                title="Error",
            )
            sys.exit(1)

        list_dates = list(filter(None, list_dates))

        # Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday
        daysRowspan = [0, 0, 0, 0, 0, 0, 0]

        data = {
            "Fecha": [],
            "Hora de inicio": [],
            "Programa": [],
        }
        df = pd.DataFrame(data)

        # Collecting Data
        # All valid cells on the table have the identifier <td.normal>
        #   * The ones that have text inside are programs
        #   * The ones that don´t have text inside are the grey cells used for padding
        #   * The ones that has a nested table inside (class='link')
        for row in table.find_all("td", {"class": "normal"}):
            textValue = ""

            if name_file.__contains__("VOLVER") or name_file.__contains__("MAGAZINE"):
                try:
                    textValue = row.contents[0].upper()
                except Exception as e:
                    print(e)
                    textValue = row.text.upper()
                    textValue = textValue.replace("\n", " ")

            else:
                
                textValue = row.text.upper()
                print(textValue)
                textValue = textValue.replace("\n", " ")

            rowspanValue = int(row.get("rowspan"))
            colspanValue = int(row.get("colspan"))

            # If we find text inside the cell, parse it following the expected format:
            #  * startTime - Title [- Id] [\n Subtitle]
            #  * Example:
            # 22:45 - EL HOTEL DE LOS FAMOSOS
            formatRegex = "^(.*?)-(.*?)(- \$(.*?))?(\n.*)?$"
            programData = {}

            if textValue:
                # Clearing text from wrong words
                textValue = re.sub("ESTRENO\S*", "", textValue)
                textValue = re.sub("REPE DEL \S*", "", textValue)
                textValue = re.sub("CAP.DEL DIA \S*", "", textValue)
                textValue = re.sub("DIA ANTERIOR\S*", "", textValue)
                textValue = re.sub("FALSO VIVO\S*", "", textValue)
                textValue = re.sub("86CAP(.*)", "", textValue)
                textValue = re.sub("ANTERIOR ANTERIOR(.*)", "", textValue)
                textValue = re.sub("DEL DIA DEL DIA(.*)", "", textValue)
                textValue = re.sub("CON CARMEN BARBIERI(.*)", "", textValue)
                textValue = re.sub("\?CAP(.*)", "", textValue)
                textValue = re.sub("CAP.(.*)", "", textValue)
                textValue = re.sub("\*(.*)", "", textValue)
                textValue = re.sub("REPE DE(.*)", "", textValue)
                matches = re.findall(formatRegex, textValue)

                try:
                    programData["startTime"] = matches[0][0].strip()
                    programData["title"] = matches[0][1].strip()
                except:
                    showinfo(
                        message="Se encontro un error al procesar el contenido a descargar. Contacte al soporte.",
                        title="Error",
                    )
                    sys.exit(1)

            # Find which day has the lower rowspan value. This is where we need to put the next cell for the schedule.
            dayWithLowestRowspan = 0
            lowestRowspanValue = daysRowspan[0]

            for [day, rowspan] in enumerate(daysRowspan):
                if rowspan < lowestRowspanValue:
                    lowestRowspanValue = rowspan
                    dayWithLowestRowspan = day

            # The colspan value determines how many days the program will be repeated (colspan 1 -> 1 day. colspan 3 -> 3 days, etc)
            for i in range(0, colspanValue):
                day = int(dayWithLowestRowspan) + i
                daysRowspan[day] += rowspanValue

                # Add program to schedule if the cell had data inside it
                if len(programData) > 0:
                    print(day, programData)
                    df.loc[len(df)] = [
                        list_dates[day],
                        programData["startTime"],
                        programData["title"],
                    ]

        # removed_df_duplicates = df.drop_duplicates(subset=['Fecha', 'Hora de inicio'], keep='first')
        df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y')
        sorted_df = df.sort_values(
            by=["Fecha", "Hora de inicio"], ascending=True)

        sorted_df['Fecha'] = sorted_df['Fecha'].dt.strftime('%d/%m/%Y')

        with pd.option_context(
            "display.max_rows",
            None,
            "display.max_columns",
            None,
            "display.precision",
            3,
        ):
            print(sorted_df)

        try:
            path = './Grillas'
            if not os.path.exists(path):
                os.makedirs(path)
        except:
            print(err)
            showinfo(
                message="No se ha podido crear el directorio donde guardar el archivo. Cambie la ubicación de este programa a una ruta con permisos e intente de nuevo.",
                title="Error",
            )

        try:
            path = './Grillas'
            initial_date = df['Fecha'].dt.strftime('%d-%m-%Y').iloc[0]
            final_date = df['Fecha'].dt.strftime('%d-%m-%Y').iloc[-1]
            writer = pd.ExcelWriter(
                f"{path}/{name_file} ({initial_date} ~ {final_date}).xlsx", engine="xlsxwriter")
            sorted_df.to_excel(writer, sheet_name="GRILLA",
                               index=False, na_rep="NaN")
            for column in sorted_df:
                column_length = max(
                    sorted_df[column].astype(str).map(len).max(), len(column)
                )
                col_idx = sorted_df.columns.get_loc(column)
                writer.sheets["GRILLA"].set_column(
                    col_idx, col_idx, column_length)
            writer.save()
            for i in range(50, 101):
                time.sleep(0.01)
                self.progressbar['value'] = i

            showinfo(
                message="Se ha descargado la grilla. Consulte la ruta de este directorio para encontrar el archivo descargado.",
                title="Completado",
            )
            self.progressbar['value'] = 0
            self.status_label['text'] = 'Estado: No hay ninguna descarga activa.'
        except Exception as err:
            print(err)
            showinfo(
                message="No se pudo guardar el archivo. Cambie la ubicación de este programa e intente de nuevo.",
                title="Error",
            )

    def main_process(self):
        thread = threading.Thread(target=self.gui_loading)
        thread.start()  # start parallel computation


if __name__ == "__main__":
    main_app = application()
    main_app.main_process()
