import json
import os
import time
import re

from PySide2 import QtWidgets, QtCore, QtGui
from oauth2client.service_account import ServiceAccountCredentials
import gspread

"""GOOGLE DRIVE API
    Permet de se connecter à google drive pour accéder aux fichiers GSHEETS.
    """

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_secret.json', scope)
client = gspread.authorize(creds)

"""CHEMIN
    Chemin des fichiers .json
    """

CUR_DIR = os.path.dirname(__file__)
DATA_FILE = os.path.join(CUR_DIR, "url.json")
DATA_FILE2 = os.path.join(CUR_DIR, "KeyW.json")


class App(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Clean_DB")
        self.setGeometry(600, 300, 600, 300)
        self.setup_ui()
        self.list_url()
        self.list_listKW()
        self.setup_connections()
        self.add_url()
        self.rem_url()
        self.add_wk()
        self.add_CO()
        self.comboChanged()
        self.view_list()
        self.add_KW()
        self.rem_KW()
        self.COKW_selected()


# SETUP

    def setup_ui(self):
        # Fenetre principale
        self.main_layout = QtWidgets.QVBoxLayout(self)

        # Ligne "entrer URL"
        self.Demo_Url = QtWidgets.QLabel(
            "Script : Clean_DB")
        self.Demo_Url.setAlignment(QtCore.Qt.AlignCenter)
        self.UrlTitle = QtWidgets.QLineEdit()
        self.UrlTitle.setPlaceholderText(
            "Enter google file's name")

        # Bouton "ajouter nom du fichier"
        self.btn_addUrl = QtWidgets.QPushButton("add")

        # Création de deux BoxLayout sur le Main Layout
        self.layoutdeux = QtWidgets.QGridLayout(self)
        # AddLayout au mainLayout
        self.main_layout.addLayout(self.layoutdeux)

        self.Info_RR = QtWidgets.QLabel(
            "GoogleSheets ")
        self.Info_RR.setAlignment(QtCore.Qt.AlignCenter)

        self.lw_URLS = QtWidgets.QListWidget()
        self.lw_URLS.setSelectionMode(QtWidgets.QListWidget.SingleSelection)

        # Bouton "Retirer nom du fichier"
        self.btn_remUrl = QtWidgets.QPushButton("delete")
        # Layout de gauche : Liste des Worksheets
        self.Info_WS = QtWidgets.QLabel(
            "Worksheets")
        self.Info_WS.setAlignment(QtCore.Qt.AlignCenter)

        self.lw_WS = QtWidgets.QListWidget()
        self.lw_WS.setSelectionMode(QtWidgets.QListWidget.SingleSelection)

        # Layout de droite : Liste des Worksheets
        self.Info_CO = QtWidgets.QLabel(
            "columns ")
        self.Info_CO.setAlignment(QtCore.Qt.AlignCenter)

        self.lw_CO = QtWidgets.QListWidget()
        self.lw_CO.setSelectionMode(QtWidgets.QListWidget.SingleSelection)

        self.Info_fleche0 = QtWidgets.QLabel(
            "->")
        self.Info_fleche1 = QtWidgets.QLabel(
            "->")
        self.Info_fleche2 = QtWidgets.QLabel(
            "->")
        # Connecter au Layoutdeux

        # AddWidget
        self.layoutdeux.addWidget(self.Demo_Url, 0, 5, 1, 3)
        self.layoutdeux.addWidget(self.UrlTitle, 1, 5, 1, 3)
        self.layoutdeux.addWidget(self.btn_addUrl, 2, 5, 1, 3)
        self.layoutdeux.addWidget(self.Info_RR, 3, 5, 1, 3)
        self.layoutdeux.addWidget(self.lw_URLS, 4, 5, 1, 3)
        self.layoutdeux.addWidget(self.btn_remUrl, 5, 5, 1, 3)

        self.layoutdeux.addWidget(self.Info_fleche1, 3, 9, 1, 1)
        self.layoutdeux.addWidget(self.Info_WS, 0, 10, 1, 3)
        self.layoutdeux.addWidget(self.lw_WS, 1, 10, 6, 3)

        self.layoutdeux.addWidget(self.Info_fleche2, 3, 14, 1, 1)
        self.layoutdeux.addWidget(self.Info_CO, 0, 15, 1, 3)
        self.layoutdeux.addWidget(self.lw_CO, 1, 15, 6, 3)

        # Layout supplémentaire : KW/CLEAN
        self.layoutKW = QtWidgets.QGridLayout(self)

        self.Info_List = QtWidgets.QLabel(
            " select a list")
        self.Info_List.setAlignment(QtCore.Qt.AlignCenter)

        self.ListKWTitle = QtWidgets.QComboBox()

        self.lw_ListKW = QtWidgets.QListWidget()
        self.lw_ListKW.setSelectionMode(QtWidgets.QListWidget.MultiSelection)

        self.Info_ADDLIST = QtWidgets.QLabel(
            "Add a list")
        self.Info_ADDLIST.setFont(QtGui.QFont("", 7, QtGui.QFont.Bold))
        self.Info_ADDLIST.setAlignment(QtCore.Qt.AlignCenter)

        self.List_Title = QtWidgets.QLineEdit()
        self.List_Title.setPlaceholderText("your list")

        self.btn_add_list = QtWidgets.QPushButton(
            "Add a list")

        self.Info_REMLIST = QtWidgets.QLabel(
            "Delete selected list")
        self.Info_REMLIST.setFont(QtGui.QFont("", 7, QtGui.QFont.Bold))
        self.Info_REMLIST.setAlignment(QtCore.Qt.AlignCenter)

        self.btn_REMLIST = QtWidgets.QPushButton(
            "Delete")

        self.Info_ADDKW = QtWidgets.QLabel(
            "Add KeyWord")
        self.Info_ADDKW.setFont(QtGui.QFont("", 7, QtGui.QFont.Bold))
        self.Info_ADDKW.setAlignment(QtCore.Qt.AlignCenter)

        self.KW = QtWidgets.QLineEdit()
        self.KW.setPlaceholderText("your KeyWord")

        self.btn_addKW = QtWidgets.QPushButton(
            "Add")

        self.Info_REMKW = QtWidgets.QLabel(
            "Delete selected KeyWord")
        self.Info_REMKW.setFont(QtGui.QFont("", 7, QtGui.QFont.Bold))
        self.Info_REMKW.setAlignment(QtCore.Qt.AlignCenter)

        self.btn_remKW = QtWidgets.QPushButton(
            "Delete")

        self.Info_CLEAN = QtWidgets.QLabel(
            " ❌ Please check your informations ❌ ")
        self.Info_CLEAN.setAlignment(QtCore.Qt.AlignCenter)

        self.INFO_SEL = QtWidgets.QLabel(
            "🔄 Select Column and KeyWord(s) 🔄")
        self.INFO_SEL.setAlignment(QtCore.Qt.AlignCenter)

        self.btn_clean_db = QtWidgets.QPushButton("🚀 Clean your file 🚀")

        self.layoutKW.addWidget(self.Info_List, 0, 0, 1, 4)
        self.layoutKW.addWidget(self.Info_ADDLIST, 1, 4, 1, 2)
        self.layoutKW.addWidget(self.List_Title, 2, 4, 1, 2)
        self.layoutKW.addWidget(self.btn_add_list, 3, 4, 1, 2)
        self.layoutKW.addWidget(self.ListKWTitle, 1, 0, 1, 4)
        self.layoutKW.addWidget(self.lw_ListKW, 2, 0, 12, 4)
        self.layoutKW.addWidget(self.Info_REMLIST, 5, 4, 1, 2)
        self.layoutKW.addWidget(self.btn_REMLIST, 6, 4, 1, 2)
        self.layoutKW.addWidget(self.Info_ADDKW, 8, 4, 1, 2)
        self.layoutKW.addWidget(self.KW, 9, 4, 1, 2)
        self.layoutKW.addWidget(self.btn_addKW, 10, 4, 1, 2)
        self.layoutKW.addWidget(self.Info_REMKW, 12, 4, 1, 2)
        self.layoutKW.addWidget(self.btn_remKW, 13, 4, 1, 2)
        self.layoutKW.addWidget(self.Info_CLEAN, 16, 0, 1, 6)

        self.layoutKW.addWidget(self.INFO_SEL, 17, 0, 1, 6)

        self.layoutKW.addWidget(self.btn_clean_db, 19, 0, 1, 6)
        self.main_layout.addLayout(self.layoutKW)
# # SETUP

# CONNECTIONS

    def setup_connections(self):

        self.UrlTitle.returnPressed.connect(self.add_url)
        self.btn_addUrl.clicked.connect(self.add_url)
        self.btn_remUrl.clicked.connect(self.rem_url)
        self.lw_URLS.clicked.connect(self.add_wk)
        self.lw_WS.clicked.connect(self.add_CO)
        self.lw_CO.clicked.connect(self.COKW_selected)
        self.lw_ListKW.clicked.connect(self.COKW_selected)
        self.ListKWTitle.currentTextChanged.connect(self.comboChanged)
        self.ListKWTitle.currentTextChanged.connect(self.view_list)
        self.KW.returnPressed.connect(self.add_KW)
        self.List_Title.returnPressed.connect(self.add_list)
        self.btn_add_list.clicked.connect(self.add_list)
        self.btn_REMLIST.clicked.connect(self.show_del_list)
        self.btn_addKW.clicked.connect(self.add_KW)
        self.btn_remKW.clicked.connect(self.rem_KW)
        self.btn_clean_db.clicked.connect(self.show_popup)


# # CONNECTIONS

# URL

    def list_url(self):
        with open(DATA_FILE, "r") as f:
            liste_json = json.load(f)
            for url in liste_json:
                self.lw_URLS.addItem(str(url))

    def add_url(self):
        url_title = self.UrlTitle.text()

        if not url_title:
            return False

        with open(DATA_FILE, "r") as f:
            liste_url = json.load(f)

            if url_title not in liste_url:
                liste_url.append(url_title)
                with open(DATA_FILE, "w") as y:
                    json.dump(liste_url, y, indent=4)
                    self.lw_URLS.addItem(url_title)

            else:
                pass
        self.UrlTitle.clear()

    def rem_url(self):
        for selecteditem in self.lw_URLS.selectedItems():
            url_cliked = selecteditem.text()
            with open(DATA_FILE, "r") as f:
                liste_url = json.load(f)

                if url_cliked in liste_url:
                    liste_url.remove(url_cliked)
                    with open(DATA_FILE, "w") as y:
                        resultat = json.dump(liste_url, y, indent=4)
                        self.lw_URLS.takeItem(self.lw_URLS.row(selecteditem))
                else:
                    pass
# # URL

# WS pour WORKSHEET

    def add_wk(self):
        self.lw_WS.clear()
        for selecteditem in self.lw_URLS.selectedItems():
            url_cliked = selecteditem.text()

            for WKSHEET in client.open(url_cliked).worksheets():

                WS_cliked = str(WKSHEET).replace(
                    "\' ", " \'").split(" \'")[1]
                self.lw_WS.addItem(WS_cliked)

# # WS pour WORKSHEET

# CO pour COLONNE

    def add_CO(self):
        self.lw_CO.clear()
        for selecteditem in self.lw_URLS.selectedItems():
            url_cliked = selecteditem.text()

            for selecteditem2 in self.lw_WS.selectedItems():
                WS_cliked = selecteditem2.text()

                for row in client.open(
                    url_cliked).worksheet(
                        WS_cliked).row_values(1):

                    CO_list = str(row).split(", ")
                    self.lw_CO.addItem(CO_list[0])
# # CO pour COLONNE

# combo Changed
    def comboChanged(self):
        text = self.ListKWTitle.currentText()

    def COKW_selected(self):
        listSI = []
        for selecteditem in self.lw_CO.selectedItems():
            if selecteditem.text():
                CO_cliked = selecteditem.text()
            else:
                CO_cliked = "rien"

            for selecteditem2 in self.lw_ListKW.selectedItems():
                if selecteditem2.text() in listSI:
                    listSI.remove(selecteditem2.text())

                else:
                    listSI.append(selecteditem2.text())

            if not listSI:
                listSI = []

            if listSI == []:
                self.INFO_SEL.setText(
                    f"⛔️ No KeyWord selected ! | ✅ Column selected: < {CO_cliked} > ")
            else:
                self.INFO_SEL.setText(
                    f' ✅ KeyWord(s) selected < {", ".join(listSI)} > | ✅ Column selected: < {CO_cliked} > ')


# combo Changed

# LIST_KW pour KEYWORD

    def list_listKW(self):

        with open(DATA_FILE2, "r") as f:
            liste_json = json.load(f)
            for dict_liste in liste_json:
                self.ListKWTitle.addItem(str(dict_liste))

    def view_list(self):
        self.lw_ListKW.clear()

        ListCB = self.ListKWTitle.currentText()

        if ListCB:
            with open(DATA_FILE2, "r") as f:
                liste_json = json.load(f)

                values = liste_json.get(ListCB, {}).values()
                self.lw_ListKW.addItems(values)

    def add_list(self):
        LE_list = self.List_Title.text()

        if not LE_list:
            return False
        with open(DATA_FILE2, "r") as f:
            liste_list = json.load(f)

            if LE_list not in liste_list:
                liste_list.update({LE_list: {}})

                with open(DATA_FILE2, "w") as y:
                    json.dump(liste_list, y, indent=4)
                    self.ListKWTitle.addItem(LE_list)

        self.ListKWTitle.setCurrentText(LE_list)
        self.List_Title.clear()

    def rem_list(self):
        ListCB = self.ListKWTitle.currentText()

        with open(DATA_FILE2, "r") as f:
            liste_json = json.load(f)
            try:
                if ListCB in liste_json:
                    del liste_json[ListCB]
                    index = self.ListKWTitle.findText(ListCB)
                    self.ListKWTitle.removeItem(index)
                    with open(DATA_FILE2, "w") as y:
                        resultat = json.dump(
                            liste_json, y, indent=4)

            except ValueError as e:
                print("oulala", e)

    def add_KW(self):
        ListCB = self.ListKWTitle.currentText()
        kw_title = self.KW.text()

        if not kw_title:
            return False

        with open(DATA_FILE2, "r") as f:
            json_kw = json.load(f)
            list_listCB = json_kw[ListCB]

            keys = json_kw.get(ListCB).keys()
            le_kw = list(keys)

            if kw_title not in list_listCB.values():

                if not le_kw:
                    New_KW = "KW1"

                    list_listCB[New_KW] = kw_title

                    with open(DATA_FILE2, "w") as y:
                        resultat = json.dump(
                            json_kw, y, indent=4, sort_keys=True)
                        self.lw_ListKW.addItem(kw_title)

                elif f'KW{(len(keys))}' == le_kw[-1]:

                    KW_index_last = f'KW{(len(keys)+1)}'

                    list_listCB[KW_index_last] = kw_title

                    with open(DATA_FILE2, "w") as y:
                        resultat = json.dump(
                            json_kw, y, indent=4, sort_keys=True)
                        self.lw_ListKW.addItem(kw_title)

                else:
                    i = 1
                    for key in le_kw:
                        KW_index = f'KW{i}'

                        if KW_index not in key and kw_title not in list_listCB:
                            list_listCB[KW_index] = kw_title

                            with open(DATA_FILE2, "w") as y:
                                json.dump(json_kw, y, indent=4,
                                          sort_keys=True)
                                self.lw_ListKW.addItem(kw_title)
                                break
                        else:
                            i += 1
            else:
                self.KW.clear()
        self.KW.clear()

    def rem_KW(self):
        ListCB = self.ListKWTitle.currentText()
        # self.ListKWTitle.currentText()

        for selecteditem in self.lw_ListKW.selectedItems():
            KW_cliked = selecteditem.text()
            with open(DATA_FILE2, "r") as f:
                liste_json = json.load(f)
                list_listCB = liste_json[ListCB]
                try:
                    for key, value in liste_json.get(ListCB).items():
                        if value == KW_cliked:
                            del list_listCB[key]

                except RuntimeError:
                    print(key, value)

                    with open(DATA_FILE2, "w") as y:
                        resultat = json.dump(
                            liste_json, y, indent=4, sort_keys=True)
                        self.lw_ListKW.takeItem(
                            self.lw_ListKW.row(selecteditem))


# # LIST_KW pourKEYWORD

# clean_db pour CLEAN DATABASE


    def show_popup(self):
        msg_CDB = QtWidgets.QMessageBox()
        msg_CDB.setIcon(QtWidgets.QMessageBox.Information)
        ans_0 = msg_CDB.question(
            self, "Clean", "Are you sure ? 💀", msg_CDB.Yes | msg_CDB.No)
        if ans_0 == msg_CDB.Yes:
            self.INFO_SEL.setText(
                f"⛔️⛔️⛔️ CLEANING ⛔️⛔️⛔️")
            ans_1 = msg_CDB.question(
                self, "Clean", "Are you really sure ? 💀💀", msg_CDB.Yes | msg_CDB.No)
            if ans_1 == msg_CDB.Yes:
                self.clean_db()
            else:
                self.COKW_selected()
                return False
        else:
            return False

    def show_del_list(self):
        msg_DL = QtWidgets.QMessageBox()
        msg_DL.setIcon(QtWidgets.QMessageBox.Information)
        ans = msg_DL.question(
            self, "Delete a list", "Are you sure ? 💀", msg_DL.Yes | msg_DL.No)
        if ans == msg_DL.Yes:
            self.rem_list()

        else:
            return False

    def clean_db(self):
        start_time = time.time()
        i = 0
        y = 0
        z = 0
        zz = 0

        for selected_item in self.lw_URLS.selectedItems():
            url_cliked = selected_item.text()

            for selected_item2 in self.lw_WS.selectedItems():
                WS_cliked = selected_item2.text()

                for selected_item3 in self.lw_CO.selectedItems():
                    CO_cliked = selected_item3.text()

                    row_value = client.open(url_cliked).worksheet(
                        WS_cliked).row_values(1)
                    if CO_cliked in row_value:

                        for CO, e in enumerate(row_value):
                            if e == CO_cliked:
                                values_list = client.open(
                                    url_cliked).worksheet(
                                    WS_cliked).col_values(CO+1)

                                for selected_item4 in self.lw_ListKW.selectedItems():
                                    KW_cliked = selected_item4.text()

                                    r = re.compile(
                                        KW_cliked, re.IGNORECASE | re.UNICODE)

                                    for k in values_list:
                                        if r.match(k) or r.search(k):
                                            try:

                                                KW_select1 = client.open(
                                                    url_cliked).worksheet(
                                                    WS_cliked).find(k)
                                                resultat = client.open(url_cliked).worksheet(
                                                    WS_cliked).delete_rows(
                                                        KW_select1.row)

                                                i += 1
                                                time. sleep(2.0)
                                            except gspread.exceptions.CellNotFound:
                                                continue
                                            except gspread.exceptions.APIError:
                                                zz += 1
                                                print(
                                                    "Attention QUOTAS : pause de 60 secondes")
                                                print(i)
                                                time. sleep(90.0)

                                                KW_select1 = client.open(
                                                    url_cliked).worksheet(
                                                    WS_cliked).find(k)
                                                resultat = client.open(url_cliked).worksheet(
                                                    WS_cliked).delete_rows(
                                                        KW_select1.row)

                                        else:
                                            replacements = [(r'[àâä]', 'a'), (r'[ÀÂÄ]', 'A'), (r'[éèê]', 'e'), (r'[ë]', 'e'), (r'[ÉÈÊ]', 'E'), (r'[Ë]', 'E'), (r'[ïî]', 'i'), (r'[ÏÎ]', 'I'), (
                                                r'[ôö]', 'o'), (r'[Ô]', 'O'), (r'[ùûü]', 'u'), (r'[ÙÛÜ]', 'U'), (r'[ÿ]', 'y'), (r'[Ÿ]', 'Y'), (r'[ç]', 'c'), (r'[Ç]', 'C')]

                                            k1 = re.sub(
                                                r'[àâä]', "a", k, re.IGNORECASE | re.UNICODE)
                                            k2 = re.sub(
                                                r'[ÀÂÄ]', "A", k1, re.IGNORECASE | re.UNICODE)
                                            k3 = re.sub(
                                                r'[éèêë]', "e", k2, re.IGNORECASE | re.UNICODE)
                                            k4 = re.sub(
                                                r'[ÉÈÊË]', "E", k3, re.IGNORECASE | re.UNICODE)
                                            k5 = re.sub(
                                                r'[ïî]', "i", k4, re.IGNORECASE | re.UNICODE)
                                            k6 = re.sub(
                                                r'[ÏÎ]', "I", k5, re.IGNORECASE | re.UNICODE)
                                            k7 = re.sub(
                                                r'[ôö]', "o", k6, re.IGNORECASE | re.UNICODE)
                                            k8 = re.sub(
                                                r'[Ô]', "O", k7, re.IGNORECASE | re.UNICODE)
                                            k9 = re.sub(
                                                r'[ùûü]', "u", k8, re.IGNORECASE | re.UNICODE)
                                            k10 = re.sub(
                                                r'[ÙÛÜ]', "U", k9, re.IGNORECASE | re.UNICODE)
                                            k11 = re.sub(
                                                r'[ÿ]', "y", k10, re.IGNORECASE | re.UNICODE)
                                            k12 = re.sub(
                                                r'[Ÿ]', "Y", k11, re.IGNORECASE | re.UNICODE)
                                            k13 = re.sub(
                                                r'[ç]', "c", k12, re.IGNORECASE | re.UNICODE)
                                            k14 = re.sub(
                                                r'[Ç]', "C", k13, re.IGNORECASE | re.UNICODE)

                                            if r.match(k14) or r.search(k14):
                                                try:
                                                    KW_select1 = client.open(
                                                        url_cliked).worksheet(
                                                        WS_cliked).find(k)
                                                    resultat = client.open(url_cliked).worksheet(
                                                        WS_cliked).delete_rows(
                                                            KW_select1.row)
                                                    print(KW_select1)
                                                    i += 1
                                                    time. sleep(2.0)

                                                except gspread.exceptions.CellNotFound:
                                                    continue

                                                except gspread.exceptions.APIError:
                                                    zz += 1
                                                    print(
                                                        "Attention QUOTAS : pause de 60 secondes")
                                                    print(i)
                                                    time. sleep(90.0)
                                                    KW_select1 = client.open(
                                                        url_cliked).worksheet(
                                                        WS_cliked).find(k)
                                                    resultat = client.open(url_cliked).worksheet(
                                                        WS_cliked).delete_rows(
                                                            KW_select1.row)
                                                continue
                            else:
                                print("rien")

        end_time = time.time()
        msg_end = QtWidgets.QMessageBox()
        msg_end.setIcon(QtWidgets.QMessageBox.Information)
        if round(end_time-start_time) < 60:
            ans = msg_end.question(
                self, "end_clean", f' My job is done here ✅ \n {round(end_time-start_time)} seconds to clean the file\n{i} lines were deleted \n{zz} break(s) (google limitations)', msg_end.Ok)
        else:
            minutes = (round(end_time-start_time) / 60)
            ans = msg_end.question(
                self, "end_clean", f' My job is done here ✅ \n {round(minutes)} seconds to clean the file\n{i} lines were deleted \n{zz} break(s) (google limitations)', msg_end.Ok)


app = QtWidgets.QApplication([])
win = App()
win.show()
app.exec_()
