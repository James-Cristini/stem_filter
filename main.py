"""
The Stem Filter is intended for use by the project management and verbal departments within Addison Whitney

The stem filter takes in a list of names (any number of names), parses the names out and screens them against
an internal database of INN stems. Once completed an excel document is built which displays the conflicts.

The excel document shows all of the names (not just those with conflicts) and outputs the conflicts in an easy to
copy/paste format for easy entering into the creative database

If the list of names are generic names which intentionally contain INN stems, the stem used can be entered and ignored

Please note that the actual database of INN stems is not included here for confidentiality purposes

This program is intended for use on Winodws only, it can run, but will not build an excel doc on Mac

"""

__author__ = "James Cristini"
__credits__ = ["James Cristini"]
__version__ = "1.1.2"
__maintainer__ = "James Cristini"
__email__ = "jacristi0428@gmail.com"

import sys
import sip
import os
import string
from PyQt4 import QtGui, QtCore, uic
from PyQt4.QtGui import QApplication, QMainWindow, QWidget, QDialog, QInputDialog, QMessageBox
from logic import get_stem_conflicts
from main_ui import Ui_StemFilter



reload(sys)
sys.setdefaultencoding('utf-8')

class MainWindow(QMainWindow):

    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_StemFilter()
        self.ui.setupUi(self)
        self.setWindowIcon(QtGui.QIcon("filter.ico"))

        self.ui.get_stems_btn.clicked.connect(self.build_stems_doc)

        self.ui.close_btn.clicked.connect(self.close_app)


    def build_stems_doc(self):
        # Get the stem to ignore
        ignore = str(self.ui.ignore_line.text())

        # Replace special characters with plain text characters then remove extra asterisks after names
        names_text = (str(self.ui.names_text.toPlainText()).replace(u"\u2018", '"').replace(u"\u2019", '"').\
            replace(u"\u201c",'"').replace(u"\u201d", '"').replace(u"\u2013", "-"))
        new_text = ""
        for x in range(len(names_text)-1) :
            if names_text[x] == "*" and names_text[x+1] not in string.letters :
                pass
            else :
                new_text += names_text[x]

        new_text += names_text[-1]
        names_list = new_text.split("\n")

        file_name = "stem_conflicts" + ".xls"

        try:
            get_stem_conflicts(names_list, ignore, file_name)
            self.open_results_doc(file_name)
        except IOError:
            print "File already open"
            QMessageBox.warning(self, "File already open", "The stem_conflicts file is currently open; close this file first and try again") % (file_name)

    def open_results_doc(self, file_name):
        try:
            os.system("start " + str(file_name))
        except:
            print file_name, "not found"


    def close_app(self):
        choice = QtGui.QMessageBox.question(self, "Quit", "Leave?", QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)

        if choice == QtGui.QMessageBox.Yes :
            sys.exit()
        else :
            print "No: Not exiting..."
            pass


def main():
    app = QtGui.QApplication(sys.argv)
    app.setStyle("plastique")
    stem_filter = MainWindow()
    stem_filter.show()
    sip.setdestroyonexit(False)
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
