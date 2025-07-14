from PySide6.QtWidgets import QApplication, QWidget, QFileDialog, QMessageBox
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, QThreadPool, QRunnable, QUrl, Qt, QFile, QDir, QObject, Signal, Slot
from PySide6.QtGui import QDesktopServices, QCursor
import qt_themes
import darkdetect
import os
import sys
import docx
from odf.opendocument import load as load_odt
import spacy
from spacy.util import get_package_path

class WorkerSignals(QObject):
    progress = Signal(int)
    message = Signal(str)
    finished = Signal()
    error = Signal(str)

def on_worker_finished():
    window.statusBar().showMessage("Job completed.")

def on_worker_progress(value):
    window.progressBar.setValue(value)

def on_worker_message(message):
    window.statusBar().showMessage(message)

def on_worker_error(message):
    window.statusBar().showMessage(message)

class Worker(QRunnable):
    def __init__(self):
        super().__init__()
        self.signals = WorkerSignals()
        self._is_interrupted = False
        self.nlp=spacy.load(self.get_model_path())

    def get_model_path(self):
    # Funziona sia in .py che in eseguibile PyInstaller
        if hasattr(sys, "_MEIPASS"):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, "en_core_web_md")

    def cancel(self):
        self._is_interrupted = True

    def run(self):
        try:
            model_path = os.path.dirname(self.nlp.path)
            self.signals.message.emit("Job started..")
            global is_working
            is_working=True
            check_buttons()
            nouns=self.extract_nouns(window.txtManuscript.text()) #gets a list of unique nouns
            self.signals.progress.emit(1)
            correlated_nouns= self.find_correlated_nouns(nouns, polished_list, 0.67)
            self.signals.progress.emit(2)
        except Exception as e:
            self.signals.error.emit("An error occurred: " + str(e))
        finally:
            is_working=False
            #self.signals.finished.emit()
            #check_buttons()

    def extract_nouns(self, filepath):
        extension = os.path.splitext(filepath)[1].lower()
        text=""
        if(extension==".docx"):
            doc = docx.Document(filepath)
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text)
            text= '\n'.join(full_text)
        else:
            odt_doc = load_odt(filepath)
            text = ""
            for element in odt_doc.getElementsByType(odt_doc.text_elements):
                if element.text:
                    text += element.text + " "
        document=self.nlp(text)
        unique_nouns = set()
        for token in document:
            if token.pos_ == "NOUN" and not token.is_punct and not token.is_space:
                unique_nouns.add(token.lemma_.lower())
        return sorted(list(unique_nouns))
    
    def find_correlated_nouns(self, nouns, keywords, threshold=0.65):
        keyword_docs = [self.nlp(keyword) for keyword in keywords]
        filtered_nouns = set()

        for noun in nouns:
            noun_doc = self.nlp(noun)
            for kw_doc in keyword_docs:
                if noun_doc.has_vector and kw_doc.has_vector:
                    similarity = noun_doc.similarity(kw_doc)
                    if similarity >= threshold:
                        filtered_nouns.add(noun)
                        break  # at least one keyword has to be  beyond threshold
        return sorted(filtered_nouns)

def resource_path(relative_path):
    try:
        # PyInstaller creates a temporary directory and saves its path to _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Path when running from terminal
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def manuscript_select():
    try:
        file_name = QFileDialog.getOpenFileName(window, "Open manuscript",  filter="*.docx | *.odt")
        if file_name:
            window.txtManuscript.setText(file_name[0])
            check_buttons()
    except Exception as e:
        window.statusBar().showMessage("An error occurred: " + str(e))

def output_select():
    try:
        file_name = QFileDialog.getSaveFileName(window, "Output index",  filter="*.txt")
        if file_name:
            window.txtOutput.setText(file_name[0])
            check_buttons()
    except Exception as e:
        window.statusBar().showMessage("An error occurred: " + str(e))

def check_buttons():
    global is_working
    if(is_working==False):
        window.btnCancel.setEnabled(False)
        if(window.txtManuscript.text() and window.txtOutput.text() and window.txtKeywords.text()):
            window.btnStart.setEnabled(True)
            window.progressBar.setEnabled(True)
        else:
            window.btnStart.setEnabled(False)
            window.progressBar.setEnabled(False)
    else:
        window.btnCancel.setEnabled(True)
        window.btnStart.setEnabled(True)

def start_job():
    try:
        #polishes keywords of excessive spaces or semicolons
        keywords_text=window.txtKeywords.text()
        tmp = keywords_text.split(';')
        global polished_list
        polished_list = [keyword.strip() for keyword in tmp if keyword.strip()]
        if(polished_list):
            global worker
            worker=Worker()
            worker.signals.progress.connect(on_worker_progress)
            worker.signals.message.connect(on_worker_message)
            worker.signals.finished.connect(on_worker_finished)
            worker.signals.error.connect(on_worker_error)
            threadpool.start(worker)
        else:
           window.statusBar().showMessage("Please specify a valid set of keywords.") 
    except Exception as e:
        window.statusBar().showMessage("An error occurred: " + str(e))

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = QWidget()

        ui_file_path = resource_path("lexifinder.ui")
        ui_file = QFile(ui_file_path)

        loader = QUiLoader()
        window = loader.load(ui_file)

        threadpool=QThreadPool()

        if darkdetect.isDark():
            qt_themes.set_theme('modern_dark')
        else:
            qt_themes.set_theme('modern_light')
        
        window.btnCancel.setEnabled(False)
        window.btnStart.setEnabled(False)
        
        window.btnManuscriptSelect.clicked.connect(manuscript_select)
        window.btnOutputSelect.clicked.connect(output_select)
        window.btnStart.clicked.connect(start_job)
        window.txtKeywords.textChanged.connect(check_buttons)

        global is_working
        is_working = False
        window.progressBar.setMaximum(4)
        window.progressBar.setValue(0)
        window.progressBar.setEnabled(False)

        window.show()
        app.exec()
    except Exception as e:
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle("Critical error")
        msg_box.setText("A critical error occurred: " + str(e))
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec()