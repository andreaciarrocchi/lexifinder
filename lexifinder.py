from PySide6.QtWidgets import QApplication, QWidget, QFileDialog, QMessageBox, QToolTip
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, QThreadPool, QRunnable, QFile, QObject, Signal, QUrl, Qt
from PySide6.QtGui import QDesktopServices, QCursor
import qt_themes
import darkdetect
import os
import sys
import spacy
import fitz  # PyMuPDF

class WorkerSignals(QObject):
    progress = Signal(int)
    message = Signal(str)
    finished = Signal()
    error = Signal(str)

def on_worker_finished():
    global word_count
    global job_cancelled

    if(not job_cancelled):
        window.statusBar().showMessage("Job completed. " + str(word_count) + " words were added to the index.")
    else:
        window.statusBar().showMessage("Job cancelled.")
    QApplication.restoreOverrideCursor()
    check_buttons()

def on_worker_progress(value):
    window.progressBar.setValue(value)

def on_worker_message(message):
    window.statusBar().showMessage(message)

def on_worker_error(message):
    QApplication.restoreOverrideCursor()
    window.statusBar().showMessage(message)

class Worker(QRunnable):
    def __init__(self):
        super().__init__()
        self.signals = WorkerSignals()
        self._is_interrupted = False
        self.nlp=spacy.load(resource_path("en_core_web_md"))

    def cancel(self):
        # Set the flag to stop the task
        self._is_interrupted = True
        global is_working
        is_working=False
        QApplication.restoreOverrideCursor()
        check_buttons()

    def run(self):
        try:
            global is_working
            is_working=True
            check_buttons()
            global job_cancelled
            self.signals.message.emit("Job started..")

            # Step 1: extract unique nouns from the manuscript PDF
            nouns=self.extract_nouns(window.txtManuscript.text()) #gets a list of unique nouns
            if self._is_interrupted:
                job_cancelled=True
                return
            self.signals.progress.emit(1)

            # Step 2: filter nouns by semantic similarity with input keywords
            correlated_nouns= self.find_correlated_nouns(nouns, polished_list, window.horizontalSlider.value() / 100)
            global word_count
            word_count = len(correlated_nouns)
            if self._is_interrupted:
                job_cancelled=True
                return
            self.signals.progress.emit(2)

             # Step 3: build an index of noun occurrences, page by page
            index=self.extract_occurrences_by_page(window.txtManuscript.text(), correlated_nouns)
            if self._is_interrupted:
                job_cancelled=True
                return
            self.signals.progress.emit(3)

            # Step 4: write the index to a text file
            if(len(index)!=0):
                self.write_on_file(index)
            if self._is_interrupted:
                job_cancelled=True
                return    
            self.signals.progress.emit(4)
        except Exception as e:
            self.signals.error.emit("An error occurred: " + str(e))
        finally:
            is_working=False
            self.signals.finished.emit()

    def write_on_file(self, index_dict):
        try:
            output_path = window.txtOutput.text()
            if not output_path:
                self.signals.error.emit("No output file was selected.")
                return
            
             # Save the index to a .txt file, each noun followed by its page numbers
            with open(output_path, "w", encoding="utf-8") as f:
                for noun in sorted(index_dict):
                    pages = ", ".join(str(p) for p in sorted(index_dict[noun]))
                    f.write(f"{noun} {pages}\n")
        except Exception as e:
            self.signals.error.emit("An error occurred while saving the output: " + str(e))


    def extract_occurrences_by_page(self, pdf_path, word_list):
        try:
            doc = fitz.open(pdf_path)
            index = {word.lower(): [] for word in word_list}

            # Check each page for the presence of each target word
            for page_num in range(len(doc)):
                page = doc[page_num]
                text = page.get_text().lower()
                for word in word_list:
                    if word.lower() in text:
                        index[word.lower()].append(page_num + 1)
            return index
        except Exception as e:
            self.signals.error.emit("An error occurred while extracting occurrences: " + str(e))

    def extract_nouns(self, filepath):
        try:
            if not filepath.lower().endswith(".pdf"):
                raise ValueError("File must be PDF.")

            doc = fitz.open(filepath)
            text = ""
            for page in doc:
                text += page.get_text() + "\n"

            # Use SpaCy to extract lemmatized nouns from the full text
            document = self.nlp(text)
            unique_nouns = set()
            for token in document:
                if token.pos_ == "NOUN" and not token.is_punct and not token.is_space:
                    unique_nouns.add(token.lemma_.lower())

            return sorted(list(unique_nouns))
        except Exception as e:
            self.signals.error.emit("An error occurred while extracting nouns: " + str(e))

    def find_correlated_nouns(self, nouns, keywords, threshold):
        try:
            # Turn keywords into SpaCy Doc objects
            keyword_docs = [self.nlp(keyword) for keyword in keywords]
            filtered_nouns = set()

            # Keep only nouns with similarity above the threshold to at least one keyword
            for noun in nouns:
                noun_doc = self.nlp(noun)
                for kw_doc in keyword_docs:
                    if noun_doc.has_vector and kw_doc.has_vector:
                        similarity = noun_doc.similarity(kw_doc)
                        if similarity >= threshold:
                            filtered_nouns.add(noun)
                            break  # only one match needed
            return sorted(filtered_nouns)
        except Exception as e:
            self.signals.error.emit("An error occurred while processing the manuscript: " + str(e))

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
        file_name = QFileDialog.getOpenFileName(window, "Select PDF", filter="PDF files (*.pdf)")
        if file_name:
            window.txtManuscript.setText(file_name[0])
            check_buttons()
    except Exception as e:
        window.statusBar().showMessage("An error occurred: " + str(e))

def output_select():
    try:
        file_name = QFileDialog.getSaveFileName(window, "Output index",  filter="*.txt")
        if file_name:
            if not file_name[0].lower().endswith(".txt") and file_name[0] != "":
                file_name_final = file_name[0] + ".txt"
            else:
                file_name_final=file_name[0]
            window.txtOutput.setText(file_name_final)
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
        window.btnStart.setEnabled(False)

def start_job():
    try:
        global job_cancelled
        job_cancelled=False
        window.progressBar.setValue(0)
        window.statusBar().showMessage("Initialization..")
        QApplication.processEvents()
        QApplication.setOverrideCursor(QCursor(Qt.BusyCursor))

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

def slider_value_changed():
    try:
        window.lblValue.setText(str(window.horizontalSlider.value()))
    except Exception as e:
        window.statusBar().showMessage("An error occurred: " + str(e))

def cancel_job():
    try:
        global worker
        if 'worker' in globals():
            worker.cancel()
            QApplication.restoreOverrideCursor()
    except Exception as e:
        window.statusBar().showMessage("An error occurred: " + str(e))

def donate_url():
    try:
        url = QUrl("https://paypal.me/ciarro85")
        QDesktopServices.openUrl(url)
    except:
        window.statusBar().showMessage("An error occurred while launching the browser.")

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
        window.btnCancel.clicked.connect(cancel_job)
        window.btnOutputSelect.clicked.connect(output_select)
        window.btnDonate.clicked.connect(donate_url)
        window.btnStart.clicked.connect(start_job)
        window.txtKeywords.textChanged.connect(check_buttons)
        window.horizontalSlider.valueChanged.connect(slider_value_changed)

        window.txtKeywords.setToolTip("Nouns separated by semicolons (e.g. myword; my other word; lastword)")

        global is_working
        is_working = False

        global word_count
        word_count=0

        global job_cancelled
        job_cancelled=False

        window.progressBar.setMaximum(4)
        window.progressBar.setValue(0)
        window.progressBar.setEnabled(False)
        window.lblValue.setText(str(window.horizontalSlider.value()))

        window.show()
        app.exec()

    except Exception as e:
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle("Critical error")
        msg_box.setText("A critical error occurred: " + str(e))
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec()