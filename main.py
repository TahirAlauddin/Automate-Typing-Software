from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from interface import Ui_MainWindow
from utils import TypeWriter, read_document
from qt_material import apply_stylesheet
import time
import platform

error_message = "The software couldn't type properly!"
success_message = "The software typed successfully!"
default_sleep_time = 5  # In seconds


class MainWindow(QMainWindow):

    documentFile = None

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.chooseDocumentFileButton.clicked.connect(
            self.openDocumentFileDialog)
        self.ui.startTypingButton.clicked.connect(self.type_document)
        self.ui.waitSpinBox.setValue(default_sleep_time)
        shortcut = QShortcut(QKeySequence('ctrl+o'), self)
        shortcut.activated.connect(self.openDocumentFileDialog)

        if platform.system() == 'Darwin':
            self.ui.chooseDocumentFileButton.setStatusTip(
                "Either press Command+O or click on it to choose Document File")
        self.show()

    def openDocumentFileDialog(self):
        """Open a window so the user can select a document file
        It sets `documentFile` variable to the path of docs."""
        options = QFileDialog.Options()
        self.documentFile, _ = QFileDialog.getOpenFileName(
            self, "Documents File", "", "Docs (*.docx)", options=options)
        if self.documentFile:
            self.ui.documentFileInput.setText(self.documentFile)

    def type_document(self):
        """Read the document from `documentFile` and start typing
        it with super fast speed after a few seconds.
        """
        try:
            writer = TypeWriter()
            document_text = read_document(self.documentFile)
            # Wait for a few seconds before typing
            sleep_for_seconds = self.ui.waitSpinBox.value()
            time.sleep(sleep_for_seconds)
            # Start typing
            writer.start_typing(document_text)
        except Exception as e:
            print(e)
            self.ui.statusBar.showMessage(error_message, 10*1000)
        # If everything went well then show a success message on status bar
        else:
            self.ui.statusBar.showMessage(success_message, 10*1000)


if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    # setup stylesheet
    apply_stylesheet(app, theme='dark_teal.xml')
    window = MainWindow()
    sys.exit(app.exec_())
