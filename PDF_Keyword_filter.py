import os
import fitz  # PyMuPDF
from timeit import default_timer as timer
from PySide6.QtWidgets import QApplication, QLabel, QMainWindow, QHBoxLayout, QVBoxLayout, QWidget, QPushButton, QLineEdit, QFileDialog, QTextEdit, QMessageBox, QDateEdit, QRadioButton
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import QByteArray, Qt, QDate
from datetime import datetime
import pandas as pd
import xlsxwriter

start = timer()
base64_image = b"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"

class PDFSearchApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("PDF Keyword Search")
        self.setGeometry(100, 100, 500, 400)

        pixmap = QPixmap()
        pixmap.loadFromData(QByteArray.fromBase64(base64_image))
        self.setWindowIcon(QIcon(pixmap))

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        layout = QVBoxLayout()
        self.result_textbox = QTextEdit(self.central_widget)
        layout.addWidget(self.result_textbox)

        # Horizontal layout for start and end date
        self.date_widget = QWidget(self.central_widget)
        date_layout = QHBoxLayout()  # Use QHBoxLayout for a horizontal layout

        # "From" label and start date
        self.from_label = QLabel("From")
        date_layout.addWidget(self.from_label)

        self.start_date_edit = QDateEdit(self.date_widget)
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setCalendarPopup(True)
        date_layout.addWidget(self.start_date_edit)

        # "To" label and end date
        self.to_label = QLabel("To")
        date_layout.addWidget(self.to_label)

        self.end_date_edit = QDateEdit(self.date_widget)
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_date_edit.setCalendarPopup(True)
        default_end_date = QDate.currentDate()
        self.end_date_edit.setDate(QDate.currentDate())

        date_layout.addWidget(self.end_date_edit)

        self.date_widget.setLayout(date_layout)
        self.date_widget.hide()

        # Radio button to toggle date visibility
        self.toggle_date_button = QRadioButton("Filter By Date", self.central_widget)
        self.toggle_date_button.toggled.connect(self.toggle_date_visibility)
        layout.addWidget(self.toggle_date_button)
        layout.addWidget(self.date_widget)

        self.keyword_textbox = QLineEdit(self.central_widget)
        self.keyword_textbox.setPlaceholderText("Enter keyword")
        layout.addWidget(self.keyword_textbox)

        self.path_textbox = QLineEdit(self.central_widget)
        self.path_textbox.setPlaceholderText("Selected path will appear here")
        layout.addWidget(self.path_textbox)

        self.browse_button = QPushButton("Select Path", self.central_widget)
        self.browse_button.clicked.connect(self.browse_path)
        layout.addWidget(self.browse_button)

        self.search_button = QPushButton("Search", self.central_widget)
        self.search_button.clicked.connect(self.start_search)
        layout.addWidget(self.search_button)

        self.central_widget.setLayout(layout)

        self.setStyleSheet("""
              QMainWindow {
                  background-color:#f5f5f5; /* White background */
              }
              QLabel {
                  color: #333333; /* Black text */
                  font-size: 12px; /* Slightly smaller font */
              }
              QLineEdit {
                  background-color: white;
                  border: 1px solid #ccc; /* Light Gray border */
                  padding: 4px;
                  width: 150px; /* Adjust as needed */
              }
              QPushButton {
                  background-color: #0d6efd; /* Dark Blue button */
                  color: white;
                  border: none;
                  padding: 6px 12px;
                  border-radius: 3px;
                  width: 150px;
              }
              QPushButton:hover {
                  background-color: #0951ba; /* Slightly darker grey on hover */
              }
              QDateEdit {
                  background-color: white;
                  border: 1px solid #ccc; /* Light Gray border */
                  padding: 4px;
                  width: 150px; /* Adjust as needed */
              }
          """)

    def browse_path(self):
        selected_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        self.path_textbox.setText(selected_path)

    def start_search(self):
        c = 0
        keyword = self.keyword_textbox.text()
        folder_path = self.path_textbox.text()

        if not keyword or not folder_path:
            self.result_textbox.clear()
            self.result_textbox.append("Please enter a keyword and select a folder.")
            return

        self.result_textbox.clear()

        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        total_files = len(pdf_files)

        if total_files == 0:
            self.result_textbox.append("No PDF files found in the selected folder.")
            return

        if not self.toggle_date_button.isChecked():
            # If the "Filter By Date" button is unchecked, search all files in the directory
            results_df, files_read, excel_path = process_local_pdf_files(folder_path, keyword, None, None)
        else:
            # If the button is checked, apply the date filter
            start_date = self.start_date_edit.date().toPython()
            end_date = self.end_date_edit.date().toPython()
            results_df, files_read, excel_path = process_local_pdf_files(folder_path, keyword, start_date, end_date)

        self.result_textbox.append(f"Searching in {total_files} PDF files. Please wait...\n")

        for i, pdf_file in enumerate(pdf_files, start=1):
            file_path = os.path.join(folder_path, pdf_file)
            last_modified_date = datetime.fromtimestamp(os.path.getmtime(file_path)).date()

            if not self.toggle_date_button.isChecked() or (self.toggle_date_button.isChecked() and start_date <= last_modified_date <= end_date):
                occurrences, count = extract_text_from_local_pdf(file_path, keyword)
                if count >= 1:
                    c += 1

                    #self.result_textbox.append(f"File: {pdf_file}, Count: {count}, Page Numbers: {occurrences}
        self.result_textbox.append(f"Out of {files_read} files, {c} files  contain the keyword.")
        self.result_textbox.repaint()
        QApplication.processEvents()

        self.result_textbox.append("\nSearch complete.")
        QMessageBox.information(self, "Search Results",
                                f"{files_read} files read \n{c} files  contain the keyword.\nExcel file saved at:\n{excel_path}")


    def toggle_date_visibility(self, checked):
            self.date_widget.setVisible(checked)


def extract_text_from_local_pdf(file_path, target_word):
    occurrences = []
    count = 0
    try:
        with fitz.open(file_path) as pdf_doc:
            for page_number in range(pdf_doc.page_count):
                page = pdf_doc[page_number]
                text = page.get_text()
                if target_word.lower() in text.lower():
                    occurrences.append(page_number + 1)
                    count += text.lower().count(target_word.lower())
    except Exception as e:
        print(f"Error reading PDF: {e}")
    return occurrences, count

def process_local_pdf_files(folder_path, keyword, start_date, end_date):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

    pdf_files_with_keyword = []
    occurrences = []
    counts = []
    total_pages_with_keyword = []

    for pdf_file in pdf_files:
        file_path = os.path.join(folder_path, pdf_file)

        last_modified_date = datetime.fromtimestamp(os.path.getmtime(file_path)).date()

        # Skip date comparison if either start_date or end_date is None
        if (start_date is None or last_modified_date >= start_date) and \
           (end_date is None or last_modified_date <= end_date):
            file_occurrences, count = extract_text_from_local_pdf(file_path, keyword)
            if count >= 1:
                pdf_files_with_keyword.append(pdf_file)
                occurrences.append(file_occurrences)
                counts.append(count)
                total_pages_with_keyword.append(len(file_occurrences))

    # Create a DataFrame for directory and keyword information
    info_df = pd.DataFrame({'Directory': [folder_path, None, None], 'Keyword': [keyword, None, None]})

    # Save directory and keyword information to Excel
    info_excel_path = os.path.join(folder_path, 'directory_and_keyword_info.xlsx')
    info_df.to_excel(info_excel_path, index=False, header=False)

    # Create a DataFrame for search results
    results_df = pd.DataFrame({
        'PDF Files': pdf_files_with_keyword,
        'Count': counts,
        'Page Numbers': occurrences,
        'Total Pages with Keyword': total_pages_with_keyword
    })

    # Save results_df starting from the 4th row and 1st column in Excel
    results_excel_path = os.path.join(folder_path, 'pdf_files_with_keyword.xlsx')
    with pd.ExcelWriter(results_excel_path, engine='xlsxwriter') as writer:
        info_df.to_excel(writer, sheet_name='Sheet1', index=False)
        results_df.to_excel(writer, sheet_name='Sheet1', startrow=3, startcol=0, index=False)

    return results_df, len(pdf_files), results_excel_path



if __name__ == '__main__':
    app = QApplication([])
    window = PDFSearchApp()
    window.show()
    app.exec()

end = timer()
print(end - start)
