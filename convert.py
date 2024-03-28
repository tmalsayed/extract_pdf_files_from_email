import sys
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidgetItem,QMessageBox
from PyQt5.uic import loadUi
import pdfplumber
import pandas as pd
import PyPDF2
import re
import os
import tempfile
import win32com.client
from PyQt5.QtGui import QColor, QIcon, QPixmap
import time
from PyQt5.QtCore import QTimer

class ConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi("resources\\convert.ui", self)
        self.pushButton_2.clicked.connect(self.display_summary)
        self.pushButton.clicked.connect(self.process_and_save)
        # self.setWindowTitle("Declaration Invoice File")
        self.setWindowIcon(QIcon('resources\\afl_logo_small.png'))
        # Enable drag and drop for the tableWidget
        self.tableWidget.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        self.file_paths = [u.toLocalFile() for u in event.mimeData().urls()]
        print(self.file_paths)
        self.display_summary()
        # for file in files:
        #     if file.lower().endswith('.pdf'):
        #         # Process each PDF file
                

    def get_email_attachments(self, search_text):
        """Fetch the attached PDFs from an email with matching search_text in Outlook."""
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 corresponds to the inbox
        # Search for the email with matching text in the subject
        email = None
        for message in inbox.Items:
            if search_text in message.Subject:
                self.update_status('email found')        
                email = message
                break
        # If the email is found, save the attachments to a temporary location
        file_paths = []
        if email:
            for attachment in email.Attachments:
                self.update_status('processing attachments') 
                # Check if the attachment is a PDF
                if attachment.FileName.lower().endswith(".pdf"):
                    temp_dir = tempfile.gettempdir()
                    temp_path = os.path.join(temp_dir, attachment.FileName)
                    attachment.SaveAsFile(temp_path)
                    file_paths.append(temp_path)
        return file_paths

    def extract_inv_date_from_pdf(self, file_path):
        with open(file_path, 'rb') as file:
            # Create a PDF reader object
            pdf_reader = PyPDF2.PdfReader(file)
            # Get the text content from the first page
            # text = pdf_reader.getPage(0).extractText()
            text = pdf_reader.pages[0].extract_text()
            # Search for a date pattern (assuming a format like DD/MM/YYYY)
            date_pattern = re.compile(r'(\d{2}.\d{2}.\d{4})')
            match = date_pattern.search(text)
            # Return the date if found
            if match:
                return match.group(1)
        return None  # If no date is found

    def display_summary(self):
        self.update_status('getting email attachment')
        # Fetch the search_text from the lineEdit
        search_text = self.lineEdit.text()
        print("search text:",search_text)

        if search_text == '-':
            self.update_status('getting email attachment')        
            # Use the search_text to fetch the email attachments
            file_paths = self.get_email_attachments(search_text)
            print(file_paths)        
            if not file_paths:
                # If no matching file is found, display a message and exit the function
                print("No matching email or attachment found!")
                # return
        
        # Assume that headers remain the same across all summaries. Use the first file to extract headers.
        first_file_summary = self.extract_pdf_summary(self.file_paths[0])
        print("first_file_summary", first_file_summary)
        headers = list(first_file_summary.keys())
        print(headers)
        self.tableWidget.setColumnCount(len(headers))
        self.tableWidget.setHorizontalHeaderLabels(headers)
        self.tableWidget.verticalHeader().setVisible(False)
        # Now, loop through all file paths to extract summaries and display them.
        self.update_status('extracting data')
        for file_path in self.file_paths:
            summary = self.extract_pdf_summary(file_path)
            print("summar2:", summary)
            row_position = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row_position)
            
            for col, value in enumerate(summary.values()):
                self.tableWidget.setItem(row_position, col, QTableWidgetItem(str(value)))
        self.update_status('finished extracting data')
        self.adjust_column_widths()
        # time.sleep(2)

    def process_and_save(self):
        """Saves the content of the tableWidget to a CSV."""

        # Get the row and column counts from the tableWidget
        row_count = self.tableWidget.rowCount()
        col_count = self.tableWidget.columnCount()

        # Prepare a container to hold the data
        table_data = []

        # Iterate over the tableWidget to extract its data
        for row in range(row_count):
            row_data = []
            for col in range(col_count):
                item = self.tableWidget.item(row, col)
                row_data.append(item.text() if item else "")  # Get the text from QTableWidgetItem or an empty string
            table_data.append(row_data)

        # Convert the list of lists into a DataFrame
        df = pd.DataFrame(table_data, columns=[self.tableWidget.horizontalHeaderItem(col).text() for col in range(col_count)])
        df = df.astype(str)
        # Use the current date and time to generate a filename in a clear, readable format
        datetime_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"Data_{datetime_str}.csv"

        # Save the DataFrame to CSV
        df['File name'] = df['File name'].astype(str)
        df.to_csv(filename, index=False)
        self.update_status('finished saving data')

    def extract_pdf_summary(self, file_path):
        """Extracts and processes data from the PDF, returning a summary."""

        def convert_custom_format(s):
            """Utility function to clean up the string format."""
            if pd.isna(s):  # Check for NaN values
                return s  # Return NaN as is
            print("convert:", s)
            return s.replace('\n', '').replace(',', '').replace(' ', '')
        # def convert_custom_format(s):
        #     """Utility function to clean up the string format."""
        #     print("convert:",s)
        #     return s.replace('\n', '').replace(',', '').replace(' ','')
        file_name = str(file_path.split('-')[-1].split('.')[0])
        all_data = self.extract_tables_from_pdf(file_path)
        # Combine tables and preprocess the data
        final_df = pd.concat(all_data, ignore_index=True)
        print(final_df)
        # print(final_df)
        # final_df.to_excel(f"{file_name}.xlsx", index=False)
        columns_to_convert = ['QTY', 'Weight\n(KG)', 'Volume\n(M3)', 'Unit\nValue\n(AED)', 'Total\nValue\n(AED)', 'CW BOE\nREF. NO.']
        
        for col in columns_to_convert:
            final_df[col] = final_df[col].apply(convert_custom_format)
            if col != 'CW BOE\nREF. NO.':
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce')
        print(final_df)
        final_df['QTY'] = final_df['QTY'].round(4)
        final_df['Weight'] = final_df['Weight\n(KG)'].sum().round(4)
        final_df['Volume'] = final_df['Volume\n(M3)'].sum().round(4)
        final_df['Total Value'] = final_df['Total\nValue\n(AED)'].sum().round(4)
        # final_df.to_excel("processed_data.xlsx", index=False)
        # Construct and return the summary
        return {
            'File name': file_name,
            'BOE Reference': self.generate_boe_reference(file_name),
            'Inv Date': self.extract_inv_date_from_pdf(file_path),
            # 'Inv Date': final_df['Inv Date'].iloc[0] if 'Inv Date' in final_df else None,
            'Plant': final_df['CW BOE\nREF. NO.'].iloc[1] if 'CW BOE\nREF. NO.' in final_df else None,
            # 'Plant': final_df['Plant'].iloc[0] if 'Plant' in final_df else None,
            'Remarks': final_df['Remarks'].iloc[0] if 'Remarks' in final_df else None,
            'QTY': final_df['QTY'].sum().round(4),
            'QTY': final_df['QTY'].sum().round(4),
            'Weight': final_df['Weight\n(KG)'].sum().round(4),
            'Volume': final_df['Volume\n(M3)'].sum().round(4),
            'Total Value': final_df['Total\nValue\n(AED)'].sum().round(4)
        }
    
    def adjust_column_widths(self):
        headers = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount())]
        if 'BOE Reference' in headers:
            col_index = headers.index('BOE Reference')
            self.tableWidget.setColumnWidth(col_index, 300)
        if 'Plant' in headers:
            col_index = headers.index('Plant')
            self.tableWidget.setColumnWidth(col_index, 300)

    def extract_tables_from_pdf(self, file_path):
        """Extracts tables from the PDF and returns a list of dataframes."""
        all_data = []
        with pdfplumber.open(file_path) as pdf:
            total_pages = len(pdf.pages)
            for page_index, page in enumerate(pdf.pages):
                table = page.extract_table()
                if table and page_index == total_pages - 1:
                    df = pd.DataFrame(table[1:-1], columns=table[0])
                else:
                    df = pd.DataFrame(table[1:], columns=table[0])
                all_data.append(df)
        return all_data

    def generate_boe_reference(self, file_name):
        """Generates the BOE reference given a filename."""
        date_str = datetime.now().strftime('%d%m%y')
        return f"CW{date_str}{file_name}"

    def update_status(self, message):
        """Updates the status label (label_3) with the given message."""
        self.label_3.setText(message)
        QApplication.processEvents()  # This will force the UI to update immediately

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ConverterApp()
    window.show()
    sys.exit(app.exec_())
