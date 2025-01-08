import os
import shutil
import win32com.client as win32
import logging


class DataDispatcher:
    def __init__(self, source_folder, target_folder):
        self.source_folder = source_folder
        self.target_folder = target_folder
        self.logger = self.setup_logger()

    def setup_logger(self):
        logger = logging.getLogger('DataDispatcher')
        logger.setLevel(logging.DEBUG)
        fh = logging.FileHandler('data_dispatcher.log')
        fh.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        logger.addHandler(fh)
        return logger

    def transfer_files(self):
        try:
            if not os.path.exists(self.target_folder):
                os.makedirs(self.target_folder)
                self.logger.info(f"Created target directory: {self.target_folder}")

            for filename in os.listdir(self.source_folder):
                source_file = os.path.join(self.source_folder, filename)
                target_file = os.path.join(self.target_folder, filename)
                shutil.copy2(source_file, target_file)
                self.logger.info(f"Copied {source_file} to {target_file}")
            self.logger.info("Data transfer completed successfully.")
        except Exception as e:
            self.logger.error(f"Error during file transfer: {e}")

    def list_files_in_folder(self, folder):
        try:
            files = os.listdir(folder)
            self.logger.info(f"Listing files in {folder}: {files}")
            return files
        except Exception as e:
            self.logger.error(f"Error listing files in folder {folder}: {e}")
            return []

    def automate_excel(self, excel_file):
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(excel_file)
            ws = wb.Sheets(1)
            # Example operation: Print the value of cell A1
            value = ws.Cells(1, 1).Value
            self.logger.info(f"Value in A1: {value}")
            wb.Close(SaveChanges=0)
            excel.Quit()
        except Exception as e:
            self.logger.error(f"Error automating Excel: {e}")


if __name__ == "__main__":
    dispatcher = DataDispatcher(source_folder='C:/source', target_folder='C:/target')
    dispatcher.transfer_files()
    dispatcher.list_files_in_folder('C:/target')
    dispatcher.automate_excel('C:/target/sample.xlsx')