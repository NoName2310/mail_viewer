import sys
from pathlib import Path
from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon
import win32com.client as win32
from email_viewer import Ui_MainWindow
import shutil
import json 
import resources

class EmailViewer(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)  # <-- Set up the UI from the generated class

        # Initialize variables
        self.base_folder = None
        self.current_folder = ""
        self.msg_files = []
        self.current_msg_path = None
        self.json_path = Path("link.json")

        # Connect signals
        self.btnLoadFolder.clicked.connect(self.load_folders)
        self.folderList.currentItemChanged.connect(self.on_folder_select)
        self.mailList.currentItemChanged.connect(self.on_mail_select)
        self.create_button.clicked.connect(self.create_email_copy)
        self.change_link.clicked.connect(self.change_folder)

        # Configure widgets
        self.contentBrowser.setOpenExternalLinks(True)

        # Set window properties
        self.setWindowTitle("Email Viewer (.MSG)")
        self.resize(1200, 800)
        self.splitter.setSizes([300, 600])

        # Load folder from link.json
        self.load_folder_from_json()
        
    def load_folders(self):
        folder_path = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select Mail Folder", "", QtWidgets.QFileDialog.ShowDirsOnly)
        if not folder_path: return

        self.base_folder = Path(folder_path)
        self.update_folder_list()
    
    def load_folder_from_json(self):
        try:
            if not self.json_path.exists():
                return  # Skip if file doesn't exist
            
            # Read JSON file with proper encoding handling
            with open(self.json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                folder_path = data.get("folder_form_mail", "")
                
                if not folder_path:
                    return  # Skip if key doesn't exist or is empty
                
                folder_path = Path(folder_path)
                if folder_path.exists():
                    self.base_folder = folder_path
                    self.update_folder_list()
                else:
                    QtWidgets.QMessageBox.warning(
                        self,
                        "Folder Not Found",
                        f"The saved folder path does not exist:\n{folder_path}"
                    )
                    
        except json.JSONDecodeError as e:
            QtWidgets.QMessageBox.critical(
                self,
                "Invalid JSON File",
                f"The JSON file is corrupted:\n{str(e)}"
            )
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                "Error",
                f"Failed to load folder:\n{str(e)}"
            )

    def update_folder_list(self):
        self.folders = [f for f in self.base_folder.iterdir() if f.is_dir()]
        if not self.folders:
            QtWidgets.QMessageBox.warning(self, "Warning", "No subfolders found!")
            return
        self.folderList.clear()
        self.folderList.addItems([folder.name for folder in self.folders])
        if self.folderList.count() > 0:
            self.folderList.setCurrentRow(0)

    def change_folder(self):
        # Show a Yes/No question dialog instead of warning
        reply = QtWidgets.QMessageBox.question(
            self,
            "Confirm Folder Change",
            "Do you want to change the default mail folder location?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No
        )
        
        # Only proceed if user clicked Yes
        if reply != QtWidgets.QMessageBox.Yes:
            return

        folder_path = QtWidgets.QFileDialog.getExistingDirectory(
            self, 
            "Select New Default Mail Folder", 
            "", 
            QtWidgets.QFileDialog.ShowDirsOnly
        )
        
        if not folder_path:
            return

        try:
            with open(self.json_path, "w", encoding="utf-8") as f:
                json.dump(
                    {"folder_form_mail": folder_path},
                    f,
                    indent=4,
                    ensure_ascii=False  # Preserves Japanese characters
                )
            QtWidgets.QMessageBox.information(
                self, "Success", "Folder link updated successfully."
            )
            self.base_folder = Path(folder_path)
            self.update_folder_list()
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to update folder link:\n{str(e)}"
            )

    def on_folder_select(self, current_item, previous_item):
        if not current_item:
            return
            
        self.current_folder = current_item.text()
        self.setWindowTitle(f"Email Viewer - {self.current_folder}")
        
        # Get .msg files in folder
        folder_path = self.base_folder / self.current_folder
        self.msg_files = [f.name for f in folder_path.glob('*.msg')]
        
        # Update mail list
        self.mailList.clear()
        
        # Add emails
        for msg_file in self.msg_files:
            item = QtWidgets.QListWidgetItem(msg_file)
            self.mailList.addItem(item)
        
        # Select first mail if available
        if self.mailList.count() > 0:
            self.mailList.setCurrentRow(0)
            
    def on_mail_select(self, current_item, previous_item):
        if not current_item:
            return
            
        msg_file = current_item.text()
        msg_path = self.base_folder / self.current_folder / msg_file
        self.current_msg_path = msg_path  # Store the current message path
        
        try:
            outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
            msg = outlook.OpenSharedItem(str(msg_path))
            
            # Display header info in HTML table
            header_html = f"""
            <table style="width:100%; border-collapse:collapse; font-family:'Meiryo UI'; font-size:10pt;">
                <tr style="border-bottom:1px solid #e5e5e5;">
                    <td style="padding:4px; vertical-align:top;"><b>Subject:</b></td>
                    <td style="padding:4px;">{self.escape_html(msg.Subject)}</td>
                </tr>
                <tr style="border-bottom:1px solid #e5e5e5;">
                    <td style="padding:4px; vertical-align:top;"><b>To:</b></td>
                    <td style="padding:4px;">{self.escape_html(msg.To)}</td>
                </tr>
                <tr style="border-bottom:1px solid #e5e5e5;">
                    <td style="padding:4px; vertical-align:top;"><b>CC:</b></td>
                    <td style="padding:4px;">{self.escape_html(msg.CC)}</td>
                </tr>
                <tr style="border-bottom:1px solid #e5e5e5;">
                    <td style="padding:4px; vertical-align:top;"><b>BCC:</b></td>
                    <td style="padding:4px;">{self.escape_html(msg.BCC)}</td>
                </tr>
            </table>
            """
            self.headerText.setHtml(header_html)
            
            # Display email content with original formatting
            if msg.HTMLBody:
                html_content = f"""
                <div style="
                    font-family:'Meiryo UI', sans-serif;
                    font-size:10pt;
                    color:#000000;
                    line-height:1.5;
                    max-width:100%;
                ">
                    {msg.HTMLBody}
                </div>
                """
                self.contentBrowser.setHtml(html_content)
            else:
                self.contentBrowser.setPlainText(msg.Body)
                
            msg.Close(0)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", 
                f"Failed to open .msg file:\n{str(e)}")
            
    def create_email_copy(self):
        if not self.current_msg_path:
            return

        try:
            # Move the .msg file to temp_folder
            original_path = self.current_msg_path
            Path(r"C:\Temp\mail_temp").mkdir(parents=True, exist_ok= True)
            move_path = Path(r"C:\Temp\mail_temp") / original_path.name

            # Move file
            shutil.move(original_path, move_path)

            # Copy the moved file back to the original folder with the same name
            save_path = original_path

            # Copy file
            shutil.copy2(move_path, save_path)

            # Show success message
            QtWidgets.QMessageBox.information(
                self, "Success",
                f"Email copy created successfully.\n{move_path.parent}"
            )

            # Open the moved file with Outlook
            try:
                outlook = win32.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                item = namespace.OpenSharedItem(str(move_path))
                item.Display()  # Open the email in a new window for editing
            except Exception as e:
                QtWidgets.QMessageBox.warning(
                    self, "Warning",
                    f"File moved but failed to open in Outlook:\n{str(e)}"
                )

        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error",
                f"Failed to create email copy:\n{str(e)}"
            )
            
    def escape_html(self, text):
        if not text:
            return ""
        return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')  # Modern style  
    app.setWindowIcon(QIcon(":/icons/email.png"))  
    window = EmailViewer()
    window.show()
    sys.exit(app.exec_())
