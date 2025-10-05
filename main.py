import sys
from PyQt6 import QtWidgets, QtCore, QtGui
from ui_DriveGoogleMultilinkDownloader import Ui_Form_DriveGoogleMultilinkDownloader
from ui_AddLink import Ui_Form_AddLink
import gdown
import os
import threading
import time
import re
import subprocess
# --- DriveDownloader Thread Class ---
class DownloadWorker(QtCore.QObject):
    finished = QtCore.pyqtSignal()
    progress_update = QtCore.pyqtSignal(int)               # % của file hiện tại
    log_message = QtCore.pyqtSignal(str, str)              # (msg, level)
    update_item_status = QtCore.pyqtSignal(int, str, str)  # (row, status, filename)
    total_update = QtCore.pyqtSignal(str)                  # "Total: n/total"
    speed_update = QtCore.pyqtSignal(str)                  # "4.10MB/s"

    def __init__(self, links_data, save_path):
        super().__init__()
        self.links_data = links_data          # [(link, row_index)]
        self.save_path = save_path
        self._is_paused = False
        self._is_stopped = False
        self._last_speed_emit = 0.0           # throttle cập nhật tốc độ

    # --- helpers ---
    def _to_direct(self, url: str) -> str:
        m = re.search(r"/file/d/([A-Za-z0-9_-]+)", url)
        if m:
            fid = m.group(1)
            return f"https://drive.google.com/uc?id={fid}&export=download"
        m = re.search(r"[?&]id=([^&]+)", url)
        return f"https://drive.google.com/uc?id={m.group(1)}&export=download" if m else url

    def _run_gdown(self, direct_url: str, out_folder: str):
        # -O <folder/> => gdown tự đặt tên file
        if not out_folder.endswith(os.sep):
            out_folder = out_folder + os.sep
        cmd = [sys.executable, "-m", "gdown", direct_url, "-O", out_folder, "--fuzzy"]
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1, encoding="utf-8", errors="replace"
        )
        return proc

    @QtCore.pyqtSlot()
    def run(self):
        total = len(self.links_data)
        done = 0
        self.log_message.emit("Starting download process...", "INFO")
        os.makedirs(self.save_path, exist_ok=True)

        for idx, (url, row) in enumerate(self.links_data, start=1):
            if self._is_stopped:
                self.log_message.emit("Download stopped by user.", "WARNING")
                break

            while self._is_paused and not self._is_stopped:
                time.sleep(0.1)

            # reset progress + speed mỗi file
            self.progress_update.emit(0)
            self.speed_update.emit("—")
            self.update_item_status.emit(row, "Downloading...", "Preparing...")
            self.log_message.emit(f"Processing link {idx}/{total}: {url}", "INFO")

            direct = self._to_direct(url)
            proc = None
            current_filename = None
            self._last_speed_emit = 0.0

            try:
                proc = self._run_gdown(direct, self.save_path)
                assert proc.stdout is not None

                for line in proc.stdout:
                    s = line.strip()

                    # nhấn Stop
                    if self._is_stopped:
                        proc.terminate()
                        raise RuntimeError("Stopped by user")

                    # bắt tên file từ 'To: ...'
                    if s.startswith("To:"):
                        tail = s.split("To:", 1)[1].strip()
                        current_filename = os.path.basename(tail.replace("\\", "/"))
                        self.update_item_status.emit(row, "Downloading...", current_filename)
                        continue

                    # bắt % + tốc độ (MB/s, KB/s, GB/s...) từ dòng progress
                    # ví dụ gdown/tqdm: "37%|█████▎ ... [00:12<00:18, 4.10MB/s]"
                    m_pct = re.match(r"^(\d+)%\|", s)
                    if m_pct:
                        self.progress_update.emit(int(m_pct.group(1)))

                        # cố gắng tách tốc độ nếu có: lấy phần ", 4.10MB/s]"
                        m_speed = re.search(r"\[\s*.*?,\s*([0-9.]+\s*(?:[KMG]?B)/s)\s*\]$", s)
                        if m_speed:
                            now = time.time()
                            if now - self._last_speed_emit > 0.3:  # throttle ~300ms
                                speed = m_speed.group(1).replace(" ", "")
                                self.speed_update.emit(speed)
                                self._last_speed_emit = now
                        continue

                    # lọc bớt log ồn
                    if (s == "Downloading..." or s == "" or
                        s.startswith("From (original):") or s.startswith("From (redirected):") or
                        s.startswith("From:") or s.startswith("To:") or
                        s.startswith("Processing") or s.startswith("Checking")):
                        continue

                    # còn lại: ghi log
                    self.log_message.emit(s, "INFO")

                proc.wait()
                if proc.returncode != 0:
                    raise RuntimeError(f"gdown exited with code {proc.returncode}")

                # hoàn tất file
                self.progress_update.emit(100)
                self.speed_update.emit("—")
                shown = current_filename or "Downloaded file"
                self.update_item_status.emit(row, "Completed", shown)
                self.log_message.emit(f"✅ Downloaded: {shown}", "SUCCESS")

                done += 1
                self.total_update.emit(f"Total: {done}/{total}")

            except Exception as e:
                self.progress_update.emit(0)
                self.speed_update.emit("—")
                self.update_item_status.emit(row, "Failed", "Error")
                self.log_message.emit(f"❌ Error: {e}", "ERROR")
                if proc and proc.poll() is None:
                    proc.kill()

        self.log_message.emit(f"All downloads attempted. Successfully downloaded {done} out of {total} links.", "INFO")
        self.finished.emit()

    # controls
    def pause(self):
        self._is_paused = True
        self.log_message.emit("Paused.", "INFO")

    def resume(self):
        self._is_paused = False
        self.log_message.emit("Resumed.", "INFO")

    def stop(self):
        self._is_stopped = True
        self.log_message.emit("Stopping...", "WARNING")

        
# --- Main Application Window ---
class DriveDownloaderMainWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Form_DriveGoogleMultilinkDownloader()
        self.ui.setupUi(self)

        self.download_thread = None
        self.worker = None
        self.is_downloading = False
        self.current_links_data = [] # Stores (link_text, row_index)
        self.ui.label_Total.setText("Total: 0/0")
        
        # Set default download directory and display in lineEdit
        self.default_save_directory = os.path.join(os.path.expanduser("~"), "Downloads", "DriveGoogleDownloads")
        os.makedirs(self.default_save_directory, exist_ok=True) # Ensure default directory exists
        self.ui.lineEdit_DestinationFolder.setText(self.default_save_directory)

        self._connect_signals()
        self._setup_table_widget()
        self._update_download_buttons_state()

    def _connect_signals(self):
        self.ui.pushButton_Add.clicked.connect(self._open_add_link_form)
        self.ui.pushButton_Delete.clicked.connect(self._delete_selected_link)
        self.ui.pushButton_DeleteAll.clicked.connect(self._delete_all_links)
        self.ui.pushButton_Download.clicked.connect(self._start_download)
        
        self.ui.pushButton_Pause.clicked.connect(self._pause_download)
        self.ui.pushButton_Stop.clicked.connect(self._stop_download)
        self.ui.pushButton_SelectFolder.clicked.connect(self._browse_save_folder)
        
        
        # Connect Up, Down, Edit buttons
        self.ui.pushButton_Up.clicked.connect(lambda: self._move_link_in_table(-1))
        self.ui.pushButton_Down.clicked.connect(lambda: self._move_link_in_table(1))
        self.ui.pushButton_Edit.clicked.connect(self._edit_selected_link)

    def _setup_table_widget(self):
        self.ui.tableWidget_ListLinkDriveGoogle.setColumnWidth(0, 400) # Link column
        self.ui.tableWidget_ListLinkDriveGoogle.setColumnWidth(1, 150) # Status column
        self.ui.tableWidget_ListLinkDriveGoogle.setColumnWidth(2, 150) # ETA/Filename column
        self.ui.tableWidget_ListLinkDriveGoogle.horizontalHeader().setStretchLastSection(True)

    def _open_add_link_form(self):
        # Pass self as parent to center the dialog
        self.add_link_window = AddLinkWindow(self) 
        # Connect the signal from AddLinkWindow to a slot in MainWindow
        self.add_link_window.links_added.connect(self._add_links_to_table)
        self.add_link_window.exec() # Use exec() for modal dialog

    def _add_links_to_table(self, new_links):
        current_row_count = self.ui.tableWidget_ListLinkDriveGoogle.rowCount()
        for i, link in enumerate(new_links):
            row_index = current_row_count + i
            self.ui.tableWidget_ListLinkDriveGoogle.insertRow(row_index)
            self.ui.tableWidget_ListLinkDriveGoogle.setItem(row_index, 0, QtWidgets.QTableWidgetItem(link))
            self.ui.tableWidget_ListLinkDriveGoogle.setItem(row_index, 1, QtWidgets.QTableWidgetItem("Pending"))
            self.ui.tableWidget_ListLinkDriveGoogle.setItem(row_index, 2, QtWidgets.QTableWidgetItem("N/A"))
            # When adding, update current_links_data
            self.current_links_data.append((link, row_index)) 

        self._log_message(f"Added {len(new_links)} new links to the list.", "INFO")

    def _delete_selected_link(self):
        selected_rows = sorted(list(set(index.row() for index in self.ui.tableWidget_ListLinkDriveGoogle.selectedIndexes())), reverse=True)
        if not selected_rows:
            self._log_message("No link selected to delete.", "WARNING")
            return
        
        for row in selected_rows:
            link_item = self.ui.tableWidget_ListLinkDriveGoogle.item(row, 0)
            if link_item:
                link_text = link_item.text()
                self.ui.tableWidget_ListLinkDriveGoogle.removeRow(row)
                self._log_message(f"Deleted link: {link_text}", "INFO")

        self._reindex_links_data() # Re-index after deletion
        
    def _reindex_links_data(self):
        # Rebuild current_links_data from the table widget after modifications
        new_links_data = []
        for r in range(self.ui.tableWidget_ListLinkDriveGoogle.rowCount()):
            link_item = self.ui.tableWidget_ListLinkDriveGoogle.item(r, 0)
            if link_item:
                link_text = link_item.text()
                new_links_data.append((link_text, r))
        self.current_links_data = new_links_data

    def _move_link_in_table(self, direction):  # -1: Up, +1: Down
            table = self.ui.tableWidget_ListLinkDriveGoogle
            selected_rows = sorted(list(set(index.row() for index in table.selectedIndexes())))

            if not selected_rows or len(selected_rows) != 1:
                self._log_message("Please select exactly one link to move.", "WARNING")
                return

            current_row = selected_rows[0]
            new_row = current_row + direction

            # Kiểm tra nếu vị trí mới nằm ngoài phạm vi của bảng
            if not (0 <= new_row < table.rowCount()):
                return

            # Hoán đổi nội dung của các ô giữa hàng hiện tại và hàng mới
            for col in range(table.columnCount()):
                current_item = table.takeItem(current_row, col)
                new_item = table.takeItem(new_row, col)

                table.setItem(current_row, col, new_item)
                table.setItem(new_row, col, current_item)
            
            # Chọn lại hàng đã được di chuyển đến vị trí mới
            table.selectRow(new_row)
            self._reindex_links_data()
            self._log_message(f"Moved link from row {current_row + 1} to {new_row + 1}.", "INFO")


    def _edit_selected_link(self):
        selected_items = self.ui.tableWidget_ListLinkDriveGoogle.selectedItems()
        if not selected_items or len(selected_items) > 1:
            self._log_message("Please select exactly one link to edit.", "WARNING")
            return
        
        current_row = selected_items[0].row()
        current_link_item = self.ui.tableWidget_ListLinkDriveGoogle.item(current_row, 0)
        
        if current_link_item:
            old_link = current_link_item.text()
            # QInputDialog for simple text input
            new_link, ok = QtWidgets.QInputDialog.getText(self, "Edit Link", "Edit Google Drive Link:", 
                                                        QtWidgets.QLineEdit.EchoMode.Normal, old_link)
            
            if ok and new_link != old_link:
                self.ui.tableWidget_ListLinkDriveGoogle.setItem(current_row, 0, QtWidgets.QTableWidgetItem(new_link))
                self.ui.tableWidget_ListLinkDriveGoogle.setItem(current_row, 1, QtWidgets.QTableWidgetItem("Pending")) # Reset status
                self.ui.tableWidget_ListLinkDriveGoogle.setItem(current_row, 2, QtWidgets.QTableWidgetItem("N/A"))     # Reset ETA
                self._reindex_links_data() # Re-index after editing
                self._log_message(f"Edited link in row {current_row}: '{old_link}' -> '{new_link}'", "INFO")
            elif ok:
                self._log_message("Link not changed.", "INFO")
            else:
                self._log_message("Edit cancelled.", "INFO")


    def _delete_all_links(self):
        reply = QtWidgets.QMessageBox.question(self, 'Delete All',
                                            "Are you sure you want to delete all links?",
                                            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                                            QtWidgets.QMessageBox.StandardButton.No)
        if reply == QtWidgets.QMessageBox.StandardButton.Yes:
            self.ui.tableWidget_ListLinkDriveGoogle.setRowCount(0)
            self.current_links_data = []
            self._log_message("All links deleted.", "INFO")

    def _log_message(self, message, level="INFO"):
        self.ui.textEdit_Log.append(f"[{level}] {message}")
        self.ui.textEdit_Log.verticalScrollBar().setValue(self.ui.textEdit_Log.verticalScrollBar().maximum()) # Auto-scroll

    def _update_progress_bar(self, value):
        self.ui.progressBar.setValue(value)

    def _update_table_item_status(self, row, status, eta_or_filename):
        # Ensure row exists before attempting to set item
        if 0 <= row < self.ui.tableWidget_ListLinkDriveGoogle.rowCount():
            self.ui.tableWidget_ListLinkDriveGoogle.setItem(row, 1, QtWidgets.QTableWidgetItem(status))
            self.ui.tableWidget_ListLinkDriveGoogle.setItem(row, 2, QtWidgets.QTableWidgetItem(eta_or_filename))
        else:
            self._log_message(f"Attempted to update non-existent row {row}. Link might have been deleted.", "WARNING")


    def _browse_save_folder(self):
        # Get initial directory from lineEdit, or default to home directory
        initial_dir = self.ui.lineEdit_DestinationFolder.text()
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~") # Fallback to home if current path is invalid
            
        folder_selected = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Destination Folder", initial_dir)
        if folder_selected:
            self.ui.lineEdit_DestinationFolder.setText(folder_selected)
            self._log_message(f"Destination folder set to: {folder_selected}", "INFO")

    def _start_download(self):
        if self.is_downloading:
            self._log_message("A download is already in progress.", "WARNING")
            return
        
        if not self.current_links_data:
            self._log_message("No links available for download. Please add links.", "WARNING")
            return

        save_path = self.ui.lineEdit_DestinationFolder.text()
        if not save_path:
            self._log_message("Please select a destination folder.", "ERROR")
            return
        
        # Ensure current_links_data is up-to-date with current table state
        self._reindex_links_data()
        
        # Reset all table items to "Pending" before starting a new download run
        for r in range(self.ui.tableWidget_ListLinkDriveGoogle.rowCount()):
            self.ui.tableWidget_ListLinkDriveGoogle.setItem(r, 1, QtWidgets.QTableWidgetItem("Pending"))
            self.ui.tableWidget_ListLinkDriveGoogle.setItem(r, 2, QtWidgets.QTableWidgetItem("N/A"))

        self.is_downloading = True
        self._update_download_buttons_state()
        self.ui.textEdit_Log.clear()
        self.ui.progressBar.setValue(0)

        # Create QThread and Worker
        self.download_thread = QtCore.QThread()
        # Pass a copy of current_links_data to worker to prevent modifications while running
        self.worker = DownloadWorker(list(self.current_links_data), save_path) 
        self.worker.moveToThread(self.download_thread)

        # Connect signals and slots
        self.download_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.download_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.download_thread.finished.connect(self.download_thread.deleteLater)
        self.download_thread.finished.connect(self._download_finished)

        self.worker.progress_update.connect(self.ui.progressBar.setValue)
        self.worker.log_message.connect(self._log_message)
        self.worker.update_item_status.connect(self._update_table_item_status)
        self.worker.total_update.connect(lambda t: self.ui.label_Total.setText(t))
        self.worker.speed_update.connect(lambda s: self.ui.label_Speed.setText(f"Speed: {s}"))
        # khi bắt đầu:
        self.ui.label_Total.setText(f"Total: 0/{len(self.current_links_data)}")
        # Start the thread
        
        self.ui.label_Speed.setText("Speed: —")
        self.ui.label_Total.setText(f"Total: 0/{len(self.current_links_data)}")

        self.download_thread.start()
        self._log_message("Download initiated.", "INFO")


    def _download_finished(self):
        self.is_downloading = False
        self._update_download_buttons_state()
        self._log_message("Download process completed or stopped.", "INFO")
        self.download_thread = None
        self.worker = None
        
    def _pause_download(self):
        if self.worker and self.is_downloading:
            if self.worker._is_paused:
                self.worker.resume()
                self.ui.pushButton_Pause.setText("PAUSE")
            else:
                self.worker.pause()
                self.ui.pushButton_Pause.setText("RESUME")
        else:
            self._log_message("No active download to pause/resume.", "WARNING")

    def _stop_download(self):
        if self.worker and self.is_downloading:
            self.worker.stop()
            # The _download_finished will be called when the worker actually stops
        else:
            self._log_message("No active download to stop.", "WARNING")

    def _update_download_buttons_state(self):
        # Enable/disable main action buttons
        self.ui.pushButton_Download.setEnabled(not self.is_downloading)
        self.ui.pushButton_Pause.setEnabled(self.is_downloading)
        self.ui.pushButton_Stop.setEnabled(self.is_downloading)
        
        # Enable/disable table modification buttons and folder selection
        can_modify_table = not self.is_downloading
        self.ui.pushButton_Add.setEnabled(can_modify_table)
        self.ui.pushButton_Delete.setEnabled(can_modify_table)
        self.ui.pushButton_DeleteAll.setEnabled(can_modify_table)
        self.ui.pushButton_Up.setEnabled(can_modify_table)
        self.ui.pushButton_Down.setEnabled(can_modify_table)
        self.ui.pushButton_Edit.setEnabled(can_modify_table)

        self.ui.pushButton_SelectFolder.setEnabled(can_modify_table)
        self.ui.lineEdit_DestinationFolder.setEnabled(can_modify_table)

        # Reset pause button text if not downloading
        if not self.is_downloading:
            self.ui.pushButton_Pause.setText("PAUSE") 

    def _update_total_label(self, text: str):
        self.ui.label_Total.setText(text)


# --- Add Link Window (now a QDialog) ---
class AddLinkWindow(QtWidgets.QDialog): # Changed from QWidget to QDialog
    # Custom signal to emit list of links back to the main window
    links_added = QtCore.pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent) # Pass parent for proper dialog behavior
        self.ui = Ui_Form_AddLink()
        self.ui.setupUi(self)
        self.setFixedSize(self.size()) # Make dialog non-resizable
        self._connect_signals()

    def _connect_signals(self):
        self.ui.pushButton_OK.clicked.connect(self._emit_links_and_close)
        self.ui.pushButton_Cancel.clicked.connect(self.close)

    def _emit_links_and_close(self):
        links_text = self.ui.textEdit_ListLink.toPlainText()
        if links_text.find(",") != -1:
            links = [link.strip() for link in links_text.split(',') if link.strip()]
        else:
            links = [link.strip() for link in links_text.split('\n') if link.strip()]
        
        if links:
            self.links_added.emit(links) # Emit the list of links
        
        self.accept() # Use accept() for QDialog when OK is pressed


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    main_window = DriveDownloaderMainWindow()
    main_window.show()
    sys.exit(app.exec())