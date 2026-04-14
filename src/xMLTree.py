#!/usr/bin/env python3
import copy
import os
import sys
import threading
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

PANDAS_AVAILABLE = False
BULK_AVAILABLE = False
DIFF_AVAILABLE = False

try:
    from PyQt6.QtCore import Qt, QTimer, pyqtSignal
    from PyQt6.QtGui import QColor, QBrush, QIcon, QPixmap, QPainter, QPen
    from PyQt6.QtWidgets import (
        QApplication,
        QAbstractItemView,
        QButtonGroup,
        QCheckBox,
        QComboBox,
        QDialog,
        QFileDialog,
        QGridLayout,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QProgressBar,
        QTableWidget,
        QTableWidgetItem,
        QTabWidget,
        QTextEdit,
        QTreeWidget,
        QTreeWidgetItem,
        QVBoxLayout,
        QWidget,
    )
except ImportError:
    from PySide6.QtCore import Qt, QTimer, Signal as pyqtSignal
    from PySide6.QtGui import QColor, QBrush, QIcon, QPixmap, QPainter, QPen
    from PySide6.QtWidgets import (
        QApplication,
        QAbstractItemView,
        QButtonGroup,
        QCheckBox,
        QComboBox,
        QDialog,
        QFileDialog,
        QGridLayout,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QProgressBar,
        QTableWidget,
        QTableWidgetItem,
        QTabWidget,
        QTextEdit,
        QTreeWidget,
        QTreeWidgetItem,
        QVBoxLayout,
        QWidget,
    )

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except Exception as e:
    print(f"Warning: pandas not available - some modes disabled. ({e})")

try:
    from bulk_update import XMLConfigurator, ConfigurationManager
    BULK_AVAILABLE = True
except Exception as e:
    print(f"Warning: bulk_update not available - Bulk Update mode disabled. ({e})")

try:
    from xml_difference import XMLComparator
    DIFF_AVAILABLE = True
except Exception as e:
    print(f"Warning: xml_difference not available - XML Difference mode disabled. ({e})")

class XMLTreeWidget(QTreeWidget):
    fileDropped = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".xml"):
                self.fileDropped.emit(path)
                break
        event.acceptProposedAction()


class XMLEditorWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.item_to_element = {}
        self.etree = None
        self.root_element = None
        self.file_path = None
        self.search_results = []
        self.current_search_index = -1

        self.treeview = XMLTreeWidget()
        self.treeview.setColumnCount(2)
        self.treeview.setHeaderLabels(["Tag / Attributes", "Text"])
        self.treeview.itemDoubleClicked.connect(self.edit_item)
        self.treeview.fileDropped.connect(self.load_xml_from_path)

        self.search_mode = "content"
        self.search_mode_group = QButtonGroup(self)

        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("Search XML...")
        self.prev_button = QPushButton("Previous")
        self.next_button = QPushButton("Next")
        self.prev_button.setEnabled(False)
        self.next_button.setEnabled(False)

        self._build_ui()

    def _build_ui(self):
        root_layout = QVBoxLayout(self)

        button_row = QHBoxLayout()
        for text, cb in [
            ("Load XML", self.load_xml),
            ("Save XML", self.save_xml),
            ("Add Child", self.add_child),
            ("Duplicate", self.duplicate_item),
            ("Delete Selected", self.delete_item),
        ]:
            btn = QPushButton(text)
            btn.clicked.connect(cb)
            button_row.addWidget(btn)
        button_row.addStretch(1)
        root_layout.addLayout(button_row)

        search_mode_row = QHBoxLayout()
        search_mode_row.addWidget(QLabel("Search Mode:"))
        content_btn = QCheckBox("Content")
        path_btn = QCheckBox("Path")
        content_btn.setChecked(True)
        self.search_mode_group.setExclusive(True)
        self.search_mode_group.addButton(content_btn)
        self.search_mode_group.addButton(path_btn)
        content_btn.toggled.connect(lambda checked: self._set_mode("content", checked))
        path_btn.toggled.connect(lambda checked: self._set_mode("path", checked))
        search_mode_row.addWidget(content_btn)
        search_mode_row.addWidget(path_btn)
        search_mode_row.addStretch(1)
        root_layout.addLayout(search_mode_row)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Search:"))
        search_row.addWidget(self.search_entry)
        search_btn = QPushButton("Search")
        search_btn.clicked.connect(self.perform_search)
        self.prev_button.clicked.connect(self.prev_match)
        self.next_button.clicked.connect(self.next_match)
        search_row.addWidget(search_btn)
        search_row.addWidget(self.prev_button)
        search_row.addWidget(self.next_button)
        root_layout.addLayout(search_row)

        root_layout.addWidget(self.treeview)
        root_layout.addWidget(QLabel("Tip: Drag and drop an XML file on the tree to open it."))

    def _set_mode(self, mode_name, checked):
        if checked:
            self.search_mode = mode_name

    def load_xml(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open XML", "", "XML files (*.xml)")
        if file_path:
            self.load_xml_from_path(file_path)

    def load_xml_from_path(self, path):
        try:
            self.file_path = path
            self.etree = ET.parse(self.file_path)
            self.root_element = self.etree.getroot()
            self.treeview.clear()
            self.item_to_element = {}
            self.populate_tree(None, self.root_element)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load XML:\n{e}")

    def populate_tree(self, parent_item, element):
        item = self._add_element_to_treeview(parent_item, element)
        for child in element:
            self.populate_tree(item, child)
        self.search_results = []
        self.current_search_index = -1
        self.prev_button.setEnabled(False)
        self.next_button.setEnabled(False)
        self.search_entry.clear()

    def edit_item(self, item, _column):
        element = self.item_to_element.get(id(item))
        if not element:
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Element")
        layout = QVBoxLayout(dialog)

        tag_entry = QLineEdit(element.tag)
        text_entry = QLineEdit(element.text or "")
        attrib_entry = QLineEdit(" ".join([f"{k}={v}" for k, v in element.attrib.items()]))

        layout.addWidget(QLabel("Tag:"))
        layout.addWidget(tag_entry)
        layout.addWidget(QLabel("Text:"))
        layout.addWidget(text_entry)
        layout.addWidget(QLabel("Attributes (key=value space-separated):"))
        layout.addWidget(attrib_entry)

        save_btn = QPushButton("Save")
        layout.addWidget(save_btn)

        def save_edit():
            element.tag = tag_entry.text().strip()
            element.text = text_entry.text()
            new_attrib = {}
            for pair in attrib_entry.text().split():
                if "=" in pair:
                    k, v = pair.split("=", 1)
                    new_attrib[k] = v
            element.attrib = new_attrib
            self._update_item_view(item, element)
            dialog.accept()

        save_btn.clicked.connect(save_edit)
        dialog.exec()

    def add_child(self):
        selected = self.treeview.selectedItems()
        if selected:
            parent_item = selected[0]
            parent_element = self.item_to_element[id(parent_item)]
            new_element = ET.SubElement(parent_element, 'new_tag')
            new_element.text = 'new_text'

            new_item = self._add_element_to_treeview(parent_item, new_element)
            parent_item.setExpanded(True)
            self.treeview.scrollToItem(new_item)

    def _add_element_to_treeview(self, parent_item, element):
        attrib_str = ' '.join([f'{k}="{v}"' for k, v in element.attrib.items()])
        display_text = element.tag
        if attrib_str:
            display_text += f' {attrib_str}'
        text_value = element.text.strip() if element.text else ''
        new_item = QTreeWidgetItem([display_text, text_value])
        if parent_item is None:
            self.treeview.addTopLevelItem(new_item)
        else:
            parent_item.addChild(new_item)
        self.item_to_element[id(new_item)] = element
        return new_item

    def _update_item_view(self, item, element):
        attrib_str = ' '.join([f'{k}="{v}"' for k, v in element.attrib.items()])
        display_text = element.tag if not attrib_str else f"{element.tag} {attrib_str}"
        item.setText(0, display_text)
        item.setText(1, element.text.strip() if element.text else "")

    def delete_item(self):
        selected = self.treeview.selectedItems()
        if selected:
            item = selected[0]
            element = self.item_to_element[id(item)]
            parent_item = item.parent()
            if parent_item:
                parent_element = self.item_to_element[id(parent_item)]
                parent_element.remove(element)
                parent_item.removeChild(item)
                del self.item_to_element[id(item)]
            else:
                QMessageBox.warning(self, "Warning", "Cannot delete root element.")

    def save_xml(self):
        if self.file_path and self.etree:
            save_path, _ = QFileDialog.getSaveFileName(
                self, "Save XML", self.file_path, "XML files (*.xml)"
            )
            if save_path:
                self.etree.write(save_path, encoding='utf-8', xml_declaration=True)
                self.file_path = save_path

    def _iter_items(self):
        stack = [self.treeview.topLevelItem(i) for i in range(self.treeview.topLevelItemCount())]
        while stack:
            item = stack.pop(0)
            yield item
            for i in range(item.childCount()):
                stack.append(item.child(i))

    def _find_by_path(self, parent_item, path_parts, current_index):
        if current_index >= len(path_parts):
            return
        current_part = path_parts[current_index]
        children = []
        if parent_item is None:
            children = [self.treeview.topLevelItem(i) for i in range(self.treeview.topLevelItemCount())]
        else:
            children = [parent_item.child(i) for i in range(parent_item.childCount())]

        for item in children:
            element = self.item_to_element[id(item)]
            if current_part.isdigit():
                expected_tag = f"i{current_part}"
                if element.tag == expected_tag:
                    if current_index == len(path_parts) - 1:
                        self.search_results.append(item)
                    else:
                        self._find_by_path(item, path_parts, current_index + 1)
            else:
                tag_name = element.tag.split('}')[-1] if '}' in element.tag else element.tag
                if current_part.lower() == tag_name.lower():
                    if current_index == len(path_parts) - 1:
                        self.search_results.append(item)
                    else:
                        self._find_by_path(item, path_parts, current_index + 1)

            self._find_by_path(item, path_parts, current_index)

    def perform_search(self):
        query = self.search_entry.text().strip()
        if not query:
            QMessageBox.information(self, "Info", "Enter search term.")
            return

        for item in self._iter_items():
            item.setBackground(0, QBrush())
            item.setBackground(1, QBrush())

        self.search_results = []
        self.current_search_index = -1

        if self.search_mode == "path":
            self._find_by_path(None, query.split("."), 0)
        else:
            self._find_matches(None, query.lower())

        if self.search_results:
            self.prev_button.setEnabled(True)
            self.next_button.setEnabled(True)
            self.next_match()
            QMessageBox.information(self, "Info", f"Found {len(self.search_results)} matches.")
        else:
            self.prev_button.setEnabled(False)
            self.next_button.setEnabled(False)
            QMessageBox.information(self, "Info", "No matches found.")

    def _find_matches(self, parent_item, query_lower):
        if parent_item is None:
            children = [self.treeview.topLevelItem(i) for i in range(self.treeview.topLevelItemCount())]
        else:
            children = [parent_item.child(i) for i in range(parent_item.childCount())]
        for item in children:
            element = self.item_to_element[id(item)]
            found_match = False

            for key in element.attrib:
                local_name = key.split('}')[-1] if '}' in key else key
                if query_lower in local_name.lower():
                    self.search_results.append(item)
                    found_match = True
                    break

            if not found_match:
                for value in element.attrib.values():
                    if query_lower in value.lower():
                        self.search_results.append(item)
                        found_match = True
                        break
            
            # 3. Check ELEMENT TAG NAME (case insensitive)
            if not found_match:
                tag_name = element.tag.split('}')[-1] if '}' in element.tag else element.tag
                if query_lower in tag_name.lower():
                    self.search_results.append(item)
                    found_match = True

            if not found_match and element.text:
                if query_lower in element.text.lower():
                    self.search_results.append(item)
                    found_match = True

            self._find_matches(item, query_lower)

    def prev_match(self):
        if self.search_results:
            self.current_search_index = (self.current_search_index - 1) % len(self.search_results)
            self._select_match()

    def next_match(self):
        if self.search_results:
            self.current_search_index = (self.current_search_index + 1) % len(self.search_results)
            self._select_match()

    def _select_match(self):
        if not self.search_results:
            return
        item = self.search_results[self.current_search_index]
        for match_item in self.search_results:
            match_item.setBackground(0, QBrush())
            match_item.setBackground(1, QBrush())
        item.setBackground(0, QBrush(QColor("#9fc5e8")))
        item.setBackground(1, QBrush(QColor("#9fc5e8")))
        self.treeview.setCurrentItem(item)

        parent = item.parent()
        while parent:
            parent.setExpanded(True)
            parent = parent.parent()
        self.treeview.scrollToItem(item)

    def duplicate_item(self):
        selected = self.treeview.selectedItems()
        if not selected:
            QMessageBox.information(self, "Info", "Please select an element to duplicate.")
            return

        item = selected[0]
        element = self.item_to_element[id(item)]
        parent_item = item.parent()

        if not parent_item:
            QMessageBox.warning(self, "Warning", "Cannot duplicate root element.")
            return

        parent_element = self.item_to_element[id(parent_item)]
        new_element = copy.deepcopy(element)

        if element.tag.startswith('i') and element.tag[1:].isdigit():
            new_index = self._get_next_available_index(parent_element, element.tag)
            new_element.tag = f"i{new_index}"
            for alias_elem in new_element.findall('.//Alias'):
                if alias_elem.text and element.tag in alias_elem.text:
                    old_alias = alias_elem.text
                    base_name = old_alias.rsplit(element.tag[1:], 1)[0]
                    alias_elem.text = f"{base_name}{new_index}"

        parent_element.append(new_element)
        new_item = self._add_element_to_treeview(parent_item, new_element)
        for child in new_element:
            self._populate_duplicate_children(new_item, child)

        parent_item.setExpanded(True)
        self.treeview.setCurrentItem(new_item)
        self.treeview.scrollToItem(new_item)
        QMessageBox.information(self, "Info", f"Successfully duplicated {element.tag} as {new_element.tag}")
    def _get_next_available_index(self, parent_element, tag_pattern):
        if not tag_pattern.startswith('i') or not tag_pattern[1:].isdigit():
            return 1
        existing_indices = []
        for child in parent_element:
            if child.tag.startswith('i') and child.tag[1:].isdigit():
                try:
                    existing_indices.append(int(child.tag[1:]))
                except ValueError:
                    continue
        if not existing_indices:
            return 1
        return max(existing_indices) + 1

    def _populate_duplicate_children(self, parent_item, element):
        item = self._add_element_to_treeview(parent_item, element)
        for child in element:
            self._populate_duplicate_children(item, child)

class ExcelExtractorWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.file_paths = []
        self.all_columns = []
        self.output_dir = str(Path("Extracted_Output"))
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        file_row = QHBoxLayout()
        self.btn_files = QPushButton("Browse Files")
        self.btn_clear = QPushButton("Clear Files")
        self.file_label = QLabel("No files selected")
        self.btn_files.clicked.connect(self.select_files)
        self.btn_clear.clicked.connect(self.clear_files)
        file_row.addWidget(self.btn_files)
        file_row.addWidget(self.btn_clear)
        file_row.addWidget(self.file_label)
        file_row.addStretch(1)
        layout.addLayout(file_row)

        self.file_list = QListWidget()
        layout.addWidget(self.file_list)

        search_row = QHBoxLayout()
        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("Search column names...")
        self.search_entry.textChanged.connect(self.filter_columns)
        search_row.addWidget(self.search_entry)
        sel_all = QPushButton("Select All")
        desel_all = QPushButton("Deselect All")
        sel_all.clicked.connect(lambda: self.toggle_all(True))
        desel_all.clicked.connect(lambda: self.toggle_all(False))
        search_row.addWidget(sel_all)
        search_row.addWidget(desel_all)
        layout.addLayout(search_row)

        self.col_list = QListWidget()
        layout.addWidget(self.col_list)

        out_row = QHBoxLayout()
        self.out_label = QLabel(self.output_dir)
        out_btn = QPushButton("Change Output Folder")
        out_btn.clicked.connect(self.change_output_dir)
        self.extract_btn = QPushButton("Extract Columns")
        self.extract_btn.clicked.connect(self.extract_columns)
        out_row.addWidget(QLabel("Output folder:"))
        out_row.addWidget(self.out_label)
        out_row.addWidget(out_btn)
        out_row.addWidget(self.extract_btn)
        out_row.addStretch(1)
        layout.addLayout(out_row)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)
        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select files", "", "Excel/CSV (*.xlsx *.xls *.csv)")
        if not files:
            return
        self.file_paths = files
        self.file_list.clear()
        for f in files:
            self.file_list.addItem(os.path.basename(f))
        self.file_label.setText(f"{len(files)} file(s) selected")
        self.load_columns()

    def clear_files(self):
        self.file_paths = []
        self.all_columns = []
        self.file_list.clear()
        self.col_list.clear()
        self.file_label.setText("No files selected")

    def load_columns(self):
        self.all_columns = []
        seen = set()
        for path in self.file_paths:
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path, nrows=0)
            else:
                df = pd.read_excel(path, nrows=0)
            for col in df.columns:
                col = str(col).strip()
                if col and col not in seen:
                    seen.add(col)
                    self.all_columns.append(col)
        self.col_list.clear()
        for col in self.all_columns:
            item = QListWidgetItem(col)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.col_list.addItem(item)

    def filter_columns(self):
        query = self.search_entry.text().strip().lower()
        for i in range(self.col_list.count()):
            item = self.col_list.item(i)
            item.setHidden(query not in item.text().lower() if query else False)

    def toggle_all(self, checked):
        target_state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for i in range(self.col_list.count()):
            item = self.col_list.item(i)
            if not item.isHidden():
                item.setCheckState(target_state)

    def change_output_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if directory:
            self.output_dir = directory
            self.out_label.setText(directory)

    def extract_columns(self):
        selected = [
            self.col_list.item(i).text()
            for i in range(self.col_list.count())
            if self.col_list.item(i).checkState() == Qt.CheckState.Checked
        ]
        if not self.file_paths:
            QMessageBox.warning(self, "Warning", "Please select files first.")
            return
        if not selected:
            QMessageBox.warning(self, "Warning", "Please select columns.")
            return
        os.makedirs(self.output_dir, exist_ok=True)
        self.progress.setValue(0)
        for idx, path in enumerate(self.file_paths, start=1):
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path)
            matched = [c for c in df.columns if str(c).strip() in selected]
            output_df = df[matched] if matched else pd.DataFrame()
            out_name = f"{Path(path).stem}_extracted.xlsx"
            output_df.to_excel(os.path.join(self.output_dir, out_name), index=False, engine="openpyxl")
            self.progress.setValue(int((idx / len(self.file_paths)) * 100))
        self.status_label.setText("Done")
        QMessageBox.information(self, "Done", f"Saved extracted files to:\n{self.output_dir}")


class BulkUpdateWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.configurator = XMLConfigurator()
        self.mappings_df = None
        self.checked_mappings = []
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        grid = QGridLayout()
        self.xml_entry = QLineEdit()
        self.map_entry = QLineEdit()
        self.path_col = QLineEdit("path")
        self.value_col = QLineEdit("value")
        browse_xml = QPushButton("Browse XML")
        browse_map = QPushButton("Browse Mapping")
        load_xml = QPushButton("Load XML")
        load_map = QPushButton("Load Mappings")
        template_btn = QPushButton("Create Template CSV")

        browse_xml.clicked.connect(self.browse_xml)
        browse_map.clicked.connect(self.browse_mapping)
        load_xml.clicked.connect(self.load_xml)
        load_map.clicked.connect(self.load_mapping)
        template_btn.clicked.connect(self.create_template)

        grid.addWidget(QLabel("XML File:"), 0, 0)
        grid.addWidget(self.xml_entry, 0, 1)
        grid.addWidget(browse_xml, 0, 2)
        grid.addWidget(load_xml, 0, 3)
        grid.addWidget(QLabel("Mapping File:"), 1, 0)
        grid.addWidget(self.map_entry, 1, 1)
        grid.addWidget(browse_map, 1, 2)
        grid.addWidget(load_map, 1, 3)
        grid.addWidget(QLabel("Path Column:"), 2, 0)
        grid.addWidget(self.path_col, 2, 1)
        grid.addWidget(QLabel("Value Column:"), 2, 2)
        grid.addWidget(self.value_col, 2, 3)
        grid.addWidget(template_btn, 2, 4)
        layout.addLayout(grid)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Status", "Path", "Current Value", "New Value"])
        layout.addWidget(self.table)

        action_row = QHBoxLayout()
        self.check_btn = QPushButton("Check Mappings")
        self.update_btn = QPushButton("Update XML")
        self.save_btn = QPushButton("Save As")
        self.check_btn.clicked.connect(self.check_mappings)
        self.update_btn.clicked.connect(self.update_xml)
        self.save_btn.clicked.connect(self.save_xml)
        action_row.addWidget(self.check_btn)
        action_row.addWidget(self.update_btn)
        action_row.addWidget(self.save_btn)
        action_row.addStretch(1)
        layout.addLayout(action_row)

        self.progress = QProgressBar()
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.progress)
        layout.addWidget(self.log)

    def append_log(self, message):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.append(f"{ts} {message}")

    def browse_xml(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select XML", "", "XML files (*.xml)")
        if path:
            self.xml_entry.setText(path)

    def browse_mapping(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Mapping", "", "CSV/Excel (*.csv *.xlsx *.xls)")
        if path:
            self.map_entry.setText(path)

    def load_xml(self):
        if self.configurator.load_xml(self.xml_entry.text().strip()):
            self.append_log("Loaded XML.")
        else:
            QMessageBox.critical(self, "Error", "Failed to load XML.")

    def load_mapping(self):
        self.mappings_df = self.configurator.load_mapping_file(self.map_entry.text().strip())
        if self.mappings_df is not None:
            self.append_log(f"Loaded mappings: {len(self.mappings_df)} rows.")
        else:
            QMessageBox.critical(self, "Error", "Failed to load mapping file.")

    def check_mappings(self):
        if self.mappings_df is None:
            QMessageBox.warning(self, "Warning", "Load mappings first.")
            return
        path_col = self.path_col.text().strip()
        value_col = self.value_col.text().strip()
        self.checked_mappings = self.configurator.check_mappings(
            self.mappings_df.to_dict("records"), path_col, value_col
        )
        self.table.setRowCount(0)
        found = 0
        for row in self.checked_mappings:
            row_idx = self.table.rowCount()
            self.table.insertRow(row_idx)
            status = "Found" if row["found"] else "Not Found"
            if row["found"]:
                found += 1
            for col_idx, value in enumerate([status, row["path"], row["current_value"], row["new_value"]]):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))
        self.append_log(f"Checked mappings. Found {found}/{len(self.checked_mappings)}.")

    def update_xml(self):
        if not self.checked_mappings:
            QMessageBox.warning(self, "Warning", "Check mappings first.")
            return
        apply_rows = [m for m in self.checked_mappings if m["found"]]
        result = self.configurator.apply_mappings(apply_rows)
        self.append_log(f"Applied: {result['applied']}, failed: {result['failed']}")
        self.progress.setValue(100)

    def save_xml(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save XML", "", "XML files (*.xml)")
        if path and self.configurator.save_xml(path):
            self.append_log(f"Saved: {path}")
            QMessageBox.information(self, "Saved", f"XML saved to:\n{path}")

    def create_template(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Template CSV", "", "CSV files (*.csv)")
        if path:
            ok = ConfigurationManager.create_template_csv(path, self.xml_entry.text().strip() or None)
            if ok:
                self.append_log(f"Template created: {path}")
            else:
                QMessageBox.critical(self, "Error", "Failed to create template CSV.")


class XMLDifferenceWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.comparator = XMLComparator()
        self.differences = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        grid = QGridLayout()
        self.xml1 = QLineEdit()
        self.xml2 = QLineEdit()
        b1 = QPushButton("Browse 1")
        b2 = QPushButton("Browse 2")
        l1 = QPushButton("Load 1")
        l2 = QPushButton("Load 2")
        b1.clicked.connect(lambda: self._browse(self.xml1))
        b2.clicked.connect(lambda: self._browse(self.xml2))
        l1.clicked.connect(lambda: self._load(1))
        l2.clicked.connect(lambda: self._load(2))
        grid.addWidget(QLabel("First XML:"), 0, 0)
        grid.addWidget(self.xml1, 0, 1)
        grid.addWidget(b1, 0, 2)
        grid.addWidget(l1, 0, 3)
        grid.addWidget(QLabel("Second XML:"), 1, 0)
        grid.addWidget(self.xml2, 1, 1)
        grid.addWidget(b2, 1, 2)
        grid.addWidget(l2, 1, 3)
        layout.addLayout(grid)

        filters = QHBoxLayout()
        self.cb_missing = QCheckBox("Missing")
        self.cb_changed = QCheckBox("Changed")
        self.cb_added = QCheckBox("Added")
        self.cb_missing.setChecked(True)
        self.cb_changed.setChecked(True)
        self.cb_added.setChecked(True)
        filters.addWidget(self.cb_missing)
        filters.addWidget(self.cb_changed)
        filters.addWidget(self.cb_added)
        filters.addStretch(1)
        layout.addLayout(filters)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Status", "Path", "Value 1", "Value 2"])
        layout.addWidget(self.table)

        btn_row = QHBoxLayout()
        compare_btn = QPushButton("Compare")
        export_btn = QPushButton("Export to Excel")
        compare_btn.clicked.connect(self.compare_xml)
        export_btn.clicked.connect(self.export_excel)
        btn_row.addWidget(compare_btn)
        btn_row.addWidget(export_btn)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

    def _browse(self, target):
        path, _ = QFileDialog.getOpenFileName(self, "Select XML", "", "XML files (*.xml)")
        if path:
            target.setText(path)

    def _load(self, which):
        target = self.xml1 if which == 1 else self.xml2
        ok = self.comparator.load_xml(target.text().strip(), which)
        if ok:
            self.log.append(f"Loaded XML {which}.")
        else:
            QMessageBox.critical(self, "Error", f"Failed to load XML {which}.")

    def compare_xml(self):
        self.differences = self.comparator.compare()
        self.table.setRowCount(0)
        for key in ["missing", "changed", "added"]:
            for row in self.differences[key]:
                idx = self.table.rowCount()
                self.table.insertRow(idx)
                self.table.setItem(idx, 0, QTableWidgetItem(row["status"]))
                self.table.setItem(idx, 1, QTableWidgetItem(row["path"]))
                self.table.setItem(idx, 2, QTableWidgetItem(str(row["value1"])))
                self.table.setItem(idx, 3, QTableWidgetItem(str(row["value2"])))
        self.log.append("Comparison complete.")

    def export_excel(self):
        if not self.differences:
            QMessageBox.warning(self, "Warning", "Run comparison first.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export Excel", "xml_differences.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        ok = self.comparator.export_to_excel(
            self.differences,
            path,
            include_missing=self.cb_missing.isChecked(),
            include_changed=self.cb_changed.isChecked(),
            include_added=self.cb_added.isChecked(),
        )
        if ok:
            self.log.append(f"Exported to {path}")
            QMessageBox.information(self, "Done", f"Exported to:\n{path}")
        else:
            QMessageBox.critical(self, "Error", "Export failed.")


class MultiAppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        app_icon_path = Path(__file__).resolve().parent.parent / "icons" / "xMLTree.png"
        if app_icon_path.exists():
            self.setWindowIcon(QIcon(str(app_icon_path)))
        self.setWindowTitle("xMLTree - Utility Tool")
        self.resize(1280, 820)
        self.is_dark_mode = True

        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.TabPosition.North)
        self.tabs.addTab(XMLEditorWidget(), "XML Editor")
        self.tabs.addTab(self._build_mode_widget(PANDAS_AVAILABLE, ExcelExtractorWidget, "pandas is required"), "Excel Column Extractor")
        self.tabs.addTab(self._build_mode_widget(BULK_AVAILABLE, BulkUpdateWidget, "bulk_update module is required"), "Bulk Update")
        self.tabs.addTab(self._build_mode_widget(DIFF_AVAILABLE, XMLDifferenceWidget, "xml_difference module is required"), "XML Difference")

        self.theme_toggle_btn = QPushButton("Day Mode")
        self.sun_icon, self.moon_icon = self._create_theme_icons()
        self.theme_toggle_btn.setFixedSize(42, 42)
        self.theme_toggle_btn.setText("")
        self.theme_toggle_btn.setIconSize(self.theme_toggle_btn.size() * 0.6)
        self.theme_toggle_btn.clicked.connect(self.toggle_theme)

        top_row = QHBoxLayout()
        top_row.addStretch(1)
        top_row.addWidget(self.theme_toggle_btn)

        root_layout = QVBoxLayout()
        root_layout.addLayout(top_row)
        root_layout.addWidget(self.tabs)

        container = QWidget()
        container.setLayout(root_layout)
        self.setCentralWidget(container)
        self._apply_theme()

    def _create_theme_icons(self):
        icons_dir = Path(__file__).resolve().parent.parent / "icons"
        day_icon_path = icons_dir / "day.png"
        night_icon_path = icons_dir / "night.png"

        if day_icon_path.exists() and night_icon_path.exists():
            return QIcon(str(day_icon_path)), QIcon(str(night_icon_path))

        def make_sun_icon():
            pixmap = QPixmap(24, 24)
            pixmap.fill(Qt.GlobalColor.transparent)
            painter = QPainter(pixmap)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            painter.setPen(QPen(QColor("#f6c343"), 2))
            painter.setBrush(QColor("#f6c343"))
            painter.drawEllipse(7, 7, 10, 10)
            for x1, y1, x2, y2 in [
                (12, 1, 12, 5),
                (12, 19, 12, 23),
                (1, 12, 5, 12),
                (19, 12, 23, 12),
                (4, 4, 7, 7),
                (17, 17, 20, 20),
                (4, 20, 7, 17),
                (17, 7, 20, 4),
            ]:
                painter.drawLine(x1, y1, x2, y2)
            painter.end()
            return QIcon(pixmap)

        def make_moon_icon():
            pixmap = QPixmap(24, 24)
            pixmap.fill(Qt.GlobalColor.transparent)
            painter = QPainter(pixmap)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor("#cfd8e6"))
            painter.drawEllipse(4, 3, 16, 16)
            painter.setBrush(QColor("#121822"))
            painter.drawEllipse(10, 2, 14, 18)
            painter.end()
            return QIcon(pixmap)

        return make_sun_icon(), make_moon_icon()

    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode
        self._apply_theme()

    def _apply_theme(self):
        if self.is_dark_mode:
            self.theme_toggle_btn.setIcon(self.sun_icon)
            self.theme_toggle_btn.setToolTip("Switch to day mode")
            self.setStyleSheet(
                """
                QWidget { background: #121822; color: #e8eef8; font-family: Segoe UI; font-size: 13px; }
                QLineEdit, QTextEdit, QListWidget, QTreeWidget, QTableWidget, QComboBox {
                    background: #1b2533; border: 1px solid #2f415b; border-radius: 8px; padding: 6px;
                }
                QPushButton {
                    background: #2d7dff; color: white; border: none; border-radius: 8px; padding: 8px 14px;
                }
                QPushButton:hover { background: #4690ff; }
                QPushButton:pressed { background: #1f67d9; }
                QPushButton:disabled { background: #4b5d78; color: #b6c2d7; }
                QTabWidget::pane { border: 1px solid #2f415b; border-radius: 8px; }
                QTabBar::tab { background: #253347; color: #d9e2f2; padding: 10px 16px; margin-right: 4px; border-top-left-radius: 8px; border-top-right-radius: 8px; }
                QTabBar::tab:selected { background: #2d7dff; color: #ffffff; }
                """
            )
        else:
            self.theme_toggle_btn.setIcon(self.moon_icon)
            self.theme_toggle_btn.setToolTip("Switch to night mode")
            self.setStyleSheet(
                """
                QWidget { background: #f4f7fc; color: #1a2533; font-family: Segoe UI; font-size: 13px; }
                QLineEdit, QTextEdit, QListWidget, QTreeWidget, QTableWidget, QComboBox {
                    background: #ffffff; border: 1px solid #c7d2e2; border-radius: 8px; padding: 6px;
                }
                QPushButton {
                    background: #2d7dff; color: white; border: none; border-radius: 8px; padding: 8px 14px;
                }
                QPushButton:hover { background: #4690ff; }
                QPushButton:pressed { background: #1f67d9; }
                QPushButton:disabled { background: #9fb6d6; color: #e8effa; }
                QTabWidget::pane { border: 1px solid #c7d2e2; border-radius: 8px; }
                QTabBar::tab { background: #dbe5f4; color: #2a3a52; padding: 10px 16px; margin-right: 4px; border-top-left-radius: 8px; border-top-right-radius: 8px; }
                QTabBar::tab:selected { background: #2d7dff; color: #ffffff; }
                """
            )

    def _build_mode_widget(self, enabled, cls, message):
        if enabled:
            return cls()
        widget = QWidget()
        layout = QVBoxLayout(widget)
        label = QLabel(message)
        layout.addWidget(label)
        layout.addStretch(1)
        return widget


if __name__ == "__main__":
    print("Arguments:", sys.argv)
    if os.name == "nt":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("xMLTree.UtilityTool")
        except Exception:
            pass
    if "--test" in sys.argv:
        try:
            print("All imports successful")
            sys.exit(0)
        except ImportError as e:
            print(f"Import error: {e}")
            sys.exit(1)
    app = QApplication(sys.argv)
    window = MultiAppWindow()
    window.show()
    sys.exit(app.exec())