#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, scrolledtext, messagebox
import xml.etree.ElementTree as ET
import sys
import os
from threading import Thread

# Try to import pandas and ExcelColumnExtractorApp
PANDAS_AVAILABLE = False
ExcelColumnExtractorApp = None
BULK_AVAILABLE = False
BulkUpdateApp = None
DIFF_AVAILABLE = False
XMLDifferenceApp = None

try:
    import pandas as pd
    # Import the Excel extractor class from the local module
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from excel_column_extractor import ExcelColumnExtractorApp
    PANDAS_AVAILABLE = True
    # Try to import bulk update functionality
    try:
        from bulk_update import XMLConfigurator, ConfigurationManager
        BULK_AVAILABLE = True
    except Exception as e:
        print(f"Warning: bulk_update not available - Bulk Update feature disabled. ({e})")
    # Try to import XML difference functionality
    try:
        from xml_difference import XMLDifferenceGUI
        DIFF_AVAILABLE = True
    except Exception as e:
        print(f"Warning: xml_difference not available - XML Difference feature disabled. ({e})")
except ImportError as e:
    print(f"Warning: pandas not available - Excel column extractor will be disabled. ({e})")
except ValueError as e:
    if "numpy.dtype size changed" in str(e):
        print(f"Warning: numpy binary incompatibility detected - Excel column extractor will be disabled.")
        print(f"  Error: {e}")
        print(f"  Solution: Create a virtual environment and install compatible numpy/packages:")
        print(f"    python3 -m venv .venv")
        print(f"    source .venv/bin/activate")
        print(f"    pip install pandas openpyxl")
    else:
        print(f"Warning: Failed to import ExcelColumnExtractorApp: {e}")
except Exception as e:
    print(f"Warning: Failed to import ExcelColumnExtractorApp: {e}")

class XMLEditor:
    def __init__(self, root):
        self.root = root
        # Only set title if root is a Tk or Toplevel window
        if hasattr(self.root, 'title') and callable(self.root.title):
            self.root.title("XML Tree Editor")
        
        # Treeview setup with columns for tag/attrib and text
        self.treeview = ttk.Treeview(self.root, columns=('text',), show='tree headings')
        self.treeview.heading('#0', text='Tag / Attributes')
        self.treeview.heading('text', text='Text')
        self.treeview.pack(expand=True, fill='both')

        self.treeview.tag_configure('search_match', background='#fff2cc')  # Light yellow highlight
        
        # Buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill='x')
        
        load_button = tk.Button(button_frame, text="Load XML", command=self.load_xml)
        load_button.pack(side='left')
        
        save_button = tk.Button(button_frame, text="Save XML", command=self.save_xml)
        save_button.pack(side='left')
        
        add_child_button = tk.Button(button_frame, text="Add Child", command=self.add_child)
        add_child_button.pack(side='left')

        duplicate_button = tk.Button(button_frame, text="Duplicate", command=self.duplicate_item)  # NEW
        duplicate_button.pack(side='left')
        
        delete_button = tk.Button(button_frame, text="Delete Selected", command=self.delete_item)
        delete_button.pack(side='left')
        
        # Bind double-click for editing
        self.treeview.bind('<Double-1>', self.edit_item)
        
        # Storage
        self.item_to_element = {}
        self.etree = None
        self.root_element = None
        self.file_path = None

        # Search
        self.search_results = []  # List of matching treeview item IDs
        self.current_search_index = -1  # Current position in search results 

        # Search frame - UPDATED
        search_frame = tk.Frame(self.root)
        search_frame.pack(fill='x', pady=5)

        # Add search mode selection
        self.search_mode = tk.StringVar(value="content")  # "content" or "path"
        
        mode_frame = tk.Frame(search_frame)
        mode_frame.pack(fill='x')
        
        tk.Label(mode_frame, text="Search Mode:").pack(side='left')
        tk.Radiobutton(mode_frame, text="Content", variable=self.search_mode, value="content").pack(side='left')
        tk.Radiobutton(mode_frame, text="Path", variable=self.search_mode, value="path").pack(side='left')
        
        tk.Label(search_frame, text="Search:").pack(side='left')
        self.search_entry = tk.Entry(search_frame, width=40)
        self.search_entry.pack(side='left', padx=5)

        search_button = tk.Button(search_frame, text="Search", command=self.perform_search)
        search_button.pack(side='left')

        self.prev_button = tk.Button(search_frame, text="Previous", command=self.prev_match, state='disabled')
        self.prev_button.pack(side='left', padx=5)

        self.next_button = tk.Button(search_frame, text="Next", command=self.next_match, state='disabled')
        self.next_button.pack(side='left')

    def load_xml(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
        if self.file_path:
            self.etree = ET.parse(self.file_path)
            self.root_element = self.etree.getroot()
            self.treeview.delete(*self.treeview.get_children())
            self.item_to_element = {}
            self.populate_tree('', self.root_element)
        # self.search_results = []
        # self.current_search_index = -1
        # self.prev_button.config(state='disabled')
        # self.next_button.config(state='disabled')
        # self.search_entry.delete(0, tk.END)

    # def populate_tree(self, parent_item, element):
    #     attrib_str = ' '.join([f'{k}="{v}"' for k, v in element.attrib.items()])
    #     display_text = element.tag
    #     if attrib_str:
    #         display_text += f' {attrib_str}'
    #     text_value = element.text.strip() if element.text else ''
    #     item = self.treeview.insert(parent_item, 'end', text=display_text, values=(text_value,))
    #     self.item_to_element[item] = element
    #     for child in element:
    #         self.populate_tree(item, child)
    #     self.search_results = []
    #     self.current_search_index = -1
    #     self.prev_button.config(state='disabled')
    #     self.next_button.config(state='disabled')
    #     if hasattr(self, 'search_entry'):
    #         self.search_entry.delete(0, tk.END)

    def populate_tree(self, parent_item, element):
        item = self._add_element_to_treeview(parent_item, element)  # Use helper method
        for child in element:
            self.populate_tree(item, child)
        self.search_results = []
        self.current_search_index = -1
        self.prev_button.config(state='disabled')
        self.next_button.config(state='disabled')
        if hasattr(self, 'search_entry'):
            self.search_entry.delete(0, tk.END)

    def edit_item(self, event):
        item = self.treeview.identify_row(event.y)
        if item:
            element = self.item_to_element[item]
            # Open edit dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Edit Element")
            
            tk.Label(dialog, text="Tag:").pack()
            tag_entry = tk.Entry(dialog)
            tag_entry.insert(0, element.tag)
            tag_entry.pack()
            
            tk.Label(dialog, text="Text:").pack()
            text_entry = tk.Entry(dialog)
            text_entry.insert(0, element.text or '')
            text_entry.pack()
            
            tk.Label(dialog, text="Attributes (key=value space-separated):").pack()
            attrib_entry = tk.Entry(dialog, width=50)
            attrib_entry.insert(0, ' '.join([f'{k}={v}' for k, v in element.attrib.items()]))
            attrib_entry.pack()
            
            def save_edit():
                element.tag = tag_entry.get()
                element.text = text_entry.get()
                new_attrib = {}
                for pair in attrib_entry.get().split():
                    if '=' in pair:
                        k, v = pair.split('=', 1)
                        new_attrib[k] = v
                element.attrib = new_attrib
                # Update treeview item
                attrib_str = ' '.join([f'{k}="{v}"' for k, v in element.attrib.items()])
                display_text = element.tag
                if attrib_str:
                    display_text += f' {attrib_str}'
                text_value = element.text.strip() if element.text else ''
                self.treeview.item(item, text=display_text, values=(text_value,))
                dialog.destroy()
            
            tk.Button(dialog, text="Save", command=save_edit).pack()

    # def add_child(self):
    #     selected = self.treeview.selection()
    #     if selected:
    #         parent_item = selected[0]
    #         parent_element = self.item_to_element[parent_item]
    #         new_element = ET.SubElement(parent_element, 'new_tag')
    #         new_element.text = 'new_text'
    #         # Insert into treeview
    #         attrib_str = ' '.join([f'{k}="{v}"' for k, v in new_element.attrib.items()])
    #         display_text = new_element.tag
    #         if attrib_str:
    #             display_text += f' {attrib_str}'
    #         text_value = new_element.text.strip() if new_element.text else ''
    #         new_item = self.treeview.insert(parent_item, 'end', text=display_text, values=(text_value,))
    #         self.item_to_element[new_item] = new_element
    #         self.treeview.see(new_item)

    def add_child(self):
        selected = self.treeview.selection()
        if selected:
            parent_item = selected[0]
            parent_element = self.item_to_element[parent_item]
            new_element = ET.SubElement(parent_element, 'new_tag')
            new_element.text = 'new_text'
            
            # Insert into treeview using the helper method
            self._add_element_to_treeview(parent_item, new_element)
            self.treeview.see(self.treeview.get_children(parent_item)[-1])  # Scroll to last child
    
    def _add_element_to_treeview(self, parent_item, element):
        """Helper method to add an element to treeview"""
        attrib_str = ' '.join([f'{k}="{v}"' for k, v in element.attrib.items()])
        display_text = element.tag
        if attrib_str:
            display_text += f' {attrib_str}'
        text_value = element.text.strip() if element.text else ''
        new_item = self.treeview.insert(parent_item, 'end', text=display_text, values=(text_value,))
        self.item_to_element[new_item] = element
        return new_item

    def delete_item(self):
        selected = self.treeview.selection()
        if selected:
            item = selected[0]
            element = self.item_to_element[item]
            parent_item = self.treeview.parent(item)
            if parent_item:
                parent_element = self.item_to_element[parent_item]
                parent_element.remove(element)
                self.treeview.delete(item)
                del self.item_to_element[item]
            else:
                tk.messagebox.showwarning("Warning", "Cannot delete root element.")

    def save_xml(self):
        if self.file_path and self.etree:
            save_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml")], initialfile=self.file_path)
            if save_path:
                self.etree.write(save_path, encoding='utf-8', xml_declaration=True)
                self.file_path = save_path  # Update file path if saved to new location
    
    def perform_path_search(self, path_query):
        path_parts = path_query.split('.')
        if not path_parts:
            return
        
        # Clear previous search results and highlighting
        for item in self.treeview.get_children():
            self.treeview.item(item, tags=())
        
        self.search_results = []
        self.current_search_index = -1
        
        # Start searching from root
        self._find_by_path('', path_parts, 0)
        
        if self.search_results:
            self.prev_button.config(state='normal')
            self.next_button.config(state='normal')
            self.next_match()  # Auto-select first match
            tk.messagebox.showinfo("Info", f"Found {len(self.search_results)} matches for path.")
        else:
            self.prev_button.config(state='disabled')
            self.next_button.config(state='disabled')
            tk.messagebox.showinfo("Info", "No matches found for the given path.")

    def _find_by_path(self, parent_item, path_parts, current_index):
        """Recursively find elements matching the path"""
        if current_index >= len(path_parts):
            return
            
        current_part = path_parts[current_index]
        
        for item in self.treeview.get_children(parent_item):
            element = self.item_to_element[item]
            
            # Check if current element matches the current path part
            # Handle numeric indices (like "1" meaning <i1> element)
            if current_part.isdigit():
                # Look for elements like <i1>, <i2>, etc.
                expected_tag = f"i{current_part}"
                if element.tag == expected_tag:
                    if current_index == len(path_parts) - 1:
                        # This is the final element in the path
                        self.search_results.append(item)
                    else:
                        # Continue searching in children
                        self._find_by_path(item, path_parts, current_index + 1)
            else:
                # Regular tag name matching
                tag_name = element.tag.split('}')[-1] if '}' in element.tag else element.tag
                if current_part.lower() == tag_name.lower():
                    if current_index == len(path_parts) - 1:
                        # This is the final element in the path
                        self.search_results.append(item)
                    else:
                        # Continue searching in children
                        self._find_by_path(item, path_parts, current_index + 1)
            
            # Also search in children regardless of match (for cases where path continues through non-matching parents)
            self._find_by_path(item, path_parts, current_index)

    def perform_search(self):
        query = self.search_entry.get().strip()
        if not query:
            tk.messagebox.showinfo("Info", "Enter search term.")
            return
        
        # Clear previous search results and highlighting
        for item in self.treeview.get_children():
            self.treeview.item(item, tags=())
        
        self.search_results = []
        self.current_search_index = -1
        
        # Choose search mode based on selection
        if self.search_mode.get() == "path":
            self.perform_path_search(query)
        else:
            self._find_matches('', query.lower())  # Existing content search
            
            if self.search_results:
                self.prev_button.config(state='normal')
                self.next_button.config(state='normal')
                self.next_match()  # Auto-select first match
                tk.messagebox.showinfo("Info", f"Found {len(self.search_results)} matches.")
            else:
                self.prev_button.config(state='disabled')
                self.next_button.config(state='disabled')
                tk.messagebox.showinfo("Info", "No matches found.")
    def get_element_path(self, element):
        """Get the dot-separated path for an element (for debugging)"""
        path_parts = []
        current = element
        while current is not None and current != self.root_element:
            # Handle indexed elements (i1, i2, etc.)
            if current.tag.startswith('i') and current.tag[1:].isdigit():
                path_parts.append(current.tag[1:])  # Just the number
            else:
                path_parts.append(current.tag)
            current = current.getparent()
        
        path_parts.reverse()
        return '.'.join(path_parts)


    # def perform_search(self):
    #     query = self.search_entry.get().strip()
    #     if not query:
    #         tk.messagebox.showinfo("Info", "Enter search term.")
    #         return
        
    #     # Clear previous search results and highlighting
    #     for item in self.treeview.get_children():
    #         self.treeview.item(item, tags=())
        
    #     self.search_results = []
    #     self.current_search_index = -1
    #     self._find_matches('', query.lower())  # Start from root with lowercase query
        
    #     if self.search_results:
    #         self.prev_button.config(state='normal')
    #         self.next_button.config(state='normal')
    #         self.next_match()  # Auto-select first match
    #         tk.messagebox.showinfo("Info", f"Found {len(self.search_results)} matches.")
    #     else:
    #         self.prev_button.config(state='disabled')
    #         self.next_button.config(state='disabled')
    #         tk.messagebox.showinfo("Info", "No matches found.")

    def _find_matches(self, parent_item, query_lower):
        for item in self.treeview.get_children(parent_item):
            element = self.item_to_element[item]
            found_match = False
            
            # 1. Check ATTRIBUTE NAMES (case insensitive)
            for key in element.attrib:
                # Remove namespace if present
                local_name = key.split('}')[-1] if '}' in key else key
                if query_lower in local_name.lower():
                    self.search_results.append(item)
                    found_match = True
                    break
            
            # 2. Check ATTRIBUTE VALUES (case insensitive)
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
            
            # 4. Check TEXT CONTENT (case insensitive)
            if not found_match and element.text:
                if query_lower in element.text.lower():
                    self.search_results.append(item)
                    found_match = True

            # Always recurse into children (even if we found a match in parent)
            self._find_matches(item, query_lower)
        
    def prev_match(self):
        if self.search_results:
            # Clear previous highlight from all search results
            for item in self.search_results:
                self.treeview.item(item, tags=())
            
            self.current_search_index = (self.current_search_index - 1) % len(self.search_results)
            self._select_match()

    def next_match(self):
        if self.search_results:
            # Clear previous highlight from all search results
            for item in self.search_results:
                self.treeview.item(item, tags=())
            
            self.current_search_index = (self.current_search_index + 1) % len(self.search_results)
            self._select_match()

    def _select_match(self):
        if not self.search_results:
            return
            
        item = self.search_results[self.current_search_index]
        self.treeview.selection_set(item)
        self.treeview.focus(item)
        self.treeview.item(item, tags=('search_match',))  # Highlight in yellow

        # Expand parents and scroll to make it visible
        parent = self.treeview.parent(item)
        while parent:
            self.treeview.item(parent, open=True)
            parent = self.treeview.parent(parent)
        self.treeview.see(item)

    def duplicate_item(self):
        selected = self.treeview.selection()
        if not selected:
            tk.messagebox.showinfo("Info", "Please select an element to duplicate.")
            return
            
        item = selected[0]
        element = self.item_to_element[item]
        parent_item = self.treeview.parent(item)
        
        if not parent_item:
            tk.messagebox.showwarning("Warning", "Cannot duplicate root element.")
            return
        
        parent_element = self.item_to_element[parent_item]
        
        # Create a deep copy of the element
        import copy
        new_element = copy.deepcopy(element)
        
        # Generate the next available index for elements like i1, i2, etc.
        if element.tag.startswith('i') and element.tag[1:].isdigit():
            new_index = self._get_next_available_index(parent_element, element.tag)
            new_element.tag = f"i{new_index}"
            
            # Also update any Alias elements if they exist to reflect the new index
            for alias_elem in new_element.findall('.//Alias'):
                if alias_elem.text and element.tag in alias_elem.text:
                    # Update alias to reflect new index (e.g., cpe-GERANFreqGroup1 -> cpe-GERANFreqGroup2)
                    old_alias = alias_elem.text
                    base_name = old_alias.rsplit(element.tag[1:], 1)[0]  # Remove the old index
                    alias_elem.text = f"{base_name}{new_index}"
        
        # Add the new element to the parent
        parent_element.append(new_element)
        
        # Add to treeview
        attrib_str = ' '.join([f'{k}="{v}"' for k, v in new_element.attrib.items()])
        display_text = new_element.tag
        if attrib_str:
            display_text += f' {attrib_str}'
        text_value = new_element.text.strip() if new_element.text else ''
        new_item = self.treeview.insert(parent_item, 'end', text=display_text, values=(text_value,))
        self.item_to_element[new_item] = new_element
        
        # Recursively populate children of the duplicated element
        for child in new_element:
            self._populate_duplicate_children(new_item, child)
        
        # Expand and scroll to the new item
        self.treeview.see(new_item)
        self.treeview.selection_set(new_item)
        
        tk.messagebox.showinfo("Info", f"Successfully duplicated {element.tag} as {new_element.tag}")
    def _get_next_available_index(self, parent_element, tag_pattern):
        """Find the next available index for elements like i1, i2, etc."""
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
        """Recursively populate children for duplicated elements"""
        attrib_str = ' '.join([f'{k}="{v}"' for k, v in element.attrib.items()])
        display_text = element.tag
        if attrib_str:
            display_text += f' {attrib_str}'
        text_value = element.text.strip() if element.text else ''
        item = self.treeview.insert(parent_item, 'end', text=display_text, values=(text_value,))
        self.item_to_element[item] = element
        
        for child in element:
            self._populate_duplicate_children(item, child)

class MultiApp:
    """Main application that switches between XML Editor and Excel Column Extractor."""
    def __init__(self, root):
        self.root = root
        self.root.title("xMLTree - XML Editor & Excel Extractor")
        self.root.geometry("1200x800")
        self.current_frame = None
        self.current_app = None
        
        # Create menu bar
        menubar = tk.Menu(root)
        root.config(menu=menubar)
        
        app_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Application", menu=app_menu)
        app_menu.add_command(label="XML Editor", command=self.show_xml_editor)
        if PANDAS_AVAILABLE and ExcelColumnExtractorApp is not None:
            app_menu.add_command(label="Excel Column Extractor", command=self.show_excel_extractor)
        else:
            app_menu.add_command(label="Excel Column Extractor (requires pandas)", state="disabled")
        if BULK_AVAILABLE:
            app_menu.add_command(label="Bulk Update", command=self.show_bulk_update)
        else:
            app_menu.add_command(label="Bulk Update (requires pandas)", state="disabled")
        if DIFF_AVAILABLE:
            app_menu.add_command(label="XML Difference", command=self.show_xml_difference)
        else:
            app_menu.add_command(label="XML Difference (requires pandas)", state="disabled")
        
        # Container frame for apps
        self.container = tk.Frame(root)
        self.container.pack(fill="both", expand=True)
        
        # Show XML Editor by default
        self.show_xml_editor()
    
    def clear_current(self):
        """Destroy current frame and app."""
        if self.current_frame:
            self.current_frame.destroy()
            self.current_frame = None
        self.current_app = None
    
    def show_xml_editor(self):
        """Switch to XML Editor."""
        self.clear_current()
        self.current_frame = tk.Frame(self.container)
        self.current_frame.pack(fill="both", expand=True)
        self.current_app = XMLEditor(self.current_frame)
    
    def show_excel_extractor(self):
        """Switch to Excel Column Extractor."""
        if not PANDAS_AVAILABLE or ExcelColumnExtractorApp is None:
            tk.messagebox.showwarning("Feature unavailable", "pandas is not installed. Please install pandas to use Excel Column Extractor.")
            return
        self.clear_current()
        self.current_frame = tk.Frame(self.container)
        self.current_frame.pack(fill="both", expand=True)
        self.current_app = ExcelColumnExtractorApp(self.current_frame)

    def show_bulk_update(self):
        """Switch to Bulk Update GUI."""
        if not BULK_AVAILABLE:
            tk.messagebox.showwarning("Feature unavailable", "bulk_update is not available. Please ensure pandas is installed and bulk_update.py is present.")
            return
        self.clear_current()
        self.current_frame = tk.Frame(self.container)
        self.current_frame.pack(fill="both", expand=True)
        self.current_app = BulkUpdateApp(self.current_frame)

    def show_xml_difference(self):
        """Switch to XML Difference GUI."""
        if not DIFF_AVAILABLE:
            tk.messagebox.showwarning("Feature unavailable", "xml_difference is not available. Please ensure pandas is installed and xml_difference.py is present.")
            return
        self.clear_current()
        self.current_frame = tk.Frame(self.container)
        self.current_frame.pack(fill="both", expand=True)
        self.current_app = XMLDifferenceGUI(self.current_frame)


class BulkUpdateApp:
    """GUI interface for XML Configurator with preview functionality"""
    
    def __init__(self, root):
        self.root = root
        if hasattr(self.root, 'title') and callable(self.root.title):
            self.root.title("Bulk XML Update")
        if hasattr(self.root, 'geometry') and callable(self.root.geometry):
            self.root.geometry("1200x700")
        
        # Configure styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Initialize configurator
        self.configurator = XMLConfigurator()
        
        # Data storage
        self.mappings_df = None
        self.checked_mappings = []
        self.xml_loaded = False
        self.mappings_loaded = False
        
        # Create GUI
        self.create_widgets()
        self.setup_layout()
        
    def create_widgets(self):
        """Create all GUI widgets"""
        # Top frame for file selection
        self.top_frame = ttk.Frame(self.root, padding="10")
        
        # XML File selection
        self.xml_label = ttk.Label(self.top_frame, text="XML File:")
        self.xml_path_var = tk.StringVar()
        self.xml_entry = ttk.Entry(self.top_frame, textvariable=self.xml_path_var, width=50)
        self.xml_browse_btn = ttk.Button(self.top_frame, text="Browse...", 
                                         command=self.browse_xml_file)
        self.xml_load_btn = ttk.Button(self.top_frame, text="Load XML", 
                                       command=self.load_xml_file)
        
        # Mapping File selection
        self.mapping_label = ttk.Label(self.top_frame, text="Mapping File (CSV/Excel):")
        self.mapping_path_var = tk.StringVar()
        self.mapping_entry = ttk.Entry(self.top_frame, textvariable=self.mapping_path_var, width=50)
        self.mapping_browse_btn = ttk.Button(self.top_frame, text="Browse...", 
                                            command=self.browse_mapping_file)
        self.mapping_load_btn = ttk.Button(self.top_frame, text="Load Mappings", 
                                          command=self.load_mapping_file)
        
        # Column selection
        self.column_frame = ttk.Frame(self.top_frame)
        self.path_col_label = ttk.Label(self.column_frame, text="Path Column:")
        self.path_col_var = tk.StringVar(value="path")
        self.path_col_entry = ttk.Entry(self.column_frame, textvariable=self.path_col_var, width=15)
        
        self.value_col_label = ttk.Label(self.column_frame, text="Value Column:")
        self.value_col_var = tk.StringVar(value="value")
        self.value_col_entry = ttk.Entry(self.column_frame, textvariable=self.value_col_var, width=15)
        
        self.create_template_btn = ttk.Button(self.column_frame, text="Create Template CSV", 
                                              command=self.create_template)
        
        # Middle frame for preview
        self.mid_frame = ttk.Frame(self.root, padding="10")
        
        # Treeview for preview
        self.tree_frame = ttk.Frame(self.mid_frame)
        self.tree_scroll = ttk.Scrollbar(self.tree_frame)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.preview_tree = ttk.Treeview(
            self.tree_frame,
            columns=('Status', 'Path', 'Current Value', 'New Value'),
            show='headings',
            yscrollcommand=self.tree_scroll.set,
            height=20
        )
        
        # Configure tree columns
        self.preview_tree.heading('Status', text='Status')
        self.preview_tree.heading('Path', text='Path')
        self.preview_tree.heading('Current Value', text='Current Value')
        self.preview_tree.heading('New Value', text='New Value')
        
        self.preview_tree.column('Status', width=100, anchor='center')
        self.preview_tree.column('Path', width=400)
        self.preview_tree.column('Current Value', width=200)
        self.preview_tree.column('New Value', width=200)
        
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_scroll.config(command=self.preview_tree.yview)
        
        # Bottom frame for buttons and status
        self.bottom_frame = ttk.Frame(self.root, padding="10")
        
        # Action buttons
        self.check_btn = ttk.Button(self.bottom_frame, text="Check Mappings", 
                                   command=self.check_mappings, state='disabled')
        self.update_btn = ttk.Button(self.bottom_frame, text="Update XML", 
                                     command=self.update_xml, state='disabled')
        self.save_as_btn = ttk.Button(self.bottom_frame, text="Save As...", 
                                      command=self.save_as_xml, state='disabled')
        
        # Status display
        self.status_text = scrolledtext.ScrolledText(
            self.bottom_frame, 
            height=8,
            width=100,
            state='disabled'
        )
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.bottom_frame, 
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        
    def setup_layout(self):
        """Arrange widgets in the window"""
        # Top frame layout
        self.top_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=5)
        
        # Row 1: XML File
        self.xml_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        self.xml_entry.grid(row=0, column=1, padx=(5, 0), pady=2, sticky=(tk.W, tk.E))
        self.xml_browse_btn.grid(row=0, column=2, padx=(5, 0), pady=2)
        self.xml_load_btn.grid(row=0, column=3, padx=(5, 5), pady=2)
        
        # Row 2: Mapping File
        self.mapping_label.grid(row=1, column=0, sticky=tk.W, pady=2)
        self.mapping_entry.grid(row=1, column=1, padx=(5, 0), pady=2, sticky=(tk.W, tk.E))
        self.mapping_browse_btn.grid(row=1, column=2, padx=(5, 0), pady=2)
        self.mapping_load_btn.grid(row=1, column=3, padx=(5, 5), pady=2)
        
        # Row 3: Column selection
        self.column_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
        self.path_col_label.grid(row=0, column=0, sticky=tk.W)
        self.path_col_entry.grid(row=0, column=1, padx=(5, 20))
        self.value_col_label.grid(row=0, column=2, sticky=tk.W)
        self.value_col_entry.grid(row=0, column=3, padx=(5, 0))
        self.create_template_btn.grid(row=0, column=4, padx=(10, 0))
        
        # Middle frame layout
        self.mid_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=5)
        self.mid_frame.grid_rowconfigure(0, weight=1)
        self.mid_frame.grid_columnconfigure(0, weight=1)
        
        self.tree_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)
        
        # Bottom frame layout
        self.bottom_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=10, pady=5)
        
        # Button row
        self.check_btn.grid(row=0, column=0, padx=(0, 5), pady=5)
        self.update_btn.grid(row=0, column=1, padx=5, pady=5)
        self.save_as_btn.grid(row=0, column=2, padx=(5, 0), pady=5)
        
        # Status and progress
        self.status_text.grid(row=1, column=0, columnspan=3, pady=(5, 0), sticky=(tk.W, tk.E))
        self.progress_bar.grid(row=2, column=0, columnspan=3, pady=(5, 0), sticky=(tk.W, tk.E))
        
        # Configure grid weights
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
    def browse_xml_file(self):
        """Open file dialog for XML file"""
        filename = filedialog.askopenfilename(
            title="Select XML File",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if filename:
            self.xml_path_var.set(filename)
            self.xml_load_btn.config(state='normal')
    
    def load_xml_file(self):
        """Load the selected XML file"""
        xml_path = self.xml_path_var.get()
        if not xml_path:
            messagebox.showerror("Error", "Please select an XML file")
            return
        
        try:
            success = self.configurator.load_xml(xml_path)
            if success:
                self.xml_loaded = True
                self.log_message(f"✓ XML loaded: {xml_path}")
                self.update_button_states()
            else:
                self.log_message(f"✗ Failed to load XML: {xml_path}", error=True)
        except Exception as e:
            self.log_message(f"Error loading XML: {e}", error=True)
    
    def browse_mapping_file(self):
        """Open file dialog for mapping file"""
        filename = filedialog.askopenfilename(
            title="Select Mapping File",
            filetypes=[
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.mapping_path_var.set(filename)
            self.mapping_load_btn.config(state='normal')
    
    def load_mapping_file(self):
        """Load the mapping file"""
        mapping_path = self.mapping_path_var.get()
        if not mapping_path:
            messagebox.showerror("Error", "Please select a mapping file")
            return
        
        try:
            self.mappings_df = self.configurator.load_mapping_file(mapping_path)
            if self.mappings_df is not None:
                self.mappings_loaded = True
                self.log_message(f"✓ Mappings loaded: {mapping_path}")
                self.update_button_states()
                
                # Show column info
                columns = list(self.mappings_df.columns)
                self.log_message(f"  Available columns: {', '.join(columns)}")
            else:
                self.log_message(f"✗ Failed to load mappings: {mapping_path}", error=True)
        except Exception as e:
            self.log_message(f"Error loading mappings: {e}", error=True)
    
    def check_mappings(self):
        """Check which mappings can be applied"""
        if not self.xml_loaded:
            messagebox.showerror("Error", "Please load an XML file first")
            return
        
        if not self.mappings_loaded:
            messagebox.showerror("Error", "Please load a mapping file first")
            return
        
        # Clear previous results
        self.preview_tree.delete(*self.preview_tree.get_children())
        
        # Get column names
        path_col = self.path_col_var.get()
        value_col = self.value_col_var.get()
        
        # Validate columns exist
        if path_col not in self.mappings_df.columns:
            messagebox.showerror("Error", f"Column '{path_col}' not found in mapping file")
            return
        if value_col not in self.mappings_df.columns:
            messagebox.showerror("Error", f"Column '{value_col}' not found in mapping file")
            return
        
        # Disable buttons during processing
        self.check_btn.config(state='disabled')
        self.update_btn.config(state='disabled')
        
        # Start checking in background thread
        self.progress_var.set(0)
        self.log_message("Checking mappings...")
        
        thread = Thread(target=self._check_mappings_thread, 
                       args=(path_col, value_col))
        thread.daemon = True
        thread.start()
    
    def _check_mappings_thread(self, path_col, value_col):
        """Background thread for checking mappings"""
        try:
            # Convert DataFrame to list of dicts
            mapping_data = self.mappings_df.to_dict('records')
            
            # Check mappings
            self.checked_mappings = self.configurator.check_mappings(
                mapping_data, path_col, value_col
            )
            
            # Update UI in main thread
            self.root.after(0, self._update_preview_results)
            
        except Exception as e:
            self.root.after(0, lambda: self.log_message(f"Error checking mappings: {e}", error=True))
            self.root.after(0, lambda: self.check_btn.config(state='normal'))
    
    def _update_preview_results(self):
        """Update UI with check results"""
        # Clear tree
        self.preview_tree.delete(*self.preview_tree.get_children())
        
        # Populate tree with results
        found_count = 0
        not_found_count = 0
        
        for mapping in self.checked_mappings:
            status = "✓ Found" if mapping['found'] else "✗ Not Found"
            tag = 'found' if mapping['found'] else 'notfound'
            
            self.preview_tree.insert(
                '', 'end',
                values=(
                    status,
                    mapping['path'],
                    mapping['current_value'],
                    mapping['new_value']
                ),
                tags=(tag,)
            )
            
            if mapping['found']:
                found_count += 1
            else:
                not_found_count += 1
        
        # Configure tags for color coding
        self.preview_tree.tag_configure('found', background='#e8f5e9')
        self.preview_tree.tag_configure('notfound', background='#ffebee')
        
        # Update status
        self.log_message(f"✓ Check complete: {found_count} found, {not_found_count} not found")
        
        # Update progress bar
        self.progress_var.set(100)
        
        # Enable buttons
        self.check_btn.config(state='normal')
        if found_count > 0:
            self.update_btn.config(state='normal')
        
        # Show summary
        if not_found_count > 0:
            self.log_message(f"⚠ {not_found_count} paths were not found in the XML", warning=True)
    
    def update_xml(self):
        """Apply the checked mappings and update XML"""
        if not self.checked_mappings:
            messagebox.showerror("Error", "No mappings to apply. Please check mappings first.")
            return
        
        # Filter only found mappings
        mappings_to_apply = [m for m in self.checked_mappings if m['found']]
        
        if not mappings_to_apply:
            messagebox.showwarning("Warning", "No found mappings to apply.")
            return
        
        # Ask for confirmation
        confirm = messagebox.askyesno(
            "Confirm Update",
            f"Update {len(mappings_to_apply)} values in the XML?\n\n"
            f"Only found mappings will be applied. Not found mappings will be ignored."
        )
        
        if not confirm:
            return
        
        # Disable buttons during update
        self.update_btn.config(state='disabled')
        self.check_btn.config(state='disabled')
        
        self.log_message("Applying updates...")
        self.progress_var.set(0)
        
        # Start update in background thread
        thread = Thread(target=self._update_xml_thread, args=(mappings_to_apply,))
        thread.daemon = True
        thread.start()
    
    def _update_xml_thread(self, mappings_to_apply):
        """Background thread for updating XML"""
        try:
            # Apply mappings
            result = self.configurator.apply_mappings(mappings_to_apply)
            
            # Update UI in main thread
            self.root.after(0, lambda: self._update_complete(result))
            
        except Exception as e:
            self.root.after(0, lambda: self.log_message(f"Error updating XML: {e}", error=True))
            self.root.after(0, lambda: self.update_btn.config(state='normal'))
            self.root.after(0, lambda: self.check_btn.config(state='normal'))
    
    def _update_complete(self, result):
        """Handle completion of XML update"""
        # Update progress bar
        self.progress_var.set(100)
        
        # Log results
        self.log_message(f"✓ Update complete!")
        self.log_message(f"  Total mappings: {result['total']}")
        self.log_message(f"  Successfully applied: {result['applied']}")
        self.log_message(f"  Failed to apply: {result['failed']}")
        self.log_message(f"  Not found (skipped): {result['not_found']}")
        
        # Enable save button
        self.save_as_btn.config(state='normal')
        
        # Re-enable check button
        self.check_btn.config(state='normal')
        
        # Ask if user wants to save
        messagebox.showinfo("Update Complete", 
                           f"XML updated successfully!\n\n"
                           f"Applied: {result['applied']} values\n"
                           f"Failed: {result['failed']}\n"
                           f"Skipped (not found): {result['not_found']}\n\n"
                           f"Click 'Save As...' to save the updated XML.")
    
    def save_as_xml(self):
        """Save the updated XML to a new file"""
        if not self.configurator.xml_path:
            messagebox.showerror("Error", "No XML file loaded")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Save XML As",
            defaultextension=".xml",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
            initialfile=f"{self.configurator.xml_path.stem}_updated.xml"
        )
        
        if filename:
            try:
                success = self.configurator.save_xml(filename)
                if success:
                    self.log_message(f"✓ XML saved to: {filename}")
                else:
                    self.log_message(f"✗ Failed to save XML", error=True)
            except Exception as e:
                self.log_message(f"Error saving XML: {e}", error=True)
    
    def create_template(self):
        """Create a template CSV file with paths from the loaded XML"""
        if not self.configurator.xml_path:
            messagebox.showwarning('No XML', 'Please load an XML file first to extract paths for template.')
            return
        out = filedialog.asksaveasfilename(title='Save template CSV', defaultextension='.csv', filetypes=[('CSV','*.csv')])
        if not out:
            return
        try:
            ok = ConfigurationManager.create_template_csv(out, str(self.configurator.xml_path))
            if ok:
                self.log_message(f"✓ Template CSV created: {out}")
                messagebox.showinfo('Template created', f'Template saved to: {out}')
            else:
                self.log_message(f"✗ Failed to create template", error=True)
                messagebox.showerror('Error', 'Failed to create template')
        except Exception as e:
            self.log_message(f"Error creating template: {e}", error=True)
            messagebox.showerror('Error', str(e))
    
    def update_button_states(self):
        """Update button states based on loaded files"""
        if self.xml_loaded and self.mappings_loaded:
            self.check_btn.config(state='normal')
        else:
            self.check_btn.config(state='disabled')
    
    def log_message(self, message: str, error: bool = False, warning: bool = False):
        """Add message to status text area"""
        self.status_text.config(state='normal')
        
        # Add timestamp and message
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if error:
            prefix = "[ERROR]"
            color = "red"
        elif warning:
            prefix = "[WARN]"
            color = "orange"
        else:
            prefix = "[INFO]"
            color = "green"
        
        # Insert message
        self.status_text.insert(tk.END, f"{timestamp} {prefix} {message}\n", color)
        
        # Scroll to bottom
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')
        
        # Also log to console
        import logging
        logger = logging.getLogger(__name__)
        if error:
            logger.error(message)
        elif warning:
            logger.warning(message)
        else:
            logger.info(message)

# Add new release
if __name__ == "__main__":
    import sys
    print("Arguments:", sys.argv)
    if "--test" in sys.argv:
        # Test imports
        try:
            import pandas as pd
            import openpyxl
            print("All imports successful")
            sys.exit(0)
        except ImportError as e:
            print(f"Import error: {e}")
            sys.exit(1)
    root = tk.Tk()
    app = MultiApp(root)
    root.mainloop()