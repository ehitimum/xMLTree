#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import xml.etree.ElementTree as ET

class XMLEditor:
    def __init__(self, root):
        self.root = root
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

        # Search frame
        search_frame = tk.Frame(self.root)
        search_frame.pack(fill='x', pady=5)

        tk.Label(search_frame, text="Search by Attribute:").pack(side='left')
        self.search_entry = tk.Entry(search_frame, width=30)
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

    def populate_tree(self, parent_item, element):
        attrib_str = ' '.join([f'{k}="{v}"' for k, v in element.attrib.items()])
        display_text = element.tag
        if attrib_str:
            display_text += f' {attrib_str}'
        text_value = element.text.strip() if element.text else ''
        item = self.treeview.insert(parent_item, 'end', text=display_text, values=(text_value,))
        self.item_to_element[item] = element
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

    def add_child(self):
        selected = self.treeview.selection()
        if selected:
            parent_item = selected[0]
            parent_element = self.item_to_element[parent_item]
            new_element = ET.SubElement(parent_element, 'new_tag')
            new_element.text = 'new_text'
            # Insert into treeview
            attrib_str = ' '.join([f'{k}="{v}"' for k, v in new_element.attrib.items()])
            display_text = new_element.tag
            if attrib_str:
                display_text += f' {attrib_str}'
            text_value = new_element.text.strip() if new_element.text else ''
            new_item = self.treeview.insert(parent_item, 'end', text=display_text, values=(text_value,))
            self.item_to_element[new_item] = new_element
            self.treeview.see(new_item)

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
        self._find_matches('', query.lower())  # Start from root with lowercase query
        
        if self.search_results:
            self.prev_button.config(state='normal')
            self.next_button.config(state='normal')
            self.next_match()  # Auto-select first match
            tk.messagebox.showinfo("Info", f"Found {len(self.search_results)} matches.")
        else:
            self.prev_button.config(state='disabled')
            self.next_button.config(state='disabled')
            tk.messagebox.showinfo("Info", "No matches found.")

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
# Add new release
if __name__ == "__main__":
    root = tk.Tk()
    app = XMLEditor(root)
    root.mainloop()