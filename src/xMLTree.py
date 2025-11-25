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

# Add new release
if __name__ == "__main__":
    root = tk.Tk()
    app = XMLEditor(root)
    root.mainloop()