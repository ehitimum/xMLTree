#!/usr/bin/env python3
"""
XML Difference Tool
Compare two XML files and generate Excel report with differences:
- Missing: Elements present in first XML but not in second
- Changed: Elements with different values/text
- Added: Elements present in second XML but not in first
"""

import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Set, Any
import pandas as pd
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from threading import Thread
import queue

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class XMLComparator:
    """Core class for comparing two XML files"""
    
    def __init__(self):
        self.tree1 = None
        self.tree2 = None
        self.root1 = None
        self.root2 = None
        self.namespace1 = ''
        self.namespace2 = ''
        
    def load_xml(self, xml_path: str, which: int = 1) -> bool:
        """Load and parse XML file (1 for first, 2 for second)"""
        try:
            path = Path(xml_path)
            if not path.exists():
                logger.error(f"XML file not found: {xml_path}")
                return False
            
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Extract namespace if present
            namespace = ''
            if '}' in root.tag:
                namespace = root.tag.split('}')[0] + '}'
            
            if which == 1:
                self.tree1 = tree
                self.root1 = root
                self.namespace1 = namespace
            else:
                self.tree2 = tree
                self.root2 = root
                self.namespace2 = namespace
            
            logger.info(f"Successfully loaded XML {which}: {xml_path}")
            return True
        except Exception as e:
            logger.error(f"Failed to load XML {which}: {e}")
            return False
    
    def _get_element_path(self, element: ET.Element, namespace: str = '') -> str:
        """Generate dot-notation path for an element"""
        path_parts = []
        current = element
        
        while current is not None:
            # Get tag without namespace
            tag = current.tag
            if namespace and tag.startswith(namespace):
                tag = tag.replace(namespace, '')
            
            # Handle numeric indices (i1, i2, etc.)
            if tag.startswith('i') and tag[1:].isdigit():
                tag = tag[1:]  # Convert i1 to 1
            
            path_parts.insert(0, tag)
            
            # Move to parent
            current = current.getparent()
        
        # Skip the root element (Device)
        if path_parts and path_parts[0] == 'Device':
            path_parts = path_parts[1:]
        
        return '.'.join(path_parts)
    
    def _get_element_value(self, element: ET.Element) -> str:
        """Get value from element (text or attribute)"""
        if element.text is not None and element.text.strip():
            return element.text.strip()
        
        # Check common value attributes
        value_attrs = ['value', 'Value', 'VALUE', 'content', 'Content']
        for attr in value_attrs:
            if attr in element.attrib:
                return element.attrib[attr]
        
        return ''
    
    def _collect_elements(self, root: ET.Element, namespace: str = '') -> Dict[str, str]:
        """Collect all leaf elements with their paths and values"""
        elements = {}
        
        def traverse(element, current_path=''):
            # Get current tag without namespace
            tag = element.tag
            if namespace and tag.startswith(namespace):
                tag = tag.replace(namespace, '')
            
            # Handle numeric indices
            if tag.startswith('i') and tag[1:].isdigit():
                tag = tag[1:]
            
            # Build path
            if current_path:
                path = f"{current_path}.{tag}"
            else:
                path = tag
            
            # Check if this is a leaf element (has value)
            value = self._get_element_value(element)
            if value:
                elements[path] = value
            
            # Recursively traverse children
            for child in element:
                traverse(child, path)
        
        traverse(root)
        return elements
    
    def compare(self) -> Dict[str, List[Dict]]:
        """
        Compare two XML files and return differences
        Returns dict with keys: 'missing', 'changed', 'added'
        """
        if not self.root1 or not self.root2:
            raise ValueError("Both XML files must be loaded before comparison")
        
        # Collect elements from both files
        elements1 = self._collect_elements(self.root1, self.namespace1)
        elements2 = self._collect_elements(self.root2, self.namespace2)
        
        # Find differences
        paths1 = set(elements1.keys())
        paths2 = set(elements2.keys())
        
        missing = []
        changed = []
        added = []
        
        # Find missing (in 1 but not in 2)
        for path in paths1 - paths2:
            missing.append({
                'path': path,
                'value1': elements1[path],
                'value2': '',
                'status': 'Missing'
            })
        
        # Find added (in 2 but not in 1)
        for path in paths2 - paths1:
            added.append({
                'path': path,
                'value1': '',
                'value2': elements2[path],
                'status': 'Added'
            })
        
        # Find changed (in both but different values)
        common_paths = paths1.intersection(paths2)
        for path in common_paths:
            if elements1[path] != elements2[path]:
                changed.append({
                    'path': path,
                    'value1': elements1[path],
                    'value2': elements2[path],
                    'status': 'Changed'
                })
        
        return {
            'missing': missing,
            'changed': changed,
            'added': added
        }
    
    def export_to_excel(self, differences: Dict[str, List[Dict]], output_path: str, 
                       include_missing: bool = True, include_changed: bool = True, 
                       include_added: bool = True) -> bool:
        """Export differences to Excel file"""
        try:
            # Combine selected difference types
            all_differences = []
            if include_missing:
                all_differences.extend(differences['missing'])
            if include_changed:
                all_differences.extend(differences['changed'])
            if include_added:
                all_differences.extend(differences['added'])
            
            if not all_differences:
                logger.warning("No differences to export")
                return False
            
            # Create DataFrame
            df = pd.DataFrame(all_differences)
            
            # Reorder columns for better readability
            df = df[['status', 'path', 'value1', 'value2']]
            
            # Save to Excel
            output_path = Path(output_path)
            if output_path.suffix.lower() != '.xlsx':
                output_path = output_path.with_suffix('.xlsx')
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Differences', index=False)
                
                # Auto-adjust column widths
                worksheet = writer.sheets['Differences']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            logger.info(f"Successfully exported differences to: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to export to Excel: {e}")
            return False


class XMLDifferenceGUI:
    """GUI interface for XML Difference Tool"""
    
    def __init__(self, root):
        self.root = root
        if hasattr(self.root, 'title') and callable(self.root.title):
            self.root.title("XML Difference Tool")
        if hasattr(self.root, 'geometry') and callable(self.root.geometry):
            self.root.geometry("1200x700")
        
        # Configure styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Initialize comparator
        self.comparator = XMLComparator()
        
        # Data storage
        self.differences = None
        self.xml1_loaded = False
        self.xml2_loaded = False
        
        # Create GUI
        self.create_widgets()
        self.setup_layout()
    
    def create_widgets(self):
        """Create all GUI widgets"""
        # Top frame for file selection
        self.top_frame = ttk.Frame(self.root, padding="10")
        
        # XML File 1 selection
        self.xml1_label = ttk.Label(self.top_frame, text="First XML File:")
        self.xml1_path_var = tk.StringVar()
        self.xml1_entry = ttk.Entry(self.top_frame, textvariable=self.xml1_path_var, width=50)
        self.xml1_browse_btn = ttk.Button(self.top_frame, text="Browse...", 
                                         command=lambda: self.browse_xml_file(1))
        self.xml1_load_btn = ttk.Button(self.top_frame, text="Load", 
                                       command=lambda: self.load_xml_file(1))
        
        # XML File 2 selection
        self.xml2_label = ttk.Label(self.top_frame, text="Second XML File:")
        self.xml2_path_var = tk.StringVar()
        self.xml2_entry = ttk.Entry(self.top_frame, textvariable=self.xml2_path_var, width=50)
        self.xml2_browse_btn = ttk.Button(self.top_frame, text="Browse...", 
                                         command=lambda: self.browse_xml_file(2))
        self.xml2_load_btn = ttk.Button(self.top_frame, text="Load", 
                                       command=lambda: self.load_xml_file(2))
        
        # Filter options frame
        self.filter_frame = ttk.LabelFrame(self.top_frame, text="Filter Differences", padding="10")
        
        self.include_missing_var = tk.BooleanVar(value=True)
        self.include_missing_cb = ttk.Checkbutton(self.filter_frame, text="Missing", 
                                                 variable=self.include_missing_var)
        
        self.include_changed_var = tk.BooleanVar(value=True)
        self.include_changed_cb = ttk.Checkbutton(self.filter_frame, text="Changed", 
                                                 variable=self.include_changed_var)
        
        self.include_added_var = tk.BooleanVar(value=True)
        self.include_added_cb = ttk.Checkbutton(self.filter_frame, text="Added", 
                                               variable=self.include_added_var)
        
        # Middle frame for preview
        self.mid_frame = ttk.Frame(self.root, padding="10")
        
        # Treeview for preview
        self.tree_frame = ttk.Frame(self.mid_frame)
        self.tree_scroll = ttk.Scrollbar(self.tree_frame)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.preview_tree = ttk.Treeview(
            self.tree_frame,
            columns=('Status', 'Path', 'Value 1', 'Value 2'),
            show='headings',
            yscrollcommand=self.tree_scroll.set,
            height=20
        )
        
        # Configure tree columns
        self.preview_tree.heading('Status', text='Status')
        self.preview_tree.heading('Path', text='Path')
        self.preview_tree.heading('Value 1', text='Value 1')
        self.preview_tree.heading('Value 2', text='Value 2')
        
        self.preview_tree.column('Status', width=100, anchor='center')
        self.preview_tree.column('Path', width=400)
        self.preview_tree.column('Value 1', width=150)
        self.preview_tree.column('Value 2', width=150)
        
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_scroll.config(command=self.preview_tree.yview)
        
        # Bottom frame for buttons and status
        self.bottom_frame = ttk.Frame(self.root, padding="10")
        
        # Action buttons
        self.compare_btn = ttk.Button(self.bottom_frame, text="Compare XML Files", 
                                     command=self.compare_xml, state='disabled')
        self.export_btn = ttk.Button(self.bottom_frame, text="Export to Excel...", 
                                    command=self.export_to_excel, state='disabled')
        
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
        
        # Row 1: XML File 1
        self.xml1_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        self.xml1_entry.grid(row=0, column=1, padx=(5, 0), pady=2, sticky=(tk.W, tk.E))
        self.xml1_browse_btn.grid(row=0, column=2, padx=(5, 0), pady=2)
        self.xml1_load_btn.grid(row=0, column=3, padx=(5, 5), pady=2)
        
        # Row 2: XML File 2
        self.xml2_label.grid(row=1, column=0, sticky=tk.W, pady=2)
        self.xml2_entry.grid(row=1, column=1, padx=(5, 0), pady=2, sticky=(tk.W, tk.E))
        self.xml2_browse_btn.grid(row=1, column=2, padx=(5, 0), pady=2)
        self.xml2_load_btn.grid(row=1, column=3, padx=(5, 5), pady=2)
        
        # Row 3: Filter options
        self.filter_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
        self.include_missing_cb.grid(row=0, column=0, padx=(0, 20))
        self.include_changed_cb.grid(row=0, column=1, padx=(0, 20))
        self.include_added_cb.grid(row=0, column=2)
        
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
        self.compare_btn.grid(row=0, column=0, padx=(0, 5), pady=5)
        self.export_btn.grid(row=0, column=1, padx=(5, 0), pady=5)
        
        # Status and progress
        self.status_text.grid(row=1, column=0, columnspan=2, pady=(5, 0), sticky=(tk.W, tk.E))
        self.progress_bar.grid(row=2, column=0, columnspan=2, pady=(5, 0), sticky=(tk.W, tk.E))
        
        # Configure grid weights
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
    
    def browse_xml_file(self, which: int):
        """Open file dialog for XML file"""
        filename = filedialog.askopenfilename(
            title=f"Select XML File {which}",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if filename:
            if which == 1:
                self.xml1_path_var.set(filename)
                self.xml1_load_btn.config(state='normal')
            else:
                self.xml2_path_var.set(filename)
                self.xml2_load_btn.config(state='normal')
    
    def load_xml_file(self, which: int):
        """Load the selected XML file"""
        if which == 1:
            xml_path = self.xml1_path_var.get()
        else:
            xml_path = self.xml2_path_var.get()
        
        if not xml_path:
            messagebox.showerror("Error", f"Please select XML file {which}")
            return
        
        try:
            success = self.comparator.load_xml(xml_path, which)
            if success:
                if which == 1:
                    self.xml1_loaded = True
                    self.log_message(f"✓ XML 1 loaded: {xml_path}")
                else:
                    self.xml2_loaded = True
                    self.log_message(f"✓ XML 2 loaded: {xml_path}")
                self.update_button_states()
            else:
                self.log_message(f"✗ Failed to load XML {which}: {xml_path}", error=True)
        except Exception as e:
            self.log_message(f"Error loading XML {which}: {e}", error=True)
    
    def compare_xml(self):
        """Compare the two loaded XML files"""
        if not self.xml1_loaded or not self.xml2_loaded:
            messagebox.showerror("Error", "Please load both XML files first")
            return
        
        # Clear previous results
        self.preview_tree.delete(*self.preview_tree.get_children())
        
        # Disable buttons during processing
        self.compare_btn.config(state='disabled')
        self.export_btn.config(state='disabled')
        
        # Start comparison in background thread
        self.progress_var.set(0)
        self.log_message("Comparing XML files...")
        
        thread = Thread(target=self._compare_xml_thread)
        thread.daemon = True
        thread.start()
    
    def _compare_xml_thread(self):
        """Background thread for comparing XML"""
        try:
            # Compare XML files
            self.differences = self.comparator.compare()
            
            # Update UI in main thread
            self.root.after(0, self._update_preview_results)
            
        except Exception as e:
            self.root.after(0, lambda: self.log_message(f"Error comparing XML: {e}", error=True))
            self.root.after(0, lambda: self.compare_btn.config(state='normal'))
    
    def _update_preview_results(self):
        """Update UI with comparison results"""
        # Clear tree
        self.preview_tree.delete(*self.preview_tree.get_children())
        
        # Populate tree with results
        missing_count = len(self.differences['missing'])
        changed_count = len(self.differences['changed'])
        added_count = len(self.differences['added'])
        total_count = missing_count + changed_count + added_count
        
        # Show all differences
        for diff_type in ['missing', 'changed', 'added']:
            for diff in self.differences[diff_type]:
                status = diff['status']
                tag = diff_type
                
                self.preview_tree.insert(
                    '', 'end',
                    values=(
                        status,
                        diff['path'],
                        diff['value1'],
                        diff['value2']
                    ),
                    tags=(tag,)
                )
        
        # Configure tags for color coding
        self.preview_tree.tag_configure('missing', background='#ffebee')  # Light red
        self.preview_tree.tag_configure('changed', background='#fff3e0')  # Light orange
        self.preview_tree.tag_configure('added', background='#e8f5e9')    # Light green
        
        # Update status
        self.log_message(f"✓ Comparison complete: {total_count} differences found")
        self.log_message(f"  Missing: {missing_count}, Changed: {changed_count}, Added: {added_count}")
        
        # Update progress bar
        self.progress_var.set(100)
        
        # Enable buttons
        self.compare_btn.config(state='normal')
        if total_count > 0:
            self.export_btn.config(state='normal')
    
    def export_to_excel(self):
        """Export differences to Excel file"""
        if not self.differences:
            messagebox.showerror("Error", "No differences to export. Please compare files first.")
            return
        
        # Get filter settings
        include_missing = self.include_missing_var.get()
        include_changed = self.include_changed_cb.instate(['selected'])
        include_added = self.include_added_cb.instate(['selected'])
        
        if not (include_missing or include_changed or include_added):
            messagebox.showerror("Error", "Please select at least one difference type to export")
            return
        
        # Ask for save location
        filename = filedialog.asksaveasfilename(
            title="Save Differences As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="xml_differences.xlsx"
        )
        
        if not filename:
            return
        
        # Disable buttons during export
        self.export_btn.config(state='disabled')
        self.compare_btn.config(state='disabled')
        
        self.log_message("Exporting to Excel...")
        self.progress_var.set(0)
        
        # Start export in background thread
        thread = Thread(target=self._export_to_excel_thread, 
                       args=(filename, include_missing, include_changed, include_added))
        thread.daemon = True
        thread.start()
    
    def _export_to_excel_thread(self, filename, include_missing, include_changed, include_added):
        """Background thread for exporting to Excel"""
        try:
            # Export to Excel
            success = self.comparator.export_to_excel(
                self.differences, filename, 
                include_missing, include_changed, include_added
            )
            
            # Update UI in main thread
            if success:
                self.root.after(0, lambda: self._export_complete(filename))
            else:
                self.root.after(0, lambda: self.log_message("✗ No differences to export", error=True))
                self.root.after(0, lambda: self.export_btn.config(state='normal'))
                self.root.after(0, lambda: self.compare_btn.config(state='normal'))
            
        except Exception as e:
            self.root.after(0, lambda: self.log_message(f"Error exporting to Excel: {e}", error=True))
            self.root.after(0, lambda: self.export_btn.config(state='normal'))
            self.root.after(0, lambda: self.compare_btn.config(state='normal'))
    
    def _export_complete(self, filename):
        """Handle completion of Excel export"""
        # Update progress bar
        self.progress_var.set(100)
        
        # Log success
        self.log_message(f"✓ Differences exported to: {filename}")
        
        # Re-enable buttons
        self.export_btn.config(state='normal')
        self.compare_btn.config(state='normal')
        
        # Show success message
        messagebox.showinfo("Export Complete", 
                           f"Differences successfully exported to:\n{filename}")
    
    def update_button_states(self):
        """Update button states based on loaded files"""
        if self.xml1_loaded and self.xml2_loaded:
            self.compare_btn.config(state='normal')
        else:
            self.compare_btn.config(state='disabled')
    
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
        if error:
            logger.error(message)
        elif warning:
            logger.warning(message)
        else:
            logger.info(message)


def main():
    """Main entry point for standalone GUI"""
    import sys
    
    root = tk.Tk()
    app = XMLDifferenceGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()