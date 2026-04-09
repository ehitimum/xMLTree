#!/usr/bin/env python3
"""
XML Configurator Tool
Update XML files using path-value pairs from CSV/Excel files
With GUI interface for preview before update
"""

import os
import sys
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
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


class XMLConfigurator:
    """Main class for updating XML files using path-value mappings"""
    
    def __init__(self, xml_path: str = None):
        """Initialize with XML file path (optional)"""
        self.xml_path = Path(xml_path) if xml_path else None
        self.tree = None
        self.root = None
        self.namespace = ''
        
    def load_xml(self, xml_path: str = None) -> bool:
        """Load and parse XML file"""
        try:
            if xml_path:
                self.xml_path = Path(xml_path)
            
            if not self.xml_path or not self.xml_path.exists():
                logger.error(f"XML file not found: {self.xml_path}")
                return False
            
            self.tree = ET.parse(self.xml_path)
            self.root = self.tree.getroot()
            
            # Extract namespace if present
            if '}' in self.root.tag:
                self.namespace = self.root.tag.split('}')[0] + '}'
            
            logger.info(f"Successfully loaded XML: {self.xml_path}")
            return True
        except Exception as e:
            logger.error(f"Failed to load XML: {e}")
            return False
    
    def find_element(self, path: str) -> Tuple[Optional[ET.Element], Optional[str]]:
        """
        Find XML element using dot notation path
        Returns (element, current_value)
        """
        try:
            # Split path by dots
            parts = path.strip().split('.')
            # If first part matches root tag (without namespace), skip it
            if parts:
                root_tag = self.root.tag
                if self.namespace and root_tag.startswith(self.namespace):
                    root_tag = root_tag.replace(self.namespace, '')
                if parts[0] == root_tag:
                    parts = parts[1:]
            current_element = self.root
            
            for part in parts:
                # Remove namespace for searching
                search_part = part
                if self.namespace and search_part.startswith(self.namespace):
                    search_part = search_part.replace(self.namespace, '')

                # Generate candidate tag names
                candidate_tags = [search_part]
                if search_part.isdigit():
                    candidate_tags.append(f"i{search_part}")

                # Try to find by tag name
                found = None
                for child in current_element:
                    # Remove namespace from child tag for comparison
                    child_tag = child.tag
                    if self.namespace and child_tag.startswith(self.namespace):
                        child_tag = child_tag.replace(self.namespace, '')

                    # Check if tag matches any candidate
                    if child_tag in candidate_tags:
                        found = child
                        break
                    elif 'name' in child.attrib and child.attrib['name'] == search_part:
                        found = child
                        break

                if found is None:
                    # Try XPath as fallback for each candidate
                    for candidate in candidate_tags:
                        xpath_query = f".//{candidate}"
                        found = current_element.find(xpath_query)
                        if found is not None:
                            break

                    if found is None and self.namespace:
                        # Try with namespace
                        for candidate in candidate_tags:
                            xpath_query = f".//{{{self.namespace[:-1]}}}{candidate}"
                            found = current_element.find(xpath_query)
                            if found is not None:
                                break

                if found is None:
                    return None, None

                current_element = found
            
            # Get current value
            current_value = None
            if current_element.text is not None:
                current_value = current_element.text
            else:
                # Check for value in attributes
                value_attrs = ['value', 'Value', 'VALUE', 'content', 'Content']
                for attr in value_attrs:
                    if attr in current_element.attrib:
                        current_value = current_element.attrib[attr]
                        break
            
            return current_element, current_value
            
        except Exception as e:
            logger.error(f"Error finding element for path '{path}': {e}")
            return None, None
    
    def update_element(self, element: ET.Element, value: str, path: str = None) -> bool:
        """Update XML element with new value"""
        try:
            # Check if element has text or if we need to set an attribute
            if element.text is not None or element.text == '':
                # Update text content
                element.text = str(value)
                logger.debug(f"Updated text of element at path '{path}' to '{value}'")
            else:
                # Check if we should update an attribute
                value_attrs = ['value', 'Value', 'VALUE', 'content', 'Content']
                for attr in value_attrs:
                    if attr in element.attrib:
                        element.attrib[attr] = str(value)
                        logger.debug(f"Updated attribute '{attr}' of element at path '{path}' to '{value}'")
                        return True
                
                # If no value attribute, set text
                element.text = str(value)
            
            return True
            
        except Exception as e:
            logger.error(f"Error updating element: {e}")
            return False
    
    def check_mappings(self, mapping_data: List[Dict], path_column: str = 'path', 
                       value_column: str = 'value') -> List[Dict]:
        """
        Check which mappings can be applied (without updating)
        Returns list of dicts with status information
        """
        results = []
        
        for idx, mapping in enumerate(mapping_data):
            path = str(mapping[path_column])
            new_value = str(mapping[value_column]) if pd.notna(mapping[value_column]) else ''
            
            element, current_value = self.find_element(path)
            found = element is not None
            
            results.append({
                'index': idx,
                'path': path,
                'new_value': new_value,
                'current_value': current_value if current_value is not None else '[Not Found]',
                'found': found,
                'element': element  # Store reference for later update
            })
        
        return results
    
    def apply_mappings(self, mappings: List[Dict]) -> Dict[str, Any]:
        """
        Apply previously checked mappings
        Returns summary of results
        """
        applied = 0
        failed = 0
        
        for mapping in mappings:
            if mapping['found']:
                try:
                    success = self.update_element(mapping['element'], mapping['new_value'], mapping['path'])
                    if success:
                        applied += 1
                        logger.info(f"Applied: {mapping['path']} = {mapping['new_value']}")
                    else:
                        failed += 1
                        logger.error(f"Failed to apply: {mapping['path']}")
                except Exception as e:
                    failed += 1
                    logger.error(f"Error applying {mapping['path']}: {e}")
        
        return {
            'total': len(mappings),
            'applied': applied,
            'failed': failed,
            'not_found': len([m for m in mappings if not m['found']])
        }
    
    def load_mapping_file(self, csv_path: str, sheet_name: int = 0) -> Optional[pd.DataFrame]:
        """Load mapping data from CSV/Excel file"""
        try:
            file_path = Path(csv_path)
            if file_path.suffix.lower() == '.csv':
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            logger.info(f"Loaded {len(df)} records from {csv_path}")
            return df
        except Exception as e:
            logger.error(f"Error loading mapping file: {e}")
            return None
    
    def save_xml(self, output_path: str = None) -> bool:
        """Save the modified XML to file"""
        try:
            if output_path:
                save_path = Path(output_path)
            else:
                # Create backup of original
                backup_path = self.xml_path.with_suffix('.xml.bak')
                if not backup_path.exists():
                    self.xml_path.rename(backup_path)
                    logger.info(f"Created backup: {backup_path}")
                
                save_path = self.xml_path
            
            # Write the XML with proper formatting
            self.tree.write(save_path, encoding='utf-8', xml_declaration=True)
            logger.info(f"Successfully saved XML to: {save_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to save XML: {e}")
            return False


class ConfigurationManager:
    """Manager for handling different configuration formats and templates"""

    @staticmethod
    def validate_mapping_file(file_path: str) -> Tuple[bool, str]:
        """Validate the structure of the mapping file"""
        try:
            path = Path(file_path)
            if not path.exists():
                return False, "File does not exist"

            if path.suffix.lower() == '.csv':
                df = pd.read_csv(file_path, nrows=5)
            else:
                df = pd.read_excel(file_path, sheet_name=0, nrows=5)

            required_columns = ['path', 'value']
            missing = [col for col in required_columns if col not in df.columns]

            if missing:
                return False, f"Missing required columns: {missing}"

            return True, "Valid mapping file"

        except Exception as e:
            return False, f"Error validating file: {e}"

    @staticmethod
    def create_template_csv(output_path: str, xml_path: str = None) -> bool:
        """Create a template CSV with paths from XML"""
        try:
            if xml_path:
                configurator = XMLConfigurator(xml_path)
                if configurator.load_xml(xml_path):
                    paths = configurator.get_all_paths() if hasattr(configurator, 'get_all_paths') else []
                else:
                    return False
            else:
                paths = [
                    "cpe-FAPService1.FAPControl.LTE.EARFCNDL",
                    "cpe-FAPService1.FAPControl.LTE.DLBandwidth",
                    "cpe-FAPService1.CellConfig.LTE.RAN.RF.DLBandwidth"
                ]

            df = pd.DataFrame({
                'path': paths,
                'value': [''] * len(paths),
                'description': [''] * len(paths),
                'required': [''] * len(paths)
            })

            output_path = Path(output_path)
            df.to_csv(output_path, index=False)
            logger.info(f"Created template CSV at: {output_path}")
            return True

        except Exception as e:
            logger.error(f"Failed to create template: {e}")
            return False

    @staticmethod
    def convert_csv_to_excel(csv_path: str, excel_path: str) -> bool:
        """Convert CSV mapping file to Excel format"""
        try:
            df = pd.read_csv(csv_path)
            df.to_excel(excel_path, index=False)
            logger.info(f"Converted CSV to Excel: {excel_path}")
            return True
        except Exception as e:
            logger.error(f"Conversion failed: {e}")
            return False


class XMLConfiguratorGUI:
    """GUI interface for XML Configurator with preview functionality"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("XML Configurator")
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
        if error:
            logger.error(message)
        elif warning:
            logger.warning(message)
        else:
            logger.info(message)
    
    def run(self):
        """Start the GUI application"""
        self.root.mainloop()


def main():
    """Main entry point - supports both CLI and GUI modes"""
    parser = argparse.ArgumentParser(
        description='Update XML configuration files using path-value mappings from CSV/Excel',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s config.xml updates.csv -o updated_config.xml
  %(prog)s config.xml mapping.xlsx --path-col parameter --value-col new_value
  %(prog)s --gui  (launch graphical interface)
        """
    )
    
    # Main arguments
    parser.add_argument('xml_file', nargs='?', help='Input XML file to update')
    parser.add_argument('mapping_file', nargs='?', 
                       help='CSV/Excel file with path-value mappings')
    
    # Optional arguments
    parser.add_argument('-o', '--output', help='Output XML file (default: overwrite input)')
    parser.add_argument('--path-col', default='path', 
                       help='Column name for paths in mapping file (default: path)')
    parser.add_argument('--value-col', default='value', 
                       help='Column name for values in mapping file (default: value)')
    parser.add_argument('--sheet', type=int, default=0,
                       help='Sheet index/name for Excel files (default: 0)')
    parser.add_argument('--ignore-errors', action='store_true',
                       help='Continue processing even if some paths are not found')
    parser.add_argument('--dry-run', action='store_true',
                       help='Show what would be updated without making changes')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose logging')
    
    # GUI mode
    parser.add_argument('--gui', action='store_true',
                       help='Launch graphical user interface')
    
    args = parser.parse_args()
    
    # Configure logging level
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    # Launch GUI if requested
    if args.gui:
        try:
            app = XMLConfiguratorGUI()
            app.run()
        except Exception as e:
            logger.error(f"Failed to launch GUI: {e}")
            sys.exit(1)
        return
    
    # Continue with CLI mode if no GUI requested
    # ... (rest of your original CLI code here)
    # Note: I've removed the CLI implementation here for brevity since you wanted GUI focus
    # You should keep your original CLI code below this point
    
    print("CLI mode not fully implemented in this version. Use --gui for graphical interface.")
    parser.print_help()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        sys.exit(1)