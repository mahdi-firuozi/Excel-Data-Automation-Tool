import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np

class ExcelAutomationTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Professional Excel Data Processor")
        self.root.geometry("900x650")
        self.root.minsize(800, 500)
        
        # Application State
        self.df = None          # Current working DataFrame
        self.file_path = None   # Loaded file path
        
        self._setup_ui()

    def _setup_ui(self):
        """Initializes the graphical user interface components."""
        # --- Top Frame: Controls ---
        control_frame = tk.LabelFrame(self.root, text="Data Controls", padx=10, pady=10)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        # 1. File Selection
        tk.Button(control_frame, text="📁 Load Excel File", command=self.load_file, 
                  bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5)
        
        self.lbl_file = tk.Label(control_frame, text="No file loaded", fg="gray")
        self.lbl_file.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # 2. Column Selection & Sorting
        tk.Label(control_frame, text="Target Column:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.col_combobox = ttk.Combobox(control_frame, state="readonly", width=25)
        self.col_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        tk.Button(control_frame, text="Sort Data", command=self.sort_data).grid(row=1, column=2, padx=5, pady=5)

        # 3. Filtering
        tk.Label(control_frame, text="Filter (contains):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.filter_entry = tk.Entry(control_frame, width=28)
        self.filter_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        tk.Button(control_frame, text="Apply Filter", command=self.filter_data).grid(row=2, column=2, padx=5, pady=5)

        # 4. Global Actions
        actions_frame = tk.Frame(control_frame)
        actions_frame.grid(row=0, column=3, rowspan=3, padx=30, pady=5, sticky="n")
        
        tk.Button(actions_frame, text="🧹 Clean Data (Drop NaN & Duplicates)", command=self.clean_data).pack(fill=tk.X, pady=2)
        tk.Button(actions_frame, text="📊 Show Statistical Summary", command=self.show_summary).pack(fill=tk.X, pady=2)
        tk.Button(actions_frame, text="💾 Export Processed Data", command=self.export_data, 
                  bg="#2196F3", fg="white", font=("Arial", 10, "bold")).pack(fill=tk.X, pady=2)

        # --- Middle Frame: Data Preview (Treeview) ---
        preview_frame = tk.LabelFrame(self.root, text="Data Preview (Highlights rows with missing values)", padx=10, pady=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Scrollbars for Treeview
        scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL)
        
        self.tree = ttk.Treeview(preview_frame, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)

        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Configure tag for missing values highlight
        self.tree.tag_configure("missing", background="#ffcccc")

    # ---------------------------------------------------------
    # CORE FUNCTIONALITIES
    # ---------------------------------------------------------

    def load_file(self):
        """Prompts user to select an Excel file and loads it into a Pandas DataFrame."""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            self.df = pd.read_excel(file_path)
            self.file_path = file_path
            self.lbl_file.config(text=file_path.split('/')[-1], fg="black")
            
            # Update UI components
            self._update_columns_dropdown()
            self._refresh_treeview()
            messagebox.showinfo("Success", f"Successfully loaded {len(self.df)} rows.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file.\nDetails: {str(e)}")

    def clean_data(self):
        """Removes duplicate rows and rows with missing values."""
        if self._is_data_empty(): return

        try:
            initial_rows = len(self.df)
            # Remove Missing and Duplicates
            self.df.dropna(inplace=True)
            self.df.drop_duplicates(inplace=True)
            
            final_rows = len(self.df)
            removed = initial_rows - final_rows
            
            self._refresh_treeview()
            messagebox.showinfo("Data Cleaned", f"Data cleaned successfully.\nRemoved {removed} invalid/duplicate rows.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clean data.\nDetails: {str(e)}")

    def sort_data(self):
        """Sorts the dataframe based on the selected column."""
        if self._is_data_empty(): return
        
        target_col = self.col_combobox.get()
        if not target_col:
            messagebox.showwarning("Warning", "Please select a column to sort by.")
            return

        try:
            self.df.sort_values(by=target_col, ascending=True, inplace=True)
            self._refresh_treeview()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to sort data.\nDetails: {str(e)}")

    def filter_data(self):
        """Filters rows where the selected column contains the user's input string."""
        if self._is_data_empty(): return
        
        target_col = self.col_combobox.get()
        filter_val = self.filter_entry.get().strip()
        
        if not target_col or not filter_val:
            messagebox.showwarning("Warning", "Please select a column and enter a filter value.")
            return

        try:
            # Case-insensitive substring match. Converts to string safely first.
            mask = self.df[target_col].astype(str).str.contains(filter_val, case=False, na=False)
            self.df = self.df[mask]
            self._refresh_treeview()
            messagebox.showinfo("Filter Applied", f"Showing {len(self.df)} matching rows.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to filter data.\nDetails: {str(e)}")

    def show_summary(self):
        """Calculates and displays basic statistics for numeric columns."""
        if self._is_data_empty(): return

        try:
            numeric_df = self.df.select_dtypes(include=[np.number])
            if numeric_df.empty:
                messagebox.showinfo("Summary", "No numeric columns available to summarize.")
                return

            summary = numeric_df.agg(['mean', 'max', 'min']).T
            summary = summary.round(2)
            
            # Format output string
            output = "Statistical Summary (Numeric Columns):\n" + "-"*45 + "\n"
            output += summary.to_string()
            
            messagebox.showinfo("Data Summary", output)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate summary.\nDetails: {str(e)}")

    def export_data(self):
        """Exports the processed DataFrame back to an Excel file."""
        if self._is_data_empty(): return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Processed Data As"
        )
        if not save_path:
            return

        try:
            self.df.to_excel(save_path, index=False, engine='openpyxl')
            messagebox.showinfo("Success", "Data exported successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file.\nDetails: {str(e)}")

    # ---------------------------------------------------------
    # HELPER METHODS
    # ---------------------------------------------------------

    def _is_data_empty(self) -> bool:
        """Helper to check if DataFrame exists and is not empty."""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first.")
            return True
        if self.df.empty:
            messagebox.showwarning("Warning", "The current dataset is empty.")
            return True
        return False

    def _update_columns_dropdown(self):
        """Populates the combobox with column headers."""
        if self.df is not None:
            cols = list(self.df.columns)
            self.col_combobox['values'] = cols
            if cols:
                self.col_combobox.current(0) # Select first column by default

    def _refresh_treeview(self):
        """Clears and repopulates the Treeview with current DataFrame data."""
        self.tree.delete(*self.tree.get_children())
        
        if self.df is None or self.df.empty:
            return

        # Setup Columns
        self.tree["column"] = list(self.df.columns)
        self.tree["show"] = "headings"
        for col in self.tree["column"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor=tk.CENTER)

        # Limit rows displayed to prevent UI freezing on massive datasets
        display_df = self.df.head(1000) 
        
        for index, row in display_df.iterrows():
            values = list(row.values)
            # Check if row contains any missing values (NaN)
            has_nan = pd.isna(values).any()
            
            # Clean display for NaNs
            values = ["" if pd.isna(x) else x for x in values]
            
            if has_nan:
                self.tree.insert("", "end", values=values, tags=("missing",))
            else:
                self.tree.insert("", "end", values=values)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelAutomationTool(root)
    root.mainloop()
