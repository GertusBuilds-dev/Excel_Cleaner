"""
Excel Cleaner Pro - Professional Data Cleaning Tool
Author: GertusBuilds
Version: 2.0
Description: A professional-grade Excel cleaning application with modern UI.

WINDOW SIZE OPTIONS:
By default, the application uses a larger window (800x900) to show all content.
If you have a smaller screen or prefer scrolling, you can enable the scrollable
interface by changing 'use_scrollable_ui = False' to 'use_scrollable_ui = True'
in the ExcelCleanerGUI __init__ method (around line 207).
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import shutil
import webbrowser
import logging
from typing import Dict, List, Optional, Callable
from dataclasses import dataclass
from datetime import datetime
from PIL import Image, ImageTk
import json


@dataclass
class CleaningConfig:
    """Configuration for cleaning operations."""
    remove_duplicates: bool = False
    remove_empty_rows: bool = False
    remove_empty_columns: bool = False
    trim_spaces: bool = False
    normalize_column_names: bool = False
    title_case_cells: bool = False


@dataclass
class AppTheme:
    """Application theme configuration."""
    name: str
    bg_color: str
    fg_color: str
    accent_color: str
    button_color: str
    button_text: str
    entry_bg: str


class ExcelCleaner:
    """Core Excel cleaning functionality."""

    def __init__(self):
        self.setup_logging()

    def setup_logging(self) -> None:
        """Setup application logging."""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('excel_cleaner.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def remove_duplicates(self, df: pd.DataFrame) -> pd.DataFrame:
        """Remove duplicate rows from DataFrame."""
        initial_count = len(df)
        df_cleaned = df.drop_duplicates()
        removed_count = initial_count - len(df_cleaned)
        self.logger.info(f"Removed {removed_count} duplicate rows")
        return df_cleaned

    def remove_empty_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        """Remove completely empty rows."""
        initial_count = len(df)
        df_cleaned = df.dropna(how="all")
        removed_count = initial_count - len(df_cleaned)
        self.logger.info(f"Removed {removed_count} empty rows")
        return df_cleaned

    def remove_empty_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Remove completely empty columns."""
        initial_cols = len(df.columns)
        df_cleaned = df.dropna(axis=1, how="all")
        removed_count = initial_cols - len(df_cleaned.columns)
        self.logger.info(f"Removed {removed_count} empty columns")
        return df_cleaned

    def trim_spaces(self, df: pd.DataFrame) -> pd.DataFrame:
        """Trim leading and trailing spaces from string cells."""
        df_cleaned = df.copy()
        for col in df_cleaned.columns:
            if df_cleaned[col].dtype == 'object':
                df_cleaned[col] = df_cleaned[col].astype(str).str.strip()
        self.logger.info("Trimmed spaces from all text cells")
        return df_cleaned

    def normalize_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """Normalize column names to Title Case."""
        df_cleaned = df.copy()
        old_columns = list(df_cleaned.columns)
        df_cleaned.columns = [str(col).strip().title().replace('_', ' ')
                              for col in df_cleaned.columns]
        self.logger.info(f"Normalized {len(old_columns)} column names")
        return df_cleaned

    def title_case_cells(self, df: pd.DataFrame) -> pd.DataFrame:
        """Convert text cells to Title Case."""
        df_cleaned = df.copy()
        for col in df_cleaned.columns:
            if df_cleaned[col].dtype == 'object':
                df_cleaned[col] = df_cleaned[col].astype(str).str.title()
        self.logger.info("Converted text cells to Title Case")
        return df_cleaned

    def clean_excel_file(self, file_path: str, config: CleaningConfig,
                         progress_callback: Optional[Callable] = None) -> tuple[str, str, dict]:
        """
        Clean Excel file based on configuration.

        Returns:
            tuple: (backup_path, cleaned_path, statistics)
        """
        try:
            # Create backup
            backup_path = self._create_backup(file_path)

            # Load data
            self.logger.info(f"Loading Excel file: {file_path}")
            df = pd.read_excel(file_path)
            original_stats = self._get_dataframe_stats(df)

            # Apply cleaning operations
            operations = [
                ("remove_duplicates", self.remove_duplicates,
                 config.remove_duplicates),
                ("remove_empty_rows", self.remove_empty_rows,
                 config.remove_empty_rows),
                ("remove_empty_columns", self.remove_empty_columns,
                 config.remove_empty_columns),
                ("trim_spaces", self.trim_spaces, config.trim_spaces),
                ("normalize_column_names", self.normalize_column_names,
                 config.normalize_column_names),
                ("title_case_cells", self.title_case_cells, config.title_case_cells),
            ]

            enabled_operations = [op for op in operations if op[2]]
            total_ops = len(enabled_operations)

            for i, (name, func, enabled) in enumerate(enabled_operations):
                if enabled:
                    df = func(df)
                    if progress_callback:
                        progress_callback(int((i + 1) / total_ops * 100))

            # Save cleaned file
            cleaned_path = self._get_cleaned_filename(file_path)
            df.to_excel(cleaned_path, index=False)

            final_stats = self._get_dataframe_stats(df)
            statistics = {
                'original': original_stats,
                'final': final_stats,
                'operations_applied': [op[0] for op in enabled_operations]
            }

            self.logger.info(f"Successfully cleaned file: {cleaned_path}")
            return backup_path, cleaned_path, statistics

        except Exception as e:
            self.logger.error(f"Error cleaning file: {str(e)}")
            raise

    def _create_backup(self, file_path: str) -> str:
        """Create backup of original file."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = file_path.replace(".xlsx", f"_backup_{timestamp}.xlsx")
        shutil.copy2(file_path, backup_path)
        return backup_path

    def _get_cleaned_filename(self, file_path: str) -> str:
        """Generate filename for cleaned file."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return file_path.replace(".xlsx", f"_cleaned_{timestamp}.xlsx")

    def _get_dataframe_stats(self, df: pd.DataFrame) -> dict:
        """Get statistics about DataFrame."""
        return {
            'rows': len(df),
            'columns': len(df.columns),
            'cells': df.size,
            'empty_cells': df.isnull().sum().sum(),
            'data_types': df.dtypes.value_counts().to_dict()
        }


class ExcelCleanerGUI:
    """Professional GUI for Excel Cleaner application."""

    def __init__(self):
        self.cleaner = ExcelCleaner()
        self.themes = self._load_themes()
        self.current_theme = "professional_light"
        self.config = CleaningConfig()
        
        # UI Configuration Options:
        # False = Larger fixed window (800x900) - RECOMMENDED for most users
        # True = Scrollable interface (works with smaller screens)
        self.use_scrollable_ui = False  # Change to True for scrollable UI
        
        self.setup_gui()
        self.apply_theme()

    def _load_themes(self) -> Dict[str, AppTheme]:
        """Load application themes."""
        return {
            "professional_light": AppTheme(
                name="Professional Light",
                bg_color="#FFFFFF",
                fg_color="#2C3E50",
                accent_color="#3498DB",
                button_color="#2980B9",
                button_text="#FFFFFF",
                entry_bg="#F8F9FA"
            ),
            "professional_dark": AppTheme(
                name="Professional Dark",
                bg_color="#2C3E50",
                fg_color="#ECF0F1",
                accent_color="#3498DB",
                button_color="#E74C3C",
                button_text="#FFFFFF",
                entry_bg="#34495E"
            ),
            "modern_blue": AppTheme(
                name="Modern Blue",
                bg_color="#F0F8FF",
                fg_color="#1E3A8A",
                accent_color="#3B82F6",
                button_color="#1E40AF",
                button_text="#FFFFFF",
                entry_bg="#DBEAFE"
            )
        }

    def setup_gui(self) -> None:
        """Setup the main GUI."""
        self.root = tk.Tk()
        self.root.title("Excel Cleaner Pro v1.0")
        self.root.geometry("800x900")  # Increased height from 750 to 900
        self.root.minsize(750, 850)   # Increased minimum size
        
        # Center the window
        self.root.eval('tk::PlaceWindow . center')

        # Set window icon - try multiple options
        try:
            # Try PNG logo first for window icon
            if os.path.exists("logo.png"):
                self.root_logo = tk.PhotoImage(file="logo.png")
                self.root.iconphoto(True, self.root_logo)
            elif os.path.exists("icon.ico"):
                self.root.iconbitmap("icon.ico")
        except Exception as e:
            print(f"Could not set window icon: {e}")

        self.create_widgets()
        self.setup_keyboard_shortcuts()

    def create_widgets(self) -> None:
        """Create all GUI widgets."""
        if self.use_scrollable_ui:
            # Create a scrollable main container
            self.create_scrollable_container()
            main_frame = self.scrollable_frame
        else:
            # Use regular container with larger window
            main_frame = tk.Frame(self.root)
            main_frame.pack(fill="both", expand=True, padx=25, pady=15)

        # Header section
        self.create_header(main_frame)

        # Separator
        ttk.Separator(main_frame, orient="horizontal").pack(
            fill="x", pady=(15, 20))

        # Options section
        self.create_options_section(main_frame)

        # Progress section
        self.create_progress_section(main_frame)

        # Actions section
        self.create_actions_section(main_frame)

        # Footer
        self.create_footer(main_frame)
        
        # Status bar (created last to stay at bottom)
        self.create_status_bar()

    def create_scrollable_container(self) -> None:
        """Create a scrollable container for the main content."""
        # Create main canvas and scrollbar
        self.canvas = tk.Canvas(self.root, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(
            self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)
        
        # Configure scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all"))
        )
        
        # Create window in canvas
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw"
        )
        
        # Configure canvas scrolling
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Bind mousewheel to canvas
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        self.canvas.bind("<MouseWheel>", _on_mousewheel)
        
        # Bind canvas resize to adjust frame width
        def _on_canvas_configure(event):
            canvas_width = event.width
            self.canvas.itemconfig(self.canvas_window, width=canvas_width)
        
        self.canvas.bind("<Configure>", _on_canvas_configure)
        
        # Pack canvas and scrollbar
        self.canvas.pack(
            side="left", fill="both", expand=True, padx=(25, 0), pady=15)
        self.scrollbar.pack(
            side="right", fill="y", padx=(0, 25), pady=15)

    def create_header(self, parent: tk.Widget) -> None:
        """Create header with logo and title."""
        header_frame = tk.Frame(parent)
        header_frame.pack(fill="x", pady=(0, 15))

        # Logo and title (left side)
        title_frame = tk.Frame(header_frame)
        title_frame.pack(side="left", fill="y")

        try:
            # Try to load logo with multiple fallback options
            logo_files = [
                "logo.png",
                "93520a39-6b10-4054-8409-ee3e05923881.png",
                "icon.png"
            ]
            logo_loaded = False
            
            for logo_file in logo_files:
                if os.path.exists(logo_file):
                    try:
                        logo_img = Image.open(logo_file).resize((40, 40))
                        self.logo = ImageTk.PhotoImage(logo_img)
                        logo_label = tk.Label(title_frame, image=self.logo)
                        logo_label.pack(side="left", padx=(0, 12))
                        logo_loaded = True
                        break
                    except Exception as e:
                        print(f"Could not load {logo_file}: {e}")
                        continue
            
            if not logo_loaded:
                # Create a simple text logo if no image found
                logo_label = tk.Label(
                    title_frame,
                    text="📊",
                    font=("Segoe UI", 24),
                    width=2
                )
                logo_label.pack(side="left", padx=(0, 12))
                
        except Exception as e:
            print(f"Logo loading failed: {e}")
            # Create a simple text logo as ultimate fallback
            logo_label = tk.Label(
                title_frame,
                text="📊",
                font=("Segoe UI", 24),
                width=2
            )
            logo_label.pack(side="left", padx=(0, 12))

        title_content = tk.Frame(title_frame)
        title_content.pack(side="left", fill="y")
        
        title_label = tk.Label(
            title_content,
            text="Excel Cleaner Pro",
            font=("Segoe UI", 20, "bold")
        )
        title_label.pack(anchor="w")

        version_label = tk.Label(
            title_content,
            text="Professional Data Cleaning Tool v2.0",
            font=("Segoe UI", 9),
            fg="#7F8C8D"
        )
        version_label.pack(anchor="w")

        # Controls (right side)
        controls_frame = tk.Frame(header_frame)
        controls_frame.pack(side="right", fill="y")
        
        # Controls organized vertically for better layout
        theme_frame = tk.Frame(controls_frame)
        theme_frame.pack(anchor="e", pady=(0, 5))
        
        theme_label = tk.Label(
            theme_frame, text="Theme:", font=("Segoe UI", 9))
        theme_label.pack(side="left", padx=(0, 5))

        self.theme_var = tk.StringVar(value=self.current_theme)
        theme_combo = ttk.Combobox(
            theme_frame,
            textvariable=self.theme_var,
            values=["professional_light", "professional_dark", "modern_blue"],
            state="readonly",
            width=18,
            font=("Segoe UI", 9)
        )
        theme_combo.pack(side="left")
        theme_combo.bind("<<ComboboxSelected>>", self.on_theme_change)

        # Help button
        help_btn = tk.Button(
            controls_frame,
            text="❓ Help",
            command=self.show_help,
            font=("Segoe UI", 9),
            width=12,
            relief="raised",
            bd=1
        )
        help_btn.pack(anchor="e")

    def create_options_section(self, parent: tk.Widget) -> None:
        """Create cleaning options section."""
        options_frame = tk.LabelFrame(
            parent,
            text="Cleaning Options",
            font=("Segoe UI", 12, "bold"),
            padx=15,
            pady=10
        )
        options_frame.pack(fill="x", pady=10)

        # Create variables for options
        self.option_vars = {
            "remove_duplicates": tk.BooleanVar(),
            "remove_empty_rows": tk.BooleanVar(),
            "remove_empty_columns": tk.BooleanVar(),
            "trim_spaces": tk.BooleanVar(),
            "normalize_column_names": tk.BooleanVar(),
            "title_case_cells": tk.BooleanVar()
        }

        options_data = [
            ("remove_duplicates", "Remove Duplicate Rows",
             "Removes identical rows from your data"),
            ("remove_empty_rows", "Remove Empty Rows",
             "Removes rows that contain no data"),
            ("remove_empty_columns", "Remove Empty Columns",
             "Removes columns that contain no data"),
            ("trim_spaces", "Trim Whitespace",
             "Removes leading and trailing spaces from text"),
            ("normalize_column_names", "Normalize Column Names",
             "Standardizes column headers to Title Case"),
            ("title_case_cells", "Title Case Text",
             "Converts text cells to Title Case format")
        ]

        # Options content frame with better spacing
        options_content = tk.Frame(options_frame)
        options_content.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Create two columns for options
        left_frame = tk.Frame(options_content)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 15))

        right_frame = tk.Frame(options_content)
        right_frame.pack(side="right", fill="both", expand=True, padx=(15, 0))

        for i, (key, text, description) in enumerate(options_data):
            frame = left_frame if i < 3 else right_frame

            option_frame = tk.Frame(frame)
            option_frame.pack(fill="x", pady=5, padx=5)

            cb = tk.Checkbutton(
                option_frame,
                text=text,
                variable=self.option_vars[key],
                font=("Segoe UI", 10, "bold"),
                anchor="w"
            )
            cb.pack(fill="x")

            desc_label = tk.Label(
                option_frame,
                text=description,
                font=("Segoe UI", 9),
                anchor="w",
                wraplength=250,
                justify="left"
            )
            desc_label.pack(fill="x", padx=(25, 0), pady=(2, 0))

        # Quick action buttons with better styling
        quick_frame = tk.Frame(options_frame)
        quick_frame.pack(fill="x", pady=(15, 5), padx=10)
        
        # Center the quick action buttons
        button_container = tk.Frame(quick_frame)
        button_container.pack()

        select_all_btn = tk.Button(
            button_container,
            text="✅ Select All",
            command=self.select_all_options,
            font=("Segoe UI", 9),
            width=12,
            relief="raised",
            bd=1
        )
        select_all_btn.pack(side="left", padx=(0, 10))

        clear_all_btn = tk.Button(
            button_container,
            text="❌ Clear All",
            command=self.clear_all_options,
            font=("Segoe UI", 9),
            width=12,
            relief="raised",
            bd=1
        )
        clear_all_btn.pack(side="left")

    def create_progress_section(self, parent: tk.Widget) -> None:
        """Create progress section."""
        progress_frame = tk.LabelFrame(
            parent,
            text="Progress",
            font=("Segoe UI", 11, "bold"),
            padx=15,
            pady=10
        )
        progress_frame.pack(fill="x", pady=(15, 10))

        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            length=500,
            mode="determinate",
            variable=self.progress_var,
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(pady=(5, 10))

        self.progress_label = tk.Label(
            progress_frame,
            text="Ready to clean your Excel files",
            font=("Segoe UI", 10),
            fg="#2C3E50"
        )
        self.progress_label.pack()

    def create_actions_section(self, parent: tk.Widget) -> None:
        """Create main action buttons."""
        actions_frame = tk.Frame(parent)
        actions_frame.pack(fill="x", pady=(20, 15))

        # Main action button with better styling
        main_button_frame = tk.Frame(actions_frame)
        main_button_frame.pack(pady=(0, 15))
        
        self.clean_btn = tk.Button(
            main_button_frame,
            text="📁 Select & Clean Excel File",
            command=self.select_and_clean_file,
            font=("Segoe UI", 14, "bold"),
            height=2,
            width=28,
            relief="raised",
            bd=2,
            cursor="hand2"
        )
        self.clean_btn.pack()

        # Additional action buttons with better organization
        button_frame = tk.LabelFrame(
            actions_frame,
            text="Additional Actions",
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=10
        )
        button_frame.pack(fill="x")
        
        # Create a centered container for the buttons
        button_container = tk.Frame(button_frame)
        button_container.pack()

        save_btn = tk.Button(
            button_container,
            text="💾 Save Settings",
            command=self.save_settings,
            font=("Segoe UI", 9),
            width=15,
            relief="raised",
            bd=1
        )
        save_btn.pack(side="left", padx=8)

        load_btn = tk.Button(
            button_container,
            text="� Load Settings",
            command=self.load_settings,
            font=("Segoe UI", 9),
            width=15,
            relief="raised",
            bd=1
        )
        load_btn.pack(side="left", padx=8)

        log_btn = tk.Button(
            button_container,
            text="� View Log",
            command=self.view_log,
            font=("Segoe UI", 9),
            width=15,
            relief="raised",
            bd=1
        )
        log_btn.pack(side="left", padx=8)

    def create_status_bar(self) -> None:
        """Create status bar."""
        self.status_frame = tk.Frame(self.root)
        self.status_frame.pack(side="bottom", fill="x")

        ttk.Separator(self.status_frame, orient="horizontal").pack(fill="x")

        status_content = tk.Frame(self.status_frame)
        status_content.pack(fill="x", padx=10, pady=2)

        self.status_label = tk.Label(
            status_content,
            text="Ready",
            font=("Segoe UI", 9),
            anchor="w"
        )
        self.status_label.pack(side="left")

        # Version info on right
        version_label = tk.Label(
            status_content,
            text="Excel Cleaner Pro v2.0",
            font=("Segoe UI", 8),
            anchor="e"
        )
        version_label.pack(side="right")

    def create_footer(self, parent: tk.Widget) -> None:
        """Create footer with links."""
        # Add some space before footer
        spacer = tk.Frame(parent)
        spacer.pack(fill="x", pady=10)
        
        footer_frame = tk.Frame(parent)
        footer_frame.pack(side="bottom", fill="x", pady=(10, 0))

        ttk.Separator(footer_frame, orient="horizontal").pack(
            fill="x", pady=(0, 8))

        links_frame = tk.Frame(footer_frame)
        links_frame.pack()

        info_label = tk.Label(
            links_frame,
            text="Developed by GertusBuilds • ",
            font=("Segoe UI", 9),
            fg="#7F8C8D"
        )
        info_label.pack(side="left")
        
        # Store references to prevent garbage collection
        self.github_btn = tk.Button(
            links_frame,
            text="🌐 GitHub",
            command=lambda: webbrowser.open(
                "https://github.com/GertusBuilds-dev"),
            font=("Segoe UI", 9, "underline"),
            relief="flat",
            bd=0,
            cursor="hand2",
            fg="#3498DB",
            bg=self.themes[self.current_theme].bg_color
        )
        self.github_btn.pack(side="left", padx=(0, 15))

        self.sponsor_btn = tk.Button(
            links_frame,
            text="❤️ Support",
            command=lambda: webbrowser.open(
                "https://buymeacoffee.com/gertusbuilds.dev"),
            font=("Segoe UI", 9, "underline"),
            relief="flat",
            bd=0,
            cursor="hand2",
            fg="#E74C3C",
            bg=self.themes[self.current_theme].bg_color
        )
        self.sponsor_btn.pack(side="left")
        
        # Add a simple test to verify links are working
        test_label = tk.Label(
            footer_frame,
            text="Links: GitHub • Support • Status: Active",
            font=("Segoe UI", 8),
            fg="#95A5A6"
        )
        test_label.pack(pady=(5, 0))

    def setup_keyboard_shortcuts(self) -> None:
        """Setup keyboard shortcuts."""
        self.root.bind("<Control-o>", lambda e: self.select_and_clean_file())
        self.root.bind("<Control-h>", lambda e: self.show_help())
        self.root.bind("<Control-s>", lambda e: self.save_settings())
        self.root.bind("<Control-l>", lambda e: self.load_settings())
        self.root.bind("<F1>", lambda e: self.show_help())

    def apply_theme(self) -> None:
        """Apply current theme to all widgets."""
        theme = self.themes[self.current_theme]
        
        # Configure root
        self.root.configure(bg=theme.bg_color)
        
        # Configure ttk style
        style = ttk.Style()
        style.configure(
            "Custom.Horizontal.TProgressbar",
            background=theme.accent_color,
            troughcolor=theme.entry_bg,
            borderwidth=1,
            lightcolor=theme.accent_color,
            darkcolor=theme.accent_color
        )
        
        def configure_widget(widget):
            """Configure individual widget styling."""
            try:
                widget_class = widget.winfo_class()
                
                if widget_class == "Button":
                    # Check if it's a main action button or regular button
                    text = widget.cget("text")
                    if "📁" in text:  # Main action button
                        widget.configure(
                            bg=theme.button_color,
                            fg=theme.button_text,
                            activebackground=theme.accent_color,
                            activeforeground=theme.button_text,
                            highlightbackground=theme.bg_color
                        )
                    elif "🌐" in text or "❤️" in text:  # Footer links
                        widget.configure(
                            bg=theme.bg_color,
                            activebackground=theme.bg_color,
                            highlightbackground=theme.bg_color
                        )
                    elif widget.cget("relief") == "flat":  # Other flat buttons
                        widget.configure(
                            bg=theme.bg_color,
                            activebackground=theme.bg_color
                        )
                    else:  # Regular buttons
                        widget.configure(
                            bg=theme.entry_bg,
                            fg=theme.fg_color,
                            activebackground=theme.accent_color,
                            activeforeground=theme.button_text,
                            highlightbackground=theme.bg_color
                        )
                elif widget_class == "Label":
                    widget.configure(
                        bg=theme.bg_color,
                        fg=theme.fg_color
                    )
                elif widget_class == "Frame":
                    widget.configure(bg=theme.bg_color)
                elif widget_class == "Labelframe":
                    widget.configure(
                        bg=theme.bg_color,
                        fg=theme.fg_color
                    )
                elif widget_class == "Checkbutton":
                    widget.configure(
                        bg=theme.bg_color,
                        fg=theme.fg_color,
                        activebackground=theme.bg_color,
                        activeforeground=theme.fg_color,
                        selectcolor=theme.entry_bg
                    )
                elif widget_class == "Toplevel":
                    widget.configure(bg=theme.bg_color)
                elif widget_class == "Text":
                    widget.configure(
                        bg=theme.entry_bg,
                        fg=theme.fg_color,
                        insertbackground=theme.fg_color
                    )
            except tk.TclError:
                pass  # Widget doesn't support the configuration
        
        # Apply theme recursively to all widgets
        def apply_to_children(parent):
            configure_widget(parent)
            for child in parent.winfo_children():
                apply_to_children(child)
        
        apply_to_children(self.root)
        
        # Update status
        self.update_status(f"Applied theme: {theme.name}")

    def on_theme_change(self, event=None) -> None:
        """Handle theme change."""
        self.current_theme = self.theme_var.get()
        self.apply_theme()

    def select_all_options(self) -> None:
        """Select all cleaning options."""
        for var in self.option_vars.values():
            var.set(True)
        self.update_status("All options selected")

    def clear_all_options(self) -> None:
        """Clear all cleaning options."""
        for var in self.option_vars.values():
            var.set(False)
        self.update_status("All options cleared")

    def get_cleaning_config(self) -> CleaningConfig:
        """Get current cleaning configuration."""
        return CleaningConfig(
            remove_duplicates=self.option_vars["remove_duplicates"].get(),
            remove_empty_rows=self.option_vars["remove_empty_rows"].get(),
            remove_empty_columns=self.option_vars["remove_empty_columns"].get(
            ),
            trim_spaces=self.option_vars["trim_spaces"].get(),
            normalize_column_names=self.option_vars["normalize_column_names"].get(
            ),
            title_case_cells=self.option_vars["title_case_cells"].get()
        )

    def update_progress(self, value: int) -> None:
        """Update progress bar."""
        self.progress_var.set(value)
        self.progress_label.configure(text=f"Processing... {value}%")
        self.root.update_idletasks()

    def update_status(self, message: str) -> None:
        """Update status bar message."""
        self.status_label.configure(text=message)
        self.root.update_idletasks()

    def select_and_clean_file(self) -> None:
        """Select and clean Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Excel file to clean",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )

        if not file_path:
            return

        config = self.get_cleaning_config()

        # Check if any options are selected
        if not any([
            config.remove_duplicates, config.remove_empty_rows,
            config.remove_empty_columns, config.trim_spaces,
            config.normalize_column_names, config.title_case_cells
        ]):
            messagebox.showwarning(
                "No Options Selected",
                "Please select at least one cleaning option before proceeding."
            )
            return

        try:
            self.clean_btn.configure(state="disabled")
            self.update_status("Cleaning in progress...")
            self.progress_var.set(0)

            backup_path, cleaned_path, stats = self.cleaner.clean_excel_file(
                file_path, config, self.update_progress
            )

            self.progress_var.set(100)
            self.progress_label.configure(
                text="Cleaning completed successfully!")

            # Show results
            self.show_results(backup_path, cleaned_path, stats)

            self.update_status("Ready")

        except Exception as e:
            messagebox.showerror(
                "Error", f"An error occurred while cleaning the file:\\n\\n{str(e)}")
            self.update_status("Error occurred")
            self.progress_var.set(0)
            self.progress_label.configure(
                text="Ready to clean your Excel files")

        finally:
            self.clean_btn.configure(state="normal")

    def show_results(self, backup_path: str, cleaned_path: str, stats: dict) -> None:
        """Show cleaning results in a dialog."""
        result_window = tk.Toplevel(self.root)
        result_window.title("Cleaning Results")
        result_window.geometry("500x400")
        result_window.transient(self.root)
        result_window.grab_set()

        # Configure scrollable text
        text_frame = tk.Frame(result_window)
        text_frame.pack(fill="both", expand=True, padx=20, pady=20)

        text_widget = tk.Text(text_frame, wrap="word", font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(
            text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)

        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Format results
        result_text = f"""EXCEL CLEANING RESULTS
{'='*50}

FILES:
• Original file: {os.path.basename(backup_path.replace('_backup_', '').replace('.xlsx', ''))}
• Backup created: {os.path.basename(backup_path)}
• Cleaned file: {os.path.basename(cleaned_path)}

STATISTICS:
Original Data:
• Rows: {stats['original']['rows']:,}
• Columns: {stats['original']['columns']:,}
• Total cells: {stats['original']['cells']:,}
• Empty cells: {stats['original']['empty_cells']:,}

Final Data:
• Rows: {stats['final']['rows']:,}
• Columns: {stats['final']['columns']:,}
• Total cells: {stats['final']['cells']:,}
• Empty cells: {stats['final']['empty_cells']:,}

OPERATIONS APPLIED:
"""

        for operation in stats['operations_applied']:
            result_text += f"• {operation.replace('_', ' ').title()}\\n"

        result_text += f"""
SUMMARY:
• Rows changed: {stats['original']['rows'] - stats['final']['rows']:,}
• Columns changed: {stats['original']['columns'] - stats['final']['columns']:,}
• Empty cells reduced: {stats['original']['empty_cells'] - stats['final']['empty_cells']:,}

The original file has been backed up and the cleaned version is ready for use.
"""

        text_widget.insert("1.0", result_text)
        text_widget.configure(state="disabled")

        # Close button
        tk.Button(
            result_window,
            text="Close",
            command=result_window.destroy,
            font=("Segoe UI", 10),
            width=15
        ).pack(pady=10)

    def save_settings(self) -> None:
        """Save current settings to file."""
        config = {
            'theme': self.current_theme,
            'cleaning_options': {key: var.get() for key, var in self.option_vars.items()}
        }

        file_path = filedialog.asksaveasfilename(
            title="Save Settings",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if file_path:
            try:
                with open(file_path, 'w') as f:
                    json.dump(config, f, indent=2)
                self.update_status(
                    f"Settings saved to {os.path.basename(file_path)}")
                messagebox.showinfo("Success", "Settings saved successfully!")
            except Exception as e:
                messagebox.showerror(
                    "Error", f"Failed to save settings:\\n{str(e)}")

    def load_settings(self) -> None:
        """Load settings from file."""
        file_path = filedialog.askopenfilename(
            title="Load Settings",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if file_path:
            try:
                with open(file_path, 'r') as f:
                    config = json.load(f)

                # Apply theme
                if 'theme' in config and config['theme'] in self.themes:
                    self.current_theme = config['theme']
                    self.theme_var.set(self.current_theme)
                    self.apply_theme()

                # Apply cleaning options
                if 'cleaning_options' in config:
                    for key, value in config['cleaning_options'].items():
                        if key in self.option_vars:
                            self.option_vars[key].set(value)

                self.update_status(
                    f"Settings loaded from {os.path.basename(file_path)}")
                messagebox.showinfo("Success", "Settings loaded successfully!")

            except Exception as e:
                messagebox.showerror(
                    "Error", f"Failed to load settings:\\n{str(e)}")

    def view_log(self) -> None:
        """View application log."""
        log_window = tk.Toplevel(self.root)
        log_window.title("Application Log")
        log_window.geometry("600x400")
        log_window.transient(self.root)

        text_frame = tk.Frame(log_window)
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)

        text_widget = tk.Text(text_frame, wrap="word", font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(
            text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)

        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        try:
            with open("excel_cleaner.log", "r") as f:
                log_content = f.read()
            text_widget.insert("1.0", log_content)
        except FileNotFoundError:
            text_widget.insert("1.0", "No log file found.")

        text_widget.configure(state="disabled")

        tk.Button(
            log_window,
            text="Close",
            command=log_window.destroy,
            font=("Segoe UI", 10)
        ).pack(pady=5)

    def show_help(self) -> None:
        """Show help dialog."""
        help_window = tk.Toplevel(self.root)
        help_window.title("Excel Cleaner Pro - Help")
        help_window.geometry("550x600")
        help_window.transient(self.root)
        help_window.grab_set()

        # Create notebook for tabs
        notebook = ttk.Notebook(help_window)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # General help tab
        general_frame = tk.Frame(notebook)
        notebook.add(general_frame, text="General")

        help_text = tk.Text(general_frame, wrap="word", font=("Segoe UI", 10))
        help_scrollbar = ttk.Scrollbar(
            general_frame, orient="vertical", command=help_text.yview)
        help_text.configure(yscrollcommand=help_scrollbar.set)

        help_text.pack(side="left", fill="both", expand=True, padx=(0, 5))
        help_scrollbar.pack(side="right", fill="y")

        help_content = """EXCEL CLEANER PRO - HELP GUIDE

OVERVIEW:
Excel Cleaner Pro is a professional tool for cleaning and standardizing Excel data files. It provides various cleaning operations to improve data quality and consistency.

CLEANING OPTIONS:

• Remove Duplicate Rows
  Identifies and removes identical rows from your data, keeping only unique entries.

• Remove Empty Rows  
  Removes rows that contain no data in any column.

• Remove Empty Columns
  Removes columns that contain no data in any row.

• Trim Whitespace
  Removes leading and trailing spaces from all text cells, cleaning up formatting issues.

• Normalize Column Names
  Standardizes column headers by converting them to Title Case and replacing underscores with spaces.

• Title Case Text
  Converts all text cells to Title Case format (e.g., "john smith" becomes "John Smith").

USAGE:
1. Select one or more cleaning options
2. Click "Select & Clean Excel File" 
3. Choose your Excel file (.xlsx or .xls)
4. The application will create a backup and generate a cleaned version
5. Review the results in the completion dialog

KEYBOARD SHORTCUTS:
• Ctrl+O: Select and clean file
• Ctrl+H: Show this help
• Ctrl+S: Save settings
• Ctrl+L: Load settings  
• F1: Show help

FEATURES:
• Automatic backup creation with timestamps
• Detailed operation logging
• Multiple professional themes
• Save/load settings configurations
• Comprehensive statistics and reporting

The application creates detailed logs of all operations and provides comprehensive statistics about the cleaning process.
"""

        help_text.insert("1.0", help_content)
        help_text.configure(state="disabled")

        # Keyboard shortcuts tab
        shortcuts_frame = tk.Frame(notebook)
        notebook.add(shortcuts_frame, text="Shortcuts")

        shortcuts_text = tk.Text(
            shortcuts_frame, wrap="word", font=("Consolas", 10))
        shortcuts_scroll = ttk.Scrollbar(
            shortcuts_frame, orient="vertical", command=shortcuts_text.yview)
        shortcuts_text.configure(yscrollcommand=shortcuts_scroll.set)

        shortcuts_text.pack(side="left", fill="both", expand=True, padx=(0, 5))
        shortcuts_scroll.pack(side="right", fill="y")

        shortcuts_content = """KEYBOARD SHORTCUTS

File Operations:
  Ctrl+O    Select and clean Excel file
  
Settings:
  Ctrl+S    Save current settings to file
  Ctrl+L    Load settings from file
  
Help & Information:
  Ctrl+H    Show help dialog
  F1        Show help dialog
  
Quick Actions:
  You can also use the mouse to:
  • Click "Select All" to enable all cleaning options
  • Click "Clear All" to disable all cleaning options
  • Use the theme dropdown to change appearance
  • Click links in the footer to visit GitHub or support page

Tips:
• Always backup important files before cleaning
• The application automatically creates timestamped backups
• Use the log viewer to track all operations
• Save your preferred settings for quick reuse
"""

        shortcuts_text.insert("1.0", shortcuts_content)
        shortcuts_text.configure(state="disabled")

        # Close button
        tk.Button(
            help_window,
            text="Close",
            command=help_window.destroy,
            font=("Segoe UI", 10),
            width=15
        ).pack(pady=10)

    def run(self) -> None:
        """Start the application."""
        self.update_status("Excel Cleaner Pro ready")
        self.root.mainloop()


def main():
    """Main application entry point."""
    app = ExcelCleanerGUI()
    app.run()


if __name__ == "__main__":
    main()
