#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Exam Scheduling Generator Application
-------------------------------------
A comprehensive GUI application for teachers to generate and manage exam schedules.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, font
import mysql.connector
from mysql.connector import Error
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime, timedelta
from ortools.sat.python import cp_model
import json
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import base64
from PIL import Image, ImageTk
import io
import configparser

# Application constants
APP_TITLE = "Exam Scheduling Generator"
APP_VERSION = "2.0.0"
CONFIG_FILE = "config.ini"

class ExamSchedulerApp:
    """Main application class for the Exam Scheduling Generator."""
    
    def __init__(self, root):
        """Initialize the application."""
        self.root = root
        self.root.title(f"{APP_TITLE} v{APP_VERSION}")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # Set application icon
        # self.root.iconbitmap("icon.ico")  # Uncomment and add icon file if available
        
        # Initialize database
        self.init_database()
        
        # Create status bar variable
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        
        # Apply theme from config if available
        self.apply_theme()
        
        # Load saved font preference if available
        self.load_font_preference()
        
        # Create main UI components
        self.create_menu()
        self.create_main_frame()
        
        # Set default tab
        self.notebook.select(0)
        
    def init_database(self):
        """Initialize the MySQL database and create tables if they don't exist."""
        try:
            # Check if config file exists, if not create it with default values
            if not os.path.exists(CONFIG_FILE):
                self.create_default_config()
            
            # Read configuration
            config = configparser.ConfigParser()
            config.read(CONFIG_FILE)
            
            # Get database configuration
            self.db_config = {
                'host': config.get('Database', 'host'),
                'user': config.get('Database', 'user'),
                'password': config.get('Database', 'password'),
                'database': config.get('Database', 'database'),
                'autocommit': True  # Enable autocommit to avoid transaction issues
            }
            
            # Connect to MySQL server without specifying a database first
            try:
                print("Connecting to MySQL server...")
                self.conn = mysql.connector.connect(
                    host=self.db_config['host'],
                    user=self.db_config['user'],
                    password=self.db_config['password'],
                    autocommit=True
                )
                self.cursor = self.conn.cursor(buffered=True)  # Use buffered cursor to avoid unread result errors
            except Error as e:
                messagebox.showerror("Database Connection Error", 
                                   f"Failed to connect to MySQL server: {str(e)}\n\nPlease check your database settings in {CONFIG_FILE}")
                sys.exit(1)
            
            # Create database if it doesn't exist
            print(f"Creating database if it doesn't exist: {self.db_config['database']}")
            self.cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{self.db_config['database']}`")
            self.cursor.execute(f"USE `{self.db_config['database']}`")
            
            # Close the initial connection and reconnect with the database specified
            print("Reconnecting with database specified...")
            self.cursor.close()
            self.conn.close()
            
            # Connect to the specific database
            self.conn = mysql.connector.connect(**self.db_config)
            self.cursor = self.conn.cursor(buffered=True)  # Use buffered cursor to avoid unread result errors
            
            # Create subjects table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS subjects (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    code VARCHAR(20) NOT NULL,
                    name VARCHAR(100) NOT NULL,
                    type VARCHAR(20) NOT NULL,
                    semester VARCHAR(20) NOT NULL,
                    difficulty VARCHAR(20) NOT NULL,
                    duration INT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                )
            ''')
            
            # Create rooms table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS rooms (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    name VARCHAR(50) NOT NULL,
                    type VARCHAR(20) NOT NULL,
                    capacity INT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                )
            ''')
            
            # Create schedules table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS schedules (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    name VARCHAR(100) NOT NULL,
                    semester VARCHAR(20) NOT NULL,
                    exam_type VARCHAR(20) NOT NULL,
                    start_date DATE NOT NULL,
                    config TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                )
            ''')
            
            # Create schedule_items table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS schedule_items (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    schedule_id INT NOT NULL,
                    subject_id INT NOT NULL,
                    room_id INT NOT NULL,
                    exam_date DATE NOT NULL,
                    start_time VARCHAR(20) NOT NULL,
                    end_time VARCHAR(20) NOT NULL,
                    FOREIGN KEY (schedule_id) REFERENCES schedules (id) ON DELETE CASCADE,
                    FOREIGN KEY (subject_id) REFERENCES subjects (id) ON DELETE CASCADE,
                    FOREIGN KEY (room_id) REFERENCES rooms (id) ON DELETE CASCADE
                )
            ''')
            
            self.conn.commit()
            
            # Verify that tables were created successfully
            self.verify_database_tables()
            
        except Error as e:
            messagebox.showerror("Database Error", f"Failed to initialize database: {str(e)}")
            sys.exit(1)
    
    def ensure_connection(self):
        """Ensure that the database connection is active, reconnect if needed."""
        try:
            # Check if connection is active
            if hasattr(self, 'conn') and self.conn.is_connected():
                return True
                
            # Reconnect to the database
            print("Reconnecting to database...")
            self.conn = mysql.connector.connect(**self.db_config)
            self.cursor = self.conn.cursor(buffered=True)
            return True
        except Error as e:
            print(f"Database connection error: {str(e)}")
            messagebox.showerror("Database Error", f"Failed to connect to database: {str(e)}")
            return False
            
    def verify_database_tables(self):
        """Verify that all required tables exist in the database and create them if they don't."""
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return False
                
            # Check if subjects table exists
            self.cursor.execute("SHOW TABLES LIKE 'subjects'")
            if not self.cursor.fetchone():
                # Create subjects table
                self.cursor.execute('''
                    CREATE TABLE IF NOT EXISTS subjects (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        code VARCHAR(20) NOT NULL,
                        name VARCHAR(100) NOT NULL,
                        type VARCHAR(20) NOT NULL,
                        semester VARCHAR(20) NOT NULL,
                        difficulty VARCHAR(20) NOT NULL,
                        duration INT NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                    )
                ''')
                
            # Check if rooms table exists
            self.cursor.execute("SHOW TABLES LIKE 'rooms'")
            if not self.cursor.fetchone():
                # Create rooms table
                self.cursor.execute('''
                    CREATE TABLE IF NOT EXISTS rooms (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        name VARCHAR(50) NOT NULL,
                        type VARCHAR(20) NOT NULL,
                        capacity INT NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                    )
                ''')
                
            # Check if schedules table exists
            self.cursor.execute("SHOW TABLES LIKE 'schedules'")
            if not self.cursor.fetchone():
                # Create schedules table
                self.cursor.execute('''
                    CREATE TABLE IF NOT EXISTS schedules (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        name VARCHAR(100) NOT NULL,
                        semester VARCHAR(20) NOT NULL,
                        exam_type VARCHAR(20) NOT NULL,
                        start_date DATE NOT NULL,
                        config TEXT NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
                    )
                ''')
                
            # Check if schedule_items table exists
            self.cursor.execute("SHOW TABLES LIKE 'schedule_items'")
            if not self.cursor.fetchone():
                # Create schedule_items table
                self.cursor.execute('''
                    CREATE TABLE IF NOT EXISTS schedule_items (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        schedule_id INT NOT NULL,
                        subject_id INT NOT NULL,
                        room_id INT NOT NULL,
                        exam_date DATE NOT NULL,
                        start_time VARCHAR(20) NOT NULL,
                        end_time VARCHAR(20) NOT NULL,
                        FOREIGN KEY (schedule_id) REFERENCES schedules (id) ON DELETE CASCADE,
                        FOREIGN KEY (subject_id) REFERENCES subjects (id) ON DELETE CASCADE,
                        FOREIGN KEY (room_id) REFERENCES rooms (id) ON DELETE CASCADE
                    )
                ''')
                
            self.conn.commit()
            return True
            
        except Error as e:
            messagebox.showerror("Database Error", f"Failed to verify database tables: {str(e)}")
            print(f"Database connection error: {str(e)}")
            
            # Connection is closed or broken, try to reconnect
            try:
                # Close the existing connection if it exists
                try:
                    if hasattr(self, 'cursor') and self.cursor:
                        self.cursor.close()
                    if hasattr(self, 'conn') and self.conn:
                        self.conn.close()
                except Error as e:
                    print(f"Error closing existing connection: {str(e)}")  # Log but continue
                    
                # Create a new connection with the database specified
                print("Attempting to reconnect to database...")
                self.conn = mysql.connector.connect(**self.db_config)
                self.cursor = self.conn.cursor(buffered=True)  # Use buffered cursor
                
                # Test the new connection
                self.cursor.execute("SELECT 1")
                # Make sure to fetch the result to avoid 'Unread result found' errors
                self.cursor.fetchall()
                print("Successfully reconnected to database")
                return True
                
            except Error as e:
                error_msg = f"Failed to reconnect to database: {str(e)}"
                print(error_msg)
                self.status_var.set(error_msg)
                return False
    
    def create_default_config(self):
        """Create a default configuration file for database connection."""
        config = configparser.ConfigParser()
        config['Database'] = {
            'host': 'localhost',
            'user': 'root',
            'password': '',
            'database': 'exam_scheduler'
        }
        
        # Add default application settings
        config['Application'] = {
            'theme': 'Default',
            'autosave_interval': '5',
            'skip_sundays': 'True'
        }
        
        with open(CONFIG_FILE, 'w') as configfile:
            config.write(configfile)
        
        messagebox.showinfo(
            "Configuration Created", 
            f"A default configuration file has been created at {os.path.abspath(CONFIG_FILE)}. \n\n"
            "Please update it with your MySQL database credentials before proceeding."
        )
    
    def create_menu(self):
        """Create the application menu bar."""
        self.menu_bar = tk.Menu(self.root)
        
        # File menu
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="New Schedule", command=self.new_schedule, accelerator="Ctrl+N")
        file_menu.add_command(label="Open Schedule", command=self.open_schedule, accelerator="Ctrl+O")
        file_menu.add_command(label="Save Schedule", command=self.save_schedule, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Import Subjects", command=self.import_subjects, accelerator="Ctrl+I")
        file_menu.add_command(label="Export Subjects", command=self.export_subjects, accelerator="Ctrl+E")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit, accelerator="Alt+F4")
        self.menu_bar.add_cascade(label="File", menu=file_menu, underline=0)
        
        # Edit menu
        edit_menu = tk.Menu(self.menu_bar, tearoff=0)
        edit_menu.add_command(label="Preferences", command=self.show_preferences, accelerator="Ctrl+P")
        
        # Add separator before font options
        edit_menu.add_separator()
        
        # Create font submenu
        font_menu = tk.Menu(edit_menu, tearoff=0)
        
        # Get available fonts
        self.available_fonts = sorted(list(font.families()))
        # Filter out some special fonts that might cause display issues
        self.available_fonts = [f for f in self.available_fonts if not f.startswith('@') and not f.startswith('HoloLens')]
        
        # Store current font
        self.current_font = tk.StringVar()
        self.current_font.set(font.nametofont("TkDefaultFont").actual()["family"])
        
        # Add common fonts at the top for easy access
        common_fonts = ["Arial", "Calibri", "Cambria", "Courier New", "Georgia", "Helvetica", "Tahoma", "Times New Roman", "Verdana"]
        for font_name in common_fonts:
            if font_name in self.available_fonts:
                font_menu.add_radiobutton(label=font_name, variable=self.current_font, value=font_name, command=self.change_font)
        
        # Add separator before showing all fonts
        font_menu.add_separator()
        
        # Add "More Fonts..." option to open font chooser dialog
        font_menu.add_command(label="More Fonts...", command=self.show_font_dialog, accelerator="Ctrl+F")
        
        # Add font submenu to Edit menu
        edit_menu.add_cascade(label="Font", menu=font_menu)
        
        self.menu_bar.add_cascade(label="Edit", menu=edit_menu, underline=0)
        
        # View menu
        view_menu = tk.Menu(self.menu_bar, tearoff=0)
        view_menu.add_command(label="List View", command=self.show_list_view, accelerator="Ctrl+L")
        view_menu.add_command(label="Calendar View", command=self.show_calendar_view, accelerator="Ctrl+D")
        
        # Create zoom submenu
        zoom_menu = tk.Menu(view_menu, tearoff=0)
        zoom_menu.add_command(label="Zoom In", command=self.zoom_in, accelerator="Ctrl+Plus")
        zoom_menu.add_command(label="Zoom Out", command=self.zoom_out, accelerator="Ctrl+Minus")
        zoom_menu.add_command(label="Restore Default Zoom", command=self.reset_zoom, accelerator="Ctrl+0")
        view_menu.add_cascade(label="Zoom", menu=zoom_menu)
        
        # Add status bar toggle
        self.show_status_bar = tk.BooleanVar(value=True)
        view_menu.add_checkbutton(label="Show Status Bar", variable=self.show_status_bar, command=self.toggle_status_bar, accelerator="Ctrl+B")
        
        self.menu_bar.add_cascade(label="View", menu=view_menu, underline=0)
        
        # Tools menu
        tools_menu = tk.Menu(self.menu_bar, tearoff=0)
        
        # System tools
        tools_menu.add_command(label="Notepad", command=self.open_notepad, accelerator="Ctrl+Alt+N")
        tools_menu.add_command(label="Calculator", command=self.open_calculator, accelerator="Ctrl+Alt+C")
        tools_menu.add_command(label="File Explorer", command=self.open_file_explorer, accelerator="Ctrl+Alt+E")
        
        # Add separator before other tools
        tools_menu.add_separator()
        
        # Office tools
        tools_menu.add_command(label="Word", command=self.open_word, accelerator="Ctrl+Alt+W")
        tools_menu.add_command(label="Excel", command=self.open_excel, accelerator="Ctrl+Alt+X")
        
        # Add separator before web tools
        tools_menu.add_separator()
        
        # Web tools
        tools_menu.add_command(label="Web Browser", command=self.open_browser, accelerator="Ctrl+Alt+B")
        tools_menu.add_command(label="Email Client", command=self.open_email, accelerator="Ctrl+Alt+M")
        
        self.menu_bar.add_cascade(label="Tools", menu=tools_menu, underline=0)
        
        # Bind keyboard shortcuts for zoom
        self.root.bind("<Control-plus>", lambda event: self.zoom_in())
        self.root.bind("<Control-minus>", lambda event: self.zoom_out())
        self.root.bind("<Control-0>", lambda event: self.reset_zoom())
        
        # Help menu
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        help_menu.add_command(label="Help Contents", command=self.show_help, accelerator="F1")
        help_menu.add_command(label="About", command=self.show_about, accelerator="F2")
        self.menu_bar.add_cascade(label="Help", menu=help_menu, underline=0)
        
        # Refresh menu
        self.menu_bar.add_command(label="Refresh", command=self.refresh_application, accelerator="F5")
        
        self.root.config(menu=self.menu_bar)
        
        # Bind all keyboard shortcuts
        # File menu shortcuts
        self.root.bind("<Control-n>", lambda event: self.new_schedule())
        self.root.bind("<Control-o>", lambda event: self.open_schedule())
        self.root.bind("<Control-s>", lambda event: self.save_schedule())
        self.root.bind("<Control-i>", lambda event: self.import_subjects())
        self.root.bind("<Control-e>", lambda event: self.export_subjects())
        
        # Edit menu shortcuts
        self.root.bind("<Control-p>", lambda event: self.show_preferences())
        self.root.bind("<Control-f>", lambda event: self.show_font_dialog())
        
        # View menu shortcuts
        self.root.bind("<Control-l>", lambda event: self.show_list_view())
        self.root.bind("<Control-d>", lambda event: self.show_calendar_view())
        self.root.bind("<Control-b>", lambda event: self.toggle_status_bar())
        
        # Tools menu shortcuts
        self.root.bind("<Control-Alt-n>", lambda event: self.open_notepad())
        self.root.bind("<Control-Alt-c>", lambda event: self.open_calculator())
        self.root.bind("<Control-Alt-e>", lambda event: self.open_file_explorer())
        self.root.bind("<Control-Alt-w>", lambda event: self.open_word())
        self.root.bind("<Control-Alt-x>", lambda event: self.open_excel())
        self.root.bind("<Control-Alt-b>", lambda event: self.open_browser())
        self.root.bind("<Control-Alt-m>", lambda event: self.open_email())
        
        # Help menu shortcuts
        self.root.bind("<F1>", lambda event: self.show_help())
        self.root.bind("<F2>", lambda event: self.show_about())
        
        # Refresh shortcut
        self.root.bind("<F5>", lambda event: self.refresh_application())
    
    def create_main_frame(self):
        """Create the main application frame with notebook tabs."""
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook with tabs
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self.dashboard_tab = ttk.Frame(self.notebook)
        self.subjects_tab = ttk.Frame(self.notebook)
        self.rooms_tab = ttk.Frame(self.notebook)
        self.generator_tab = ttk.Frame(self.notebook)
        self.schedule_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.dashboard_tab, text="Dashboard")
        self.notebook.add(self.subjects_tab, text="Subjects")
        self.notebook.add(self.rooms_tab, text="Rooms")
        self.notebook.add(self.generator_tab, text="Generate Schedule")
        self.notebook.add(self.schedule_tab, text="View Schedule")
        
        # Store tab indices for easy reference
        self.tab_indices = {
            "dashboard": 0,
            "subjects": 1,
            "rooms": 2,
            "generator": 3,
            "schedule": 4
        }
        
        # Initialize each tab
        self.init_dashboard_tab()
        self.init_subjects_tab()
        self.init_rooms_tab()
        self.init_generator_tab()
        self.init_schedule_tab()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Update dashboard counts after all tabs are initialized
        # This ensures database connection is fully established
        self.root.after(500, self.update_dashboard_counts)
    
    def init_dashboard_tab(self):
        """Initialize the dashboard tab with summary information."""
        # Dashboard header
        header_frame = ttk.Frame(self.dashboard_tab)
        header_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(header_frame, text="Exam Scheduler Dashboard - Developed By Krutarth Raychura", font=("Helvetica", 16, "bold")).pack()
        ttk.Separator(self.dashboard_tab, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # Dashboard content
        content_frame = ttk.Frame(self.dashboard_tab)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Create a grid layout for dashboard widgets
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(0, weight=1)
        content_frame.rowconfigure(1, weight=1)
        
        # Summary cards
        self.create_summary_card(content_frame, 0, 0, "Subjects", "subjects_count", "Manage subjects for exams")
        self.create_summary_card(content_frame, 0, 1, "Rooms", "rooms_count", "Manage exam rooms and labs")
        self.create_summary_card(content_frame, 1, 0, "Schedules", "schedules_count", "View and manage exam schedules")
        self.create_summary_card(content_frame, 1, 1, "Generate", "generate", "Create a new exam schedule")
        
        # Update summary counts - we'll do this after all tabs are initialized
        # to ensure database connection is fully established
        # self.update_dashboard_counts() - moved to end of create_main_frame
    
    def create_summary_card(self, parent, row, col, title, card_id, description):
        """Create a summary card widget for the dashboard."""
        card = ttk.Frame(parent, relief=tk.RAISED, borderwidth=1)
        card.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        
        # Card title
        ttk.Label(card, text=title, font=("Helvetica", 14, "bold")).pack(pady=(10, 5))
        
        # Card count/icon
        if card_id == "generate":
            # For the generate card, add a button instead of a count
            generate_btn = ttk.Button(card, text="Create New Schedule", 
                                     command=lambda: self.notebook.select(3))
            generate_btn.pack(pady=20)
        else:
            # For count cards, add a count label
            count_var = tk.StringVar()
            count_var.set("0")
            setattr(self, f"{card_id}_var", count_var)
            ttk.Label(card, textvariable=count_var, font=("Helvetica", 24)).pack(pady=10)
        
        # Card description
        ttk.Label(card, text=description).pack(pady=(5, 10))
        
        # Add action button
        if card_id == "subjects_count":
            ttk.Button(card, text="Manage Subjects", 
                      command=lambda: self.notebook.select(1)).pack(pady=(0, 10))
        elif card_id == "rooms_count":
            ttk.Button(card, text="Manage Rooms", 
                      command=lambda: self.notebook.select(2)).pack(pady=(0, 10))
        elif card_id == "schedules_count":
            ttk.Button(card, text="View Schedules", 
                      command=lambda: self.notebook.select(4)).pack(pady=(0, 10))
    
    def update_dashboard_counts(self):
        """Update the count displays on the dashboard."""
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                # If connection failed, try again after a delay
                self.root.after(1000, self.update_dashboard_counts)
                return
                
            # Count subjects
            self.cursor.execute("SELECT COUNT(*) FROM subjects")
            result = self.cursor.fetchone()
            subjects_count = result[0] if result else 0
            self.subjects_count_var.set(str(subjects_count))
            
            # Count rooms
            self.cursor.execute("SELECT COUNT(*) FROM rooms")
            result = self.cursor.fetchone()
            rooms_count = result[0] if result else 0
            self.rooms_count_var.set(str(rooms_count))
            
            # Count schedules
            self.cursor.execute("SELECT COUNT(*) FROM schedules")
            result = self.cursor.fetchone()
            schedules_count = result[0] if result else 0
            self.schedules_count_var.set(str(schedules_count))
            
            # Update status to show success
            self.status_var.set("Dashboard updated successfully")
            
        except Error as e:
            self.status_var.set(f"Error updating dashboard: {e}")
            # Try again after a delay
            self.root.after(2000, self.update_dashboard_counts)
    
    def init_subjects_tab(self):
        """Initialize the subjects management tab."""
        # Create frames for the subjects tab
        control_frame = ttk.Frame(self.subjects_tab)
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        list_frame = ttk.Frame(self.subjects_tab)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add subject button
        add_btn = ttk.Button(control_frame, text="Add Subject", command=self.show_add_subject_dialog)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # Edit subject button
        edit_btn = ttk.Button(control_frame, text="Edit Subject", command=self.edit_selected_subject)
        edit_btn.pack(side=tk.LEFT, padx=5)
        
        # Delete subject button
        delete_btn = ttk.Button(control_frame, text="Delete Subject", command=self.delete_selected_subject)
        delete_btn.pack(side=tk.LEFT, padx=5)
        
        # Import/Export buttons
        import_btn = ttk.Button(control_frame, text="Import", command=self.import_subjects)
        import_btn.pack(side=tk.RIGHT, padx=5)
        
        export_pdf_btn = ttk.Button(control_frame, text="Export to PDF", command=self.export_subjects_pdf)
        export_pdf_btn.pack(side=tk.RIGHT, padx=5)
        
        export_btn = ttk.Button(control_frame, text="Export", command=self.export_subjects)
        export_btn.pack(side=tk.RIGHT, padx=5)
        
        # Search frame
        search_frame = ttk.Frame(control_frame)
        search_frame.pack(side=tk.RIGHT, padx=20)
        
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT)
        self.subject_search_var = tk.StringVar()
        self.subject_search_var.trace("w", self.filter_subjects)
        search_entry = ttk.Entry(search_frame, textvariable=self.subject_search_var, width=20)
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Filter by semester
        ttk.Label(search_frame, text="Semester:").pack(side=tk.LEFT, padx=(10, 0))
        self.semester_filter_var = tk.StringVar()
        self.semester_filter_var.set("All")
        self.semester_filter_var.trace("w", self.filter_subjects)
        semester_combo = ttk.Combobox(search_frame, textvariable=self.semester_filter_var, width=10)
        semester_combo.pack(side=tk.LEFT, padx=5)
        
        # Filter by type
        ttk.Label(search_frame, text="Type:").pack(side=tk.LEFT, padx=(10, 0))
        self.type_filter_var = tk.StringVar()
        self.type_filter_var.set("All")
        self.type_filter_var.trace("w", self.filter_subjects)
        type_combo = ttk.Combobox(search_frame, textvariable=self.type_filter_var, 
                                 values=["All", "Theory", "Practical"], width=10)
        type_combo.pack(side=tk.LEFT, padx=5)
        
        # Create treeview for subjects list
        columns = ("id", "code", "name", "type", "semester", "difficulty", "duration")
        self.subjects_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        # Define column headings
        self.subjects_tree.heading("id", text="ID")
        self.subjects_tree.heading("code", text="Course Code")
        self.subjects_tree.heading("name", text="Subject Name")
        self.subjects_tree.heading("type", text="Type")
        self.subjects_tree.heading("semester", text="Semester")
        self.subjects_tree.heading("difficulty", text="Difficulty")
        self.subjects_tree.heading("duration", text="Duration (min)")
        
        # Define column widths
        self.subjects_tree.column("id", width=50)
        self.subjects_tree.column("code", width=100)
        self.subjects_tree.column("name", width=200)
        self.subjects_tree.column("type", width=100)
        self.subjects_tree.column("semester", width=100)
        self.subjects_tree.column("difficulty", width=100)
        self.subjects_tree.column("duration", width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.subjects_tree.yview)
        self.subjects_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack treeview and scrollbar
        self.subjects_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Double-click to edit
        self.subjects_tree.bind("<Double-1>", lambda event: self.edit_selected_subject())
        
        # Load subjects data
        self.load_subjects()
        self.update_semester_filter_options()
    
    def show_add_subject_dialog(self):
        """Show dialog to add a new subject."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Subject")
        dialog.geometry("400x400")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog on parent window
        dialog.geometry("+%d+%d" % (
            self.root.winfo_rootx() + (self.root.winfo_width() / 2) - (400 / 2),
            self.root.winfo_rooty() + (self.root.winfo_height() / 2) - (400 / 2)
        ))
        
        # Create form fields
        ttk.Label(dialog, text="Course Code:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        code_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=code_var, width=30).grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Subject Name:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=30).grid(row=1, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Type:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        type_var = tk.StringVar()
        type_var.set("Theory")
        ttk.Combobox(dialog, textvariable=type_var, values=["Theory", "Practical"], 
                    width=28, state="readonly").grid(row=2, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Semester:").grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)
        semester_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=semester_var, width=30).grid(row=3, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Difficulty:").grid(row=4, column=0, padx=10, pady=10, sticky=tk.W)
        difficulty_var = tk.StringVar()
        difficulty_var.set("Medium")
        ttk.Combobox(dialog, textvariable=difficulty_var, values=["Easy", "Medium", "Hard"], 
                    width=28, state="readonly").grid(row=4, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Duration (minutes):").grid(row=5, column=0, padx=10, pady=10, sticky=tk.W)
        duration_var = tk.IntVar()
        duration_var.set(120)
        ttk.Spinbox(dialog, from_=30, to=240, increment=30, textvariable=duration_var, 
                   width=28).grid(row=5, column=1, padx=10, pady=10)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
        
        def save_subject():
            # Validate input
            if not code_var.get() or not name_var.get() or not semester_var.get():
                messagebox.showerror("Input Error", "Course code, name, and semester are required fields.")
                return
            
            try:
                # Ensure database connection is active
                if not self.ensure_connection():
                    return
                    
                # Insert new subject into database
                self.cursor.execute('''
                    INSERT INTO subjects (code, name, type, semester, difficulty, duration)
                    VALUES (%s, %s, %s, %s, %s, %s)
                ''', (
                    code_var.get(),
                    name_var.get(),
                    type_var.get(),
                    semester_var.get(),
                    difficulty_var.get(),
                    duration_var.get()
                ))
                self.conn.commit()
                
                # Refresh subjects list
                self.load_subjects()
                self.update_semester_filter_options()
                self.update_dashboard_counts()
                
                # Close dialog
                dialog.destroy()
                
                # Show success message
                self.status_var.set(f"Subject '{name_var.get()}' added successfully.")
                
            except Error as e:
                messagebox.showerror("Database Error", f"Failed to add subject: {e}")
        
        ttk.Button(button_frame, text="Save", command=save_subject).pack(side=tk.LEFT, padx=10)
    
    def edit_selected_subject(self):
        """Edit the selected subject."""
        # Get selected item
        selected_item = self.subjects_tree.selection()
        if not selected_item:
            messagebox.showinfo("Selection", "Please select a subject to edit.")
            return
        
        # Get subject data
        subject_id = self.subjects_tree.item(selected_item[0], "values")[0]
        
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return
                
            # Fetch subject data from database
            self.cursor.execute("SELECT * FROM subjects WHERE id = %s", (subject_id,))
            subject = self.cursor.fetchone()
            
            if not subject:
                messagebox.showerror("Error", "Subject not found.")
                return
            
            # Create edit dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Edit Subject")
            dialog.geometry("400x400")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Center dialog on parent window
            dialog.geometry("+%d+%d" % (
                self.root.winfo_rootx() + (self.root.winfo_width() / 2) - (400 / 2),
                self.root.winfo_rooty() + (self.root.winfo_height() / 2) - (400 / 2)
            ))
            
            # Create form fields
            ttk.Label(dialog, text="Course Code:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
            code_var = tk.StringVar(value=subject[1])
            ttk.Entry(dialog, textvariable=code_var, width=30).grid(row=0, column=1, padx=10, pady=10)
            
            ttk.Label(dialog, text="Subject Name:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
            name_var = tk.StringVar(value=subject[2])
            ttk.Entry(dialog, textvariable=name_var, width=30).grid(row=1, column=1, padx=10, pady=10)
            
            ttk.Label(dialog, text="Type:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
            type_var = tk.StringVar(value=subject[3])
            ttk.Combobox(dialog, textvariable=type_var, values=["Theory", "Practical"], 
                        width=28, state="readonly").grid(row=2, column=1, padx=10, pady=10)
            
            ttk.Label(dialog, text="Semester:").grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)
            semester_var = tk.StringVar(value=subject[4])
            ttk.Entry(dialog, textvariable=semester_var, width=30).grid(row=3, column=1, padx=10, pady=10)
            
            ttk.Label(dialog, text="Difficulty:").grid(row=4, column=0, padx=10, pady=10, sticky=tk.W)
            difficulty_var = tk.StringVar(value=subject[5])
            ttk.Combobox(dialog, textvariable=difficulty_var, values=["Easy", "Medium", "Hard"], 
                        width=28, state="readonly").grid(row=4, column=1, padx=10, pady=10)
            
            ttk.Label(dialog, text="Duration (minutes):").grid(row=5, column=0, padx=10, pady=10, sticky=tk.W)
            duration_var = tk.IntVar(value=subject[6])
            ttk.Spinbox(dialog, from_=30, to=240, increment=30, textvariable=duration_var, 
                       width=28).grid(row=5, column=1, padx=10, pady=10)
            
            # Buttons
            button_frame = ttk.Frame(dialog)
            button_frame.grid(row=6, column=0, columnspan=2, pady=20)
            
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
            
            def update_subject():
                # Validate input
                if not code_var.get() or not name_var.get() or not semester_var.get():
                    messagebox.showerror("Input Error", "Course code, name, and semester are required fields.")
                    return
                
                try:
                    # Ensure database connection is active
                    if not self.ensure_connection():
                        return
                        
                    # Update subject in database
                    self.cursor.execute('''
                        UPDATE subjects
                        SET code = %s, name = %s, type = %s, semester = %s, difficulty = %s, duration = %s, updated_at = CURRENT_TIMESTAMP
                        WHERE id = %s
                    ''', (
                        code_var.get(),
                        name_var.get(),
                        type_var.get(),
                        semester_var.get(),
                        difficulty_var.get(),
                        duration_var.get(),
                        subject_id
                    ))
                    self.conn.commit()
                    
                    # Refresh subjects list
                    self.load_subjects()
                    self.update_semester_filter_options()
                    
                    # Close dialog
                    dialog.destroy()
                    
                    # Show success message
                    self.status_var.set(f"Subject '{name_var.get()}' updated successfully.")
                    
                except Error as e:
                    messagebox.showerror("Database Error", f"Failed to update subject: {e}")
            
            ttk.Button(button_frame, text="Save", command=update_subject).pack(side=tk.LEFT, padx=10)
            
        except Error as e:
            messagebox.showerror("Database Error", f"Failed to fetch subject: {e}")
    
    def delete_selected_subject(self):
        """Delete the selected subject."""
        # Get selected item
        selected_item = self.subjects_tree.selection()
        if not selected_item:
            messagebox.showinfo("Selection", "Please select a subject to delete.")
            return
        
        # Get subject data
        subject_id = self.subjects_tree.item(selected_item[0], "values")[0]
        subject_name = self.subjects_tree.item(selected_item[0], "values")[2]
        
        # Confirm deletion
        if not messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete '{subject_name}'?"):
            return
        
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return
                
            # Check if subject is used in any schedule
            self.cursor.execute("SELECT COUNT(*) FROM schedule_items WHERE subject_id = %s", (subject_id,))
            count = self.cursor.fetchone()[0]
            
            if count > 0:
                if not messagebox.askyesno("Warning", 
                                         f"This subject is used in {count} schedule(s). Deleting it will also remove it from those schedules. Continue?"):
                    return
                
                # Delete from schedule_items first
                self.cursor.execute("DELETE FROM schedule_items WHERE subject_id = %s", (subject_id,))
            
            # Delete subject from database
            self.cursor.execute("DELETE FROM subjects WHERE id = %s", (subject_id,))
            self.conn.commit()
            
            # Refresh subjects list
            self.load_subjects()
            self.update_semester_filter_options()
            self.update_dashboard_counts()
            
            # Show success message
            self.status_var.set(f"Subject '{subject_name}' deleted successfully.")
            
        except Error as e:
            messagebox.showerror("Database Error", f"Failed to delete subject: {e}")
    
    def load_subjects(self):
        """Load subjects from database and display in treeview."""
        # Clear existing items
        for item in self.subjects_tree.get_children():
            self.subjects_tree.delete(item)
        
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return
                
            # Fetch subjects from database - no parameters needed here
            try:
                print("Executing SELECT query for subjects...")
                self.cursor.execute("SELECT id, code, name, type, semester, difficulty, duration FROM subjects ORDER BY semester, name")
                subjects = self.cursor.fetchall()
                if subjects is None:  # Handle None result
                    subjects = []
                print(f"Found {len(subjects)} subjects")
            except Error as e:
                print(f"Error fetching subjects: {str(e)}")
                messagebox.showerror("Database Error", f"Failed to fetch subjects: {str(e)}")
                return
            
            # Filter subjects if search or filters are active
            filtered_subjects = []
            search_term = self.subject_search_var.get().lower()
            semester_filter = self.semester_filter_var.get()
            type_filter = self.type_filter_var.get()
            
            for subject in subjects:
                try:
                    # Apply search filter
                    if search_term and not (
                        search_term in str(subject[0]).lower() or  # ID
                        search_term in str(subject[1]).lower() or  # Code - ensure it's a string
                        search_term in str(subject[2]).lower() or  # Name - ensure it's a string
                        search_term in str(subject[3]).lower() or  # Type - ensure it's a string
                        search_term in str(subject[4]).lower()     # Semester - ensure it's a string
                    ):
                        continue
                    
                    # Apply semester filter
                    if semester_filter != "All" and str(subject[4]) != semester_filter:
                        continue
                    
                    # Apply type filter
                    if type_filter != "All" and str(subject[3]) != type_filter:
                        continue
                    
                    filtered_subjects.append(subject)
                except Exception as e:
                    print(f"Error filtering subject: {str(e)}")
                    continue
            
            # Add filtered subjects to treeview
            for subject in filtered_subjects:
                self.subjects_tree.insert("", tk.END, values=subject)
            
            # Update status
            self.status_var.set(f"Loaded {len(filtered_subjects)} subjects.")
            
        except Error as e:
            messagebox.showerror("Database Error", f"Failed to load subjects: {e}")
    
    def filter_subjects(self, *args):
        """Filter subjects based on search term and filters."""
        # This method is called when search term or filters change
        # Simply reload the subjects with the new filters applied
        self.load_subjects()
    
    def update_semester_filter_options(self):
        """Update the semester filter dropdown with available semesters."""
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return
                
            # Fetch unique semesters from database
            self.cursor.execute("SELECT DISTINCT semester FROM subjects ORDER BY semester")
            semesters = self.cursor.fetchall() or []
            
            # Create list of semester options
            semester_options = ["All"] + [semester[0] for semester in semesters if semester[0]]
            
            # Update combobox values
            try:
                # Find the semester filter combobox
                for widget in self.subjects_tab.winfo_children():
                    if isinstance(widget, ttk.Frame):
                        for child in widget.winfo_children():
                            if isinstance(child, ttk.Frame):
                                for grandchild in child.winfo_children():
                                    if isinstance(grandchild, ttk.Combobox):
                                        # Update the combobox values
                                        grandchild["values"] = semester_options
                                        break
            except Exception as e:
                print(f"Error updating combobox values: {str(e)}")
                # Don't show error to user for this UI update
        
        except Error as e:
            self.status_var.set(f"Error updating semester filter: {e}")
    
    def init_rooms_tab(self):
        """Initialize the rooms management tab."""
        # Create frames for the rooms tab
        control_frame = ttk.Frame(self.rooms_tab)
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        list_frame = ttk.Frame(self.rooms_tab)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add room button
        add_btn = ttk.Button(control_frame, text="Add Room", command=self.show_add_room_dialog)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # Edit room button
        edit_btn = ttk.Button(control_frame, text="Edit Room", command=self.edit_selected_room)
        edit_btn.pack(side=tk.LEFT, padx=5)
        
        # Delete room button
        delete_btn = ttk.Button(control_frame, text="Delete Room", command=self.delete_selected_room)
        delete_btn.pack(side=tk.LEFT, padx=5)
        
        # Search frame
        search_frame = ttk.Frame(control_frame)
        search_frame.pack(side=tk.RIGHT, padx=20)
        
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT)
        self.room_search_var = tk.StringVar()
        self.room_search_var.trace("w", self.filter_rooms)
        search_entry = ttk.Entry(search_frame, textvariable=self.room_search_var, width=20)
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Filter by type
        ttk.Label(search_frame, text="Type:").pack(side=tk.LEFT, padx=(10, 0))
        self.room_type_filter_var = tk.StringVar()
        self.room_type_filter_var.set("All")
        self.room_type_filter_var.trace("w", self.filter_rooms)
        type_combo = ttk.Combobox(search_frame, textvariable=self.room_type_filter_var, 
                                  values=["All", "Classroom", "Lab"], width=10, state="readonly")
        type_combo.pack(side=tk.LEFT, padx=5)
        
        # Import button (after type dropdown)
        import_btn = ttk.Button(search_frame, text="Import", command=self.import_rooms)
        import_btn.pack(side=tk.LEFT, padx=5)
        
        # Export button
        export_btn = ttk.Button(search_frame, text="Export", command=self.export_rooms)
        export_btn.pack(side=tk.LEFT, padx=5)
        
        # Export as PDF button
        export_pdf_btn = ttk.Button(search_frame, text="Export as PDF", command=self.export_rooms_pdf)
        export_pdf_btn.pack(side=tk.LEFT, padx=5)
        
        # Create treeview for rooms list
        columns = ("id", "name", "type", "capacity")
        self.rooms_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        # Define column headings
        self.rooms_tree.heading("id", text="ID")
        self.rooms_tree.heading("name", text="Room Name")
        self.rooms_tree.heading("type", text="Type")
        self.rooms_tree.heading("capacity", text="Capacity")
        
        # Define column widths
        self.rooms_tree.column("id", width=50)
        self.rooms_tree.column("name", width=200)
        self.rooms_tree.column("type", width=100)
        self.rooms_tree.column("capacity", width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.rooms_tree.yview)
        self.rooms_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack treeview and scrollbar
        self.rooms_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Double-click to edit
        self.rooms_tree.bind("<Double-1>", lambda event: self.edit_selected_room())
        
        # Load rooms data
        self.load_rooms()
    
    def show_add_room_dialog(self):
        """Show dialog to add a new room."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Room")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog on parent window
        dialog.geometry("+%d+%d" % (
            self.root.winfo_rootx() + (self.root.winfo_width() / 2) - (400 / 2),
            self.root.winfo_rooty() + (self.root.winfo_height() / 2) - (300 / 2)
        ))
        
        # Create form fields
        ttk.Label(dialog, text="Room Name:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=30).grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Type:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        type_var = tk.StringVar()
        type_var.set("Classroom")
        ttk.Combobox(dialog, textvariable=type_var, values=["Classroom", "Lab"], 
                    width=28, state="readonly").grid(row=1, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="Capacity:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        capacity_var = tk.IntVar()
        capacity_var.set(30)
        ttk.Spinbox(dialog, from_=10, to=200, increment=5, textvariable=capacity_var, 
                   width=28).grid(row=2, column=1, padx=10, pady=10)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
        
        def save_room():
            # Validate input
            if not name_var.get():
                messagebox.showerror("Input Error", "Room name is required.")
                return
            
            try:
                # Ensure database connection is active
                if not self.ensure_connection():
                    return
                    
                # Insert room into database
                self.cursor.execute('''
                    INSERT INTO rooms (name, type, capacity)
                    VALUES (%s, %s, %s)
                ''', (
                    name_var.get(),
                    type_var.get(),
                    capacity_var.get()
                ))
                self.conn.commit()
                
                # Refresh rooms list
                self.load_rooms()
                self.update_dashboard_counts()
                
                # Close dialog
                dialog.destroy()
                
                # Show success message
                self.status_var.set(f"Room '{name_var.get()}' added successfully.")
                
            except Error as e:
                messagebox.showerror("Database Error", f"Failed to add room: {e}")
        
        ttk.Button(button_frame, text="Save", command=save_room).pack(side=tk.LEFT, padx=10)
    
    def edit_selected_room(self):
        """Edit the selected room."""
        # Get selected item
        selected_item = self.rooms_tree.selection()
        if not selected_item:
            messagebox.showinfo("Selection", "Please select a room to edit.")
            return
        
        # Get room data
        room_id = self.rooms_tree.item(selected_item[0], "values")[0]
        
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return
                
            # Fetch room data from database
            self.cursor.execute("SELECT * FROM rooms WHERE id = %s", (room_id,))
            room = self.cursor.fetchone()
            
            if not room:
                messagebox.showerror("Error", "Room not found.")
                return
            
            # Create edit dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Edit Room")
            dialog.geometry("400x300")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Center dialog on parent window
            dialog.geometry("+%d+%d" % (
                self.root.winfo_rootx() + (self.root.winfo_width() / 2) - (400 / 2),
                self.root.winfo_rooty() + (self.root.winfo_height() / 2) - (300 / 2)
            ))
            
            # Create form fields
            ttk.Label(dialog, text="Room Name:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
            name_var = tk.StringVar(value=room[1])
            ttk.Entry(dialog, textvariable=name_var, width=30).grid(row=0, column=1, padx=10, pady=10)
            
            ttk.Label(dialog, text="Type:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
            type_var = tk.StringVar(value=room[2])
            ttk.Combobox(dialog, textvariable=type_var, values=["Classroom", "Lab"], 
                        width=28, state="readonly").grid(row=1, column=1, padx=10, pady=10)
            
            ttk.Label(dialog, text="Capacity:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
            capacity_var = tk.IntVar(value=room[3])
            ttk.Spinbox(dialog, from_=10, to=200, increment=5, textvariable=capacity_var, 
                       width=28).grid(row=2, column=1, padx=10, pady=10)
            
            # Buttons
            button_frame = ttk.Frame(dialog)
            button_frame.grid(row=3, column=0, columnspan=2, pady=20)
            
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
            
            def update_room():
                # Validate input
                if not name_var.get():
                    messagebox.showerror("Input Error", "Room name is required.")
                    return
                
                try:
                    # Ensure database connection is active
                    if not self.ensure_connection():
                        return
                        
                    # Update room in database
                    self.cursor.execute('''
                        UPDATE rooms
                        SET name = %s, type = %s, capacity = %s, updated_at = CURRENT_TIMESTAMP
                        WHERE id = %s
                    ''', (
                        name_var.get(),
                        type_var.get(),
                        capacity_var.get(),
                        room_id
                    ))
                    self.conn.commit()
                    
                    # Refresh rooms list
                    self.load_rooms()
                    
                    # Close dialog
                    dialog.destroy()
                    
                    # Show success message
                    self.status_var.set(f"Room '{name_var.get()}' updated successfully.")
                    
                except Error as e:
                    messagebox.showerror("Database Error", f"Failed to update room: {e}")
            
            ttk.Button(button_frame, text="Save", command=update_room).pack(side=tk.LEFT, padx=10)
            
        except Error as e:
            messagebox.showerror("Database Error", f"Failed to fetch room: {e}")
    
    def delete_selected_room(self):
        """Delete the selected room."""
        # Get selected item
        selected_item = self.rooms_tree.selection()
        if not selected_item:
            messagebox.showinfo("Selection", "Please select a room to delete.")
            return
        
        # Get room data
        room_id = self.rooms_tree.item(selected_item[0], "values")[0]
        room_name = self.rooms_tree.item(selected_item[0], "values")[1]
        
        # Confirm deletion
        if not messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete '{room_name}'?"):
            return
        
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return
                
            # Check if room is used in any schedule
            self.cursor.execute("SELECT COUNT(*) FROM schedule_items WHERE room_id = %s", (room_id,))
            count = self.cursor.fetchone()[0]
            
            if count > 0:
                if not messagebox.askyesno("Warning", 
                                         f"This room is used in {count} schedule(s). Deleting it will also remove it from those schedules. Continue?"):
                    return
                
                # Delete from schedule_items first
                self.cursor.execute("DELETE FROM schedule_items WHERE room_id = %s", (room_id,))
            
            # Delete room from database
            self.cursor.execute("DELETE FROM rooms WHERE id = %s", (room_id,))
            self.conn.commit()
            
            # Refresh rooms list
            self.load_rooms()
            self.update_dashboard_counts()
            
            # Show success message
            self.status_var.set(f"Room '{room_name}' deleted successfully.")
            
        except Error as e:
            messagebox.showerror("Database Error", f"Failed to delete room: {e}")
    
    def import_rooms(self):
        """Import rooms from a CSV file."""
        try:
            # Ask user to select a CSV file
            file_path = filedialog.askopenfilename(
                title="Import Rooms",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
            )
            
            if not file_path:
                return
            
            # Read CSV file
            df = pd.read_csv(file_path)
            
            # Check required columns
            required_columns = ["name", "type", "capacity"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                messagebox.showerror(
                    "Error",
                    f"CSV file is missing required columns: {', '.join(missing_columns)}"
                )
                return
            
            # Ensure database connection
            if not self.ensure_connection():
                return
            
            # Import rooms
            imported_count = 0
            updated_count = 0
            
            for _, row in df.iterrows():
                # Check if room already exists
                self.cursor.execute(
                    "SELECT id FROM rooms WHERE name = %s",
                    (row["name"],)
                )
                existing = self.cursor.fetchone()
                
                if existing:
                    # Update existing room
                    self.cursor.execute(
                        """
                        UPDATE rooms
                        SET type = %s, capacity = %s, updated_at = CURRENT_TIMESTAMP
                        WHERE name = %s
                        """,
                        (row["type"], row["capacity"], row["name"])
                    )
                    updated_count += 1
                else:
                    # Insert new room
                    self.cursor.execute(
                        """
                        INSERT INTO rooms (name, type, capacity)
                        VALUES (%s, %s, %s)
                        """,
                        (row["name"], row["type"], row["capacity"])
                    )
                    imported_count += 1
            
            self.conn.commit()
            self.load_rooms()
            self.update_dashboard_counts()
            messagebox.showinfo("Success", f"Imported {imported_count} new rooms and updated {updated_count} existing rooms from CSV file.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import rooms: {str(e)}")
    
    def export_subjects(self):
        """Export subjects to a CSV file."""
        try:
            # Ensure database connection
            if not self.ensure_connection():
                return
            
            # Get all subjects
            self.cursor.execute("SELECT code, name, type, semester, difficulty, duration FROM subjects ORDER BY semester, name")
            subjects = self.cursor.fetchall()
            
            if not subjects:
                messagebox.showinfo("Info", "No subjects to export.")
                return
            
            # Create DataFrame
            df = pd.DataFrame(subjects, columns=["code", "name", "type", "semester", "difficulty", "duration"])
            
            # Ask user for save location
            file_path = filedialog.asksaveasfilename(
                title="Export Subjects",
                defaultextension=".csv",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
            )
            
            if not file_path:
                return
            
            # Save to CSV
            df.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Exported {len(subjects)} subjects to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export subjects: {str(e)}")
            
    def export_rooms(self):
        """Export rooms to a CSV file."""
        try:
            # Ensure database connection
            if not self.ensure_connection():
                return
            
            # Get all rooms
            self.cursor.execute("SELECT name, type, capacity FROM rooms ORDER BY name")
            rooms = self.cursor.fetchall()
            
            if not rooms:
                messagebox.showinfo("Info", "No rooms to export.")
                return
            
            # Create DataFrame
            df = pd.DataFrame(rooms, columns=["name", "type", "capacity"])
            
            # Ask user for save location
            file_path = filedialog.asksaveasfilename(
                title="Export Rooms",
                defaultextension=".csv",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
            )
            
            if not file_path:
                return
            
            # Save to CSV
            df.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Exported {len(rooms)} rooms to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export rooms: {str(e)}")
    
    def export_subjects_pdf(self):
        """Export subjects to a PDF file."""
        try:
            # Ensure database connection
            if not self.ensure_connection():
                return
            
            # Get all subjects including ID
            self.cursor.execute("SELECT id, code, name, type, semester, difficulty, duration FROM subjects ORDER BY semester, name")
            subjects = self.cursor.fetchall()
            
            if not subjects:
                messagebox.showinfo("Info", "No subjects to export.")
                return
            
            # Ask user for save location
            file_path = filedialog.asksaveasfilename(
                title="Export Subjects as PDF",
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
            )
            
            if not file_path:
                return
            
            # Create PDF document
            from reportlab.lib.pagesizes import letter, landscape
            from reportlab.lib import colors
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            
            # Use landscape orientation for better fit of all columns
            doc = SimpleDocTemplate(file_path, pagesize=landscape(letter))
            elements = []
            
            # Add title
            styles = getSampleStyleSheet()
            title = Paragraph("Subjects List", styles["Title"])
            elements.append(title)
            elements.append(Spacer(1, 20))
            
            # Add date
            date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            date_paragraph = Paragraph(f"Generated on: {date_str}", styles["Normal"])
            elements.append(date_paragraph)
            elements.append(Spacer(1, 20))
            
            # Create table data
            data = [["ID", "Course Code", "Subject Name", "Type", "Semester", "Difficulty", "Duration (min)"]]
            for subject in subjects:
                data.append([str(subject[0]), subject[1], subject[2], subject[3], str(subject[4]), subject[5], str(subject[6])])
            
            # Create table with adjusted column widths
            table = Table(data, colWidths=[40, 80, 180, 70, 70, 80, 80])
            
            # Style the table
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                # Left-align subject names for better readability
                ('ALIGN', (2, 1), (2, -1), 'LEFT')
            ]))
            
            # Add table to elements
            elements.append(table)
            
            # Add summary and grouping information
            elements.append(Spacer(1, 20))
            
            # Count subjects by semester
            semester_counts = {}
            for subject in subjects:
                semester = subject[4]
                if semester not in semester_counts:
                    semester_counts[semester] = 0
                semester_counts[semester] += 1
            
            # Add semester summary
            summary_text = f"Total Subjects: {len(subjects)}\n\nSubjects by Semester:\n"
            for semester, count in sorted(semester_counts.items()):
                summary_text += f"Semester {semester}: {count} subjects\n"
            
            summary = Paragraph(summary_text, styles["Normal"])
            elements.append(summary)
            
            # Build PDF
            doc.build(elements)
            
            messagebox.showinfo("Success", f"Exported {len(subjects)} subjects to {file_path}")
            
        except Exception as e:
            print(f"Error exporting subjects to PDF: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to export subjects as PDF: {str(e)}")
    
    def export_rooms_pdf(self):
        """Export rooms to a PDF file."""
        try:
            # Ensure database connection
            if not self.ensure_connection():
                return
            
            # Get all rooms including ID
            self.cursor.execute("SELECT id, name, type, capacity FROM rooms ORDER BY type, name")
            rooms = self.cursor.fetchall()
            
            if not rooms:
                messagebox.showinfo("Info", "No rooms to export.")
                return
            
            # Ask user for save location
            file_path = filedialog.asksaveasfilename(
                title="Export Rooms as PDF",
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
            )
            
            if not file_path:
                return
            
            # Create PDF document
            from reportlab.lib.pagesizes import letter
            from reportlab.lib import colors
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            
            doc = SimpleDocTemplate(file_path, pagesize=letter)
            elements = []
            
            # Add title
            styles = getSampleStyleSheet()
            title = Paragraph("Rooms List", styles["Title"])
            elements.append(title)
            elements.append(Spacer(1, 20))
            
            # Add date
            date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            date_paragraph = Paragraph(f"Generated on: {date_str}", styles["Normal"])
            elements.append(date_paragraph)
            elements.append(Spacer(1, 20))
            
            # Create table data
            data = [["ID", "Room Name", "Type", "Capacity"]]
            for room in rooms:
                data.append([room[0], room[1], room[2], room[3]])
            
            # Create table with adjusted column widths to accommodate ID column
            table = Table(data, colWidths=[50, 200, 150, 100])
            
            # Style the table
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            # Add table to elements
            elements.append(table)
            
            # Add summary
            elements.append(Spacer(1, 20))
            summary = Paragraph(f"Total Rooms: {len(rooms)}", styles["Normal"])
            elements.append(summary)
            
            # Build PDF
            doc.build(elements)
            
            messagebox.showinfo("Success", f"Exported {len(rooms)} rooms to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export rooms as PDF: {str(e)}")
    
    def load_rooms(self):
        """Load rooms from database and display in treeview."""
        # Clear existing items
        for item in self.rooms_tree.get_children():
            self.rooms_tree.delete(item)
        
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return
                
            # Fetch rooms from database - no parameters needed here
            try:
                print("Executing SELECT query for rooms...")
                self.cursor.execute("SELECT id, name, type, capacity FROM rooms ORDER BY type, name")
                rooms = self.cursor.fetchall()
                if rooms is None:  # Handle None result
                    rooms = []
                print(f"Found {len(rooms)} rooms")
            except Error as e:
                print(f"Error fetching rooms: {str(e)}")
                messagebox.showerror("Database Error", f"Failed to fetch rooms: {str(e)}")
                return
            
            # Filter rooms if search or filters are active
            filtered_rooms = []
            search_term = self.room_search_var.get().lower() if hasattr(self, 'room_search_var') else ""
            type_filter = self.room_type_filter_var.get() if hasattr(self, 'room_type_filter_var') else "All"
            
            for room in rooms:
                try:
                    # Apply search filter
                    if search_term and not (
                        search_term in str(room[0]).lower() or  # ID
                        search_term in str(room[1]).lower() or  # Name - ensure it's a string
                        search_term in str(room[2]).lower()     # Type - ensure it's a string
                    ):
                        continue
                    
                    # Apply type filter
                    if type_filter != "All" and str(room[2]) != type_filter:
                        continue
                    
                    filtered_rooms.append(room)
                except Exception as e:
                    print(f"Error filtering room: {str(e)}")
                    continue
            
            # Add filtered rooms to treeview
            for room in filtered_rooms:
                self.rooms_tree.insert("", tk.END, values=room)
            
            # Update status
            self.status_var.set(f"Loaded {len(filtered_rooms)} rooms.")
            
        except Error as e:
            # Log the error for debugging
            print(f"Error loading rooms: {str(e)}")
            messagebox.showerror("Database Error", f"Failed to load rooms: {e}")
    
    def filter_rooms(self, *args):
        """Filter rooms based on search term and filters."""
        self.load_rooms()
    
    def toggle_room_selection(self, event):
        """Toggle room selection in the room selection treeview."""
        try:
            region = self.room_select_tree.identify_region(event.x, event.y)
            if region != "cell":
                return
                
            column = self.room_select_tree.identify_column(event.x)
            if column != "#5":  # The 'Selected' column is at index 5 (column #5)
                return
                
            item = self.room_select_tree.identify_row(event.y)
            if not item:
                return
                
            # Get current values
            values = self.room_select_tree.item(item, "values")
            
            # Toggle selection
            if values[4] == "No":  # The 'Selected' column is at index 4
                new_values = values[:4] + ("Yes",)
                self.room_select_tree.item(item, values=new_values)
            else:
                new_values = values[:4] + ("No",)
                self.room_select_tree.item(item, values=new_values)
                
        except Exception as e:
            print(f"Error toggling room selection: {str(e)}")
            # Don't show error message to user for this minor UI interaction
    
    def get_selected_subjects(self):
        """Get the list of selected subjects for scheduling."""
        selected_subjects = []
        
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return selected_subjects
                
            print("Getting selected subjects for scheduling...")
            
            # Use the selected_subjects list directly since it contains tuples of (id, display_string)
            if hasattr(self, 'selected_subjects') and self.selected_subjects:
                print(f"Found {len(self.selected_subjects)} subjects in the selected_subjects list")
                
                for subject_tuple in self.selected_subjects:
                    try:
                        subject_id = subject_tuple[0]  # The ID is the first element in the tuple
                        
                        # Get subject details from database
                        self.cursor.execute(
                            "SELECT id, code, name, type, semester, difficulty, duration FROM subjects WHERE id = %s",
                            (subject_id,)
                        )
                        subject = self.cursor.fetchone()
                        
                        if subject:
                            selected_subjects.append(subject)
                            print(f"Added subject to schedule: {subject[2]} (ID: {subject[0]})")
                    except Exception as e:
                        print(f"Error processing subject with ID {subject_id}: {str(e)}")
                        continue
            else:
                print("No subjects found in the selected_subjects list")
                
                # As a fallback, try to get selected subjects from the listbox
                try:
                    # Get all items from the selected subjects listbox
                    all_subjects = self.selected_subjects_list.get(0, tk.END)
                    print(f"Found {len(all_subjects)} subjects in the listbox")
                    
                    for subject_str in all_subjects:
                        try:
                            # Try different formats to extract the subject ID
                            if " - " in subject_str and "(" in subject_str:
                                # Format: "CODE - NAME (TYPE)"
                                code = subject_str.split(" - ")[0].strip()
                                self.cursor.execute(
                                    "SELECT id, code, name, type, semester, difficulty, duration FROM subjects WHERE code = %s",
                                    (code,)
                                )
                            else:
                                # Try to find any number at the beginning which might be the ID
                                import re
                                match = re.search(r'^\d+', subject_str)
                                if match:
                                    subject_id = int(match.group())
                                    self.cursor.execute(
                                        "SELECT id, code, name, type, semester, difficulty, duration FROM subjects WHERE id = %s",
                                        (subject_id,)
                                    )
                                else:
                                    # Just try to find any subject that contains this string
                                    self.cursor.execute(
                                        "SELECT id, code, name, type, semester, difficulty, duration FROM subjects WHERE name LIKE %s OR code LIKE %s LIMIT 1",
                                        (f"%{subject_str}%", f"%{subject_str}%")
                                    )
                            
                            subject = self.cursor.fetchone()
                            if subject:
                                selected_subjects.append(subject)
                                print(f"Added subject to schedule (from listbox): {subject[2]} (ID: {subject[0]})")
                        except Exception as e:
                            print(f"Error processing subject string '{subject_str}': {str(e)}")
                            continue
                except Exception as e:
                    print(f"Error getting subjects from listbox: {str(e)}")
        except Exception as e:
            print(f"Error getting selected subjects: {str(e)}")
            messagebox.showerror("Error", f"Failed to get selected subjects: {str(e)}")
        
        if not selected_subjects:
            print("No subjects were selected for scheduling")
            
        return selected_subjects
    
    def get_selected_rooms(self):
        """Get the list of selected rooms for scheduling."""
        selected_rooms = []
        
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return selected_rooms
                
            # Get all rooms from the treeview
            for item in self.room_select_tree.get_children():
                try:
                    values = self.room_select_tree.item(item, "values")
                    
                    # Check if room is selected
                    if values[4] == "Yes":  # The 'Selected' column is at index 4
                        try:
                            room_id = int(values[0])  # The 'ID' column is at index 0
                            
                            # Get room details from database
                            self.cursor.execute(
                                "SELECT id, name, type, capacity FROM rooms WHERE id = %s",
                                (room_id,)
                            )
                            room = self.cursor.fetchone()
                            
                            if room:
                                selected_rooms.append(room)
                        except (ValueError, IndexError) as e:
                            print(f"Error getting room ID: {str(e)}")
                            continue
                except Exception as e:
                    print(f"Error processing room item: {str(e)}")
                    continue
        except Exception as e:
            print(f"Error getting selected rooms: {str(e)}")
            messagebox.showerror("Error", f"Failed to get selected rooms: {str(e)}")
        
        return selected_rooms
    
    def generate_schedule(self):
        """Generate an exam schedule using constraint programming."""
        # Get selected subjects and rooms
        selected_subjects = self.get_selected_subjects()
        
        # If no subjects are selected, try to add all available subjects
        if not selected_subjects and hasattr(self, 'available_subjects') and self.available_subjects:
            print("No subjects selected, automatically adding all available subjects")
            try:
                # Add all available subjects to selected subjects
                self.selected_subjects.extend(self.available_subjects)
                self.available_subjects = []
                
                # Update both listboxes
                self.available_subjects_var.set([])
                self.selected_subjects_var.set([subject[1] for subject in self.selected_subjects])
                
                # Try to get selected subjects again
                selected_subjects = self.get_selected_subjects()
            except Exception as e:
                print(f"Error auto-selecting subjects: {str(e)}")
        
        # Get selected rooms
        selected_rooms = self.get_selected_rooms()
        
        # If no rooms are selected, try to select all rooms
        if not selected_rooms:
            print("No rooms selected, trying to automatically select all rooms")
            try:
                # Execute a query to get all rooms
                self.cursor.execute("SELECT id, name, type, capacity FROM rooms ORDER BY type, name")
                all_rooms = self.cursor.fetchall() or []
                
                if all_rooms:
                    selected_rooms = all_rooms
                    print(f"Auto-selected {len(all_rooms)} rooms")
            except Exception as e:
                print(f"Error auto-selecting rooms: {str(e)}")
        
        # Final check for subjects and rooms
        if not selected_subjects:
            messagebox.showerror("Error", "No subjects selected for scheduling. Please add subjects in the Subjects tab and select them for scheduling.")
            return None
        
        if not selected_rooms:
            messagebox.showerror("Error", "No rooms selected for scheduling. Please add rooms in the Rooms tab and select them for scheduling.")
            return None
        
        # Get scheduling parameters using the day, month, year variables
        try:
            # Create date objects from the individual components
            from datetime import date
            
            # Start date
            start_day = self.start_day_var.get()
            start_month = self.start_month_var.get()
            start_year = self.start_year_var.get()
            start_date = date(start_year, start_month, start_day)
            
            # For end date, use start date + 14 days if not specified
            try:
                if hasattr(self, 'end_day_var') and hasattr(self, 'end_month_var') and hasattr(self, 'end_year_var'):
                    end_day = self.end_day_var.get()
                    end_month = self.end_month_var.get()
                    end_year = self.end_year_var.get()
                    end_date = date(end_year, end_month, end_day)
                else:
                    # Default to 14 days after start date
                    end_date = start_date + timedelta(days=14)
            except Exception as e:
                print(f"Error getting end date: {str(e)}, using default")
                end_date = start_date + timedelta(days=14)
            
            print(f"Scheduling period: {start_date} to {end_date}")
            
            if start_date > end_date:
                messagebox.showerror("Error", "Start date must be before end date.")
                return None
        except Exception as e:
            print(f"Error processing dates: {str(e)}")
            messagebox.showerror("Error", f"Invalid date format: {str(e)}")
            return None
        
        # Create model
        model = cp_model.CpModel()
        
        # Define time slots (days between start and end date)
        days = []
        current_date = start_date
        
        # Always skip Sundays when generating the schedule
        print("Skip Sundays: True (Always skip Sundays)")
        
        while current_date <= end_date:
            # Skip Sundays (6 is Sunday in Python's weekday() where Monday is 0)
            if current_date.weekday() == 6:  # 6 is Sunday
                print(f"Skipping Sunday: {current_date}")
                current_date += timedelta(days=1)
                continue
            days.append(current_date)
            current_date += timedelta(days=1)
        
        if not days:
            messagebox.showerror("Error", "No valid days available for scheduling.")
            return None
        
        # Create variables
        # For each subject, create a variable for the day and room
        subject_day = {}
        subject_room = {}
        
        for subject in selected_subjects:
            subject_id = subject[0]
            subject_day[subject_id] = model.NewIntVar(0, len(days) - 1, f"subject_{subject_id}_day")
            subject_room[subject_id] = model.NewIntVar(0, len(selected_rooms) - 1, f"subject_{subject_id}_room")
        
        # Add constraints
        
        # 1. Room type constraint - subjects can only be assigned to rooms of matching type
        for subject in selected_subjects:
            subject_id = subject[0]
            subject_type = subject[3]  # Assuming index 3 is the type (Theory or Practical)
            
            # For each room, add constraint that subject can only be assigned to room if room type matches
            for room_idx, room in enumerate(selected_rooms):
                room_type = room[2]  # Assuming index 2 is the type
                
                # If room type doesn't match subject type, add constraint that subject can't be assigned to this room
                if subject_type == "Theory" and room_type == "Lab":
                    model.Add(subject_room[subject_id] != room_idx)
                elif subject_type == "Practical" and room_type == "Classroom":
                    model.Add(subject_room[subject_id] != room_idx)
        
        # 2. No two exams on the same day in the same room (unless allow_multiple_exams is enabled)
        # Check if allow_multiple_exams_var exists, default to False if not
        allow_multiple_exams = False
        if hasattr(self, 'allow_multiple_exams_var'):
            try:
                allow_multiple_exams = self.allow_multiple_exams_var.get()
            except Exception as e:
                print(f"Error getting allow_multiple_exams_var: {str(e)}")
                allow_multiple_exams = False
        
        print(f"Allow multiple exams: {allow_multiple_exams}")
        
        if not allow_multiple_exams:
            try:
                for subject1 in selected_subjects:
                    subject1_id = subject1[0]
                    for subject2 in selected_subjects:
                        subject2_id = subject2[0]
                        if subject1_id != subject2_id:
                            # If two subjects are on the same day and in the same room, add constraint
                            same_day = model.NewBoolVar(f"same_day_{subject1_id}_{subject2_id}")
                            same_room = model.NewBoolVar(f"same_room_{subject1_id}_{subject2_id}")
                            
                            model.Add(subject_day[subject1_id] == subject_day[subject2_id]).OnlyEnforceIf(same_day)
                            model.Add(subject_day[subject1_id] != subject_day[subject2_id]).OnlyEnforceIf(same_day.Not())
                            
                            model.Add(subject_room[subject1_id] == subject_room[subject2_id]).OnlyEnforceIf(same_room)
                            model.Add(subject_room[subject1_id] != subject_room[subject2_id]).OnlyEnforceIf(same_room.Not())
                            
                            # If both same_day and same_room are true, that's a conflict
                            conflict = model.NewBoolVar(f"conflict_{subject1_id}_{subject2_id}")
                            model.AddBoolAnd([same_day, same_room]).OnlyEnforceIf(conflict)
                            model.AddBoolOr([same_day.Not(), same_room.Not()]).OnlyEnforceIf(conflict.Not())
                            
                            # Prevent conflicts
                            model.Add(conflict == 0)
            except Exception as e:
                print(f"Error adding exam conflict constraints: {str(e)}")
                # Continue without these constraints if there's an error
        
        # 3. Gap between difficult exams
        for subject1 in selected_subjects:
            subject1_id = subject1[0]
            subject1_difficulty = subject1[5]  # Assuming index 5 is difficulty
            
            for subject2 in selected_subjects:
                subject2_id = subject2[0]
                subject2_difficulty = subject2[5]
                
                if subject1_id != subject2_id:
                    # If both subjects are hard, enforce gap
                    if subject1_difficulty == "Hard" and subject2_difficulty == "Hard":
                        gap = self.hard_gap_var.get()
                        if gap > 0:
                            # Either subject1 is at least 'gap' days before subject2, or vice versa
                            gap_constraint = model.NewBoolVar(f"gap_constraint_{subject1_id}_{subject2_id}")
                            model.Add(subject_day[subject1_id] + gap <= subject_day[subject2_id]).OnlyEnforceIf(gap_constraint)
                            model.Add(subject_day[subject2_id] + gap <= subject_day[subject1_id]).OnlyEnforceIf(gap_constraint.Not())
                    
                    # If both subjects are medium, enforce gap
                    elif subject1_difficulty == "Medium" and subject2_difficulty == "Medium":
                        gap = self.medium_gap_var.get()
                        if gap > 0:
                            # Either subject1 is at least 'gap' days before subject2, or vice versa
                            gap_constraint = model.NewBoolVar(f"gap_constraint_{subject1_id}_{subject2_id}")
                            model.Add(subject_day[subject1_id] + gap <= subject_day[subject2_id]).OnlyEnforceIf(gap_constraint)
                            model.Add(subject_day[subject2_id] + gap <= subject_day[subject1_id]).OnlyEnforceIf(gap_constraint.Not())
        
        # Create solver and solve
        solver = cp_model.CpSolver()
        status = solver.Solve(model)
        
        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            # Create schedule
            schedule = []
            for subject in selected_subjects:
                subject_id = subject[0]
                day_idx = solver.Value(subject_day[subject_id])
                room_idx = solver.Value(subject_room[subject_id])
                
                exam_date = days[day_idx]
                room = selected_rooms[room_idx]
                
                # Determine time slot based on subject type
                # Default time slots if variables don't exist
                if subject[3] == "Theory":  # Assuming index 3 is type
                    if hasattr(self, 'theory_time_var'):
                        try:
                            time_slot = self.theory_time_var.get()
                        except Exception as e:
                            print(f"Error getting theory_time_var: {str(e)}")
                            time_slot = "09:00 AM - 12:00 PM"  # Default time for theory exams
                    else:
                        time_slot = "09:00 AM - 12:00 PM"  # Default time for theory exams
                else:  # Practical
                    if hasattr(self, 'practical_time_var'):
                        try:
                            time_slot = self.practical_time_var.get()
                        except Exception as e:
                            print(f"Error getting practical_time_var: {str(e)}")
                            time_slot = "02:00 PM - 05:00 PM"  # Default time for practical exams
                    else:
                        time_slot = "02:00 PM - 05:00 PM"  # Default time for practical exams
                
                print(f"Using time slot: {time_slot} for subject type: {subject[3]}")
                
                # Split time slot into start and end times
                try:
                    start_time, end_time = time_slot.split(" - ")
                except Exception as e:
                    print(f"Error splitting time slot: {str(e)}")
                    start_time = time_slot.split(" - ")[0] if " - " in time_slot else "09:00 AM"
                    end_time = time_slot.split(" - ")[1] if " - " in time_slot and len(time_slot.split(" - ")) > 1 else "12:00 PM"
                
                schedule.append({
                    "subject_id": subject_id,
                    "subject_code": subject[1],
                    "subject_name": subject[2],
                    "room_id": room[0],
                    "room_name": room[1],
                    "exam_date": exam_date,
                    "start_time": start_time,
                    "end_time": end_time
                })
            
            return schedule
        else:
            messagebox.showerror("Error", "No feasible schedule found with the given constraints. Try relaxing some constraints.")
            return None
    
    def generate_preview(self):
        """Generate a preview of the schedule based on selected subjects and rooms."""
        try:
            # Clear existing preview
            for item in self.preview_tree.get_children():
                self.preview_tree.delete(item)
            
            # Generate schedule
            print("Generating schedule for preview...")
            schedule = self.generate_schedule()
            
            if not schedule:
                messagebox.showinfo("Schedule Generation", "No schedule could be generated with the current settings. Please check your subject and room selections.")
                return
            
            print(f"Schedule generated with {len(schedule)} items")
            
            # Sort schedule by date
            schedule.sort(key=lambda x: (x["exam_date"], x["start_time"]))
            
            # Display schedule in preview tree
            for item in schedule:
                # Make sure we have all required fields
                date_str = item["exam_date"].strftime("%d-%m-%Y") if hasattr(item["exam_date"], "strftime") else str(item["exam_date"])
                
                self.preview_tree.insert("", "end", values=(
                    date_str,
                    item["subject_code"],
                    item["subject_name"],
                    item["room_name"],
                    f"{item['start_time']} - {item['end_time']}"
                ))
            
            # Update preview info if the variable exists
            if hasattr(self, 'preview_info_var'):
                try:
                    start_date = min(item["exam_date"] for item in schedule)
                    end_date = max(item["exam_date"] for item in schedule)
                    start_date_str = start_date.strftime("%d-%m-%Y") if hasattr(start_date, "strftime") else str(start_date)
                    end_date_str = end_date.strftime("%d-%m-%Y") if hasattr(end_date, "strftime") else str(end_date)
                    
                    self.preview_info_var.set(
                        f"Preview generated with {len(schedule)} exams from {start_date_str} to {end_date_str}."
                    )
                except Exception as e:
                    print(f"Error updating preview info: {str(e)}")
            
            # Store the schedule for later use
            self.current_schedule = schedule
            
            # Update the schedule tree in the View Schedule tab
            self.update_schedule_view(schedule)
            
            # Automatically save the schedule to the database
            if self.schedule_name_var.get().strip():
                try:
                    self.save_schedule()
                    print("Schedule automatically saved to database")
                except Exception as e:
                    print(f"Error auto-saving schedule: {str(e)}")
                    # Continue even if auto-save fails
            else:
                # Prompt user to enter a name for the schedule
                name_dialog = tk.Toplevel(self.root)
                name_dialog.title("Save Schedule")
                name_dialog.geometry("400x150")
                name_dialog.transient(self.root)
                name_dialog.grab_set()
                
                ttk.Label(name_dialog, text="Enter a name for this schedule:").pack(pady=10)
                name_var = tk.StringVar()
                ttk.Entry(name_dialog, textvariable=name_var, width=40).pack(pady=10)
                
                def save_with_name():
                    if name_var.get().strip():
                        self.schedule_name_var.set(name_var.get().strip())
                        name_dialog.destroy()
                        try:
                            self.save_schedule()
                            print("Schedule saved to database with user-provided name")
                        except Exception as e:
                            print(f"Error saving schedule with name: {str(e)}")
                            messagebox.showerror("Error", f"Failed to save schedule: {str(e)}")
                    else:
                        messagebox.showwarning("Warning", "Please enter a name for the schedule.")
                
                button_frame = ttk.Frame(name_dialog)
                button_frame.pack(pady=10)
                ttk.Button(button_frame, text="Save", command=save_with_name).pack(side=tk.LEFT, padx=10)
                ttk.Button(button_frame, text="Cancel", command=name_dialog.destroy).pack(side=tk.LEFT)
            
            # Show success message
            messagebox.showinfo("Schedule Generation", f"Successfully generated a schedule with {len(schedule)} exams.")
            
            # First, switch to the generator tab
            self.notebook.select(self.tab_indices["generator"])
            
            # Then, switch to the Preview tab within the generator notebook
            self.root.after(100, lambda: self.generator_notebook.select(self.preview_tab_index))
            
        except Exception as e:
            print(f"Error generating preview: {str(e)}")
            messagebox.showerror("Error", f"Failed to generate preview: {str(e)}")
            return None
    
    def delete_schedule(self):
        """Delete the selected schedule from the UI and database."""
        # Check if a schedule is selected in the treeview
        selected_item = self.schedule_tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a schedule to delete.")
            return
            
        # Get schedule ID and name
        schedule_id = self.schedule_tree.item(selected_item[0])["values"][0]
        schedule_name = self.schedule_tree.item(selected_item[0])["values"][1]
        
        # Confirm deletion
        confirm = messagebox.askyesno("Confirm Deletion", 
                                    f"Are you sure you want to delete the schedule '{schedule_name}'?\n\n" +
                                    "This will permanently delete the schedule and all its exam data from the database.")
        
        if not confirm:
            return
            
        try:
            # Ensure database connection
            if not self.ensure_connection():
                return
                
            # Set autocommit to False to start a transaction
            self.conn.autocommit = False
            
            try:
                # Delete schedule items first (should cascade automatically due to foreign key constraints,
                # but we'll do it explicitly to be safe)
                self.cursor.execute("DELETE FROM schedule_items WHERE schedule_id = %s", (schedule_id,))
                
                # Delete the schedule
                self.cursor.execute("DELETE FROM schedules WHERE id = %s", (schedule_id,))
                
                # Commit the transaction
                self.conn.commit()
                
                # Set autocommit back to True
                self.conn.autocommit = True
                
                # Remove from treeview
                self.schedule_tree.delete(selected_item[0])
                
                # Clear current schedule if it was the deleted one
                if hasattr(self, 'current_schedule') and self.current_schedule:
                    # Switch back to list view
                    self.view_var.set("list")
                    self.show_list_view()
                    # Clear current schedule
                    self.current_schedule = None
                
                messagebox.showinfo("Success", f"Schedule '{schedule_name}' has been deleted.")
                
            except Exception as e:
                # Rollback in case of error
                try:
                    self.conn.rollback()
                except Exception as rollback_error:
                    print(f"Error during rollback: {str(rollback_error)}")
                
                # Reset autocommit to True
                try:
                    self.conn.autocommit = True
                except Exception:
                    pass
                    
                print(f"Error deleting schedule: {str(e)}")
                messagebox.showerror("Error", f"Failed to delete schedule: {str(e)}")
                
        except Exception as e:
            print(f"Error in delete_schedule: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def export_schedule_csv(self):
        """Export the current schedule to a CSV file."""
        if not hasattr(self, 'current_schedule') or not self.current_schedule:
            messagebox.showerror("Error", "No schedule to export.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export Schedule to CSV",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Prepare data for CSV
            data = []
            for item in self.current_schedule:
                data.append({
                    "Date": item["exam_date"].strftime("%d-%m-%Y"),
                    "Start Time": item["start_time"],
                    "End Time": item["end_time"],
                    "Subject Code": item["subject_code"],
                    "Subject Name": item["subject_name"],
                    "Room": item["room_name"]
                })
            
            # Create DataFrame
            df = pd.DataFrame(data)
            
            # Export to CSV
            df.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Exported schedule to {file_path}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export schedule: {str(e)}")
    
    def export_schedule_pdf(self):
        """Export the current schedule to a PDF file."""
        if not hasattr(self, 'current_schedule') or not self.current_schedule:
            messagebox.showerror("Error", "No schedule to export.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export Schedule to PDF",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Create PDF document
            doc = SimpleDocTemplate(file_path, pagesize=A4)
            elements = []
            
            # Add title
            styles = getSampleStyleSheet()
            title_style = styles["Title"]
            title_text = f"Exam Schedule: {self.schedule_name_var.get() or 'Untitled'}"
            elements.append(Paragraph(title_text, title_style))
            elements.append(Spacer(1, 12))
            
            # Add date range
            if self.current_schedule:
                start_date = min(item["exam_date"] for item in self.current_schedule)
                end_date = max(item["exam_date"] for item in self.current_schedule)
                date_range = f"From {start_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')}"
                elements.append(Paragraph(date_range, styles["Normal"]))
                elements.append(Spacer(1, 12))
            
            # Sort schedule by date and time
            sorted_schedule = sorted(self.current_schedule, key=lambda x: (x["exam_date"], x["start_time"]))
            
            # Create table data
            table_data = [
                ["Date", "Time", "Subject Code", "Subject Name", "Room"]
            ]
            
            for item in sorted_schedule:
                table_data.append([
                    item["exam_date"].strftime("%d-%m-%Y"),
                    f"{item['start_time']} - {item['end_time']}",
                    item["subject_code"],
                    item["subject_name"],
                    item["room_name"]
                ])
            
            # Create table
            table = Table(table_data, repeatRows=1)
            
            # Style table
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ])
            
            # Add alternating row colors
            for i in range(1, len(table_data)):
                if i % 2 == 0:
                    table_style.add('BACKGROUND', (0, i), (-1, i), colors.lightgrey)
            
            table.setStyle(table_style)
            elements.append(table)
            
            # Build PDF
            doc.build(elements)
            messagebox.showinfo("Success", f"Exported schedule to {file_path}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export schedule to PDF: {str(e)}")
    
    def init_schedule_tab(self):
        """Initialize the schedule view tab."""
        # Use the existing schedule tab instead of creating a new one
        schedule_frame = self.schedule_tab
        
        # Create top frame for schedule name and buttons
        top_frame = ttk.Frame(schedule_frame)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Add search frame
        search_frame = ttk.Frame(top_frame)
        search_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(search_frame, text="Search Schedule:").pack(side=tk.LEFT, padx=5)
        self.schedule_search_var = tk.StringVar()
        self.schedule_search_var.trace('w', self.filter_schedules)
        ttk.Entry(search_frame, textvariable=self.schedule_search_var).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Button frame
        button_frame = ttk.Frame(top_frame)
        button_frame.pack(side=tk.RIGHT)
        
        ttk.Button(button_frame, text="Save Schedule", command=self.save_schedule).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Delete Schedule", command=self.delete_schedule, style="Danger.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Export to CSV", command=self.export_schedule_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Export to PDF", command=self.export_schedule_pdf).pack(side=tk.LEFT, padx=5)
        
        # Register a custom style for the delete button
        style = ttk.Style()
        style.configure("Danger.TButton", foreground="red", font=("Helvetica", 9, "bold"))
        
        # Create view selection frame
        view_frame = ttk.Frame(schedule_frame)
        view_frame.pack(fill=tk.X, padx=5)
        
        ttk.Label(view_frame, text="View:").pack(side=tk.LEFT, padx=5)
        self.view_var = tk.StringVar(value="list")
        ttk.Radiobutton(view_frame, text="List View", variable=self.view_var, value="list", 
                       command=self.show_list_view).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(view_frame, text="Calendar View", variable=self.view_var, value="calendar", 
                       command=self.show_calendar_view).pack(side=tk.LEFT, padx=5)
        
        # Create content frame to hold either list view or calendar view
        self.content_frame = ttk.Frame(schedule_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create list view frame
        self.list_frame = ttk.Frame(self.content_frame)
        
        # Create schedule list
        columns = ("id", "name", "semester", "exam_type", "start_date", "created_at")
        self.schedule_tree = ttk.Treeview(self.list_frame, columns=columns, show="headings", height=20)
        
        # Configure columns
        self.schedule_tree.heading("id", text="ID")
        self.schedule_tree.heading("name", text="Name")
        self.schedule_tree.heading("semester", text="Semester")
        self.schedule_tree.heading("exam_type", text="Exam Type")
        self.schedule_tree.heading("start_date", text="Start Date")
        self.schedule_tree.heading("created_at", text="Created At")
        
        # Add scrollbars
        y_scrollbar = ttk.Scrollbar(self.list_frame, orient=tk.VERTICAL, command=self.schedule_tree.yview)
        self.schedule_tree.configure(yscrollcommand=y_scrollbar.set)
        
        # Pack widgets
        self.schedule_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Hide ID column and set other column widths
        self.schedule_tree.column("id", width=0, stretch=False)
        self.schedule_tree.column("name", width=150)
        self.schedule_tree.column("semester", width=100)
        self.schedule_tree.column("exam_type", width=100)
        self.schedule_tree.column("start_date", width=100)
        self.schedule_tree.column("created_at", width=150)
        
        # Bind double-click event
        self.schedule_tree.bind('<Double-1>', self.on_schedule_select)
        
        # Create calendar view frame (but don't pack it yet)
        self.calendar_frame = ttk.Frame(self.content_frame)
        
        # Create canvas for calendar view
        self.calendar_canvas = tk.Canvas(self.calendar_frame, bg="white", width=800, height=600)
        
        # Create scrollbars for calendar
        x_scrollbar = ttk.Scrollbar(self.calendar_frame, orient=tk.HORIZONTAL, command=self.calendar_canvas.xview)
        y_scrollbar = ttk.Scrollbar(self.calendar_frame, orient=tk.VERTICAL, command=self.calendar_canvas.yview)
        self.calendar_canvas.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
        
        # Pack calendar components
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.calendar_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Show list view by default
        self.show_list_view()
        
        # Load all schedules
        self.load_all_schedules()
        
    def update_schedule_view(self, schedule):
        """Update the schedule view with the given schedule."""
        if not schedule:
            return
            
        try:
            # Store the current schedule
            self.current_schedule = schedule
            
            # If we're in calendar view, update the calendar
            if self.view_var.get() == "calendar":
                self.show_calendar_view()
                
        except Exception as e:
            print(f"Error updating schedule view: {str(e)}")
            
    def load_all_schedules(self):
        """Load all schedules from the database."""
        try:
            if not self.ensure_connection():
                return
            
            # Clear existing items
            for item in self.schedule_tree.get_children():
                self.schedule_tree.delete(item)
            
            # Get all schedules
            self.cursor.execute("""
                SELECT id, name, semester, exam_type, 
                       DATE_FORMAT(start_date, '%Y-%m-%d') as start_date,
                       DATE_FORMAT(created_at, '%Y-%m-%d %H:%i') as created_at
                FROM schedules
                ORDER BY created_at DESC
            """)
            
            schedules = self.cursor.fetchall()
            for schedule in schedules:
                self.schedule_tree.insert("", "end", values=schedule)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load schedules: {str(e)}")
    
    def filter_schedules(self, *args):
        """Filter schedules based on search term."""
        search_term = self.schedule_search_var.get().lower()
        
        # Clear tree
        for item in self.schedule_tree.get_children():
            self.schedule_tree.delete(item)
        
        try:
            if not self.ensure_connection():
                return
            
            # Get filtered schedules
            self.cursor.execute("""
                SELECT id, name, semester, exam_type, 
                       DATE_FORMAT(start_date, '%Y-%m-%d') as start_date,
                       DATE_FORMAT(created_at, '%Y-%m-%d %H:%i') as created_at
                FROM schedules
                WHERE LOWER(name) LIKE %s OR 
                      LOWER(semester) LIKE %s OR
                      LOWER(exam_type) LIKE %s
                ORDER BY created_at DESC
            """, ("%" + search_term + "%", "%" + search_term + "%", "%" + search_term + "%"))
            
            schedules = self.cursor.fetchall()
            for schedule in schedules:
                self.schedule_tree.insert("", "end", values=schedule)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to filter schedules: {str(e)}")
    
    def select_schedule_tab(self):
        """Select the View Schedule tab."""
        self.notebook.select(self.tab_indices["schedule"])
        # Force update to ensure the tab change is processed
        self.root.update()
    
    def on_schedule_select(self, event):
        """Handle double-click on schedule item."""
        selected_item = self.schedule_tree.selection()
        if not selected_item:
            return
            
        # Get schedule ID and other details
        values = self.schedule_tree.item(selected_item[0])["values"]
        schedule_id = values[0]
        schedule_name = values[1]
        semester = values[2]
        exam_type = values[3]
        
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return
                
            # Fetch schedule items from database
            self.cursor.execute("""
                SELECT si.subject_id, si.room_id, si.exam_date, si.start_time, si.end_time,
                       s.code as subject_code, s.name as subject_name,
                       r.name as room_name
                FROM schedule_items si
                JOIN subjects s ON si.subject_id = s.id
                JOIN rooms r ON si.room_id = r.id
                WHERE si.schedule_id = %s
                ORDER BY si.exam_date, si.start_time
            """, (schedule_id,))
            
            items = []
            for item in self.cursor.fetchall():
                # Convert to dictionary for easier access
                items.append({
                    "subject_id": item[0],
                    "room_id": item[1],
                    "exam_date": item[2],
                    "start_time": item[3],
                    "end_time": item[4],
                    "subject_code": item[5],
                    "subject_name": item[6],
                    "room_name": item[7]
                })
            
            # Store the loaded schedule and its ID
            self.current_schedule = items
            self.current_schedule_id = schedule_id
            
            # First ensure we're on the View Schedule tab
            self.select_schedule_tab()
            
            # Update the schedule name in the generator tab
            self.schedule_name_var.set(schedule_name)
            
            # Update semester and exam type if possible
            if hasattr(self, 'semester_var'):
                self.semester_var.set(semester)
            if hasattr(self, 'exam_type_var'):
                self.exam_type_var.set(exam_type)
            
            # Update the preview tab in the generator section
            self.update_preview_with_schedule(items)
            
            # Switch to calendar view in the View Schedule tab
            self.view_var.set("calendar")
            self.show_calendar_view()
            
            # Also show the schedule in the Preview tab of the Generate Schedule section
            # First, switch to the generator tab
            self.root.after(500, lambda: self.show_schedule_in_generator_preview())
            
        except Exception as e:
            print(f"Error loading schedule: {str(e)}")
            messagebox.showerror("Error", f"Failed to load schedule: {str(e)}")
    
    def update_preview_with_schedule(self, schedule):
        """Update the preview tab with the loaded schedule."""
        try:
            # Clear existing preview
            for item in self.preview_tree.get_children():
                self.preview_tree.delete(item)
            
            # Switch to the preview tab
            self.generator_notebook.select(self.preview_tab_index)
            
            # Display schedule in preview tree
            for item in schedule:
                # Make sure we have all required fields
                date_str = item["exam_date"].strftime("%d-%m-%Y") if hasattr(item["exam_date"], "strftime") else str(item["exam_date"])
                
                self.preview_tree.insert("", "end", values=(
                    date_str,
                    item["subject_code"],
                    item["subject_name"],
                    item["room_name"],
                    f"{item['start_time']} - {item['end_time']}"
                ))
                
            # Store the schedule for later use
            self.current_schedule = schedule
            
        except Exception as e:
            print(f"Error updating preview with schedule: {e}")
            import traceback
            traceback.print_exc()
            
    def on_preview_item_double_click(self, event):
        """Handle double-click on a preview item to edit it."""
        try:
            # Get the selected item
            selected_item = self.preview_tree.selection()
            if not selected_item:
                return
                
            # Get the item values
            item_values = self.preview_tree.item(selected_item[0], 'values')
            if not item_values or len(item_values) < 5:
                return
                
            # Find the corresponding item in the current_schedule
            selected_date_str = item_values[0]
            selected_subject_code = item_values[1]
            selected_subject_name = item_values[2]
            selected_room_name = item_values[3]
            selected_time = item_values[4]
            
            # Find the matching item in the current schedule
            matching_item = None
            for item in self.current_schedule:
                date_str = item["exam_date"].strftime("%d-%m-%Y") if hasattr(item["exam_date"], "strftime") else str(item["exam_date"])
                time_str = f"{item['start_time']} - {item['end_time']}"
                
                if (date_str == selected_date_str and 
                    item["subject_code"] == selected_subject_code and 
                    item["subject_name"] == selected_subject_name and 
                    item["room_name"] == selected_room_name and 
                    time_str == selected_time):
                    matching_item = item
                    break
            
            if not matching_item:
                messagebox.showwarning("Edit Error", "Could not find the selected item in the schedule.")
                return
                
            # Create a modal dialog for editing
            self.open_edit_dialog(matching_item, selected_item[0])
            
        except Exception as e:
            print(f"Error handling preview item double-click: {e}")
            import traceback
            traceback.print_exc()
            
    def open_edit_dialog(self, item, tree_item_id):
        """Open a dialog to edit a schedule item."""
        try:
            # Create a top-level dialog
            edit_dialog = tk.Toplevel(self.root)
            edit_dialog.title("Edit Schedule Item")
            edit_dialog.geometry("500x400")
            edit_dialog.resizable(False, False)
            edit_dialog.transient(self.root)  # Set to be on top of the main window
            edit_dialog.grab_set()  # Make it modal
            
            # Add padding
            frame = ttk.Frame(edit_dialog, padding=20)
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Configure grid
            frame.columnconfigure(0, weight=0)
            frame.columnconfigure(1, weight=1)
            
            # Create form fields
            ttk.Label(frame, text="Subject Code:").grid(row=0, column=0, sticky=tk.W, pady=10)
            subject_code_var = tk.StringVar(value=item["subject_code"])
            ttk.Entry(frame, textvariable=subject_code_var, width=30).grid(row=0, column=1, sticky=tk.W, pady=10)
            
            ttk.Label(frame, text="Subject Name:").grid(row=1, column=0, sticky=tk.W, pady=10)
            subject_name_var = tk.StringVar(value=item["subject_name"])
            ttk.Entry(frame, textvariable=subject_name_var, width=30).grid(row=1, column=1, sticky=tk.W, pady=10)
            
            ttk.Label(frame, text="Room:").grid(row=2, column=0, sticky=tk.W, pady=10)
            
            # Get all available rooms for the dropdown
            rooms = []
            try:
                if self.ensure_connection():
                    self.cursor.execute("SELECT name FROM rooms ORDER BY name")
                    rooms = [row[0] for row in self.cursor.fetchall()]
            except Exception as e:
                print(f"Error fetching rooms for edit dialog: {e}")
                rooms = [item["room_name"]]  # Fallback to current room
                
            room_var = tk.StringVar(value=item["room_name"])
            room_combo = ttk.Combobox(frame, textvariable=room_var, values=rooms, width=28, state="readonly")
            room_combo.grid(row=2, column=1, sticky=tk.W, pady=10)
            
            ttk.Label(frame, text="Date (DD-MM-YYYY):").grid(row=3, column=0, sticky=tk.W, pady=10)
            
            # Format the date
            date_str = item["exam_date"].strftime("%d-%m-%Y") if hasattr(item["exam_date"], "strftime") else str(item["exam_date"])
            date_var = tk.StringVar(value=date_str)
            ttk.Entry(frame, textvariable=date_var, width=30).grid(row=3, column=1, sticky=tk.W, pady=10)
            
            ttk.Label(frame, text="Start Time:").grid(row=4, column=0, sticky=tk.W, pady=10)
            start_time_var = tk.StringVar(value=item["start_time"])
            ttk.Entry(frame, textvariable=start_time_var, width=30).grid(row=4, column=1, sticky=tk.W, pady=10)
            
            ttk.Label(frame, text="End Time:").grid(row=5, column=0, sticky=tk.W, pady=10)
            end_time_var = tk.StringVar(value=item["end_time"])
            ttk.Entry(frame, textvariable=end_time_var, width=30).grid(row=5, column=1, sticky=tk.W, pady=10)
            
            # Add a note about time format
            ttk.Label(frame, text="Time format: HH:MM AM/PM", font=("Helvetica", 8), foreground="gray").grid(
                row=6, column=1, sticky=tk.W)
            
            # Buttons
            button_frame = ttk.Frame(frame)
            button_frame.grid(row=7, column=0, columnspan=2, pady=20)
            
            def save_changes():
                try:
                    # Validate inputs
                    if not subject_code_var.get().strip() or not subject_name_var.get().strip() or \
                       not room_var.get().strip() or not date_var.get().strip() or \
                       not start_time_var.get().strip() or not end_time_var.get().strip():
                        messagebox.showwarning("Validation Error", "All fields are required.")
                        return
                    
                    # Parse date
                    try:
                        new_date = datetime.strptime(date_var.get(), "%d-%m-%Y").date()
                    except ValueError:
                        messagebox.showwarning("Validation Error", "Invalid date format. Use DD-MM-YYYY.")
                        return
                    
                    # Update the item in the current_schedule
                    item["subject_code"] = subject_code_var.get().strip()
                    item["subject_name"] = subject_name_var.get().strip()
                    item["room_name"] = room_var.get().strip()
                    item["exam_date"] = new_date
                    item["start_time"] = start_time_var.get().strip()
                    item["end_time"] = end_time_var.get().strip()
                    
                    # Update the treeview
                    self.preview_tree.item(tree_item_id, values=(
                        date_var.get(),
                        subject_code_var.get().strip(),
                        subject_name_var.get().strip(),
                        room_var.get().strip(),
                        f"{start_time_var.get().strip()} - {end_time_var.get().strip()}"
                    ))
                    
                    # Update the database if the schedule has been saved
                    if hasattr(self, 'current_schedule_id') and self.current_schedule_id:
                        try:
                            # Ensure database connection
                            if self.ensure_connection():
                                # Get the room_id based on room_name
                                self.cursor.execute("SELECT id FROM rooms WHERE name = %s", (item["room_name"],))
                                room_result = self.cursor.fetchone()
                                if room_result:
                                    room_id = room_result[0]
                                    # Update the room_id in the item
                                    item["room_id"] = room_id
                                    
                                    # Update the schedule_item in the database
                                    self.cursor.execute("""
                                        UPDATE schedule_items 
                                        SET room_id = %s, exam_date = %s, start_time = %s, end_time = %s 
                                        WHERE schedule_id = %s AND subject_id = %s
                                    """, (
                                        room_id,
                                        new_date,
                                        item["start_time"],
                                        item["end_time"],
                                        self.current_schedule_id,
                                        item["subject_id"]
                                    ))
                                    
                                    # Commit the changes
                                    self.conn.commit()
                                    print(f"Updated schedule item in database for subject {item['subject_code']}")
                                else:
                                    print(f"Error: Room '{item['room_name']}' not found in database")
                        except Exception as e:
                            print(f"Error updating schedule item in database: {e}")
                            import traceback
                            traceback.print_exc()
                            self.conn.rollback()
                            messagebox.showwarning("Database Warning", f"Changes saved in application but could not be updated in database: {str(e)}")
                    
                    # Close the dialog
                    edit_dialog.destroy()
                    
                    # Show success message
                    messagebox.showinfo("Success", "Schedule item updated successfully.")
                    
                except Exception as e:
                    print(f"Error saving changes: {e}")
                    messagebox.showerror("Error", f"Failed to save changes: {str(e)}")
            
            ttk.Button(button_frame, text="Save Changes", command=save_changes, style="Accent.TButton").pack(side=tk.LEFT, padx=10)
            ttk.Button(button_frame, text="Cancel", command=edit_dialog.destroy).pack(side=tk.LEFT, padx=10)
            
            # Center the dialog on the screen
            edit_dialog.update_idletasks()
            width = edit_dialog.winfo_width()
            height = edit_dialog.winfo_height()
            x = (edit_dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (edit_dialog.winfo_screenheight() // 2) - (height // 2)
            edit_dialog.geometry(f"{width}x{height}+{x}+{y}")
            
            # Make dialog modal
            edit_dialog.focus_set()
            edit_dialog.wait_window()
            
        except Exception as e:
            print(f"Error opening edit dialog: {e}")
            import traceback
            traceback.print_exc()
            
    def show_preview_context_menu(self, event):
        """Show context menu on right-click in preview tree."""
        try:
            # Select the item under cursor
            item = self.preview_tree.identify_row(event.y)
            if item:
                self.preview_tree.selection_set(item)
                
                # Create context menu
                context_menu = tk.Menu(self.root, tearoff=0)
                context_menu.add_command(label="Edit", command=lambda: self.on_preview_item_double_click(None))
                context_menu.add_command(label="Delete", command=self.delete_schedule_item)
                
                # Show context menu
                context_menu.tk_popup(event.x_root, event.y_root)
        except Exception as e:
            print(f"Error showing context menu: {e}")
    
    def delete_schedule_item(self):
        """Delete the selected schedule item from the preview and database."""
        try:
            # Get the selected item
            selected_item = self.preview_tree.selection()
            if not selected_item:
                return
                
            # Get the item values
            item_values = self.preview_tree.item(selected_item[0], 'values')
            if not item_values or len(item_values) < 5:
                return
                
            # Find the corresponding item in the current_schedule
            selected_date_str = item_values[0]
            selected_subject_code = item_values[1]
            selected_subject_name = item_values[2]
            selected_room_name = item_values[3]
            selected_time = item_values[4]
            
            # Confirm deletion
            confirm = messagebox.askyesno(
                "Confirm Deletion",
                f"Are you sure you want to delete this exam?\n\nSubject: {selected_subject_code} - {selected_subject_name}\nDate: {selected_date_str}\nRoom: {selected_room_name}\nTime: {selected_time}"
            )
            
            if not confirm:
                return
            
            # Find the matching item in the current schedule
            matching_item = None
            matching_index = -1
            for i, item in enumerate(self.current_schedule):
                date_str = item["exam_date"].strftime("%d-%m-%Y") if hasattr(item["exam_date"], "strftime") else str(item["exam_date"])
                time_str = f"{item['start_time']} - {item['end_time']}"
                
                if (date_str == selected_date_str and 
                    item["subject_code"] == selected_subject_code and 
                    item["subject_name"] == selected_subject_name and 
                    item["room_name"] == selected_room_name and 
                    time_str == selected_time):
                    matching_item = item
                    matching_index = i
                    break
            
            if matching_item is None:
                messagebox.showwarning("Delete Error", "Could not find the selected item in the schedule.")
                return
            
            # Delete from database if schedule has been saved
            if hasattr(self, 'current_schedule_id') and self.current_schedule_id:
                try:
                    # Ensure database connection
                    if self.ensure_connection():
                        # Delete the schedule item from the database
                        self.cursor.execute("""
                            DELETE FROM schedule_items 
                            WHERE schedule_id = %s AND subject_id = %s AND room_id = %s
                        """, (
                            self.current_schedule_id,
                            matching_item["subject_id"],
                            matching_item["room_id"]
                        ))
                        self.conn.commit()
                except Exception as e:
                    print(f"Error deleting schedule item from database: {e}")
                    messagebox.showerror("Database Error", f"Failed to delete schedule item: {str(e)}")
                    return
            
            # Remove from current_schedule list
            if matching_index >= 0:
                self.current_schedule.pop(matching_index)
            
            # Remove from treeview
            self.preview_tree.delete(selected_item[0])
            
            # Show success message
            self.status_var.set("Schedule item deleted successfully.")
            
        except Exception as e:
            print(f"Error deleting schedule item: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to delete schedule item: {str(e)}")
    
    def open_add_dialog(self):
        """Open a dialog to add a new exam item to the schedule."""
        try:
            # Check if a schedule is loaded
            if not hasattr(self, 'current_schedule_id') or not self.current_schedule_id:
                messagebox.showwarning("Add Error", "Please load a schedule first before adding items.")
                return
                
            # Create a top-level dialog
            add_dialog = tk.Toplevel(self.root)
            add_dialog.title("Add Exam Item")
            add_dialog.geometry("500x400")
            add_dialog.resizable(False, False)
            add_dialog.transient(self.root)  # Set to be on top of the main window
            add_dialog.grab_set()  # Make it modal
            
            # Add padding
            frame = ttk.Frame(add_dialog, padding=20)
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Configure grid
            frame.columnconfigure(0, weight=0)
            frame.columnconfigure(1, weight=1)
            
            # Create form fields
            ttk.Label(frame, text="Subject Code:").grid(row=0, column=0, sticky=tk.W, pady=10)
            subject_code_var = tk.StringVar()
            ttk.Entry(frame, textvariable=subject_code_var, width=30).grid(row=0, column=1, sticky=tk.W, pady=10)
            
            ttk.Label(frame, text="Subject Name:").grid(row=1, column=0, sticky=tk.W, pady=10)
            subject_name_var = tk.StringVar()
            ttk.Entry(frame, textvariable=subject_name_var, width=30).grid(row=1, column=1, sticky=tk.W, pady=10)
            
            ttk.Label(frame, text="Room:").grid(row=2, column=0, sticky=tk.W, pady=10)
            
            # Get all available rooms for the dropdown
            rooms = []
            try:
                if self.ensure_connection():
                    self.cursor.execute("SELECT name FROM rooms ORDER BY name")
                    rooms = [row[0] for row in self.cursor.fetchall()]
            except Exception as e:
                print(f"Error fetching rooms for add dialog: {e}")
                
            room_var = tk.StringVar()
            if rooms:
                room_var.set(rooms[0])
            room_combo = ttk.Combobox(frame, textvariable=room_var, values=rooms, width=28, state="readonly")
            room_combo.grid(row=2, column=1, sticky=tk.W, pady=10)
            
            ttk.Label(frame, text="Date (DD-MM-YYYY):").grid(row=3, column=0, sticky=tk.W, pady=10)
            
            # Default to today's date
            date_str = datetime.now().strftime("%d-%m-%Y")
            date_var = tk.StringVar(value=date_str)
            ttk.Entry(frame, textvariable=date_var, width=30).grid(row=3, column=1, sticky=tk.W, pady=10)
            
            ttk.Label(frame, text="Start Time:").grid(row=4, column=0, sticky=tk.W, pady=10)
            start_time_var = tk.StringVar(value="10:00 AM")
            ttk.Entry(frame, textvariable=start_time_var, width=30).grid(row=4, column=1, sticky=tk.W, pady=10)
            
            ttk.Label(frame, text="End Time:").grid(row=5, column=0, sticky=tk.W, pady=10)
            end_time_var = tk.StringVar(value="01:00 PM")
            ttk.Entry(frame, textvariable=end_time_var, width=30).grid(row=5, column=1, sticky=tk.W, pady=10)
            
            # Add a note about time format
            ttk.Label(frame, text="Time format: HH:MM AM/PM", font=("Helvetica", 8), foreground="gray").grid(
                row=6, column=1, sticky=tk.W)
            
            # Buttons
            button_frame = ttk.Frame(frame)
            button_frame.grid(row=7, column=0, columnspan=2, pady=20)
            
            def save_new_item():
                try:
                    # Validate inputs
                    if not subject_code_var.get().strip() or not subject_name_var.get().strip() or \
                       not room_var.get().strip() or not date_var.get().strip() or \
                       not start_time_var.get().strip() or not end_time_var.get().strip():
                        messagebox.showwarning("Validation Error", "All fields are required.")
                        return
                    
                    # Parse date
                    try:
                        exam_date = datetime.strptime(date_var.get(), "%d-%m-%Y").date()
                    except ValueError:
                        messagebox.showwarning("Validation Error", "Invalid date format. Use DD-MM-YYYY.")
                        return
                    
                    # Ensure database connection
                    if not self.ensure_connection():
                        messagebox.showerror("Database Error", "Could not connect to database.")
                        return
                    
                    # Get the room_id based on room_name
                    self.cursor.execute("SELECT id FROM rooms WHERE name = %s", (room_var.get().strip(),))
                    room_result = self.cursor.fetchone()
                    if not room_result:
                        messagebox.showwarning("Validation Error", f"Room '{room_var.get().strip()}' not found in database.")
                        return
                    
                    room_id = room_result[0]
                    
                    # Check if subject exists, if not create it
                    self.cursor.execute("SELECT id FROM subjects WHERE code = %s", (subject_code_var.get().strip(),))
                    subject_result = self.cursor.fetchone()
                    
                    if subject_result:
                        subject_id = subject_result[0]
                    else:
                        # Create a new subject
                        self.cursor.execute("""
                            INSERT INTO subjects (code, name, type, difficulty, semester) 
                            VALUES (%s, %s, %s, %s, %s)
                        """, (
                            subject_code_var.get().strip(),
                            subject_name_var.get().strip(),
                            "Regular",  # Default type
                            "Medium",   # Default difficulty
                            "1"         # Default semester
                        ))
                        self.conn.commit()
                        subject_id = self.cursor.lastrowid
                    
                    # Insert the new item into the schedule_items table
                    self.cursor.execute("""
                        INSERT INTO schedule_items (schedule_id, subject_id, room_id, exam_date, start_time, end_time) 
                        VALUES (%s, %s, %s, %s, %s, %s)
                    """, (
                        self.current_schedule_id,
                        subject_id,
                        room_id,
                        exam_date,
                        start_time_var.get().strip(),
                        end_time_var.get().strip()
                    ))
                    
                    # Commit the changes
                    self.conn.commit()
                    
                    # Add the new item to the treeview
                    self.preview_tree.insert("", tk.END, values=(
                        date_var.get(),
                        subject_code_var.get().strip(),
                        subject_name_var.get().strip(),
                        room_var.get().strip(),
                        f"{start_time_var.get().strip()} - {end_time_var.get().strip()}"
                    ))
                    
                    # Close the dialog
                    add_dialog.destroy()
                    
                    # Show success message
                    messagebox.showinfo("Success", "New exam item added successfully.")
                    
                except Exception as e:
                    print(f"Error adding new exam item: {e}")
                    import traceback
                    traceback.print_exc()
                    self.conn.rollback()
                    messagebox.showerror("Error", f"Failed to add new exam item: {str(e)}")
            
            ttk.Button(button_frame, text="Add Item", command=save_new_item, style="Accent.TButton").pack(side=tk.LEFT, padx=10)
            ttk.Button(button_frame, text="Cancel", command=add_dialog.destroy).pack(side=tk.LEFT, padx=10)
            
            # Center the dialog on the screen
            add_dialog.update_idletasks()
            width = add_dialog.winfo_width()
            height = add_dialog.winfo_height()
            x = (add_dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (add_dialog.winfo_screenheight() // 2) - (height // 2)
            add_dialog.geometry(f"{width}x{height}+{x}+{y}")
            
            # Make dialog modal
            add_dialog.focus_set()
            add_dialog.wait_window()
            
        except Exception as e:
            print(f"Error opening add dialog: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to open add dialog: {str(e)}")
    
    def show_schedule_in_generator_preview(self):
        """Navigate to the Preview tab in the Generate Schedule section."""
        try:
            # First, switch to the generator tab
            if hasattr(self, 'tab_indices') and 'generator' in self.tab_indices:
                self.notebook.select(self.tab_indices["generator"])
                
                # Then, switch to the Preview tab within the generator notebook
                if hasattr(self, 'generator_notebook') and hasattr(self, 'preview_tab_index'):
                    self.generator_notebook.select(self.preview_tab_index)
                    
                    # Update the preview tree with the current schedule data
                    if hasattr(self, 'current_schedule') and self.current_schedule and hasattr(self, 'preview_tree'):
                        # Clear existing preview
                        for item in self.preview_tree.get_children():
                            self.preview_tree.delete(item)
                        
                        # Display schedule in preview tree
                        for item in self.current_schedule:
                            # Make sure we have all required fields
                            if isinstance(item["exam_date"], str):
                                try:
                                    date_obj = datetime.strptime(item["exam_date"], "%Y-%m-%d")
                                    date_str = date_obj.strftime("%d-%m-%Y")
                                except:
                                    date_str = item["exam_date"]
                            else:
                                date_str = item["exam_date"].strftime("%d-%m-%Y") if hasattr(item["exam_date"], "strftime") else str(item["exam_date"])
                            
                            self.preview_tree.insert("", "end", values=(
                                date_str,
                                item["subject_code"],
                                item["subject_name"],
                                item["room_name"],
                                f"{item['start_time']} - {item['end_time']}"
                            ))
                    
                    # Update the UI to reflect the changes
                    self.root.update()
                    
                    # Show a message to the user
                    self.status_var.set(f"Schedule loaded in the Preview tab.")
        except Exception as e:
            print(f"Error navigating to preview tab: {str(e)}")
            # Don't show error to user as this is a secondary operation
    
    def show_list_view(self):
        """Show the list view of the schedule."""
        try:
            # First, navigate to the View Schedule tab
            if hasattr(self, 'tab_indices') and 'schedule' in self.tab_indices:
                self.notebook.select(self.tab_indices["schedule"])
                
            # If we have a content_frame, update the view
            if hasattr(self, 'content_frame'):
                # Remove all widgets from content frame
                for widget in self.content_frame.winfo_children():
                    widget.pack_forget()
                
                # Show list frame if it exists
                if hasattr(self, 'list_frame'):
                    self.list_frame.pack(fill=tk.BOTH, expand=True)
                    
                # Update view variable if it exists
                if hasattr(self, 'view_var'):
                    self.view_var.set("list")
        except Exception as e:
            print(f"Error showing list view: {e}")
            import traceback
            traceback.print_exc()
    
    def show_calendar_view(self):
        """Show the calendar view of the schedule."""
        try:
            # First, navigate to the View Schedule tab
            if hasattr(self, 'tab_indices') and 'schedule' in self.tab_indices:
                self.notebook.select(self.tab_indices["schedule"])
                
            # Check if we have a schedule to display
            if hasattr(self, 'current_schedule') and self.current_schedule:
                # If we have a content_frame, update the view
                if hasattr(self, 'content_frame'):
                    # Remove all widgets from content frame
                    for widget in self.content_frame.winfo_children():
                        widget.pack_forget()
                    
                    # Show calendar frame if it exists
                    if hasattr(self, 'calendar_frame'):
                        self.calendar_frame.pack(fill=tk.BOTH, expand=True)
                        
                        # Update calendar with current schedule
                        self.update_calendar_view(self.current_schedule)
                    
                    # Update view variable if it exists
                    if hasattr(self, 'view_var'):
                        self.view_var.set("calendar")
            else:
                # If no schedule is selected, show an error and switch to list view
                messagebox.showerror("Error", "Please select a schedule first by double-clicking on it in the list view.")
                if hasattr(self, 'view_var'):
                    self.view_var.set("list")
                self.show_list_view()
        except Exception as e:
            print(f"Error showing calendar view: {e}")
            import traceback
            traceback.print_exc()
            
    def update_calendar_view(self, schedule):
        """Update the calendar view with the given schedule."""
        try:
            # Clear the canvas
            self.calendar_canvas.delete("all")
            
            if not schedule:
                return
                
            # Make sure we have the right structure for the schedule items
            if not all(isinstance(item, dict) for item in schedule):
                print("Error: schedule items are not in dictionary format")
                return
                
            # Sort schedule by date
            schedule.sort(key=lambda x: (x["exam_date"], x["start_time"]))
            
            # Get unique dates
            dates = sorted(list(set([item["exam_date"] for item in schedule])))
            if not dates:
                return
                
            # Set canvas size and properties
            day_width = 220  # Increased width for better spacing
            hour_height = 65  # Increased height for better visibility
            header_height = 70  # Increased header height
            day_start_hour = 8  # 8 AM
            day_end_hour = 20   # 8 PM
            total_hours = day_end_hour - day_start_hour
            
            canvas_width = len(dates) * day_width
            canvas_height = total_hours * hour_height + header_height
            
            # Set background color for the entire canvas
            self.calendar_canvas.config(scrollregion=(0, 0, canvas_width, canvas_height), bg="#F5F7FA")
            
            # Draw background grid
            # First draw a light background for the entire grid
            self.calendar_canvas.create_rectangle(0, header_height, canvas_width, canvas_height, 
                                               fill="#FFFFFF", outline="")
            
            # Draw time labels and horizontal hour lines
            for hour in range(day_start_hour, day_end_hour + 1):
                y = header_height + (hour - day_start_hour) * hour_height
                
                # Create alternating background for better readability
                if (hour - day_start_hour) % 2 == 0:
                    self.calendar_canvas.create_rectangle(0, y, canvas_width, y + hour_height, 
                                                      fill="#F9F9F9", outline="")
                
                # Draw hour line
                self.calendar_canvas.create_line(0, y, canvas_width, y, fill="#D0D0D0", width=1)
                
                # Format time with AM/PM
                if hour < 12:
                    time_str = f"{hour:02d}:00 AM"
                elif hour == 12:
                    time_str = "12:00 PM"
                else:
                    time_str = f"{hour-12:02d}:00 PM"
                    
                # Draw time label with better styling
                self.calendar_canvas.create_text(15, y + 10, text=time_str, 
                                             anchor="w", font=("Helvetica", 9), fill="#555555")
            
            # Draw day headers with improved styling
            for i, date_str in enumerate(dates):
                try:
                    date_obj = datetime.strptime(str(date_str), "%Y-%m-%d")
                    day_name = date_obj.strftime("%A")
                    date_display = date_obj.strftime("%d-%m-%Y")
                    
                    # Calculate positions
                    x_start = i * day_width
                    x_end = (i + 1) * day_width
                    x_center = x_start + day_width // 2
                    
                    # Draw header background with gradient-like effect
                    header_colors = ["#4A89DC", "#5D9CEC", "#4FC1E9", "#48CFAD", "#A0D468", "#FFCE54", "#FC6E51"]
                    color_index = i % len(header_colors)
                    header_color = header_colors[color_index]
                    
                    # Main header background
                    self.calendar_canvas.create_rectangle(x_start, 0, x_end, header_height, 
                                                      fill=header_color, outline="")
                    
                    # Add a slight shadow effect at the bottom of the header
                    self.calendar_canvas.create_rectangle(x_start, header_height-3, x_end, header_height, 
                                                      fill="#E0E0E0", outline="")
                    
                    # Draw day name with shadow effect for better readability
                    self.calendar_canvas.create_text(x_center+1, header_height//3+1, 
                                                 text=day_name, font=("Helvetica", 12, "bold"),
                                                 fill="#AAAAAA")
                    self.calendar_canvas.create_text(x_center, header_height//3, 
                                                 text=day_name, font=("Helvetica", 12, "bold"),
                                                 fill="white")
                    
                    # Draw date with shadow effect
                    self.calendar_canvas.create_text(x_center+1, header_height*2//3+1, 
                                                 text=date_display, font=("Helvetica", 10),
                                                 fill="#AAAAAA")
                    self.calendar_canvas.create_text(x_center, header_height*2//3, 
                                                 text=date_display, font=("Helvetica", 10),
                                                 fill="white")
                    
                    # Draw vertical line for day separation (slightly thinner and more subtle)
                    if i > 0:  # Don't draw line for the first column
                        self.calendar_canvas.create_line(x_start, 0, x_start, canvas_height, 
                                                      fill="#D0D0D0", width=1, dash=(4, 4))
                except Exception as e:
                    print(f"Error drawing day header: {e}")
            
            # Draw exams with improved styling
            # Define a color palette for different subjects
            subject_colors = [
                ("#8CC152", "#A0D468", "#2C3E50"),  # Green
                ("#4A89DC", "#5D9CEC", "#2C3E50"),  # Blue
                ("#967ADC", "#AC92EC", "#2C3E50"),  # Purple
                ("#D770AD", "#EC87C0", "#2C3E50"),  # Pink
                ("#F5D76E", "#FFCE54", "#2C3E50"),  # Yellow
                ("#FC6E51", "#E9573F", "#FFFFFF"),  # Orange/Red
                ("#3BAFDA", "#4FC1E9", "#2C3E50"),  # Light Blue
                ("#37BC9B", "#48CFAD", "#2C3E50"),  # Teal
            ]
            
            # Create a dictionary to consistently assign colors to subjects
            subject_color_map = {}
            color_index = 0
            
            for item in schedule:
                try:
                    date_str = item["exam_date"]
                    start_time_str = str(item["start_time"])
                    end_time_str = str(item["end_time"])
                    
                    # Parse times
                    if ":" in start_time_str:
                        if "AM" in start_time_str or "PM" in start_time_str:
                            start_time = datetime.strptime(start_time_str, "%I:%M %p")
                        else:
                            start_time = datetime.strptime(start_time_str, "%H:%M:%S")
                    else:
                        # Default time if parsing fails
                        start_time = datetime.strptime("08:00:00", "%H:%M:%S")
                        
                    if ":" in end_time_str:
                        if "AM" in end_time_str or "PM" in end_time_str:
                            end_time = datetime.strptime(end_time_str, "%I:%M %p")
                        else:
                            end_time = datetime.strptime(end_time_str, "%H:%M:%S")
                    else:
                        # Default time if parsing fails
                        end_time = datetime.strptime("10:00:00", "%H:%M:%S")
                    
                    # Calculate position
                    date_index = dates.index(date_str)
                    start_hour = start_time.hour + start_time.minute / 60
                    end_hour = end_time.hour + end_time.minute / 60
                    
                    # Adjust for calendar display (8 AM to 8 PM)
                    if start_hour < day_start_hour:
                        start_hour = day_start_hour
                    if end_hour > day_end_hour:
                        end_hour = day_end_hour
                    
                    # Calculate coordinates with padding
                    x1 = date_index * day_width + 8
                    y1 = header_height + (start_hour - day_start_hour) * hour_height + 5
                    x2 = (date_index + 1) * day_width - 8
                    y2 = header_height + (end_hour - day_start_hour) * hour_height - 5
                    
                    # Get subject details
                    subject_code = item.get("subject_code", "Unknown Code")
                    subject_name = item.get("subject_name", "Unknown Subject")
                    room_name = item.get("room_name", "Unknown Room")
                    
                    # Assign consistent colors to subjects
                    if subject_code not in subject_color_map:
                        subject_color_map[subject_code] = subject_colors[color_index % len(subject_colors)]
                        color_index += 1
                    
                    border_color, fill_color, text_color = subject_color_map[subject_code]
                    
                    # Draw rounded rectangle for exam (simulate rounded corners with arcs)
                    corner_radius = 10
                    
                    # Create shadow effect
                    shadow_offset = 3
                    self.calendar_canvas.create_rectangle(
                        x1 + shadow_offset, y1 + shadow_offset, 
                        x2 + shadow_offset, y2 + shadow_offset,
                        fill="#E0E0E0", outline=""
                    )
                    
                    # Main rectangle
                    self.calendar_canvas.create_rectangle(x1, y1, x2, y2, fill=fill_color, outline=border_color, width=2)
                    
                    # Add a subtle highlight at the top of the exam block
                    gradient_height = (y2 - y1) // 4
                    highlight_color = "#FFFFFF" if "#" in fill_color else "white"
                    self.calendar_canvas.create_rectangle(
                        x1 + 2, y1 + 2, x2 - 2, y1 + gradient_height,
                        fill=highlight_color, outline="", width=0
                    )
                    
                    # Calculate text position
                    text_x = x1 + (x2-x1)/2
                    text_y = y1 + (y2-y1)/2 - 5  # Slight adjustment for visual balance
                    
                    # Format time for display (12-hour format with AM/PM)
                    if isinstance(start_time, datetime):
                        start_time_display = start_time.strftime("%I:%M %p").lstrip("0")
                    else:
                        start_time_display = str(start_time)
                        
                    if isinstance(end_time, datetime):
                        end_time_display = end_time.strftime("%I:%M %p").lstrip("0")
                    else:
                        end_time_display = str(end_time)
                    
                    # Format the display text with improved layout
                    display_text = f"{subject_code}\n{subject_name}\nRoom: {room_name}\n{start_time_display} - {end_time_display}"
                    
                    # Draw text with shadow for better readability
                    self.calendar_canvas.create_text(
                        text_x+1, text_y+1, 
                        text=display_text, 
                        width=day_width-30, 
                        justify="center", 
                        font=("Helvetica", 9),
                        fill="#AAAAAA"
                    )
                    
                    self.calendar_canvas.create_text(
                        text_x, text_y, 
                        text=display_text, 
                        width=day_width-30, 
                        justify="center", 
                        font=("Helvetica", 9, "bold"),
                        fill=text_color
                    )
                    
                except Exception as e:
                    print(f"Error drawing exam on calendar: {e}")
                    print(f"Problem item: {item}")
                    
        except Exception as e:
            print(f"Error updating calendar view: {e}")
            import traceback
            traceback.print_exc()

        
    def init_generator_tab(self):
        """Initialize the generator tab."""
        self.generator_frame = ttk.Frame(self.generator_tab)
        self.generator_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create a notebook for the generator sections
        self.generator_notebook = ttk.Notebook(self.generator_frame)
        self.generator_notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs for different sections of the generator
        basic_settings_frame = ttk.Frame(self.generator_notebook)
        advanced_settings_frame = ttk.Frame(self.generator_notebook)
        preview_frame = ttk.Frame(self.generator_notebook)
        
        self.generator_notebook.add(basic_settings_frame, text="Basic Settings")
        self.generator_notebook.add(advanced_settings_frame, text="Advanced Settings")
        self.generator_notebook.add(preview_frame, text="Preview")
        
        # Store the index of the preview tab
        self.preview_tab_index = 2  # Index of the Preview tab
        
        # Initialize each section
        self.init_basic_settings(basic_settings_frame)
        self.init_advanced_settings(advanced_settings_frame)
        self.init_preview_section(preview_frame)
        
        # Create bottom control panel with generate button
        control_frame = ttk.Frame(self.generator_frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        # Progress bar
        self.progress_var = tk.IntVar()
        self.progress_var.set(0)
        progress_bar = ttk.Progressbar(control_frame, variable=self.progress_var, maximum=100)
        progress_bar.pack(fill=tk.X, padx=10, pady=5)
        
        # Button frame for generate and reset buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(pady=10)
        
        # Generate button
        generate_btn = ttk.Button(button_frame, text="Generate Exam Schedule", 
                                command=self.generate_preview, style="Accent.TButton")
        generate_btn.pack(side=tk.LEFT, padx=10)
        
        # Reset button
        reset_btn = ttk.Button(button_frame, text="Reset Form", 
                             command=self.reset_form, style="Reset.TButton")
        reset_btn.pack(side=tk.LEFT, padx=10)
        
        # Register custom styles for buttons
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Helvetica", 12, "bold"), padding=10)
        style.configure("Reset.TButton", font=("Helvetica", 12), padding=10, foreground="red")
    
    def reset_form(self):
        """Reset all form fields in the generator tab."""
        try:
            # Confirm reset with user
            confirm = messagebox.askyesno("Confirm Reset", "Are you sure you want to reset all form fields? This will clear all your current selections.")
            if not confirm:
                return
                
            # Reset basic settings
            if hasattr(self, 'schedule_name_var'):
                self.schedule_name_var.set("")
                
            if hasattr(self, 'exam_type_var'):
                self.exam_type_var.set("Regular")
                
            if hasattr(self, 'semester_var') and hasattr(self, 'semester_combo') and self.semester_combo['values']:
                self.semester_var.set(self.semester_combo['values'][0])
                
            # Reset date settings
            today = datetime.now()
            if hasattr(self, 'start_day_var'):
                self.start_day_var.set(today.day)
            if hasattr(self, 'start_month_var'):
                self.start_month_var.set(today.month)
            if hasattr(self, 'start_year_var'):
                self.start_year_var.set(today.year)
                
            # Reset subject selections
            if hasattr(self, 'selected_subjects'):
                # Move all selected subjects back to available
                if hasattr(self, 'available_subjects'):
                    self.available_subjects.extend(self.selected_subjects)
                    self.selected_subjects = []
                    
                    # Update listboxes
                    if hasattr(self, 'available_subjects_var'):
                        self.available_subjects_var.set([subject[1] for subject in self.available_subjects])
                    if hasattr(self, 'selected_subjects_var'):
                        self.selected_subjects_var.set([])
            
            # Reset room selections
            if hasattr(self, 'room_select_tree'):
                for item in self.room_select_tree.get_children():
                    values = list(self.room_select_tree.item(item, 'values'))
                    if values[-1] == "Yes":
                        values[-1] = "No"
                        self.room_select_tree.item(item, values=values)
            
            # Reset advanced settings
            if hasattr(self, 'allow_multiple_exams_var'):
                self.allow_multiple_exams_var.set(False)
                
            if hasattr(self, 'hard_gap_var'):
                self.hard_gap_var.set(1)
                
            if hasattr(self, 'medium_gap_var'):
                self.medium_gap_var.set(0)
                
            if hasattr(self, 'theory_time_var'):
                self.theory_time_var.set("09:00 AM - 12:00 PM")
                
            if hasattr(self, 'practical_time_var'):
                self.practical_time_var.set("02:00 PM - 05:00 PM")
                
            # Clear preview tree
            if hasattr(self, 'preview_tree'):
                for item in self.preview_tree.get_children():
                    self.preview_tree.delete(item)
                    
            # Reset to first tab
            if hasattr(self, 'generator_notebook'):
                self.generator_notebook.select(0)
                
            messagebox.showinfo("Reset Complete", "All form fields have been reset to their default values.")
            
        except Exception as e:
            print(f"Error resetting form: {str(e)}")
            messagebox.showerror("Error", f"An error occurred while resetting the form: {str(e)}")
    
    def init_basic_settings(self, parent):
        """Initialize the basic settings section of the generator."""
        # Create a frame with padding
        frame = ttk.Frame(parent, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Schedule name and type
        ttk.Label(frame, text="Schedule Information", font=("Helvetica", 12, "bold")).grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(frame, text="Schedule Name:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.schedule_name_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.schedule_name_var, width=40).grid(
            row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Exam Type:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.exam_type_var = tk.StringVar(value="Regular")
        ttk.Combobox(frame, textvariable=self.exam_type_var, 
                   values=["Regular", "Remedial", "Internal", "External"], 
                   width=38, state="readonly").grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Semester:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.semester_var = tk.StringVar()
        
        # Get available semesters from subjects table
        try:
            self.cursor.execute("SELECT DISTINCT semester FROM subjects ORDER BY semester")
            semesters = [sem[0] for sem in self.cursor.fetchall()]
            if semesters:
                self.semester_var.set(semesters[0])
        except Error as e:
            semesters = []
        
        semester_combo = ttk.Combobox(frame, textvariable=self.semester_var, 
                                    values=semesters, width=38, state="readonly")
        semester_combo.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Date settings
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(
            row=4, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        ttk.Label(frame, text="Date Settings", font=("Helvetica", 12, "bold")).grid(
            row=5, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(frame, text="Start Date:").grid(row=6, column=0, sticky=tk.W, pady=5)
        
        # Create a date picker frame
        date_frame = ttk.Frame(frame)
        date_frame.grid(row=6, column=1, sticky=tk.W, pady=5)
        
        # Day, Month, Year dropdowns
        self.start_day_var = tk.IntVar(value=datetime.now().day)
        self.start_month_var = tk.IntVar(value=datetime.now().month)
        self.start_year_var = tk.IntVar(value=datetime.now().year)
        
        days = list(range(1, 32))
        months = list(range(1, 13))
        years = list(range(datetime.now().year, datetime.now().year + 5))
        
        ttk.Combobox(date_frame, textvariable=self.start_day_var, 
                    values=days, width=5, state="readonly").pack(side=tk.LEFT, padx=2)
        
        ttk.Combobox(date_frame, textvariable=self.start_month_var, 
                    values=months, width=5, state="readonly").pack(side=tk.LEFT, padx=2)
        
        ttk.Combobox(date_frame, textvariable=self.start_year_var, 
                    values=years, width=7, state="readonly").pack(side=tk.LEFT, padx=2)
        
        # Subject selection
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(
            row=7, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        ttk.Label(frame, text="Subject Selection", font=("Helvetica", 12, "bold")).grid(
            row=8, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Create a frame for the subject selection
        subjects_frame = ttk.Frame(frame)
        subjects_frame.grid(row=9, column=0, columnspan=2, sticky=tk.NSEW, pady=5)
        
        # Make the subjects frame expandable
        frame.rowconfigure(9, weight=1)
        frame.columnconfigure(1, weight=1)
        
        # Create two listboxes with available and selected subjects
        list_frame = ttk.Frame(subjects_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Available subjects list
        ttk.Label(list_frame, text="Available Subjects:").grid(row=0, column=0, pady=5)
        self.available_subjects_var = tk.StringVar()
        self.available_subjects_list = tk.Listbox(list_frame, listvariable=self.available_subjects_var, 
                                               selectmode=tk.MULTIPLE, height=10, width=40)
        self.available_subjects_list.grid(row=1, column=0, padx=5, pady=5, sticky=tk.NSEW)
        avail_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, 
                                      command=self.available_subjects_list.yview)
        avail_scrollbar.grid(row=1, column=1, sticky=tk.NS)
        self.available_subjects_list.configure(yscrollcommand=avail_scrollbar.set)
        
        # Buttons to move subjects between lists
        btn_frame = ttk.Frame(list_frame)
        btn_frame.grid(row=1, column=2, padx=10)
        
        ttk.Button(btn_frame, text=">", command=self.add_selected_subjects).pack(pady=5)
        ttk.Button(btn_frame, text=">>", command=self.add_all_subjects).pack(pady=5)
        ttk.Button(btn_frame, text="<", command=self.remove_selected_subjects).pack(pady=5)
        ttk.Button(btn_frame, text="<<", command=self.remove_all_subjects).pack(pady=5)
        
        # Selected subjects list
        ttk.Label(list_frame, text="Selected Subjects:").grid(row=0, column=3, pady=5)
        self.selected_subjects_var = tk.StringVar()
        self.selected_subjects_list = tk.Listbox(list_frame, listvariable=self.selected_subjects_var, 
                                              selectmode=tk.MULTIPLE, height=10, width=40)
        self.selected_subjects_list.grid(row=1, column=3, padx=5, pady=5, sticky=tk.NSEW)
        sel_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, 
                                    command=self.selected_subjects_list.yview)
        sel_scrollbar.grid(row=1, column=4, sticky=tk.NS)
        self.selected_subjects_list.configure(yscrollcommand=sel_scrollbar.set)
        
        # Make the listboxes expandable
        list_frame.columnconfigure(0, weight=1)
        list_frame.columnconfigure(3, weight=1)
        list_frame.rowconfigure(1, weight=1)
        
        # Update subject lists when semester changes
        self.semester_var.trace("w", self.update_subject_lists)
        
        # Initial update of subject lists
        self.update_subject_lists()
    
    def init_advanced_settings(self, frame):
        """Initialize the advanced settings section of the generator tab."""
        # Create a frame for settings with padding
        settings_frame = ttk.Frame(frame, padding=10)
        settings_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid
        settings_frame.columnconfigure(0, weight=0)
        settings_frame.columnconfigure(1, weight=1)
        
        # Gap between difficult exams
        ttk.Label(settings_frame, text="Exam Spacing Settings", font=("Helvetica", 12, "bold")).grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(settings_frame, text="Gap between difficult exams (days):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.difficult_gap_var = tk.IntVar(value=2)
        ttk.Spinbox(settings_frame, from_=0, to=10, textvariable=self.difficult_gap_var, width=5).grid(
            row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(settings_frame, text="Gap between medium difficulty exams (days):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.medium_gap_var = tk.IntVar(value=1)
        ttk.Spinbox(settings_frame, from_=0, to=5, textvariable=self.medium_gap_var, width=5).grid(
            row=2, column=1, sticky=tk.W, pady=5)
        
        # Room allocation settings
        ttk.Separator(settings_frame, orient=tk.HORIZONTAL).grid(
            row=3, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        ttk.Label(settings_frame, text="Room Allocation Settings", font=("Helvetica", 12, "bold")).grid(
            row=4, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Room selection
        ttk.Label(settings_frame, text="Select Rooms:").grid(row=5, column=0, sticky=tk.W, pady=5)
        
        # Create a frame for the room selection treeview
        room_frame = ttk.Frame(settings_frame)
        room_frame.grid(row=6, column=0, columnspan=2, sticky=tk.NSEW, pady=5)
        
        # Make the room frame expandable
        settings_frame.rowconfigure(6, weight=1)
        
        # Room selection treeview
        columns = ("id", "name", "type", "capacity", "selected")
        self.room_select_tree = ttk.Treeview(room_frame, columns=columns, show="headings", height=10)
        
        # Configure columns
        self.room_select_tree.heading("id", text="ID")
        self.room_select_tree.heading("name", text="Room Name")
        self.room_select_tree.heading("type", text="Type")
        self.room_select_tree.heading("capacity", text="Capacity")
        self.room_select_tree.heading("selected", text="Selected")
        
        # Set column widths
        self.room_select_tree.column("id", width=50)
        self.room_select_tree.column("name", width=150)
        self.room_select_tree.column("type", width=100)
        self.room_select_tree.column("capacity", width=80)
        self.room_select_tree.column("selected", width=80)
        
        # Add scrollbar
        room_scrollbar = ttk.Scrollbar(room_frame, orient=tk.VERTICAL, command=self.room_select_tree.yview)
        self.room_select_tree.configure(yscrollcommand=room_scrollbar.set)
        
        # Pack treeview and scrollbar
        self.room_select_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        room_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind click event to toggle selection
        self.room_select_tree.bind("<ButtonRelease-1>", self.toggle_room_selection)
        
        # Load rooms
        self.load_rooms_for_selection()
    
    def init_preview_section(self, frame):
        """Initialize the preview section of the generator tab."""
        # Create a frame for the preview with padding
        preview_frame = ttk.Frame(frame, padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)
        
        # Preview header
        ttk.Label(preview_frame, text="Schedule Preview", font=("Helvetica", 12, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Preview treeview
        preview_tree_frame = ttk.Frame(preview_frame)
        preview_tree_frame.grid(row=1, column=0, sticky=tk.NSEW)
        
        columns = ("date", "subject_code", "subject_name", "room", "time")
        self.preview_tree = ttk.Treeview(preview_tree_frame, columns=columns, show="headings", height=15)
        
        # Configure columns
        self.preview_tree.heading("date", text="Date")
        self.preview_tree.heading("subject_code", text="Subject Code")
        self.preview_tree.heading("subject_name", text="Subject Name")
        self.preview_tree.heading("room", text="Room")
        self.preview_tree.heading("time", text="Time")
        
        # Set column widths
        self.preview_tree.column("date", width=100)
        self.preview_tree.column("subject_code", width=100)
        self.preview_tree.column("subject_name", width=200)
        self.preview_tree.column("room", width=100)
        self.preview_tree.column("time", width=100)
        
        # Add scrollbars
        y_scrollbar = ttk.Scrollbar(preview_tree_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=y_scrollbar.set)
        
        x_scrollbar = ttk.Scrollbar(preview_tree_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        self.preview_tree.configure(xscrollcommand=x_scrollbar.set)
        
        # Pack treeview and scrollbars
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Bind double-click event to edit schedule item
        self.preview_tree.bind("<Double-1>", self.on_preview_item_double_click)
        
        # Bind right-click event to show context menu
        self.preview_tree.bind("<Button-3>", self.show_preview_context_menu)
        
        # Preview controls
        controls_frame = ttk.Frame(preview_frame)
        controls_frame.grid(row=2, column=0, sticky=tk.EW, pady=10)
        
        # Generate preview button
        ttk.Button(controls_frame, text="Generate Preview", command=self.generate_preview).pack(side=tk.LEFT, padx=5)
        
        # Add button
        ttk.Button(controls_frame, text="Add Exam", command=self.open_add_dialog).pack(side=tk.LEFT, padx=5)
        
        # Export buttons
        ttk.Button(controls_frame, text="Export to CSV", command=self.export_schedule_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(controls_frame, text="Export to PDF", command=self.export_schedule_pdf).pack(side=tk.LEFT, padx=5)
        
    def load_rooms_for_selection(self):
        """Load rooms for selection in the advanced settings tab."""
        # Clear existing items
        for item in self.room_select_tree.get_children():
            self.room_select_tree.delete(item)
        
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return
                
            # Fetch rooms from database - no parameters needed here
            try:
                print("Executing SELECT query for rooms in selection (main)...")
                self.cursor.execute("SELECT id, name, type, capacity FROM rooms ORDER BY type, name")
                rooms = self.cursor.fetchall()
                if rooms is None:  # Handle None result
                    rooms = []
                print(f"Found {len(rooms)} rooms for selection (main)")
            except Error as e:
                print(f"Error fetching rooms for selection (main): {str(e)}")
                messagebox.showerror("Database Error", f"Failed to fetch rooms for selection: {str(e)}")
                return
            
            # Add rooms to treeview with selection status
            for room in rooms:
                try:
                    self.room_select_tree.insert("", tk.END, values=(*room, "No"))
                except Exception as e:
                    print(f"Error adding room to treeview (main): {str(e)}")
                    continue
            
        except Error as e:
            print(f"Error loading rooms for selection (main): {str(e)}")
            messagebox.showerror("Database Error", f"Failed to load rooms: {str(e)}")
    
    def update_subject_lists(self, *args):
        """Update the available and selected subjects lists based on the selected semester."""
        try:
            # Ensure database connection is active
            if not self.ensure_connection():
                return
                
            # Clear current lists
            self.available_subjects = []
            self.selected_subjects = []
            
            # Get subjects for the selected semester
            semester = self.semester_var.get()
            if semester:
                try:
                    print(f"Fetching subjects for semester: {semester}")
                    self.cursor.execute(
                        "SELECT id, code, name, type FROM subjects WHERE semester = %s ORDER BY name", 
                        (semester,))
                    subjects = self.cursor.fetchall()
                    if subjects is None:  # Handle None result
                        subjects = []
                    print(f"Found {len(subjects)} subjects for semester {semester}")
                    
                    # If no subjects found, try to get all subjects
                    if not subjects:
                        print("No subjects found for selected semester, fetching all subjects")
                        self.cursor.execute(
                            "SELECT id, code, name, type FROM subjects ORDER BY name")
                        subjects = self.cursor.fetchall() or []
                        print(f"Found {len(subjects)} subjects total")
                    
                    # Format subjects for display
                    for subject in subjects:
                        # Make sure we have valid data
                        if subject and len(subject) >= 4:
                            # Create a tuple with (id, display_string)
                            subject_id = subject[0]
                            code = subject[1] or ""
                            name = subject[2] or ""
                            subject_type = subject[3] or ""
                            display_string = f"{code} - {name} ({subject_type})"
                            
                            self.available_subjects.append((subject_id, display_string))
                    
                    # Update available subjects listbox
                    self.available_subjects_var.set([subject[1] for subject in self.available_subjects])
                    
                    # Clear selected subjects listbox
                    self.selected_subjects_var.set([])
                    
                    print(f"Populated available subjects list with {len(self.available_subjects)} subjects")
                    
                    # If there are 2 or fewer subjects, automatically select them all
                    if len(self.available_subjects) <= 2:
                        print("Auto-selecting all available subjects since there are only a few")
                        self.add_all_subjects()
                except Error as e:
                    print(f"Error fetching subjects for semester: {str(e)}")
                    messagebox.showerror("Database Error", f"Failed to fetch subjects for semester: {str(e)}")
        
        except Exception as e:
            print(f"Error updating subject lists: {str(e)}")
            messagebox.showerror("Database Error", f"Failed to load subjects: {str(e)}")
    
    def add_selected_subjects(self):
        """Add selected subjects from available list to selected list."""
        selected_indices = self.available_subjects_list.curselection()
        for i in reversed(selected_indices):  # Reverse to avoid index shifting
            if i < len(self.available_subjects):
                subject = self.available_subjects.pop(i)
                self.selected_subjects.append(subject)
        
        # Update both listboxes
        self.available_subjects_var.set([subject[1] for subject in self.available_subjects])
        self.selected_subjects_var.set([subject[1] for subject in self.selected_subjects])
    
    def add_all_subjects(self):
        """Add all subjects from available list to selected list."""
        self.selected_subjects.extend(self.available_subjects)
        self.available_subjects = []
        
        # Update both listboxes
        self.available_subjects_var.set([])
        self.selected_subjects_var.set([subject[1] for subject in self.selected_subjects])
    
    def remove_selected_subjects(self):
        """Remove selected subjects from selected list back to available list."""
        selected_indices = self.selected_subjects_list.curselection()
        for i in reversed(selected_indices):  # Reverse to avoid index shifting
            if i < len(self.selected_subjects):
                subject = self.selected_subjects.pop(i)
                self.available_subjects.append(subject)
        
        # Update both listboxes
        self.available_subjects_var.set([subject[1] for subject in self.available_subjects])
        self.selected_subjects_var.set([subject[1] for subject in self.selected_subjects])
    
    def remove_all_subjects(self):
        """Remove all subjects from selected list back to available list."""
        self.available_subjects.extend(self.selected_subjects)
        self.selected_subjects = []
        
        # Update both listboxes
        self.available_subjects_var.set([subject[1] for subject in self.available_subjects])
        self.selected_subjects_var.set([])
    
    def add_selected_subjects(self):
        """Add selected subjects from available list to selected list."""
        selected_indices = self.available_subjects_list.curselection()
        for i in reversed(selected_indices):  # Reverse to avoid index shifting
            if i < len(self.available_subjects):
                subject = self.available_subjects.pop(i)
                self.selected_subjects.append(subject)
        
        # Update both listboxes
        self.available_subjects_var.set([subject[1] for subject in self.available_subjects])
        self.selected_subjects_var.set([subject[1] for subject in self.selected_subjects])

    def add_all_subjects(self):
        """Add all subjects from available list to selected list."""
        self.selected_subjects.extend(self.available_subjects)
        self.available_subjects = []
        
        # Update both listboxes
        self.available_subjects_var.set([])
        self.selected_subjects_var.set([subject[1] for subject in self.selected_subjects])

    def remove_selected_subjects(self):
        """Remove selected subjects from selected list back to available list."""
        selected_indices = self.selected_subjects_list.curselection()
        for i in reversed(selected_indices):  # Reverse to avoid index shifting
            if i < len(self.selected_subjects):
                subject = self.selected_subjects.pop(i)
                self.available_subjects.append(subject)
                
        # Update both listboxes
        self.available_subjects_var.set([subject[1] for subject in self.available_subjects])
        self.selected_subjects_var.set([subject[1] for subject in self.selected_subjects])
        
    def new_schedule(self):
        """Create a new schedule."""
        self.notebook.select(3)  # Switch to generator tab
        self.reset_form()  # Reset the generator form
        
    def reset_form(self):
        """Reset all form fields in the generator tab."""
        try:
            # Reset basic settings
            if hasattr(self, 'schedule_name_var'):
                self.schedule_name_var.set('')
            if hasattr(self, 'semester_var'):
                self.semester_var.set('')
            if hasattr(self, 'exam_type_var'):
                self.exam_type_var.set('Final')
            if hasattr(self, 'start_date_var'):
                self.start_date_var.set(datetime.now().strftime('%Y-%m-%d'))
            if hasattr(self, 'gap_days_var'):
                self.gap_days_var.set(1)
            
            # Reset advanced settings
            if hasattr(self, 'exclude_weekends_var'):
                self.exclude_weekends_var.set(True)
            if hasattr(self, 'max_exams_per_day_var'):
                self.max_exams_per_day_var.set(2)
            
            # Reset subject selections
            if hasattr(self, 'available_subjects') and hasattr(self, 'selected_subjects'):
                # Move all selected subjects back to available
                self.available_subjects.extend(self.selected_subjects)
                self.selected_subjects = []
                
                # Update listboxes
                if hasattr(self, 'available_subjects_var') and hasattr(self, 'selected_subjects_var'):
                    self.available_subjects_var.set([subject[1] for subject in self.available_subjects])
                    self.selected_subjects_var.set([])
            
            # Clear the preview treeview
            if hasattr(self, 'preview_tree'):
                for item in self.preview_tree.get_children():
                    self.preview_tree.delete(item)
            
            # Update status
            self.status_var.set("Ready to create a new schedule")
            
        except Exception as e:
            print(f"Error resetting form: {e}")
            import traceback
            traceback.print_exc()
    
    def open_schedule(self):
        """Open an existing schedule."""
        try:
            if not self.ensure_connection():
                return
            
            # Get list of saved schedules
            self.cursor.execute("""
                SELECT id, name, semester, exam_type, DATE_FORMAT(start_date, '%Y-%m-%d') as start_date
                FROM schedules
                ORDER BY created_at DESC
            """)
            schedules = self.cursor.fetchall()
            
            if not schedules:
                messagebox.showinfo("Info", "No saved schedules found.")
                return
            
            # Create dialog to select a schedule
            dialog = tk.Toplevel(self.root)
            dialog.title("Open Schedule")
            dialog.geometry("400x300")
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Create treeview
            tree = ttk.Treeview(dialog, columns=("id", "name", "semester", "exam_type", "start_date"), show="headings")
            tree.heading("id", text="ID")
            tree.heading("name", text="Name")
            tree.heading("semester", text="Semester")
            tree.heading("exam_type", text="Exam Type")
            tree.heading("start_date", text="Start Date")
            
            # Hide ID column
            tree.column("id", width=0, stretch=False)
            
            # Add schedules to treeview
            for schedule in schedules:
                tree.insert("", "end", values=schedule)
            
            tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            def on_select():
                selected = tree.selection()
                if not selected:
                    messagebox.showwarning("Warning", "Please select a schedule.")
                    return
                
                schedule_id = tree.item(selected[0])["values"][0]
                dialog.destroy()
                self.load_schedule(schedule_id)
            
            # Add buttons
            button_frame = ttk.Frame(dialog)
            button_frame.pack(fill=tk.X, padx=5, pady=5)
            
            ttk.Button(button_frame, text="Open", command=on_select).pack(side=tk.RIGHT, padx=5)
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load schedules: {str(e)}")
            
    def load_schedule(self, schedule_id):
        """Load a specific schedule by ID."""
        try:
            if not self.ensure_connection():
                return
            
            # Get schedule details
            self.cursor.execute("""
                SELECT name, semester, exam_type, DATE_FORMAT(start_date, '%Y-%m-%d') as start_date, config
                FROM schedules
                WHERE id = %s
            """, (schedule_id,))
            
            schedule = self.cursor.fetchone()
            if not schedule:
                messagebox.showerror("Error", "Schedule not found.")
                return
            
            name, semester, exam_type, start_date, config = schedule
            config = json.loads(config)
            
            # Get schedule items
            self.cursor.execute("""
                SELECT si.subject_id, si.room_id, si.exam_date, si.start_time, si.end_time, 
                       s.code as subject_code, s.name as subject_name, r.name as room_name
                FROM schedule_items si
                JOIN subjects s ON si.subject_id = s.id
                JOIN rooms r ON si.room_id = r.id
                WHERE si.schedule_id = %s
                ORDER BY si.exam_date, si.start_time
            """, (schedule_id,))
            
            items = self.cursor.fetchall()
            
            # Set form values
            self.schedule_name_var.set(name)
            self.semester_var.set(semester)
            self.exam_type_var.set(exam_type)
            
            # Parse start date and set individual day, month, year variables
            try:
                start_date_obj = datetime.strptime(start_date, "%Y-%m-%d")
                if hasattr(self, 'start_day_var'):
                    self.start_day_var.set(start_date_obj.day)
                if hasattr(self, 'start_month_var'):
                    self.start_month_var.set(start_date_obj.month)
                if hasattr(self, 'start_year_var'):
                    self.start_year_var.set(start_date_obj.year)
            except Exception as e:
                print(f"Error parsing start date: {str(e)}")
                # Continue even if date parsing fails
            
            # Set working hours and durations from config with safety checks
            if 'working_hours' in config and 'start' in config['working_hours'] and hasattr(self, 'start_time_var'):
                self.start_time_var.set(config['working_hours']['start'])
            
            if 'working_hours' in config and 'end' in config['working_hours'] and hasattr(self, 'end_time_var'):
                self.end_time_var.set(config['working_hours']['end'])
                
            if 'exam_duration' in config and hasattr(self, 'exam_duration_var'):
                self.exam_duration_var.set(config['exam_duration'])
                
            if 'break_duration' in config and hasattr(self, 'break_duration_var'):
                self.break_duration_var.set(config['break_duration'])
            
            # Create schedule items list with proper dictionary structure
            self.current_schedule = []
            for item in items:
                self.current_schedule.append({
                    'subject_id': item[0],
                    'room_id': item[1],
                    'exam_date': item[2],
                    'start_time': item[3],
                    'end_time': item[4],
                    'subject_code': item[5],
                    'subject_name': item[6],
                    'room_name': item[7]
                })
            
            # Update the schedule view
            self.update_schedule_view(self.current_schedule)
            
            # Store the current schedule ID for editing/adding items
            self.current_schedule_id = schedule_id
            
            # Navigate to the Preview tab in the Generator section
            self.show_schedule_in_generator_preview()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load schedule: {str(e)}")
    
    def save_schedule(self):
        """Save the current schedule."""
        if not hasattr(self, 'current_schedule') or not self.current_schedule:
            messagebox.showerror("Error", "No schedule to save.")
            return
        
        # Get schedule name
        schedule_name = self.schedule_name_var.get().strip()
        if not schedule_name:
            messagebox.showerror("Error", "Please enter a schedule name.")
            return
        
        try:
            if not self.ensure_connection():
                return
            
            # Get schedule parameters
            semester = self.semester_var.get()
            exam_type = self.exam_type_var.get()
            
            # Construct start date from individual components
            try:
                start_day = self.start_day_var.get()
                start_month = self.start_month_var.get()
                start_year = self.start_year_var.get()
                start_date = f"{start_year}-{start_month:02d}-{start_day:02d}"
            except Exception as e:
                print(f"Error constructing start date: {str(e)}")
                # Fallback to first exam date if available
                if hasattr(self, 'current_schedule') and self.current_schedule:
                    start_date = min(item["exam_date"] for item in self.current_schedule)
                    if hasattr(start_date, 'strftime'):
                        start_date = start_date.strftime("%Y-%m-%d")
                    else:
                        start_date = str(start_date)
                else:
                    # Last resort fallback to today's date
                    start_date = datetime.now().strftime("%Y-%m-%d")
            
            # Construct config object with available variables
            config_dict = {'working_hours': {}}
            
            # Add available settings to config
            if hasattr(self, 'start_time_var'):
                config_dict['working_hours']['start'] = self.start_time_var.get()
            else:
                config_dict['working_hours']['start'] = "09:00"
                
            if hasattr(self, 'end_time_var'):
                config_dict['working_hours']['end'] = self.end_time_var.get()
            else:
                config_dict['working_hours']['end'] = "17:00"
                
            if hasattr(self, 'exam_duration_var'):
                config_dict['exam_duration'] = self.exam_duration_var.get()
            else:
                config_dict['exam_duration'] = 180  # 3 hours default
                
            if hasattr(self, 'break_duration_var'):
                config_dict['break_duration'] = self.break_duration_var.get()
            else:
                config_dict['break_duration'] = 30  # 30 minutes default
                
            config = json.dumps(config_dict)
            
            # Insert into schedules table
            self.cursor.execute("""
                INSERT INTO schedules 
                (name, semester, exam_type, start_date, config, created_at, updated_at)
                VALUES (%s, %s, %s, %s, %s, NOW(), NOW())
            """, (schedule_name, semester, exam_type, start_date, config))
            
            schedule_id = self.cursor.lastrowid
            
            # Insert schedule items
            for item in self.current_schedule:
                self.cursor.execute("""
                    INSERT INTO schedule_items 
                    (schedule_id, subject_id, room_id, exam_date, start_time, end_time)
                    VALUES (%s, %s, %s, %s, %s, %s)
                """, (
                    schedule_id,
                    item['subject_id'],
                    item['room_id'],
                    item['exam_date'],
                    item['start_time'],
                    item['end_time']
                ))
            
            self.conn.commit()
            messagebox.showinfo("Success", "Schedule saved successfully!")
            
            # Refresh the schedule view
            self.load_schedule(schedule_id)
            
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Error", f"Failed to save schedule: {str(e)}")
            return
            exam_type = self.exam_type_var.get()
            start_date = min(item["exam_date"] for item in self.current_schedule).strftime("%Y-%m-%d")
            
            # Get configuration
            config = {
                "skip_sunday": self.skip_sunday_var.get(),
                "allow_multiple_exams": self.allow_multiple_exams_var.get(),
                "hard_gap": self.hard_gap_var.get(),
                "medium_gap": self.medium_gap_var.get(),
                "theory_time": self.theory_time_var.get(),
                "practical_time": self.practical_time_var.get()
            }
            
            # Check if schedule with same name exists
            self.cursor.execute(
                "SELECT id FROM schedules WHERE name = %s",
                (schedule_name,)
            )
            existing = self.cursor.fetchone()
            
            if existing:
                # Ask for confirmation to overwrite
                confirm = messagebox.askyesno(
                    "Confirm Overwrite",
                    f"Schedule '{schedule_name}' already exists. Do you want to overwrite it?"
                )
                
                if not confirm:
                    return
                
                # Delete existing schedule items
                self.cursor.execute(
                    "DELETE FROM schedule_items WHERE schedule_id = %s",
                    (existing[0],)
                )
                
                # Update schedule
                self.cursor.execute(
                    """
                    UPDATE schedules
                    SET semester = %s, exam_type = %s, start_date = %s, config = %s, updated_at = CURRENT_TIMESTAMP
                    WHERE id = %s
                    """,
                    (semester, exam_type, start_date, json.dumps(config), existing[0])
                )
                
                schedule_id = existing[0]
            else:
                # Insert new schedule
                self.cursor.execute(
                    """
                    INSERT INTO schedules (name, semester, exam_type, start_date, config)
                    VALUES (%s, %s, %s, %s, %s)
                    """,
                    (schedule_name, semester, exam_type, start_date, json.dumps(config))
                )
                
                schedule_id = self.cursor.lastrowid
            
            # Insert schedule items
            for item in self.current_schedule:
                self.cursor.execute(
                    """
                    INSERT INTO schedule_items (schedule_id, subject_id, room_id, exam_date, start_time, end_time)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    """,
                    (
                        schedule_id,
                        item["subject_id"],
                        item["room_id"],
                        item["exam_date"].strftime("%Y-%m-%d"),
                        item["start_time"],
                        item["end_time"]
                    )
                )
            
            self.conn.commit()
            messagebox.showinfo("Success", f"Schedule '{schedule_name}' saved successfully.")
            self.update_dashboard_counts()
        
        except Error as e:
            messagebox.showerror("Error", f"Failed to save schedule: {str(e)}")
    
    def import_subjects(self):
        """Import subjects from a CSV file."""
        file_path = filedialog.askopenfilename(
            title="Import Subjects",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Read CSV file
            df = pd.read_csv(file_path)
            
            # Check required columns
            required_columns = ["code", "name", "type", "semester", "difficulty", "duration"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                messagebox.showerror(
                    "Error",
                    f"CSV file is missing required columns: {', '.join(missing_columns)}"
                )
                return
            
            # Import subjects
            imported_count = 0
            updated_count = 0
            
            for _, row in df.iterrows():
                # Check if subject already exists
                self.cursor.execute(
                    "SELECT id FROM subjects WHERE code = %s",
                    (row["code"],)
                )
                existing = self.cursor.fetchone()
                
                if existing:
                    # Update existing subject
                    self.cursor.execute(
                        """
                        UPDATE subjects
                        SET name = %s, type = %s, semester = %s, difficulty = %s, duration = %s, updated_at = CURRENT_TIMESTAMP
                        WHERE code = %s
                        """,
                        (row["name"], row["type"], row["semester"], row["difficulty"], row["duration"], row["code"])
                    )
                    updated_count += 1
                else:
                    # Insert new subject
                    self.cursor.execute(
                        """
                        INSERT INTO subjects (code, name, type, semester, difficulty, duration)
                        VALUES (%s, %s, %s, %s, %s, %s)
                        """,
                        (row["code"], row["name"], row["type"], row["semester"], row["difficulty"], row["duration"])
                    )
                    imported_count += 1
            
            self.conn.commit()
            self.load_subjects()
            self.update_semester_filter_options()
            self.update_dashboard_counts()
            messagebox.showinfo("Success", f"Imported {imported_count} new subjects and updated {updated_count} existing subjects from CSV file.")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import subjects: {str(e)}")
    
    def export_subjects(self):
        """Export subjects to a CSV file."""
        file_path = filedialog.asksaveasfilename(
            title="Export Subjects",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Get all subjects
            self.cursor.execute(
                """
                SELECT code, name, type, semester, difficulty, duration
                FROM subjects
                ORDER BY semester, code
                """
            )
            subjects = self.cursor.fetchall()
            
            # Create DataFrame
            df = pd.DataFrame(
                subjects,
                columns=["code", "name", "type", "semester", "difficulty", "duration"]
            )
            
            # Export to CSV
            df.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Exported {len(df)} subjects to CSV file.")
        
        except Error as e:
            messagebox.showerror("Error", f"Failed to export subjects: {str(e)}")
    
    def show_preferences(self):
        """Show the preferences dialog."""
        try:
            # Create a top-level dialog
            pref_dialog = tk.Toplevel(self.root)
            pref_dialog.title("Preferences")
            pref_dialog.geometry("500x450")
            pref_dialog.resizable(False, False)
            pref_dialog.transient(self.root)  # Set to be on top of the main window
            pref_dialog.grab_set()  # Make it modal
            
            # Create a notebook for different preference categories
            pref_notebook = ttk.Notebook(pref_dialog)
            pref_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Database settings tab
            db_frame = ttk.Frame(pref_notebook, padding=20)
            pref_notebook.add(db_frame, text="Database Settings")
            
            # Configure grid
            db_frame.columnconfigure(0, weight=0)
            db_frame.columnconfigure(1, weight=1)
            
            # Read current settings
            config = configparser.ConfigParser()
            config.read(CONFIG_FILE)
            
            # Database host
            ttk.Label(db_frame, text="Database Host:").grid(row=0, column=0, sticky=tk.W, pady=10)
            host_var = tk.StringVar(value=config.get('Database', 'host', fallback='localhost'))
            ttk.Entry(db_frame, textvariable=host_var, width=30).grid(row=0, column=1, sticky=tk.W, pady=10)
            
            # Database user
            ttk.Label(db_frame, text="Database User:").grid(row=1, column=0, sticky=tk.W, pady=10)
            user_var = tk.StringVar(value=config.get('Database', 'user', fallback='root'))
            ttk.Entry(db_frame, textvariable=user_var, width=30).grid(row=1, column=1, sticky=tk.W, pady=10)
            
            # Database password
            ttk.Label(db_frame, text="Database Password:").grid(row=2, column=0, sticky=tk.W, pady=10)
            password_var = tk.StringVar(value=config.get('Database', 'password', fallback=''))
            password_entry = ttk.Entry(db_frame, textvariable=password_var, width=30, show='*')
            password_entry.grid(row=2, column=1, sticky=tk.W, pady=10)
            
            # Show/hide password checkbox
            show_password_var = tk.BooleanVar(value=False)
            
            def toggle_password_visibility():
                if show_password_var.get():
                    password_entry.config(show='')
                else:
                    password_entry.config(show='*')
            
            ttk.Checkbutton(db_frame, text="Show Password", variable=show_password_var, 
                           command=toggle_password_visibility).grid(row=3, column=1, sticky=tk.W)
            
            # Database name
            ttk.Label(db_frame, text="Database Name:").grid(row=4, column=0, sticky=tk.W, pady=10)
            db_name_var = tk.StringVar(value=config.get('Database', 'database', fallback='exam_scheduler'))
            ttk.Entry(db_frame, textvariable=db_name_var, width=30).grid(row=4, column=1, sticky=tk.W, pady=10)
            
            # Test connection button
            def test_connection():
                try:
                    # Create a temporary connection with the new settings
                    temp_conn = mysql.connector.connect(
                        host=host_var.get(),
                        user=user_var.get(),
                        password=password_var.get(),
                        database=db_name_var.get()
                    )
                    temp_conn.close()
                    messagebox.showinfo("Connection Test", "Database connection successful!")
                except Error as e:
                    messagebox.showerror("Connection Test", f"Failed to connect to database: {str(e)}")
            
            ttk.Button(db_frame, text="Test Connection", command=test_connection).grid(
                row=5, column=1, sticky=tk.W, pady=10)
            
            # Application settings tab
            app_frame = ttk.Frame(pref_notebook, padding=20)
            pref_notebook.add(app_frame, text="Application Settings")
            
            # Configure grid
            app_frame.columnconfigure(0, weight=0)
            app_frame.columnconfigure(1, weight=1)
            
            # Theme selection
            ttk.Label(app_frame, text="Theme:").grid(row=0, column=0, sticky=tk.W, pady=10)
            theme_var = tk.StringVar(value=config.get('Application', 'theme', fallback='Default'))
            theme_combo = ttk.Combobox(app_frame, textvariable=theme_var, 
                                     values=["Default", "Light", "Dark", "Blue Ocean", "Forest Green", "Midnight Purple", "Sunset Orange",
                                             "Neon Cyberpunk", "Pastel Dream", "Coffee Cream", "Royal Navy", "Cherry Blossom"], 
                                     width=28, state="readonly")
            theme_combo.grid(row=0, column=1, sticky=tk.W, pady=10)
            
            # Preview theme button
            def preview_theme():
                self.apply_theme(theme_var.get())
            
            ttk.Button(app_frame, text="Preview Theme", command=preview_theme).grid(
                row=0, column=2, sticky=tk.W, padx=5, pady=10)
            
            # Auto-save interval
            ttk.Label(app_frame, text="Auto-save Interval (minutes):").grid(row=1, column=0, sticky=tk.W, pady=10)
            autosave_var = tk.IntVar(value=int(config.get('Application', 'autosave_interval', fallback='5')))
            ttk.Spinbox(app_frame, from_=0, to=60, textvariable=autosave_var, width=5).grid(
                row=1, column=1, sticky=tk.W, pady=10)
            
            # Skip Sundays option
            ttk.Label(app_frame, text="Skip Sundays in Schedule:").grid(row=2, column=0, sticky=tk.W, pady=10)
            skip_sundays_var = tk.BooleanVar(value=config.get('Application', 'skip_sundays', fallback='True').lower() == 'true')
            ttk.Checkbutton(app_frame, variable=skip_sundays_var).grid(row=2, column=1, sticky=tk.W, pady=10)
            
            # Buttons at the bottom
            button_frame = ttk.Frame(pref_dialog)
            button_frame.pack(fill=tk.X, padx=10, pady=20)
            
            def save_preferences():
                try:
                    # Update config file
                    if not config.has_section('Database'):
                        config.add_section('Database')
                    
                    config.set('Database', 'host', host_var.get())
                    config.set('Database', 'user', user_var.get())
                    config.set('Database', 'password', password_var.get())
                    config.set('Database', 'database', db_name_var.get())
                    
                    # Add Application section if it doesn't exist
                    if not config.has_section('Application'):
                        config.add_section('Application')
                    
                    config.set('Application', 'theme', theme_var.get())
                    config.set('Application', 'autosave_interval', str(autosave_var.get()))
                    config.set('Application', 'skip_sundays', str(skip_sundays_var.get()))
                    
                    # Save to file
                    with open(CONFIG_FILE, 'w') as configfile:
                        config.write(configfile)
                    
                    # Update database connection with new settings
                    self.db_config = {
                        'host': host_var.get(),
                        'user': user_var.get(),
                        'password': password_var.get(),
                        'database': db_name_var.get(),
                        'autocommit': True
                    }
                    
                    # Close current connection if exists
                    if hasattr(self, 'conn') and self.conn:
                        try:
                            self.conn.close()
                        except:
                            pass
                    
                    # Connect with new settings
                    self.conn = mysql.connector.connect(**self.db_config)
                    self.cursor = self.conn.cursor(buffered=True)
                    
                    # Apply the selected theme
                    self.apply_theme(theme_var.get())
                    
                    messagebox.showinfo("Success", "Preferences saved successfully. New settings will take effect immediately.")
                    pref_dialog.destroy()
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save preferences: {str(e)}")
            
            ttk.Button(button_frame, text="Save", command=save_preferences, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
            ttk.Button(button_frame, text="Cancel", command=pref_dialog.destroy).pack(side=tk.RIGHT, padx=5)
            
            # Center the dialog on the screen
            pref_dialog.update_idletasks()
            width = pref_dialog.winfo_width()
            height = pref_dialog.winfo_height()
            x = (pref_dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (pref_dialog.winfo_screenheight() // 2) - (height // 2)
            pref_dialog.geometry(f"{width}x{height}+{x}+{y}")
            
            # Make dialog modal
            pref_dialog.focus_set()
            pref_dialog.wait_window()
            
        except Exception as e:
            print(f"Error showing preferences dialog: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to open preferences dialog: {str(e)}")

    
    def show_help(self):
        """Show the help documentation."""
        try:
            # Create a custom help dialog
            help_dialog = tk.Toplevel(self.root)
            help_dialog.title("Help")
            help_dialog.geometry("700x600")
            help_dialog.transient(self.root)
            help_dialog.grab_set()
            
            # Create a frame with padding
            main_frame = ttk.Frame(help_dialog, padding=10)
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # Create a notebook for different help sections
            help_notebook = ttk.Notebook(main_frame)
            help_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Getting Started tab
            getting_started_frame = ttk.Frame(help_notebook, padding=10)
            help_notebook.add(getting_started_frame, text="Getting Started")
            
            # Add scrollbar to Getting Started tab
            gs_canvas = tk.Canvas(getting_started_frame)
            gs_scrollbar = ttk.Scrollbar(getting_started_frame, orient="vertical", command=gs_canvas.yview)
            gs_scrollable_frame = ttk.Frame(gs_canvas)
            
            gs_scrollable_frame.bind(
                "<Configure>",
                lambda e: gs_canvas.configure(scrollregion=gs_canvas.bbox("all"))
            )
            
            gs_canvas.create_window((0, 0), window=gs_scrollable_frame, anchor="nw")
            gs_canvas.configure(yscrollcommand=gs_scrollbar.set)
            
            gs_canvas.pack(side="left", fill="both", expand=True)
            gs_scrollbar.pack(side="right", fill="y")
            
            # Getting Started content
            ttk.Label(gs_scrollable_frame, text="Getting Started with Exam Scheduling Generator", 
                      font=("TkDefaultFont", 14, "bold")).pack(pady=(0, 10), anchor="w")
            
            intro_text = (
                "Welcome to the Exam Scheduling Generator application! This tool helps educational \n"
                "institutions create optimized exam schedules while considering various constraints.\n\n"
                "To get started with the application, follow these steps:\n\n"
                "1. First, add your subjects using the Subjects tab\n"
                "2. Add examination rooms in the Rooms tab\n"
                "3. Navigate to the Generate tab to create a new schedule\n"
                "4. Configure the schedule parameters and select subjects\n"
                "5. Generate the schedule and review it in the preview\n"
                "6. Save your schedule when you're satisfied with it\n"
                "7. View and manage your schedules in the View Schedule tab"
            )
            
            ttk.Label(gs_scrollable_frame, text=intro_text, justify="left").pack(pady=10, anchor="w")
            
            # Database Setup section
            ttk.Label(gs_scrollable_frame, text="Database Setup", font=("TkDefaultFont", 12, "bold")).pack(pady=(10, 5), anchor="w")
            
            db_text = (
                "This application uses MySQL for data storage. To configure your database connection:\n\n"
                "1. Go to Edit → Preferences\n"
                "2. Enter your MySQL server details (host, username, password)\n"
                "3. Click 'Test Connection' to verify your settings\n"
                "4. Save your preferences\n\n"
                "The application will automatically create the necessary database and tables."
            )
            
            ttk.Label(gs_scrollable_frame, text=db_text, justify="left").pack(pady=10, anchor="w")
            
            # Features tab
            features_frame = ttk.Frame(help_notebook, padding=10)
            help_notebook.add(features_frame, text="Features")
            
            # Add scrollbar to Features tab
            f_canvas = tk.Canvas(features_frame)
            f_scrollbar = ttk.Scrollbar(features_frame, orient="vertical", command=f_canvas.yview)
            f_scrollable_frame = ttk.Frame(f_canvas)
            
            f_scrollable_frame.bind(
                "<Configure>",
                lambda e: f_canvas.configure(scrollregion=f_canvas.bbox("all"))
            )
            
            f_canvas.create_window((0, 0), window=f_scrollable_frame, anchor="nw")
            f_canvas.configure(yscrollcommand=f_scrollbar.set)
            
            f_canvas.pack(side="left", fill="both", expand=True)
            f_scrollbar.pack(side="right", fill="y")
            
            # Features content
            ttk.Label(f_scrollable_frame, text="Key Features", font=("TkDefaultFont", 14, "bold")).pack(pady=(0, 10), anchor="w")
            
            features_text = (
                "The Exam Scheduling Generator offers the following key features:\n\n"
                "• Subject Management: Add, edit, and delete subjects with details\n"
                "• Room Management: Manage examination rooms and their capacities\n"
                "• Schedule Generation: Create optimized exam schedules using constraint programming\n"
                "• Schedule Visualization: View schedules in list or calendar format\n"
                "• Export Options: Export schedules, subjects, and rooms to CSV or PDF\n"
                "• Data Import: Import subjects and rooms from CSV files\n"
                "• Manual Scheduling: Add or edit schedule items manually\n"
                "• Theme Customization: Choose from 12 different visual themes\n"
                "• Font Selection: Customize the application font\n"
                "• Zoom Controls: Adjust text size for better readability"
            )
            
            ttk.Label(f_scrollable_frame, text=features_text, justify="left").pack(pady=10, anchor="w")
            
            # Keyboard Shortcuts section
            ttk.Label(f_scrollable_frame, text="Keyboard Shortcuts", font=("TkDefaultFont", 12, "bold")).pack(pady=(10, 5), anchor="w")
            
            shortcuts_text = (
                "The application supports the following keyboard shortcuts:\n\n"
                "File Menu:\n"
                "• Ctrl+N: New Schedule\n"
                "• Ctrl+O: Open Schedule\n"
                "• Ctrl+S: Save Schedule\n"
                "• Ctrl+I: Import Subjects\n"
                "• Ctrl+E: Export Subjects\n\n"
                "Edit Menu:\n"
                "• Ctrl+P: Preferences\n"
                "• Ctrl+F: Font Selection\n\n"
                "View Menu:\n"
                "• Ctrl+L: List View\n"
                "• Ctrl+D: Calendar View\n"
                "• Ctrl+Plus: Zoom In\n"
                "• Ctrl+Minus: Zoom Out\n"
                "• Ctrl+0: Reset Zoom\n"
                "• Ctrl+B: Toggle Status Bar\n\n"
                "Other:\n"
                "• F1: Help\n"
                "• F2: About\n"
                "• F5: Refresh"
            )
            
            ttk.Label(f_scrollable_frame, text=shortcuts_text, justify="left").pack(pady=10, anchor="w")
            
            # Troubleshooting tab
            troubleshooting_frame = ttk.Frame(help_notebook, padding=10)
            help_notebook.add(troubleshooting_frame, text="Troubleshooting")
            
            # Add scrollbar to Troubleshooting tab
            t_canvas = tk.Canvas(troubleshooting_frame)
            t_scrollbar = ttk.Scrollbar(troubleshooting_frame, orient="vertical", command=t_canvas.yview)
            t_scrollable_frame = ttk.Frame(t_canvas)
            
            t_scrollable_frame.bind(
                "<Configure>",
                lambda e: t_canvas.configure(scrollregion=t_canvas.bbox("all"))
            )
            
            t_canvas.create_window((0, 0), window=t_scrollable_frame, anchor="nw")
            t_canvas.configure(yscrollcommand=t_scrollbar.set)
            
            t_canvas.pack(side="left", fill="both", expand=True)
            t_scrollbar.pack(side="right", fill="y")
            
            # Troubleshooting content
            ttk.Label(t_scrollable_frame, text="Troubleshooting", font=("TkDefaultFont", 14, "bold")).pack(pady=(0, 10), anchor="w")
            
            troubleshooting_text = (
                "If you encounter issues with the application, try these solutions:\n\n"
                "Database Connection Issues:\n"
                "• Verify your MySQL server is running\n"
                "• Check your connection credentials in Edit → Preferences\n"
                "• Ensure your MySQL user has appropriate permissions\n"
                "• Try the 'Test Connection' button to diagnose issues\n\n"
                "Schedule Generation Problems:\n"
                "• Ensure you have subjects and rooms added to the database\n"
                "• Check that your constraints aren't too restrictive\n"
                "• Try with fewer subjects or more rooms if generation fails\n"
                "• Use the 'Refresh' option (F5) to reload all data\n\n"
                "Display Issues:\n"
                "• Use the View → Zoom options to adjust text size\n"
                "• Try changing the application font in Edit → Font\n"
                "• Select a different theme in Edit → Preferences\n\n"
                "If problems persist, try restarting the application or reinstalling it."
            )
            
            ttk.Label(t_scrollable_frame, text=troubleshooting_text, justify="left").pack(pady=10, anchor="w")
            
            # About tab
            about_frame = ttk.Frame(help_notebook, padding=10)
            help_notebook.add(about_frame, text="About")
            
            # Add scrollbar to About tab
            a_canvas = tk.Canvas(about_frame)
            a_scrollbar = ttk.Scrollbar(about_frame, orient="vertical", command=a_canvas.yview)
            a_scrollable_frame = ttk.Frame(a_canvas)
            
            a_scrollable_frame.bind(
                "<Configure>",
                lambda e: a_canvas.configure(scrollregion=a_canvas.bbox("all"))
            )
            
            a_canvas.create_window((0, 0), window=a_scrollable_frame, anchor="nw")
            a_canvas.configure(yscrollcommand=a_scrollbar.set)
            
            a_canvas.pack(side="left", fill="both", expand=True)
            a_scrollbar.pack(side="right", fill="y")
            
            # About content
            ttk.Label(a_scrollable_frame, text="About Exam Scheduling Generator", font=("TkDefaultFont", 14, "bold")).pack(pady=(0, 10), anchor="w")
            
            about_text = (
                f"Exam Scheduling Generator v{APP_VERSION}\n\n"
                "A comprehensive exam scheduling application for educational institutions.\n\n"
                "This application was developed to simplify the complex task of creating exam schedules \n"
                "while considering various constraints such as room availability, subject requirements, \n"
                "and scheduling preferences.\n\n"
                "Technologies Used:\n"
                "• Python 3.x\n"
                "• tkinter for GUI\n"
                "• OR-Tools for constraint programming\n"
                "• MySQL for database\n"
                "• Pandas for data manipulation\n"
                "• ReportLab for PDF generation\n"
                "• Matplotlib for visualizations\n\n"
                "Developed By Krutarth Raychura\n"
                "© 2025 All Rights Reserved"
            )
            
            ttk.Label(a_scrollable_frame, text=about_text, justify="left").pack(pady=10, anchor="w")
            
            # Add close button at the bottom
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill=tk.X, pady=(10, 0))
            
            ttk.Button(button_frame, text="Close", command=help_dialog.destroy).pack(side=tk.RIGHT)
            
            # Center the dialog on the screen
            help_dialog.update_idletasks()
            width = help_dialog.winfo_width()
            height = help_dialog.winfo_height()
            x = (help_dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (help_dialog.winfo_screenheight() // 2) - (height // 2)
            help_dialog.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception as e:
            # Fallback to simple messagebox if there's an error
            print(f"Error showing help dialog: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showinfo("Help", "Help documentation will be available in a future update.")
    
    def apply_theme(self, theme_name=None):
        """Apply the selected theme to the application."""
        try:
            # If no theme specified, read from config
            if theme_name is None:
                config = configparser.ConfigParser()
                config.read(CONFIG_FILE)
                theme_name = config.get('Application', 'theme', fallback='Default')
            
            # Apply the theme
            style = ttk.Style()
            
            if theme_name == "Light":
                # Light theme
                style.theme_use('clam')
                style.configure('TFrame', background='#f0f0f0')
                style.configure('TLabel', background='#f0f0f0', foreground='#000000')
                style.configure('TButton', background='#e0e0e0', foreground='#000000')
                style.configure('Accent.TButton', background='#4a7dfc', foreground='#ffffff')
                style.configure('TNotebook', background='#f0f0f0')
                style.configure('TNotebook.Tab', background='#e0e0e0', foreground='#000000')
                style.map('TNotebook.Tab', background=[('selected', '#ffffff')])
                style.configure('Treeview', background='#ffffff', fieldbackground='#ffffff', foreground='#000000')
                style.map('Treeview', background=[('selected', '#4a7dfc')], foreground=[('selected', '#ffffff')])
                
                # Set main window background
                self.root.configure(background='#f0f0f0')
                
            elif theme_name == "Dark":
                # Dark theme
                style.theme_use('clam')
                style.configure('TFrame', background='#2d2d30')
                style.configure('TLabel', background='#2d2d30', foreground='#ffffff')
                style.configure('TButton', background='#3e3e42', foreground='#ffffff')
                style.configure('Accent.TButton', background='#0078d7', foreground='#ffffff')
                style.configure('TNotebook', background='#2d2d30')
                style.configure('TNotebook.Tab', background='#3e3e42', foreground='#ffffff')
                style.map('TNotebook.Tab', background=[('selected', '#1e1e1e')])
                style.configure('Treeview', background='#1e1e1e', fieldbackground='#1e1e1e', foreground='#ffffff')
                style.map('Treeview', background=[('selected', '#0078d7')], foreground=[('selected', '#ffffff')])
                
                # Set main window background
                self.root.configure(background='#2d2d30')
                
            elif theme_name == "Blue Ocean":
                # Blue Ocean theme - cool blue gradient effect
                style.theme_use('clam')
                style.configure('TFrame', background='#1a3c5e')
                style.configure('TLabel', background='#1a3c5e', foreground='#e0f0ff')
                style.configure('TButton', background='#2c5f8e', foreground='#ffffff')
                style.configure('Accent.TButton', background='#00ccff', foreground='#000000')
                style.configure('TNotebook', background='#1a3c5e')
                style.configure('TNotebook.Tab', background='#2c5f8e', foreground='#e0f0ff')
                style.map('TNotebook.Tab', background=[('selected', '#0d2b4d')])
                style.configure('Treeview', background='#0d2b4d', fieldbackground='#0d2b4d', foreground='#e0f0ff')
                style.map('Treeview', background=[('selected', '#00ccff')], foreground=[('selected', '#000000')])
                
                # Set main window background
                self.root.configure(background='#1a3c5e')
                
            elif theme_name == "Forest Green":
                # Forest Green theme - natural and calming
                style.theme_use('clam')
                style.configure('TFrame', background='#2e4a35')
                style.configure('TLabel', background='#2e4a35', foreground='#e0ffe0')
                style.configure('TButton', background='#3e6346', foreground='#ffffff')
                style.configure('Accent.TButton', background='#7cba59', foreground='#000000')
                style.configure('TNotebook', background='#2e4a35')
                style.configure('TNotebook.Tab', background='#3e6346', foreground='#e0ffe0')
                style.map('TNotebook.Tab', background=[('selected', '#1e3525')])
                style.configure('Treeview', background='#1e3525', fieldbackground='#1e3525', foreground='#e0ffe0')
                style.map('Treeview', background=[('selected', '#7cba59')], foreground=[('selected', '#000000')])
                
                # Set main window background
                self.root.configure(background='#2e4a35')
                
            elif theme_name == "Midnight Purple":
                # Midnight Purple theme - elegant and modern
                style.theme_use('clam')
                style.configure('TFrame', background='#2e1a47')
                style.configure('TLabel', background='#2e1a47', foreground='#f0e0ff')
                style.configure('TButton', background='#4a2c6d', foreground='#ffffff')
                style.configure('Accent.TButton', background='#9966cc', foreground='#ffffff')
                style.configure('TNotebook', background='#2e1a47')
                style.configure('TNotebook.Tab', background='#4a2c6d', foreground='#f0e0ff')
                style.map('TNotebook.Tab', background=[('selected', '#1e0f33')])
                style.configure('Treeview', background='#1e0f33', fieldbackground='#1e0f33', foreground='#f0e0ff')
                style.map('Treeview', background=[('selected', '#9966cc')], foreground=[('selected', '#ffffff')])
                
                # Set main window background
                self.root.configure(background='#2e1a47')
                
            elif theme_name == "Sunset Orange":
                # Sunset Orange theme - warm and energetic
                style.theme_use('clam')
                style.configure('TFrame', background='#4a2c1e')
                style.configure('TLabel', background='#4a2c1e', foreground='#ffe0c0')
                style.configure('TButton', background='#6d462c', foreground='#ffffff')
                style.configure('Accent.TButton', background='#ff7f2a', foreground='#000000')
                style.configure('TNotebook', background='#4a2c1e')
                style.configure('TNotebook.Tab', background='#6d462c', foreground='#ffe0c0')
                style.map('TNotebook.Tab', background=[('selected', '#331e14')])
                style.configure('Treeview', background='#331e14', fieldbackground='#331e14', foreground='#ffe0c0')
                style.map('Treeview', background=[('selected', '#ff7f2a')], foreground=[('selected', '#000000')])
                
                # Set main window background
                self.root.configure(background='#4a2c1e')
                
            elif theme_name == "Neon Cyberpunk":
                # Neon Cyberpunk theme - futuristic and vibrant
                style.theme_use('clam')
                style.configure('TFrame', background='#0f0f2d')
                style.configure('TLabel', background='#0f0f2d', foreground='#00ffff')
                style.configure('TButton', background='#1a1a3a', foreground='#00ffff')
                style.configure('Accent.TButton', background='#ff00ff', foreground='#000000')
                style.configure('TNotebook', background='#0f0f2d')
                style.configure('TNotebook.Tab', background='#1a1a3a', foreground='#00ffff')
                style.map('TNotebook.Tab', background=[('selected', '#000033')])
                style.configure('Treeview', background='#000033', fieldbackground='#000033', foreground='#00ffff')
                style.map('Treeview', background=[('selected', '#ff00ff')], foreground=[('selected', '#000000')])
                
                # Set main window background
                self.root.configure(background='#0f0f2d')
                
            elif theme_name == "Pastel Dream":
                # Pastel Dream theme - soft and gentle colors
                style.theme_use('clam')
                style.configure('TFrame', background='#f0e6f5')
                style.configure('TLabel', background='#f0e6f5', foreground='#5a4a66')
                style.configure('TButton', background='#d9c7e0', foreground='#5a4a66')
                style.configure('Accent.TButton', background='#b19cd9', foreground='#ffffff')
                style.configure('TNotebook', background='#f0e6f5')
                style.configure('TNotebook.Tab', background='#d9c7e0', foreground='#5a4a66')
                style.map('TNotebook.Tab', background=[('selected', '#ffffff')])
                style.configure('Treeview', background='#ffffff', fieldbackground='#ffffff', foreground='#5a4a66')
                style.map('Treeview', background=[('selected', '#b19cd9')], foreground=[('selected', '#ffffff')])
                
                # Set main window background
                self.root.configure(background='#f0e6f5')
                
            elif theme_name == "Coffee Cream":
                # Coffee Cream theme - warm and cozy
                style.theme_use('clam')
                style.configure('TFrame', background='#e6d9cc')
                style.configure('TLabel', background='#e6d9cc', foreground='#4a3c2d')
                style.configure('TButton', background='#d1bfa8', foreground='#4a3c2d')
                style.configure('Accent.TButton', background='#8c7158', foreground='#ffffff')
                style.configure('TNotebook', background='#e6d9cc')
                style.configure('TNotebook.Tab', background='#d1bfa8', foreground='#4a3c2d')
                style.map('TNotebook.Tab', background=[('selected', '#ffffff')])
                style.configure('Treeview', background='#ffffff', fieldbackground='#ffffff', foreground='#4a3c2d')
                style.map('Treeview', background=[('selected', '#8c7158')], foreground=[('selected', '#ffffff')])
                
                # Set main window background
                self.root.configure(background='#e6d9cc')
                
            elif theme_name == "Royal Navy":
                # Royal Navy theme - professional and authoritative
                style.theme_use('clam')
                style.configure('TFrame', background='#1c3144')
                style.configure('TLabel', background='#1c3144', foreground='#d6e8f6')
                style.configure('TButton', background='#2d4b69', foreground='#d6e8f6')
                style.configure('Accent.TButton', background='#d4af37', foreground='#000000')
                style.configure('TNotebook', background='#1c3144')
                style.configure('TNotebook.Tab', background='#2d4b69', foreground='#d6e8f6')
                style.map('TNotebook.Tab', background=[('selected', '#0a1622')])
                style.configure('Treeview', background='#0a1622', fieldbackground='#0a1622', foreground='#d6e8f6')
                style.map('Treeview', background=[('selected', '#d4af37')], foreground=[('selected', '#000000')])
                
                # Set main window background
                self.root.configure(background='#1c3144')
                
            elif theme_name == "Cherry Blossom":
                # Cherry Blossom theme - delicate and beautiful
                style.theme_use('clam')
                style.configure('TFrame', background='#fbe4f0')
                style.configure('TLabel', background='#fbe4f0', foreground='#814f64')
                style.configure('TButton', background='#f8c1dd', foreground='#814f64')
                style.configure('Accent.TButton', background='#e77fb3', foreground='#ffffff')
                style.configure('TNotebook', background='#fbe4f0')
                style.configure('TNotebook.Tab', background='#f8c1dd', foreground='#814f64')
                style.map('TNotebook.Tab', background=[('selected', '#ffffff')])
                style.configure('Treeview', background='#ffffff', fieldbackground='#ffffff', foreground='#814f64')
                style.map('Treeview', background=[('selected', '#e77fb3')], foreground=[('selected', '#ffffff')])
                
                # Set main window background
                self.root.configure(background='#fbe4f0')
                
            else:  # Default theme
                # Reset to default theme
                style.theme_use('clam')  # Use system default theme
                style.configure('Accent.TButton', background='#0078d7', foreground='#ffffff')
                
                # Reset main window background
                self.root.configure(background=style.lookup('TFrame', 'background'))
            
            # Create or update the Accent.TButton style for primary actions
            if theme_name != "Default":
                style.configure('Accent.TButton', font=('Helvetica', 9, 'bold'))
            
            # Update config file with the new theme if it was explicitly set
            if theme_name is not None:
                config = configparser.ConfigParser()
                config.read(CONFIG_FILE)
                
                if not config.has_section('Application'):
                    config.add_section('Application')
                
                config.set('Application', 'theme', theme_name)
                
                with open(CONFIG_FILE, 'w') as configfile:
                    config.write(configfile)
            
            return True
            
        except Exception as e:
            print(f"Error applying theme: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def refresh_application(self):
        """Refresh the entire application, reloading all data from the database."""
        try:
            # Ask for confirmation
            confirm = messagebox.askyesno(
                "Confirm Refresh", 
                "This will refresh the entire application and reload all data. Any unsaved changes will be lost. Continue?"
            )
            
            if not confirm:
                return
                
            # Update status
            self.status_var.set("Refreshing application...")
            self.root.update()
            
            # Reconnect to database
            if hasattr(self, 'conn') and self.conn:
                try:
                    self.conn.close()
                except:
                    pass
            self.init_database()
            
            # Reload data in each tab
            self.update_dashboard_counts()
            
            if hasattr(self, 'load_subjects'):
                self.load_subjects()
                
            if hasattr(self, 'load_rooms'):
                self.load_rooms()
                
            if hasattr(self, 'load_all_schedules'):
                self.load_all_schedules()
                
            # Reset current schedule if any
            if hasattr(self, 'current_schedule'):
                self.current_schedule = []
                
            if hasattr(self, 'current_schedule_id'):
                self.current_schedule_id = None
                
            # Clear preview tree if it exists
            if hasattr(self, 'preview_tree'):
                for item in self.preview_tree.get_children():
                    self.preview_tree.delete(item)
            
            # Reset form fields in generator tab
            if hasattr(self, 'schedule_name_var'):
                self.schedule_name_var.set("")
                
            # Update semester filters
            if hasattr(self, 'update_semester_filter_options'):
                self.update_semester_filter_options()
                
            # Switch to dashboard tab
            if hasattr(self, 'notebook') and hasattr(self, 'tab_indices') and 'dashboard' in self.tab_indices:
                self.notebook.select(self.tab_indices["dashboard"])
                
            # Update status
            self.status_var.set("Application refreshed successfully")
            
            # Show success message
            messagebox.showinfo("Refresh Complete", "The application has been refreshed successfully.")
            
        except Exception as e:
            print(f"Error refreshing application: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Refresh Error", f"Failed to refresh application: {str(e)}")
            self.status_var.set("Refresh failed")
    
    def zoom_in(self):
        """Increase the font size for better readability."""
        try:
            # Get current font sizes
            current_font = font.nametofont("TkDefaultFont")
            current_size = current_font.cget("size")
            
            # Increase font size (max size: 24)
            new_size = min(current_size + 1, 24)
            
            # Update default font
            current_font.configure(size=new_size)
            
            # Update other fonts
            font.nametofont("TkTextFont").configure(size=new_size)
            font.nametofont("TkFixedFont").configure(size=new_size)
            font.nametofont("TkMenuFont").configure(size=new_size)
            font.nametofont("TkHeadingFont").configure(size=new_size)
            
            # Update status bar
            self.status_var.set(f"Zoom level: {new_size}")
            
            # Store current zoom level
            self.current_zoom = new_size
            
        except Exception as e:
            print(f"Error in zoom_in: {e}")
    
    def zoom_out(self):
        """Decrease the font size."""
        try:
            # Get current font sizes
            current_font = font.nametofont("TkDefaultFont")
            current_size = current_font.cget("size")
            
            # Decrease font size (min size: 8)
            new_size = max(current_size - 1, 8)
            
            # Update default font
            current_font.configure(size=new_size)
            
            # Update other fonts
            font.nametofont("TkTextFont").configure(size=new_size)
            font.nametofont("TkFixedFont").configure(size=new_size)
            font.nametofont("TkMenuFont").configure(size=new_size)
            font.nametofont("TkHeadingFont").configure(size=new_size)
            
            # Update status bar
            self.status_var.set(f"Zoom level: {new_size}")
            
            # Store current zoom level
            self.current_zoom = new_size
            
        except Exception as e:
            print(f"Error in zoom_out: {e}")
    
    def reset_zoom(self):
        """Reset font size to default."""
        try:
            # Default font size is typically 9 or 10
            default_size = 10
            
            # Update default font
            font.nametofont("TkDefaultFont").configure(size=default_size)
            
            # Update other fonts
            font.nametofont("TkTextFont").configure(size=default_size)
            font.nametofont("TkFixedFont").configure(size=default_size)
            font.nametofont("TkMenuFont").configure(size=default_size)
            font.nametofont("TkHeadingFont").configure(size=default_size)
            
            # Update status bar
            self.status_var.set(f"Zoom level: Default ({default_size})")
            
            # Store current zoom level
            self.current_zoom = default_size
            
        except Exception as e:
            print(f"Error in reset_zoom: {e}")
    
    def toggle_status_bar(self):
        """Toggle the visibility of the status bar."""
        try:
            if self.show_status_bar.get():
                # Show status bar
                if hasattr(self, 'status_bar'):
                    self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
            else:
                # Hide status bar
                if hasattr(self, 'status_bar'):
                    self.status_bar.pack_forget()
        except Exception as e:
            print(f"Error toggling status bar: {e}")
            
    def change_font(self):
        """Change the application font based on the selected font."""
        try:
            # Get the selected font
            selected_font = self.current_font.get()
            
            # Get the current size to maintain it
            current_size = font.nametofont("TkDefaultFont").cget("size")
            
            # Update all fonts
            font.nametofont("TkDefaultFont").configure(family=selected_font)
            font.nametofont("TkTextFont").configure(family=selected_font)
            font.nametofont("TkFixedFont").configure(family=selected_font)
            font.nametofont("TkMenuFont").configure(family=selected_font)
            font.nametofont("TkHeadingFont").configure(family=selected_font)
            
            # Update status bar
            self.status_var.set(f"Font changed to {selected_font}")
            
            # Save font preference to config file
            self.save_font_preference(selected_font)
            
        except Exception as e:
            print(f"Error changing font: {e}")
            import traceback
            traceback.print_exc()
    
    def show_font_dialog(self):
        """Show a dialog with all available fonts."""
        try:
            # Create dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Select Font")
            dialog.geometry("400x500")
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Create frame with padding
            frame = ttk.Frame(dialog, padding=10)
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Add search field
            search_frame = ttk.Frame(frame)
            search_frame.pack(fill=tk.X, pady=(0, 10))
            
            ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT)
            search_var = tk.StringVar()
            search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
            search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            # Create listbox with scrollbar
            list_frame = ttk.Frame(frame)
            list_frame.pack(fill=tk.BOTH, expand=True)
            
            scrollbar = ttk.Scrollbar(list_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Create font listbox
            font_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, selectmode=tk.SINGLE, font=("TkDefaultFont", 10))
            font_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            scrollbar.config(command=font_listbox.yview)
            
            # Add all fonts to listbox
            for font_name in self.available_fonts:
                font_listbox.insert(tk.END, font_name)
            
            # Try to select the current font
            current_font = self.current_font.get()
            if current_font in self.available_fonts:
                index = self.available_fonts.index(current_font)
                font_listbox.selection_set(index)
                font_listbox.see(index)
            
            # Preview frame
            preview_frame = ttk.LabelFrame(frame, text="Preview", padding=10)
            preview_frame.pack(fill=tk.X, pady=10)
            
            # Preview text
            preview_text = "AaBbCcDdEeFfGg 123456789"
            preview_label = ttk.Label(preview_frame, text=preview_text, font=(current_font, 12))
            preview_label.pack(pady=10)
            
            # Function to update preview when a font is selected
            def on_font_select(event):
                selected_indices = font_listbox.curselection()
                if selected_indices:
                    selected_index = selected_indices[0]
                    selected_font = font_listbox.get(selected_index)
                    preview_label.config(font=(selected_font, 12))
            
            # Function to filter fonts based on search
            def filter_fonts(*args):
                search_term = search_var.get().lower()
                font_listbox.delete(0, tk.END)
                for font_name in self.available_fonts:
                    if search_term in font_name.lower():
                        font_listbox.insert(tk.END, font_name)
            
            # Bind events
            font_listbox.bind("<<ListboxSelect>>", on_font_select)
            search_var.trace("w", filter_fonts)
            
            # Button frame
            button_frame = ttk.Frame(frame)
            button_frame.pack(fill=tk.X, pady=(10, 0))
            
            # Apply button
            def apply_font():
                selected_indices = font_listbox.curselection()
                if selected_indices:
                    selected_index = selected_indices[0]
                    selected_font = font_listbox.get(selected_index)
                    self.current_font.set(selected_font)
                    self.change_font()
                    dialog.destroy()
            
            ttk.Button(button_frame, text="Apply", command=apply_font).pack(side=tk.RIGHT, padx=5)
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
            
            # Focus search entry
            search_entry.focus_set()
            
        except Exception as e:
            print(f"Error showing font dialog: {e}")
            import traceback
            traceback.print_exc()
    
    def save_font_preference(self, font_name):
        """Save the selected font to the configuration file."""
        try:
            # Create config parser
            import configparser
            config = configparser.ConfigParser()
            
            # Load existing config if it exists
            config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
            config.read(config_file)
            
            # Make sure the Application section exists
            if not config.has_section('Application'):
                config.add_section('Application')
            
            # Set the font preference
            config.set('Application', 'font', font_name)
            
            # Save the config
            with open(config_file, 'w') as f:
                config.write(f)
                
        except Exception as e:
            print(f"Error saving font preference: {e}")
            import traceback
            traceback.print_exc()
    
    def load_font_preference(self):
        """Load the saved font preference from the configuration file."""
        try:
            # Create config parser
            import configparser
            config = configparser.ConfigParser()
            
            # Load existing config if it exists
            config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
            if not os.path.exists(config_file):
                return  # No config file exists yet
                
            config.read(config_file)
            
            # Check if font preference exists
            if config.has_section('Application') and config.has_option('Application', 'font'):
                font_name = config.get('Application', 'font')
                
                # Verify the font exists on the system
                available_fonts = list(font.families())
                if font_name in available_fonts:
                    # Apply the font
                    font.nametofont("TkDefaultFont").configure(family=font_name)
                    font.nametofont("TkTextFont").configure(family=font_name)
                    font.nametofont("TkFixedFont").configure(family=font_name)
                    font.nametofont("TkMenuFont").configure(family=font_name)
                    font.nametofont("TkHeadingFont").configure(family=font_name)
                    
                    # Store current font for the menu
                    if hasattr(self, 'current_font'):
                        self.current_font.set(font_name)
                    
                    print(f"Loaded font preference: {font_name}")
                    
        except Exception as e:
            print(f"Error loading font preference: {e}")
            import traceback
            traceback.print_exc()
    
    def open_notepad(self):
        """Open Windows Notepad."""
        try:
            import subprocess
            subprocess.Popen(["notepad.exe"])
            self.status_var.set("Opened Notepad")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Notepad: {str(e)}")
    
    def open_calculator(self):
        """Open Windows Calculator."""
        try:
            import subprocess
            subprocess.Popen(["calc.exe"])
            self.status_var.set("Opened Calculator")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Calculator: {str(e)}")
    
    def open_file_explorer(self):
        """Open Windows File Explorer."""
        try:
            import subprocess
            subprocess.Popen(["explorer.exe"])
            self.status_var.set("Opened File Explorer")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open File Explorer: {str(e)}")
    
    def open_word(self):
        """Open Microsoft Word if installed."""
        try:
            import subprocess
            # Try to open Word - this will only work if Word is installed
            subprocess.Popen(["start", "winword"], shell=True)
            self.status_var.set("Opening Microsoft Word")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Word: {str(e)}")
    
    def open_excel(self):
        """Open Microsoft Excel if installed."""
        try:
            import subprocess
            # Try to open Excel - this will only work if Excel is installed
            subprocess.Popen(["start", "excel"], shell=True)
            self.status_var.set("Opening Microsoft Excel")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel: {str(e)}")
    
    def open_browser(self):
        """Open the default web browser."""
        try:
            import webbrowser
            webbrowser.open("https://www.google.com")
            self.status_var.set("Opened Web Browser")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Web Browser: {str(e)}")
    
    def open_email(self):
        """Open the default email client."""
        try:
            import webbrowser
            webbrowser.open("mailto:")
            self.status_var.set("Opened Email Client")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Email Client: {str(e)}")
    
    def show_about(self):
        """Show the about dialog with logo."""
        try:
            # Create a custom about dialog
            about_dialog = tk.Toplevel(self.root)
            about_dialog.title("About")
            about_dialog.geometry("400x450")  # Increased height to fit all content
            about_dialog.resizable(False, False)
            about_dialog.transient(self.root)
            about_dialog.grab_set()
            
            # Create a frame with padding
            frame = ttk.Frame(about_dialog, padding=20)
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Load and display the logo
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
            
            if os.path.exists(logo_path):
                # Use PIL to open and resize the image
                from PIL import Image, ImageTk
                
                # Open the image and resize it to fit nicely
                img = Image.open(logo_path)
                img = img.resize((150, 150), Image.LANCZOS)
                
                # Convert to PhotoImage
                logo_img = ImageTk.PhotoImage(img)
                
                # Create a label to display the image
                logo_label = ttk.Label(frame, image=logo_img)
                logo_label.image = logo_img  # Keep a reference to prevent garbage collection
                logo_label.pack(pady=(0, 20))
            
            # Add title and version
            title_label = ttk.Label(frame, text=f"{APP_TITLE}", font=("TkDefaultFont", 16, "bold"))
            title_label.pack(pady=(0, 5))
            
            version_label = ttk.Label(frame, text=f"Version {APP_VERSION}")
            version_label.pack(pady=(0, 15))
            
            # Add description
            desc_label = ttk.Label(frame, text="A comprehensive exam scheduling application\nfor educational institutions.", justify=tk.CENTER)
            desc_label.pack(pady=(0, 15))
            
            # Add copyright with developer name prominently displayed
            developer_label = ttk.Label(frame, text="Developed By Krutarth Raychura", font=("TkDefaultFont", 12, "bold"))
            developer_label.pack(pady=(5, 5))
            
            copyright_label = ttk.Label(frame, text="© 2025 All Rights Reserved")
            copyright_label.pack(pady=(0, 25))
            
            # Add close button
            close_button = ttk.Button(frame, text="Close", command=about_dialog.destroy)
            close_button.pack()
            
            # Center the dialog on the screen
            about_dialog.update_idletasks()
            width = about_dialog.winfo_width()
            height = about_dialog.winfo_height()
            x = (about_dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (about_dialog.winfo_screenheight() // 2) - (height // 2)
            about_dialog.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception as e:
            # Fallback to simple messagebox if there's an error
            print(f"Error showing about dialog: {e}")
            import traceback
            traceback.print_exc()
            
            about_text = f"{APP_TITLE} v{APP_VERSION}\n\n"
            about_text += "A comprehensive exam scheduling application for educational institutions.\n\n"
            about_text += "Developed By Krutarth Raychura © 2025 All Rights Reserved"
            messagebox.showinfo("About", about_text)

# Application entry point
if __name__ == "__main__":
    root = tk.Tk()
    app = ExamSchedulerApp(root)
    root.mainloop()
